import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import AdmZip from "adm-zip";
import mammoth from "mammoth";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import "dotenv/config";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cors());
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  app.use((req, res, next) => {
    console.log(`[Server] ${req.method} ${req.url}`);
    next();
  });

  const upload = multer({ dest: path.join(process.cwd(), "uploads") });

  type TranslateEngine = "9router" | "gemini" | "libretranslate";

  const languageAliases: Record<string, string> = {
    auto: "Tự động phát hiện",
    vi: "Tiếng Việt",
    zh: "Tiếng Trung",
    "zh-Hans": "Tiếng Trung giản thể",
    "zh-Hant": "Tiếng Trung phồn thể",
    en: "Tiếng Anh",
  };

  const decodeXml = (value: string) =>
    value
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, "\"")
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");

  const escapeXml = (value: string) =>
    value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");

  const isTranslatableText = (value: string) => {
    const text = decodeXml(value).trim();
    return text.length > 0 && /[\p{L}\p{Script=Han}]/u.test(text);
  };

  const toBilingual = (source: string, translated: string) => {
    const cleanSource = source.trim();
    const cleanTranslated = translated.trim();
    if (!cleanTranslated || cleanTranslated === cleanSource) return cleanSource;
    return `${cleanSource}\n${cleanTranslated}`;
  };

  const call9Router = async (
    texts: string[],
    sourceLang: string,
    targetLang: string,
    apiKey?: string,
    baseUrl?: string,
    model?: string,
  ): Promise<string[]> => {
    const key = apiKey || process.env.NINEROUTER_API_KEY || process.env.OPENAI_API_KEY || "";
    if (!key) throw new Error("Thiếu 9router API key");

    const endpoint = (baseUrl || process.env.NINEROUTER_BASE_URL || "https://5b4c-15-135-214-185.ngrok-free.app/v1")
      .replace(/\/$/, "");
    const selectedModel = model || process.env.NINEROUTER_MODEL || "cx/gpt-5.5";
    const from = languageAliases[sourceLang] || sourceLang || "Tự động phát hiện";
    const to = languageAliases[targetLang] || targetLang || "Tiếng Việt";

    const response = await fetch(`${endpoint}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${key}`,
        "ngrok-skip-browser-warning": "true",
      },
      body: JSON.stringify({
        model: selectedModel,
        messages: [
          {
            role: "system",
            content:
              "Bạn là chuyên gia dịch thuật Việt-Trung. Dịch chính xác, tự nhiên, giữ nguyên số lượng phần tử, thứ tự, mã số, placeholder và thuật ngữ riêng. Chỉ trả JSON hợp lệ dạng {\"translations\":[...]}",
          },
          {
            role: "user",
            content: JSON.stringify({
              task: "translate_array",
              source_language: from,
              target_language: to,
              output: "JSON object with translations array only",
              texts,
            }),
          },
        ],
        temperature: 0.2,
        response_format: { type: "json_object" },
      }),
    });

    if (!response.ok) {
      const detail = await response.text().catch(() => "");
      throw new Error(`9router lỗi HTTP ${response.status}${detail ? `: ${detail.slice(0, 300)}` : ""}`);
    }

    const data: any = await response.json();
    const raw = String(data?.choices?.[0]?.message?.content || "").trim();
    const cleaned = raw.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
    const parsed = JSON.parse(cleaned);
    const arr = Array.isArray(parsed) ? parsed : parsed.translations || parsed.result || parsed.texts;
    if (!Array.isArray(arr) || arr.length !== texts.length) {
      throw new Error("9router trả về JSON không đúng số lượng đoạn dịch");
    }
    return arr.map((item: unknown) => String(item ?? ""));
  };

  const translateTexts = async (
    texts: string[],
    options: {
      sourceLang: string;
      targetLang: string;
      apiKey?: string;
      baseUrl?: string;
      model?: string;
    },
  ) => {
    const unique = [...new Set(texts.map(t => t.trim()).filter(Boolean))];
    const translated = new Map<string, string>();
    const batchSize = 24;

    for (let i = 0; i < unique.length; i += batchSize) {
      const batch = unique.slice(i, i + batchSize);
      const result = await call9Router(
        batch,
        options.sourceLang,
        options.targetLang,
        options.apiKey,
        options.baseUrl,
        options.model,
      );
      result.forEach((value, index) => translated.set(batch[index], value));
    }

    return translated;
  };

  const translateOfficeXml = async (
    xml: string,
    tagName: "w:t" | "a:t" | "t",
    options: Parameters<typeof translateTexts>[1],
  ) => {
    const escapedTag = tagName.replace(":", "\\:");
    const openTag = tagName === "t" ? `<${escapedTag}(\\s[^>]*)?>` : `<${escapedTag}([^>]*)>`;
    const regex = new RegExp(`${openTag}([\\s\\S]*?)<\\/${escapedTag}>`, "g");
    const items = [...xml.matchAll(regex)]
      .map(match => decodeXml(match[tagName === "t" ? 2 : 2]))
      .filter(text => isTranslatableText(text));
    if (items.length === 0) return { xml, count: 0 };

    const translated = await translateTexts(items, options);
    const nextXml = xml.replace(regex, (match, attrs = "", rawContent) => {
      if (!isTranslatableText(rawContent)) return match;
      const original = decodeXml(rawContent);
      const output = translated.get(original.trim());
      if (!output) return match;
      return `<${tagName}${attrs}>${escapeXml(toBilingual(original, output))}</${tagName}>`;
    });

    return { xml: nextXml, count: items.length };
  };

  const extractPdfText = async (filePath: string) => {
    const dataBuffer = fs.readFileSync(filePath);
    const uint8Array = new Uint8Array(dataBuffer);
    const loadingTask = pdfjs.getDocument({
      data: uint8Array,
      useSystemFonts: true,
      disableFontFace: true,
    });
    const pdfDoc = await loadingTask.promise;
    const pages: string[] = [];
    for (let i = 1; i <= pdfDoc.numPages; i++) {
      const page = await pdfDoc.getPage(i);
      const textContent = await page.getTextContent();
      pages.push(textContent.items.map((item: any) => item.str).join(" "));
    }
    return pages;
  };

  const appendPdfTranslationPages = async (
    filePath: string,
    options: Parameters<typeof translateTexts>[1],
  ) => {
    const sourceBytes = fs.readFileSync(filePath);
    const sourcePdf = await PDFDocument.load(sourceBytes, { ignoreEncryption: true });
    const outputPdf = await PDFDocument.create();
    const copiedPages = await outputPdf.copyPages(sourcePdf, sourcePdf.getPageIndices());
    copiedPages.forEach(page => outputPdf.addPage(page));

    const pages = await extractPdfText(filePath);
    const translated = await translateTexts(pages.filter(Boolean), options);
    const font = await outputPdf.embedFont(StandardFonts.Helvetica);

    pages.forEach((pageText, index) => {
      const page = outputPdf.addPage();
      const { width, height } = page.getSize();
      const translation = translated.get(pageText.trim()) || "";
      const lines = [`Page ${index + 1}`, "", pageText, "", translation].flatMap(line => {
        const words = line.split(/\s+/);
        const wrapped: string[] = [];
        let current = "";
        for (const word of words) {
          const candidate = current ? `${current} ${word}` : word;
          if (font.widthOfTextAtSize(candidate, 10) > width - 80) {
            if (current) wrapped.push(current);
            current = word;
          } else {
            current = candidate;
          }
        }
        if (current) wrapped.push(current);
        return wrapped.length ? wrapped : [""];
      });

      let y = height - 48;
      for (const line of lines) {
        if (y < 48) break;
        page.drawText(line, { x: 40, y, size: 10, font, color: rgb(0.1, 0.1, 0.1) });
        y -= 14;
      }
    });

    return Buffer.from(await outputPdf.save());
  };

  // API route for health check
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok", timestamp: new Date().toISOString() });
  });

  app.post("/api/translate-file", upload.single("file"), async (req, res) => {
    console.log("[API] /api/translate-file hit");
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }

      const filePath = req.file.path;
      const originalName = req.file.originalname;
      const lowerName = originalName.toLowerCase();
      const options = {
        sourceLang: String(req.body.sourceLang || "auto"),
        targetLang: String(req.body.targetLang || "vi"),
        apiKey: String(req.body.apiKey || ""),
        baseUrl: String(req.body.baseUrl || ""),
        model: String(req.body.model || ""),
      };

      let outputBuffer: Buffer;
      let contentType = "application/octet-stream";
      let outputName = `Bilingual_${originalName}`;

      if (lowerName.endsWith(".docx") || lowerName.endsWith(".pptx") || lowerName.endsWith(".xlsx")) {
        const zip = new AdmZip(filePath);
        const zipEntries = zip.getEntries();
        let changedSegments = 0;

        if (lowerName.endsWith(".docx")) {
          contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
          const entry = zip.getEntry("word/document.xml");
          if (!entry) throw new Error("Không tìm thấy word/document.xml");
          const result = await translateOfficeXml(entry.getData().toString("utf8"), "w:t", options);
          zip.updateFile(entry.entryName, Buffer.from(result.xml, "utf8"));
          changedSegments += result.count;
        } else if (lowerName.endsWith(".pptx")) {
          contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
          const slideEntries = zipEntries
            .filter(entry => entry.entryName.startsWith("ppt/slides/slide") && entry.entryName.endsWith(".xml"))
            .sort((a, b) => {
              const aNum = parseInt(a.entryName.match(/\d+/)?.[0] || "0", 10);
              const bNum = parseInt(b.entryName.match(/\d+/)?.[0] || "0", 10);
              return aNum - bNum;
            });

          for (const entry of slideEntries) {
            const result = await translateOfficeXml(entry.getData().toString("utf8"), "a:t", options);
            zip.updateFile(entry.entryName, Buffer.from(result.xml, "utf8"));
            changedSegments += result.count;
          }
        } else {
          contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
          const sharedStrings = zip.getEntry("xl/sharedStrings.xml");
          if (sharedStrings) {
            const result = await translateOfficeXml(sharedStrings.getData().toString("utf8"), "t", options);
            zip.updateFile(sharedStrings.entryName, Buffer.from(result.xml, "utf8"));
            changedSegments += result.count;
          }

          const sheetEntries = zipEntries.filter(entry =>
            /^xl\/worksheets\/sheet\d+\.xml$/.test(entry.entryName)
          );
          for (const entry of sheetEntries) {
            const result = await translateOfficeXml(entry.getData().toString("utf8"), "t", options);
            zip.updateFile(entry.entryName, Buffer.from(result.xml, "utf8"));
            changedSegments += result.count;
          }
        }

        if (changedSegments === 0) throw new Error("Không tìm thấy văn bản để dịch trong file");
        outputBuffer = zip.toBuffer();
      } else if (lowerName.endsWith(".pdf")) {
        contentType = "application/pdf";
        outputName = `Bilingual_${originalName.replace(/\.pdf$/i, "")}.pdf`;
        outputBuffer = await appendPdfTranslationPages(filePath, options);
      } else {
        return res.status(400).json({ error: "Định dạng không hỗ trợ. Chỉ nhận .docx, .xlsx, .pptx, .pdf" });
      }

      fs.unlinkSync(filePath);
      res.setHeader("Content-Type", contentType);
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`);
      res.send(outputBuffer);
    } catch (error: any) {
      console.error("[API] Translate file error:", error);
      if (req.file?.path && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
      res.status(500).json({ error: error.message || "Failed to translate file" });
    }
  });

  // API route for file text extraction
  app.post("/api/extract-text", upload.single("file"), async (req, res) => {
    console.log("[API] /api/extract-text hit");
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }

      const filePath = req.file.path;
      const originalName = req.file.originalname.toLowerCase();
      let extractedText = "";

      if (originalName.endsWith(".pdf")) {
        const dataBuffer = fs.readFileSync(filePath);
        const uint8Array = new Uint8Array(dataBuffer);
        const loadingTask = pdfjs.getDocument({
          data: uint8Array,
          useSystemFonts: true,
          disableFontFace: true,
        });
        const pdfDoc = await loadingTask.promise;
        let fullText = "";
        for (let i = 1; i <= pdfDoc.numPages; i++) {
          const page = await pdfDoc.getPage(i);
          const textContent = await page.getTextContent();
          const pageText = textContent.items.map((item: any) => item.str).join(" ");
          fullText += pageText + "\n";
        }
        extractedText = fullText;
      } else if (originalName.endsWith(".docx")) {
        const result = await mammoth.extractRawText({ path: filePath });
        extractedText = result.value;
      } else if (originalName.endsWith(".pptx")) {
        const zip = new AdmZip(filePath);
        const zipEntries = zip.getEntries();
        let pptText = "";
        const slideEntries = zipEntries.filter(entry => 
          entry.entryName.startsWith("ppt/slides/slide") && entry.entryName.endsWith(".xml")
        );
        slideEntries.sort((a, b) => {
          const aNum = parseInt(a.entryName.match(/\d+/)![0]);
          const bNum = parseInt(b.entryName.match(/\d+/)![0]);
          return aNum - bNum;
        });
        for (const entry of slideEntries) {
          const content = entry.getData().toString("utf8");
          const matches = content.match(/<a:t>([^<]*)<\/a:t>/g);
          if (matches) {
            pptText += matches.map(m => m.replace(/<\/?a:t>/g, "")).join(" ") + "\n";
          }
        }
        extractedText = pptText;
      } else {
        return res.status(400).json({ error: "Unsupported file format" });
      }

      fs.unlinkSync(filePath);
      res.json({ text: extractedText });
    } catch (error: any) {
      console.error("[API] Extraction error:", error);
      res.status(500).json({ error: error.message || "Failed to extract text" });
    }
  });

  // API route for preparing "Deep Translation"
  app.post("/api/prepare-deep-translate", upload.single("file"), async (req, res) => {
    console.log("[API] /api/prepare-deep-translate hit");
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }
      const filePath = req.file.path;
      const originalName = req.file.originalname.toLowerCase();
      const zip = new AdmZip(filePath);
      const zipEntries = zip.getEntries();
      const parts: { path: string, content: string }[] = [];

      if (originalName.endsWith(".docx")) {
        const docXml = zip.getEntry("word/document.xml");
        if (docXml) {
          parts.push({ path: "word/document.xml", content: docXml.getData().toString("utf8") });
        }
      } else if (originalName.endsWith(".pptx")) {
        const slideEntries = zipEntries.filter(entry => 
          entry.entryName.startsWith("ppt/slides/slide") && entry.entryName.endsWith(".xml")
        );
        for (const entry of slideEntries) {
          parts.push({ path: entry.entryName, content: entry.getData().toString("utf8") });
        }
      } else {
        return res.status(400).json({ error: "Format không hỗ trợ giữ nguyên định dạng (chỉ Word/PPTX)" });
      }

      res.json({ 
        parts, 
        tempId: path.basename(filePath),
        originalName: req.file.originalname 
      });
    } catch (error: any) {
      console.error("[API] Prepare error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // API route to finalize "Deep Translation"
  app.post("/api/finalize-deep-translate", async (req, res) => {
    console.log("[API] /api/finalize-deep-translate hit");
    try {
      const { tempId, parts, originalName } = req.body;
      const filePath = path.join(process.cwd(), "uploads", tempId);
      
      if (!fs.existsSync(filePath)) {
        return res.status(400).json({ error: "Session expired or file not found" });
      }

      const zip = new AdmZip(filePath);
      for (const part of parts) {
        const entry = zip.getEntry(part.path);
        if (entry) {
          zip.updateFile(part.path, Buffer.from(part.content, "utf8"));
        } else {
          zip.addFile(part.path, Buffer.from(part.content, "utf8"));
        }
      }

      const outputBuffer = zip.toBuffer();
      const outputName = `Translated_${originalName}`;
      
      fs.unlinkSync(filePath);

      let contentType = "application/octet-stream";
      if (originalName.endsWith(".docx")) {
        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      } else if (originalName.endsWith(".pptx")) {
        contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
      }

      res.setHeader("Content-Type", contentType);
      // Use RFC 5987 for proper filename encoding in headers. 
      // Avoid double encoding or incorrect quoting that causes issues in some browsers/Word.
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`);
      res.send(outputBuffer);
    } catch (error: any) {
      console.error("[API] Finalize error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // Catch-all for /api to prevent falling through to Vite/SPA
  app.all("/api/*", (req, res) => {
    console.warn(`[API] Route not found: ${req.method} ${req.url}`);
    res.status(404).json({ 
      error: "API route not found", 
      method: req.method, 
      path: req.url 
    });
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { 
        middlewareMode: true,
        hmr: false, // Explicitly disable HMR to prevent port 24678 conflicts
      },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });

  // Custom error handler to ensure JSON responses for API errors
  app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
    console.error("Server Error:", err);
    if (res.headersSent) {
      return next(err);
    }
    res.status(err.status || 500).json({
      error: err.message || "Internal Server Error",
      code: err.code
    });
  });
}

startServer();
