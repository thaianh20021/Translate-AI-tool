/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useCallback } from "react";
import { GoogleGenAI } from "@google/genai";
import {
  FileText, Upload, Languages, ArrowRightLeft, Download,
  Copy, Loader2, FileWarning, X, FileCheck, Settings, Key, Cpu
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Progress } from "@/components/ui/progress";
import { Toaster, toast } from "sonner";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Document, Packer, Paragraph, TextRun } from "docx";

// ─── Types ─────────────────────────────────────────────────────────────────
type TranslationState = "idle" | "extracting" | "translating" | "completed" | "error";
type Engine = "9router" | "gemini" | "libretranslate";

const GEMINI_MODEL = "gemini-2.0-flash";
const ROUTER_DEFAULT_MODEL = "cx/gpt-5.5";
const ROUTER_MODEL_OPTIONS = [
  "cx/gpt-5.5",
  "cx/gpt-5.5-review",
  "cx/gpt-5.4",
  "cx/gpt-5.4-review",
  "cx/gpt-5.4-mini",
  "cx/gpt-5.4-mini-review",
  "cx/gpt-5.3-codex",
  "cx/gpt-5.3-codex-high",
  "cx/gpt-5.3-codex-xhigh",
  "cx/gpt-5.3-codex-low",
  "cx/gpt-5.3-codex-none",
  "cx/gpt-5.3-codex-spark",
  "ag/gemini-3-flash",
  "ag/gemini-3.5-flash-low",
  "ag/gemini-3.5-flash-extra-low",
  "ag/gemini-pro-agent",
  "ag/claude-sonnet-4-6",
  "ag/claude-opus-4-6-thinking",
  "ag/gpt-oss-120b-medium",
  "qd/auto",
  "qd/ultimate",
  "qd/performance",
  "qd/efficient",
  "qd/qmodel_latest",
  "qd/qmodel",
  "nvidia/qwen/qwen3-coder-480b-a35b-instruct",
];

const LANG_OPTIONS = [
  "Tự động phát hiện", "Tiếng Anh", "Tiếng Việt",
  "Tiếng Nhật", "Tiếng Trung", "Tiếng Pháp",
  "Tiếng Đức", "Tiếng Hàn", "Tiếng Tây Ban Nha",
];

const LIBRE_LANG_MAP: Record<string, string> = {
  "Tiếng Việt": "vi", "Tiếng Anh": "en", "Tiếng Nhật": "ja",
  "Tiếng Trung": "zh-Hans", "Tiếng Pháp": "fr", "Tiếng Đức": "de",
  "Tiếng Hàn": "ko", "Tiếng Tây Ban Nha": "es", "Tự động phát hiện": "auto",
};

// ─── JSZip lazy loader (CDN) ────────────────────────────────────────────────
let _JSZip: any = null;
async function getJSZip() {
  if (_JSZip) return _JSZip;
  if ((window as any).JSZip) { _JSZip = (window as any).JSZip; return _JSZip; }
  await new Promise<void>((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = () => { _JSZip = (window as any).JSZip; resolve(); };
    s.onerror = reject;
    document.head.appendChild(s);
  });
  return _JSZip;
}

// ─── PDF.js lazy loader (CDN) ───────────────────────────────────────────────
async function getPdfjsLib(): Promise<any> {
  if ((window as any).pdfjsLib) return (window as any).pdfjsLib;
  // Load pdf.js as a regular script (legacy build)
  await new Promise<void>((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => resolve();
    s.onerror = reject;
    document.head.appendChild(s);
  });
  const lib = (window as any).pdfjsLib;
  lib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  return lib;
}

// ─── File text extractors (browser-side, no server) ────────────────────────
async function extractPDF(file: File): Promise<string> {
  const lib = await getPdfjsLib();
  const buf = await file.arrayBuffer();
  const pdf = await lib.getDocument({ data: new Uint8Array(buf) }).promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map((it: any) => it.str).join(" ") + "\n";
  }
  return text;
}

async function extractDOCX(file: File): Promise<string> {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const xml: string = await zip.file("word/document.xml").async("string");
  return [...xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map(m => m[1]).join(" ");
}

async function extractPPTX(file: File): Promise<string> {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slides = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/\d+/)![0]) - parseInt(b.match(/\d+/)![0]));
  let text = "";
  for (const sf of slides) {
    const xml: string = await zip.file(sf).async("string");
    text += [...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map(m => m[1]).join(" ") + "\n";
  }
  return text;
}

// ─── XML parts (keep-format mode) ──────────────────────────────────────────
async function getXmlParts(file: File): Promise<{ path: string; content: string }[]> {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  if (file.name.toLowerCase().endsWith(".docx")) {
    const xml: string = await zip.file("word/document.xml").async("string");
    return [{ path: "word/document.xml", content: xml }];
  }
  // PPTX
  const slides = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/\d+/)![0]) - parseInt(b.match(/\d+/)![0]));
  const parts: { path: string; content: string }[] = [];
  for (const sf of slides) {
    parts.push({ path: sf, content: await zip.file(sf).async("string") });
  }
  return parts;
}

async function rebuildZip(file: File, parts: { path: string; content: string }[]): Promise<Blob> {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  for (const p of parts) zip.file(p.path, p.content);
  return zip.generateAsync({
    type: "blob",
    mimeType: file.name.toLowerCase().endsWith(".docx")
      ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      : "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    compression: "DEFLATE",
  });
}

// ─── App ────────────────────────────────────────────────────────────────────
export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [sourceLang, setSourceLang] = useState("Tự động phát hiện");
  const [targetLang, setTargetLang] = useState("Tiếng Việt");
  const [status, setStatus] = useState<TranslationState>("idle");
  const [progress, setProgress] = useState(0);
  const [progressLabel, setProgressLabel] = useState("");
  const [originalText, setOriginalText] = useState("");
  const [translatedText, setTranslatedText] = useState("");
  const [keepFormat, setKeepFormat] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [logs, setLogs] = useState<string[]>([]);
  const logsEndRef = useRef<HTMLDivElement>(null);
  const libreRouteRef = useRef<"direct" | "via-en" | null>(null);

  // Settings
  const [showSettings, setShowSettings] = useState(false);
  const [engine, setEngine] = useState<Engine>(
    () => (localStorage.getItem("trans_engine") as Engine) || "9router"
  );
  const [apiKey, setApiKey] = useState(() => localStorage.getItem("gemini_api_key") || "");
  const [routerApiKey, setRouterApiKey] = useState(() => localStorage.getItem("9router_api_key") || "");
  const [routerUrl, setRouterUrl] = useState(() => localStorage.getItem("9router_url") || "https://5b4c-15-135-214-185.ngrok-free.app/v1");
  const [routerModel, setRouterModel] = useState(() => {
    const saved = localStorage.getItem("9router_model");
    return saved === "cx/gpt5.5" || saved === "codex/gpt5.5" || saved === "gpt-4.1"
      ? ROUTER_DEFAULT_MODEL
      : saved || ROUTER_DEFAULT_MODEL;
  });
  const [libreUrl, setLibreUrl] = useState(() => localStorage.getItem("libre_url") || "http://localhost:5000");
  const [libreApiKey, setLibreApiKey] = useState(() => localStorage.getItem("libre_api_key") || "");

  const addLog = useCallback((msg: string) => {
    const time = new Date().toLocaleTimeString("vi-VN");
    setLogs(prev => [...prev, `[${time}] ${msg}`]);
    setTimeout(() => logsEndRef.current?.scrollIntoView({ behavior: "smooth" }), 50);
  }, []);

  const saveSettings = () => {
    localStorage.setItem("trans_engine", engine);
    localStorage.setItem("gemini_api_key", apiKey);
    localStorage.setItem("9router_api_key", routerApiKey);
    localStorage.setItem("9router_url", routerUrl);
    localStorage.setItem("9router_model", routerModel);
    localStorage.setItem("libre_url", libreUrl);
    localStorage.setItem("libre_api_key", libreApiKey);
    toast.success("Đã lưu cài đặt");
    setShowSettings(false);
  };

  const resetState = () => {
    setStatus("idle"); setProgress(0); setProgressLabel("");
    setOriginalText(""); setTranslatedText(""); setError(null); setLogs([]);
    libreRouteRef.current = null; // re-detect route on next translation
  };

  const handleFileChange = (f: File) => {
    setFile(f); resetState();
    setKeepFormat(true);
  };

  // ── LibreTranslate ────────────────────────────────────────────────────────
  // Low-level single request with timeout + retry
  const libreRequest = async (q: string, src: string, tgt: string, retries = 2): Promise<string> => {
    let lastErr: Error = new Error("Unknown");
    for (let attempt = 0; attempt < retries; attempt++) {
      try {
        const body: any = { q, source: src, target: tgt, format: "text" };
        if (libreApiKey) body.api_key = libreApiKey;
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 25000);
        try {
          const res = await fetch(`${libreUrl}/translate`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
            signal: controller.signal,
          });
          clearTimeout(timer);
          if (!res.ok) {
            const errData = await res.json().catch(() => ({}));
            throw new Error(errData.error || `HTTP ${res.status}`);
          }
          return (await res.json()).translatedText as string;
        } finally {
          clearTimeout(timer);
        }
      } catch (e: any) {
        lastErr = e;
        if (attempt < retries - 1) await new Promise(r => setTimeout(r, 800));
      }
    }
    throw lastErr;
  };

  // Detect route once: direct (zh→vi), or two-step (zh→en→vi)
  const ensureLibreRoute = async (src: string, tgt: string): Promise<"direct" | "via-en"> => {
    if (libreRouteRef.current !== null) return libreRouteRef.current;
    addLog(`   🔍 Kiểm tra route LibreTranslate...`);
    try {
      await libreRequest("测试", src, tgt, 1);
      libreRouteRef.current = "direct";
      addLog(`   ✅ Route: ${src}→${tgt} (trực tiếp)`);
    } catch {
      addLog(`   ⚠️ ${src}→${tgt} không được hỗ trợ trực tiếp.`);
      try {
        await libreRequest("测试", src, "en", 1);
        libreRouteRef.current = "via-en";
        addLog(`   🔀 Dùng 2 bước: ${src}→en→${tgt}`);
      } catch {
        libreRouteRef.current = "direct"; // fallback
        addLog(`   ❌ Không xác định được route, thử trực tiếp.`);
      }
    }
    return libreRouteRef.current;
  };

  // Translate a sub-batch by joining with separator → 1 API call → split
  const SEP = "\n⟦S⟧\n";
  const libreTranslateBulk = async (
    texts: string[], src: string, tgt: string
  ): Promise<(string | null)[]> => {
    const joined = texts.join(SEP);
    try {
      const translated = await libreRequest(joined, src, tgt);
      const parts = translated.split(/\n?⟦S⟧\n?/);
      if (parts.length === texts.length) return parts;
      // If split count mismatch, fall back to individual
    } catch { /* fall through to individual */ }
    // Individual fallback
    const results: (string | null)[] = [];
    for (const t of texts) {
      try { results.push(await libreRequest(t, src, tgt)); }
      catch { results.push(null); }
      await new Promise(r => setTimeout(r, 80));
    }
    return results;
  };

  const libreTranslate = async (texts: string[]): Promise<(string | null)[]> => {
    const src = LIBRE_LANG_MAP[sourceLang] ?? "auto";
    const tgt = LIBRE_LANG_MAP[targetLang] ?? "vi";
    const actualSrc = src === "auto" ? "zh-Hans" : src;
    const route = await ensureLibreRoute(actualSrc, tgt);

    const SUB = 8;          // segments per API call
    const CONCURRENCY = 3;  // parallel API calls
    const results: (string | null)[] = new Array(texts.length).fill(null);

    // Build sub-batches
    const subBatches: { idx: number[]; texts: string[] }[] = [];
    for (let i = 0; i < texts.length; i += SUB)
      subBatches.push({ idx: Array.from({ length: Math.min(SUB, texts.length - i) }, (_, k) => i + k), texts: texts.slice(i, i + SUB) });

    // Process with concurrency
    for (let i = 0; i < subBatches.length; i += CONCURRENCY) {
      const lane = subBatches.slice(i, i + CONCURRENCY);
      await Promise.all(lane.map(async ({ idx, texts: sub }) => {
        let out: (string | null)[];
        if (route === "via-en") {
          // Step 1: src→en
          const inEn = await libreTranslateBulk(sub, actualSrc, "en");
          // Step 2: en→tgt
          const nonNull = inEn.map(s => s ?? "");
          const final = await libreTranslateBulk(nonNull, "en", tgt);
          out = final.map((f, k) => (inEn[k] === null ? null : f));
        } else {
          out = await libreTranslateBulk(sub, actualSrc, tgt);
        }
        out.forEach((val, k) => { results[idx[k]] = val; });
      }));
      // Small pause between concurrent waves
      if (i + CONCURRENCY < subBatches.length)
        await new Promise(r => setTimeout(r, 100));
    }
    return results;
  };

  // ── Gemini batch ─────────────────────────────────────────────────────────
  const geminiTranslateBatch = async (batch: string[]): Promise<string[]> => {
    const ai = new GoogleGenAI({ apiKey: apiKey || (process.env as any).GEMINI_API_KEY || "" });
    const prompt = `Bạn là chuyên gia dịch thuật. Dịch danh sách sau từ ${sourceLang} sang ${targetLang}.
YÊU CẦU: Trả về JSON array chuỗi đã dịch. Giữ nguyên số lượng và thứ tự. Không giải thích.
${JSON.stringify(batch)}`;
    const res = await ai.models.generateContent({
      model: GEMINI_MODEL, contents: prompt,
      config: { responseMimeType: "application/json", maxOutputTokens: 8192 },
    });
    let raw = res.text.trim()
      .replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) throw new Error("Không phải JSON array");
    return parsed;
  };

  const callEngine = (batch: string[]) =>
    engine === "libretranslate" ? libreTranslate(batch) : geminiTranslateBatch(batch);

  // Retry with halved batch on failure
  const translateBatch = async (batch: string[], label: string, depth = 0): Promise<Map<string, string>> => {
    const map = new Map<string, string>();
    try {
      const results = await callEngine(batch);
      // Skip null/empty — empty string would corrupt the DOCX by blanking XML nodes
      batch.forEach((orig, i) => {
        const r = results[i];
        if (r !== undefined && r !== null && r !== "") map.set(orig, r as string);
      });
      return map;
    } catch (e) {
      if (engine === "libretranslate" || depth >= 3 || batch.length <= 1) {
        addLog(`   ⚠ Bỏ qua ${batch.length} đoạn sau ${depth} lần thử`);
        return map;
      }
      const half = Math.ceil(batch.length / 2);
      addLog(`   ↩ ${label} lỗi, thử lại (${half} đoạn)...`);
      const [left, right] = await Promise.all([
        translateBatch(batch.slice(0, half), label + "a", depth + 1),
        translateBatch(batch.slice(half), label + "b", depth + 1),
      ]);
      left.forEach((v, k) => map.set(k, v));
      right.forEach((v, k) => map.set(k, v));
      return map;
    }
  };

  // ── XML translation ───────────────────────────────────────────────────────
  const translateXml = async (xml: string, onLog?: (m: string) => void): Promise<string> => {
    const isWord = xml.includes("<w:t");
    const tagName = isWord ? "w:t" : "a:t";
    const regex = isWord
      ? /<w:t([^>]*)>([\s\S]*?)<\/w:t>/g
      : /<a:t([^>]*)>([\s\S]*?)<\/a:t>/g;

    const unique = [...new Set(
      [...xml.matchAll(regex)].map(m => m[2]).filter(t => t.trim().length > 0)
    )];
    if (unique.length === 0) return xml;
    onLog?.(`Tìm thấy ${unique.length} đoạn văn bản.`);

    const BATCH = engine === "libretranslate" ? 20 : 30;
    const total = Math.ceil(unique.length / BATCH);
    const translatedMap = new Map<string, string>();

    for (let i = 0; i < unique.length; i += BATCH) {
      const batch = unique.slice(i, i + BATCH);
      const bn = Math.floor(i / BATCH) + 1;
      onLog?.(`Đang dịch batch ${bn}/${total} (${batch.length} đoạn)...`);
      const result = await translateBatch(batch, `${bn}`);
      result.forEach((v, k) => translatedMap.set(k, v));
      onLog?.(`✓ Batch ${bn}/${total} xong (${result.size}/${batch.length}).`);
    }

    onLog?.("Đang ghi bản dịch vào XML...");
    return xml.replace(regex, (match, attrs, content) => {
      const t = translatedMap.get(content);
      // Guard: undefined or empty → keep original (prevents blank XML nodes → corrupt DOCX)
      if (!t) return match;
      const esc = t.replace(/&/g, "&amp;").replace(/</g, "&lt;")
        .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
      return `<${tagName}${attrs}>${esc}</${tagName}>`;
    });
  };

  // ── Plain text translation ────────────────────────────────────────────────
  const translatePlainText = async (text: string): Promise<string> => {
    if (engine === "libretranslate") {
      const chunks: string[] = [];
      for (let i = 0; i < text.length; i += 2000) chunks.push(text.slice(i, i + 2000));
      addLog(`🔤 LibreTranslate: ${chunks.length} đoạn...`);
      const results = await libreTranslate(chunks);
      return results.join("\n");
    }
    const ai = new GoogleGenAI({ apiKey: apiKey || (process.env as any).GEMINI_API_KEY || "" });
    const chunks: string[] = [];
    for (let i = 0; i < text.length; i += 15000) chunks.push(text.slice(i, i + 15000));
    let full = "";
    for (let i = 0; i < chunks.length; i++) {
      addLog(`Dịch đoạn ${i + 1}/${chunks.length}...`);
      const res = await ai.models.generateContent({
        model: GEMINI_MODEL,
        contents: `Dịch phần ${i + 1}/${chunks.length} từ ${sourceLang} sang ${targetLang}. Chỉ trả về bản dịch.\n\n${chunks[i]}`,
      });
      full += res.text + "\n";
      setProgress(50 + Math.round(((i + 1) / chunks.length) * 48));
    }
    return full.trim();
  };

  // ── Main handler ──────────────────────────────────────────────────────────
  const handleTranslate = async () => {
    if (!file) return;
    if (engine === "gemini" && !apiKey && !(process.env as any).GEMINI_API_KEY) {
      toast.error("Thiếu Gemini API Key! Mở Settings để nhập.");
      setShowSettings(true);
      return;
    }

    const name = file.name.toLowerCase();
    const isPDF = name.endsWith(".pdf");
    const isDocx = name.endsWith(".docx");
    const isPptx = name.endsWith(".pptx");
    const isXlsx = name.endsWith(".xlsx");

    resetState();
    setStatus("extracting");
    setProgress(5);
    setProgressLabel("Đang đọc file...");
    addLog(`📂 File: ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`);
    addLog(`Engine: ${engine === "9router" ? `9router ${routerModel}` : engine === "gemini" ? "Gemini AI" : "LibreTranslate"}`);

    try {
      if (engine === "9router" && (isDocx || isPptx || isXlsx || isPDF)) {
        addLog("9router: dich song ngu truc tiep tren file");
        setProgressLabel("Dang gui file len server...");
        setProgress(15);

        const formData = new FormData();
        formData.append("file", file);
        formData.append("sourceLang", sourceLang);
        formData.append("targetLang", targetLang);
        formData.append("apiKey", routerApiKey);
        formData.append("baseUrl", routerUrl);
        formData.append("model", routerModel);

        setStatus("translating");
        setProgressLabel("9router dang dich va ghi lai file...");
        setProgress(45);
        const res = await fetch("/api/translate-file", { method: "POST", body: formData });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.error || `HTTP ${res.status}`);
        }

        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `Bilingual_${file.name}`;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);

        setOriginalText(file.name);
        setTranslatedText("File song ngu da duoc tao va tai xuong.");
        setStatus("completed"); setProgress(100); setProgressLabel("Hoan tat!");
        addLog("File song ngu da tai xuong thanh cong.");
        toast.success("Dich song ngu va tai xuong thanh cong!");
        return;
      }

      // ── Keep-format mode ────────────────────────────────────────────────
      if (keepFormat && (isDocx || isPptx)) {
        addLog("📐 Chế độ: Giữ nguyên định dạng");
        setProgressLabel("Đang phân tích cấu trúc file...");
        const parts = await getXmlParts(file);
        addLog(`✅ Tìm thấy ${parts.length} phần XML.`);
        setProgress(20);
        setStatus("translating");

        const translated: { path: string; content: string }[] = [];
        for (let i = 0; i < parts.length; i++) {
          addLog(`\n🔄 Phần ${i + 1}/${parts.length}: ${parts[i].path}`);
          setProgressLabel(`Dịch phần ${i + 1}/${parts.length}`);
          const content = await translateXml(parts[i].content, m => addLog(`   ${m}`));
          translated.push({ path: parts[i].path, content });
          setProgress(20 + Math.round(((i + 1) / parts.length) * 65));
        }

        addLog("\n📦 Đang đóng gói file...");
        setProgressLabel("Đóng gói file...");
        const blob = await rebuildZip(file, translated);
        addLog(`✅ Xong! File: ${(blob.size / 1024).toFixed(1)} KB`);

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = `Translated_${file.name}`;
        document.body.appendChild(a); a.click();
        setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);

        setStatus("completed"); setProgress(100); setProgressLabel("Hoàn tất!");
        addLog("🎉 File đã tải xuống thành công!");
        toast.success("Dịch và tải xuống thành công!");
        return;
      }

      // ── Extract text mode ───────────────────────────────────────────────
      addLog("📝 Chế độ: Trích xuất văn bản");
      setProgressLabel("Đang trích xuất văn bản...");

      let text = "";
      if (isPDF) { addLog("📄 Đọc PDF..."); text = await extractPDF(file); }
      else if (isDocx) { addLog("📝 Đọc DOCX..."); text = await extractDOCX(file); }
      else if (isPptx) { addLog("📊 Đọc PPTX..."); text = await extractPPTX(file); }
      else throw new Error("Định dạng không hỗ trợ (.pdf, .docx, .pptx)");

      if (!text.trim()) throw new Error("Không tìm thấy văn bản trong file.");
      addLog(`✅ Trích xuất: ${text.length} ký tự`);
      setOriginalText(text);
      setProgress(40);
      setStatus("translating");
      setProgressLabel("Đang dịch thuật...");
      addLog("🤖 Bắt đầu dịch thuật...");
      const result = await translatePlainText(text);
      setTranslatedText(result);
      setStatus("completed"); setProgress(100); setProgressLabel("Hoàn tất!");
      addLog("🎉 Dịch thuật hoàn tất!");
      toast.success("Dịch thuật hoàn tất!");
    } catch (err: any) {
      addLog(`❌ LỖI: ${err.message}`);
      setError(err.message);
      setStatus("error");
      toast.error("Lỗi: " + err.message);
    }
  };

  const downloadAsDocx = async () => {
    if (!translatedText) return;
    const doc = new Document({
      sections: [{ properties: {}, children: translatedText.split("\n").map(line =>
        new Paragraph({ children: [new TextRun(line)] })
      )}],
    });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `Translated_${file?.name.split(".")[0] || "Document"}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  };

  const isBusy = status === "extracting" || status === "translating";

  // ── UI ────────────────────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans">
      <Toaster position="top-center" />

      {/* Header */}
      <header className="border-b bg-white/80 backdrop-blur-md sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 p-2 rounded-lg">
              <Languages className="w-6 h-6 text-white" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">DocTrans AI</h1>
          </div>
          <div className="flex items-center gap-2">
            <Button variant="ghost" size="icon"
              onClick={() => setShowSettings(!showSettings)}
              className={showSettings ? "bg-blue-50 text-blue-600" : "text-gray-500"}>
              <Settings className="w-5 h-5" />
            </Button>
            <Badge variant="outline" className={engine === "libretranslate"
              ? "bg-green-50 text-green-700 border-green-200"
              : engine === "9router"
                ? "bg-violet-50 text-violet-700 border-violet-200"
                : "bg-blue-50 text-blue-700 border-blue-200"}>
              {engine === "9router" ? `9router ${routerModel}` : engine === "libretranslate" ? "LibreTranslate" : "Gemini AI"}
            </Badge>
          </div>
        </div>
      </header>

      {/* Settings Panel */}
      <AnimatePresence>
        {showSettings && (
          <motion.div initial={{ height: 0, opacity: 0 }} animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }} className="bg-white border-b overflow-hidden">
            <div className="max-w-6xl mx-auto px-4 py-6">
              <div className="max-w-2xl space-y-5">
                <div>
                  <h3 className="text-sm font-semibold mb-3 flex items-center gap-2">
                    <Cpu className="w-4 h-4 text-blue-600" /> Công cụ dịch thuật
                  </h3>
                  <div className="grid grid-cols-2 gap-3">
                    {(["9router", "gemini", "libretranslate"] as Engine[]).map(e => (
                      <button key={e} onClick={() => setEngine(e)}
                        className={`p-4 rounded-xl border-2 text-left transition-all ${
                          engine === e
                            ? e === "gemini" ? "border-blue-500 bg-blue-50" : "border-green-500 bg-green-50"
                            : "border-gray-200 hover:border-gray-300"}`}>
                        <div className="font-semibold text-sm">
                          {e === "9router" ? "9router" : e === "gemini" ? "Gemini AI" : "LibreTranslate"}
                        </div>
                        <div className="text-xs text-gray-500 mt-1">
                          {e === "9router" ? "Song ngu Viet-Trung, giu cau truc file" : e === "gemini" ? "Chat luong cao, can API Key" : "Offline/Local, mien phi"}
                        </div>
                      </button>
                    ))}
                  </div>
                </div>

                {engine === "9router" && (
                  <div className="space-y-3">
                    <div>
                      <Label className="text-xs text-gray-500 mb-1 block">9router Endpoint</Label>
                      <Input placeholder="https://.../v1" value={routerUrl}
                        onChange={e => setRouterUrl(e.target.value)} className="h-10 font-mono text-sm" />
                    </div>
                    <div>
                      <Label className="text-xs text-gray-500 mb-1 block">Model</Label>
                      <select value={ROUTER_MODEL_OPTIONS.includes(routerModel) ? routerModel : "__custom"}
                        onChange={e => {
                          if (e.target.value !== "__custom") setRouterModel(e.target.value);
                        }}
                        className="w-full h-10 bg-gray-50 border border-gray-200 rounded-md px-3 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-500">
                        {ROUTER_MODEL_OPTIONS.map(model => <option key={model} value={model}>{model}</option>)}
                        <option value="__custom">Custom...</option>
                      </select>
                      <Input placeholder="Nhap model custom neu can" value={routerModel}
                        onChange={e => setRouterModel(e.target.value.trim())} className="h-10 font-mono text-sm mt-2" />
                    </div>
                    <div>
                      <Label className="text-xs text-gray-500 mb-1 block">API Key</Label>
                      <Input type="password" placeholder="sk-..." value={routerApiKey}
                        onChange={e => setRouterApiKey(e.target.value)} className="h-10" />
                    </div>
                  </div>
                )}

                {engine === "gemini" && (
                  <div>
                    <Label className="text-xs text-gray-500 flex items-center gap-1 mb-2">
                      <Key className="w-3 h-3" /> Gemini API Key
                    </Label>
                    <Input type="password" placeholder="AIza..." value={apiKey}
                      onChange={e => setApiKey(e.target.value)} className="h-10" />
                    <p className="text-[10px] text-gray-400 mt-1">
                      Lấy tại{" "}
                      <a href="https://aistudio.google.com/apikey" target="_blank"
                        className="text-blue-500 underline">aistudio.google.com/apikey</a>
                    </p>
                  </div>
                )}

                {engine === "libretranslate" && (
                  <div className="space-y-3">
                    <div>
                      <Label className="text-xs text-gray-500 mb-1 block">URL Server</Label>
                      <Input placeholder="http://localhost:5000" value={libreUrl}
                        onChange={e => setLibreUrl(e.target.value)} className="h-10 font-mono text-sm" />
                    </div>
                    <div>
                      <Label className="text-xs text-gray-500 mb-1 block">API Key (nếu có)</Label>
                      <Input type="password" placeholder="Để trống nếu không cần" value={libreApiKey}
                        onChange={e => setLibreApiKey(e.target.value)} className="h-10" />
                    </div>
                    <p className="text-[10px] text-gray-400 font-mono bg-gray-50 p-2 rounded">
                      pip install libretranslate &amp;&amp; libretranslate
                    </p>
                  </div>
                )}

                <Button onClick={saveSettings} className="bg-blue-600 hover:bg-blue-700 text-white">
                  Lưu cài đặt
                </Button>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      <main className="max-w-6xl mx-auto px-4 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">

          {/* Left Controls */}
          <div className="lg:col-span-4 space-y-6">
            <Card className="border-none shadow-sm bg-white">
              <CardHeader>
                <CardTitle className="text-lg">Tải lên tài liệu</CardTitle>
                <CardDescription>Ho tro Word (.docx), Excel (.xlsx), PPTX, PDF</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                {/* Drop zone */}
                <div
                  className={`border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer
                    ${file ? "border-blue-500 bg-blue-50" : "border-gray-200 hover:border-blue-400 hover:bg-gray-50"}`}
                  onClick={() => document.getElementById("file-upload")?.click()}
                  onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) handleFileChange(f); }}
                  onDragOver={e => e.preventDefault()}
                >
                  <input id="file-upload" type="file" className="hidden"
                    accept=".pdf,.docx,.xlsx,.pptx"
                    onChange={e => { const f = e.target.files?.[0]; if (f) handleFileChange(f); }} />
                  {file ? (
                    <div className="flex flex-col items-center gap-2">
                      <FileCheck className="w-12 h-12 text-blue-600" />
                      <p className="font-medium text-sm truncate max-w-full">{file.name}</p>
                      <p className="text-xs text-gray-500">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
                      <Button variant="ghost" size="sm"
                        className="mt-2 text-red-500 hover:text-red-600 hover:bg-red-50"
                        onClick={e => { e.stopPropagation(); setFile(null); resetState(); }}>
                        <X className="w-4 h-4 mr-1" /> Gỡ bỏ
                      </Button>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center gap-2">
                      <Upload className="w-12 h-12 text-gray-400" />
                      <p className="font-medium">Nhấp để tải lên</p>
                      <p className="text-xs text-gray-500">hoặc kéo và thả vào đây</p>
                    </div>
                  )}
                </div>

                {/* Language selectors */}
                <div className="flex items-center gap-2">
                  <div className="flex-1 space-y-1">
                    <Label className="text-xs text-gray-500 uppercase tracking-wider">Từ</Label>
                    <select value={sourceLang} onChange={e => setSourceLang(e.target.value)}
                      className="w-full bg-gray-50 border border-gray-200 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      {LANG_OPTIONS.map(l => <option key={l}>{l}</option>)}
                    </select>
                  </div>
                  <Button variant="ghost" size="icon" className="mt-6 rounded-full hover:bg-blue-50 hover:text-blue-600"
                    onClick={() => { if (sourceLang !== "Tự động phát hiện") { setSourceLang(targetLang); setTargetLang(sourceLang); } }}
                    disabled={sourceLang === "Tự động phát hiện"}>
                    <ArrowRightLeft className="w-4 h-4" />
                  </Button>
                  <div className="flex-1 space-y-1">
                    <Label className="text-xs text-gray-500 uppercase tracking-wider">Sang</Label>
                    <select value={targetLang} onChange={e => setTargetLang(e.target.value)}
                      className="w-full bg-gray-50 border border-gray-200 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      {LANG_OPTIONS.filter(l => l !== "Tự động phát hiện").map(l => <option key={l}>{l}</option>)}
                    </select>
                  </div>
                </div>

                {/* Keep format */}
                <div className="flex items-center space-x-2 pt-1">
                  <input type="checkbox" id="keep-format" checked={keepFormat}
                    disabled={false}
                    onChange={e => setKeepFormat(e.target.checked)}
                    className="w-4 h-4 text-blue-600 border-gray-300 rounded" />
                  <Label htmlFor="keep-format" className="text-sm cursor-pointer">
                    Dich song ngu truc tiep, giu cau truc file
                  </Label>
                </div>
              </CardContent>
              <CardFooter>
                <Button
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white py-6 text-lg font-semibold shadow-md shadow-blue-200"
                  disabled={!file || isBusy} onClick={handleTranslate}>
                  {isBusy
                    ? <><Loader2 className="w-5 h-5 mr-2 animate-spin" /> Đang xử lý...</>
                    : "Dịch ngay"}
                </Button>
              </CardFooter>
            </Card>

            {/* Progress + Logs */}
            <AnimatePresence>
              {(isBusy || status === "completed") && (
                <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 10 }}>
                  <Card className="border-none shadow-sm bg-white">
                    <CardHeader className="pb-2">
                      <div className="flex justify-between items-center">
                        <CardTitle className="text-sm font-medium">
                          {isBusy ? "Đang xử lý..." : "✅ Hoàn tất!"}
                        </CardTitle>
                        <span className="text-xs font-bold text-blue-600">{progress}%</span>
                      </div>
                      {progressLabel && <p className="text-xs text-gray-500 truncate">{progressLabel}</p>}
                    </CardHeader>
                    <CardContent className="space-y-3">
                      <Progress value={progress} className="h-2 bg-blue-100" />
                      {logs.length > 0 && (
                        <div className="bg-gray-900 rounded-lg p-3 max-h-52 overflow-y-auto font-mono text-xs">
                          {logs.map((log, i) => (
                            <div key={i} className={`leading-relaxed whitespace-pre-wrap ${
                              log.includes("❌") ? "text-red-400" :
                              log.includes("✅") || log.includes("🎉") ? "text-green-400" :
                              log.includes("⚠") ? "text-yellow-400" :
                              log.includes("↩") ? "text-orange-400" :
                              log.includes("🔄") || log.includes("📦") ? "text-blue-300" :
                              "text-gray-300"
                            }`}>{log}</div>
                          ))}
                          <div ref={logsEndRef} />
                        </div>
                      )}
                    </CardContent>
                  </Card>
                </motion.div>
              )}
            </AnimatePresence>

            {status === "error" && (
              <div className="bg-red-50 border border-red-200 rounded-xl p-4 flex gap-3 items-start">
                <FileWarning className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                <div>
                  <p className="text-sm font-semibold text-red-800">Đã xảy ra lỗi</p>
                  <p className="text-xs text-red-600 mt-1">{error}</p>
                </div>
              </div>
            )}
          </div>

          {/* Right: Results */}
          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm bg-white min-h-[500px] flex flex-col">
              <CardHeader className="flex flex-row items-center justify-between border-b pb-4">
                <div>
                  <CardTitle className="text-lg">Kết quả dịch thuật</CardTitle>
                  <CardDescription>Bản dịch sẽ xuất hiện tại đây</CardDescription>
                </div>
                <div className="flex gap-2">
                  <Button variant="outline" size="sm" disabled={!translatedText}
                    onClick={() => { navigator.clipboard.writeText(translatedText); toast.success("Đã sao chép!"); }}
                    className="h-9">
                    <Copy className="w-4 h-4 mr-2" /> Sao chép
                  </Button>
                  <Button variant="outline" size="sm" disabled={!translatedText}
                    onClick={downloadAsDocx}
                    className="h-9 border-blue-200 text-blue-700 hover:bg-blue-50">
                    <Download className="w-4 h-4 mr-2" /> Tải về .docx
                  </Button>
                </div>
              </CardHeader>
              <CardContent className="flex-1 p-0 relative">
                {!translatedText && !isBusy && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center text-gray-400 p-8 text-center">
                    <div className="bg-gray-50 p-6 rounded-full mb-4">
                      <FileText className="w-12 h-12" />
                    </div>
                    <p className="text-lg font-medium">Chưa có dữ liệu</p>
                    <p className="text-sm max-w-xs">Tải lên file và nhấn "Dịch ngay" để bắt đầu.</p>
                  </div>
                )}
                {isBusy && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center bg-white/60 backdrop-blur-[2px] z-10">
                    <Loader2 className="w-10 h-10 text-blue-600 animate-spin mb-4" />
                    <p className="font-medium text-blue-800 animate-pulse">
                      {status === "extracting" ? "Đang đọc tài liệu..." : "Đang dịch thuật..."}
                    </p>
                  </div>
                )}
                {translatedText && (
                  <div className="p-6 h-full max-h-[600px] overflow-y-auto">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div className="space-y-3">
                        <Badge variant="secondary" className="bg-gray-100 text-gray-600">
                          Gốc ({sourceLang})
                        </Badge>
                        <div className="text-sm text-gray-600 whitespace-pre-wrap leading-relaxed bg-gray-50 p-4 rounded-lg border border-gray-100">
                          {originalText}
                        </div>
                      </div>
                      <div className="space-y-3">
                        <Badge variant="secondary" className="bg-blue-100 text-blue-700">
                          Bản dịch ({targetLang})
                        </Badge>
                        <div className="text-sm whitespace-pre-wrap leading-relaxed bg-white p-4 rounded-lg border border-blue-100 shadow-sm">
                          {translatedText}
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>
      </main>

      <footer className="max-w-6xl mx-auto px-4 py-8 text-center text-gray-500 text-xs">
        <Separator className="mb-6" />
        <p>© 2026 DocTrans AI — Xử lý hoàn toàn trên trình duyệt, không cần server.</p>
      </footer>
    </div>
  );
}
