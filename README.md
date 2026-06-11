# Translate AI Tool

Ứng dụng web dịch tài liệu bằng AI, hỗ trợ dịch văn bản và file tài liệu với nhiều engine khác nhau.

## Tính năng

- Dịch tài liệu `PDF`, `DOCX`, `PPTX` và văn bản trích xuất từ file.
- Hỗ trợ engine `9router`, `Gemini` và `LibreTranslate`.
- Tùy chọn giữ định dạng cho `DOCX` và `PPTX`.
- Xuất kết quả dịch ra file tải xuống.
- Lưu cấu hình API trong trình duyệt bằng `localStorage`.
- Có server Express để xử lý upload, trích xuất nội dung và dịch theo batch.

## Yêu cầu

- Node.js 18 trở lên.
- API key cho engine muốn dùng:
  - `NINEROUTER_API_KEY` cho 9router.
  - `GEMINI_API_KEY` cho Gemini.
  - LibreTranslate server nếu dùng LibreTranslate.

## Cài đặt

```bash
npm install
```

Tạo file `.env` từ `.env.example`, rồi điền key cần dùng:

```env
NINEROUTER_BASE_URL="https://your-router-url/v1"
NINEROUTER_MODEL="cx/gpt-5.5"
NINEROUTER_API_KEY=""

GEMINI_API_KEY=""
APP_URL="http://localhost:3000"
```

## Chạy local

```bash
npm run dev
```

Mở trình duyệt tại:

```text
http://localhost:3000
```

## Scripts

```bash
npm run dev      # chạy Express + Vite dev server
npm run build    # build frontend
npm run preview  # preview bản build
npm run lint     # kiểm tra TypeScript
```

## Cấu hình trong app

Trong giao diện, mở phần Settings để chọn engine dịch:

- `9router`: nhập base URL, model và API key.
- `Gemini`: nhập Gemini API key.
- `LibreTranslate`: nhập URL server và API key nếu server yêu cầu.

## Ghi chú bảo mật

Không commit file `.env` hoặc API key thật lên GitHub. Repo đã ignore `.env*` và chỉ giữ `.env.example` làm mẫu cấu hình.
