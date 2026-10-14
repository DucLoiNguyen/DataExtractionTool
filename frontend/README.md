# Frontend — Tạo danh sách tài khoản

Vite + React 19 + TypeScript, Tailwind CSS v4, thành phần kiểu shadcn/ui (Radix), TanStack Table + Virtual, Motion.
Thiết kế theo [`../DESIGN.md`](../DESIGN.md). Hướng dẫn chạy/build, cấu hình và API: [`../docs/web.md`](../docs/web.md).

```bash
npm install
npm run dev       # http://localhost:5173, /api được chuyển sang http://127.0.0.1:8000 (đổi bằng VITE_API_PROXY)
npm run build     # -> dist/ (tệp tĩnh)
npm run lint
npm test
```

Cần Node 20.19+ hoặc 22.12+.

```
src/
  main.tsx, App.tsx       điểm vào, trạng thái trang (chế độ, nguồn, tác vụ)
  index.css               token màu/chữ (khớp DESIGN.md) + Tailwind
  lib/api.ts, types.ts    gọi backend (chỉ dùng đường dẫn tương đối /api/...)
  lib/text.ts             bỏ dấu tiếng Việt cho tìm kiếm/sắp xếp
  components/             UploadZone, SourceList, ResultView, ResultTable (ảo hoá hàng), CountUp
  components/ui/          Button, Input, Tabs, Checkbox (kiểu shadcn/ui, viết tay)
```
