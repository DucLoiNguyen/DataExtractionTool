# Phiên bản web

Web gồm **hai phần tách biệt**, giao tiếp qua HTTP/JSON:

| Phần | Thư mục | Công nghệ | Vai trò |
|---|---|---|---|
| Backend (API) | `src/dataextraction/web/app.py` | FastAPI, Uvicorn | Nhận tệp, gọi thẳng `core.py` (Chế độ 1) và `mode2.py` (Chế độ 2), trả JSON/tệp `.xlsx`. **Không phục vụ giao diện.** |
| Frontend | `frontend/` | Vite + React + TypeScript, Tailwind CSS v4, thành phần kiểu shadcn/ui, TanStack Table/Virtual, Motion | Giao diện; build ra tệp tĩnh trong `frontend/dist/` |

Sửa quy tắc trích xuất ở `core.py`/`mode2.py` là áp dụng cho GUI Tkinter lẫn web. Giao diện theo `DESIGN.md` ở gốc repo.

Việc đóng gói và triển khai (Docker, Nginx, CI, Kubernetes...) do chủ dự án tự làm; repo cố ý không có các tệp đó.

## Chạy khi phát triển (2 cửa sổ dòng lệnh)

```bash
# 1) Backend: http://127.0.0.1:8000
pip install -e ".[web]"
python -m dataextraction.web            # thêm --reload khi sửa code

# 2) Frontend: http://localhost:5173 (tự chuyển /api sang backend)
cd frontend
npm install
npm run dev
```

Mở http://localhost:5173. Nếu backend chạy ở nơi khác: `VITE_API_PROXY=http://host:cổng npm run dev`.

## Build frontend
```bash
cd frontend
npm run build      # tsc + vite build -> frontend/dist/ (tệp tĩnh, JS ~150 KB sau gzip)
npm run lint
npm test           # Vitest
```
Trình duyệt chỉ gọi đường dẫn tương đối `/api/...`, nên khi triển khai thật cần một lớp phía trước phục vụ `dist/` và chuyển `/api/*` sang backend (cùng domain, không cần CORS). Backend không tự phục vụ giao diện.

## Cấu hình backend (biến môi trường)
| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `DATAEXTRACTION_MAX_UPLOAD_MB` | 100 | Dung lượng tối đa mỗi tệp tải lên |
| `DATAEXTRACTION_TTL_SECONDS` | 3600 | Thời gian giữ tệp tải lên và kết quả rồi xoá |
| `DATAEXTRACTION_WORK_DIR` | thư mục tạm của hệ điều hành | Nơi ghi tệp tải lên/kết quả (cần ghi được) |

Tham số dòng lệnh: `--host`, `--port`, `--reload`. Cổng mặc định 8000.

## API
| Phương thức | Đường dẫn | Mô tả |
|---|---|---|
| GET | `/api/health` | Kiểm tra sống (`{"ok": true}`), dùng làm healthcheck |
| GET | `/api/config` | Mật khẩu mặc định, dung lượng tối đa |
| POST | `/api/uploads` | Tải **1 tệp** lên 1 lần (multipart `file`) → `file_id`, danh sách sheet có email, gợi ý Tổ chức |
| DELETE | `/api/uploads/{file_id}` | Xoá tệp đã tải |
| POST | `/api/jobs` | Tạo việc (trả `202` ngay): `{"mode": "mode1" hoặc "mode2", "sources": [{"file_id", "sheet", "to_chuc"}], "dedupe", "password"}` → `job_id` |
| GET | `/api/jobs/{job_id}` | `{"status": "running"}`, hoặc `done` kèm `result`, hoặc `error` kèm thông báo |
| GET | `/api/download/{job_id}/result` hoặc `/issues` | Tải tệp Excel kết quả / danh sách cần kiểm tra |

Chế độ 1 chỉ nhận 1 nguồn. Chế độ 2 nhận nhiều nguồn; Excel nhiều sheet là mỗi sheet một nguồn.

## Hiệu năng (đã làm)
- Tải lên đúng 1 lần mỗi tệp (trước đây tệp được gửi 3 lần: sheet, gợi ý, trích xuất); chạy lại với mật khẩu/Tổ chức khác không cần tải lại.
- Việc nặng chạy trong luồng riêng, `/api/health` và các yêu cầu nhẹ vẫn trả lời ngay khi đang đọc tệp lớn (có test). Giao diện hỏi tiến độ bằng `/api/jobs/{id}` thay vì giữ một kết nối HTTP mở lâu (proxy phía trước thường ngắt kết nối sau ~60 giây).
- JSON kết quả được nén gzip.
- Bảng kết quả ảo hoá hàng (chỉ vẽ ~28 hàng đang thấy dù có ~6000 dòng), chuỗi tìm kiếm không dấu tính một lần khi nhận dữ liệu.

## Giới hạn đã biết (để bạn quyết khi triển khai)
- **Chỉ chạy được 1 bản backend**: danh sách tệp tải lên và việc đang chạy nằm trong bộ nhớ của tiến trình. Chạy 2 bản trở lên thì yêu cầu có thể rơi vào bản không có dữ liệu. Muốn mở rộng cần chuyển trạng thái ra ngoài (Redis hoặc thư mục dùng chung) và dùng hàng đợi.
- Mỗi lúc chỉ xử lý một việc trích xuất (khoá toàn cục `_work_lock`; cần thiết trên Windows vì Word qua COM, trên Linux có thể nới sau khi đo).
- Giao diện chưa hiện tiến độ chi tiết ("đang đọc tệp 2/5"): `core`/`mode2` chưa có hàm báo tiến độ, backend mới biết `running`/`done`/`error`.
- Không có đăng nhập; tệp chứa dữ liệu cá nhân của cán bộ, chỉ nên chạy trong mạng nội bộ hoặc đặt sau lớp xác thực.
- Tệp `.doc` cần LibreOffice trên máy chủ Linux (Word/pywin32 chỉ có trên Windows).

## Kiểm thử
```bash
pip install -e ".[dev]"
pytest -q                     # API (gọi API thật, không mock logic trích xuất)
cd frontend && npm test       # Vitest (hàm tiện ích)
```
