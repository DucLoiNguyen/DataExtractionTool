# Phiên bản web

Giao diện web cho cùng bộ máy trích xuất của `gui.py`. `webapp/app.py` (FastAPI) chỉ nhận tệp tải lên rồi gọi thẳng
`extract_contacts.py` (Chế độ 1) và `process_template_v2.py` (Chế độ 2), nên sửa quy tắc ở đó là áp dụng cho cả GUI lẫn web.

## Chạy

```bash
pip install -r webapp/requirements.txt
uvicorn webapp.app:app --host 0.0.0.0 --port 8000     # chạy từ thư mục gốc của repo
```

Mở http://localhost:8000.

## Tính năng
- 2 chế độ như GUI. Chế độ 2 nhận nhiều tệp, tệp Excel nhiều sheet tách theo sheet, mỗi nguồn 1 Tổ chức mặc định
  (tự gợi ý từ tên miền email), sheet trùng ≥ 50% email với sheet trước mặc định bị bỏ (bật lại được).
- Xem trước và "Cần kiểm tra" có tìm kiếm không dấu, sắp xếp theo cột; cảnh báo hiển thị dưới thanh thống kê.
- Tải tệp kết quả `.xlsx` theo file mẫu, và `.xlsx` các dòng cần kiểm tra (Chế độ 2).

## API
`POST /api/extract` (multipart: `mode`, `files`, `sources` JSON, `dedupe`, `password`), `POST /api/sheets`,
`POST /api/suggest-to-chuc`, `GET /api/download/{job}/{result|issues}`.

## Lưu ý
- Kết quả giữ trong thư mục tạm 1 giờ rồi xoá; tệp tải lên chứa dữ liệu cá nhân, không có đăng nhập nên chỉ nên chạy
  trong mạng nội bộ hoặc đặt sau lớp xác thực.
- Tệp `.doc` cần LibreOffice trên máy chủ (Word/pywin32 chỉ có trên Windows).
- Mỗi lần chỉ xử lý một yêu cầu (khoá toàn cục), vì các hàm đọc tệp dùng chung tài nguyên.

## Kiểm thử
```bash
pip install -r requirements-dev.txt
pytest -q
```
