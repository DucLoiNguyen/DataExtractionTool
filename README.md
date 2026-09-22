# Bộ công cụ tạo danh sách tài khoản

**1 giao diện duy nhất**, cho phép chọn 1 trong 2 chế độ trích xuất tuỳ theo
loại dữ liệu đầu vào bạn có.

## Cấu trúc thư mục

```
📁 (thư mục này)
├── gui.py                       🖥️  CHẠY FILE NÀY — giao diện duy nhất
├── extract_contacts.py          ⚙️  Logic dùng chung (KHÔNG chạy trực tiếp)
├── process_template_v2.py       ⚙️  Logic riêng cho chế độ 2 (KHÔNG chạy trực tiếp)
├── cls_template_users.xlsx      📄  File mẫu cho chế độ 1
├── TemplateV2.xlsx              📄  File mẫu cho chế độ 2
└── README.md                    📖  File này
```

**Quan trọng**: cả 5 file trên phải luôn nằm **cùng 1 thư mục**. Chỉ cần chạy
`gui.py` — không cần chạy 2 file logic kia trực tiếp.

## Cài đặt (chỉ cần làm 1 lần)

```bash
pip install openpyxl python-docx pdfplumber xlrd
```

## Chạy

```bash
python3 gui.py
```

## Chọn chế độ

Ngay khi mở tool, chọn 1 trong 2 chế độ ở đầu trang:

| | Chế độ 1 | Chế độ 2 |
|---|---|---|
| **Dùng khi** | Có công văn/danh sách tự do dạng PDF, Word, Excel/CSV | Đã có bảng Excel với cột Họ tên/Đơn vị công tác/Email/SĐT |
| **Đầu vào** | Nhiều tệp PDF/Word/Excel/CSV | 1 tệp Excel |
| **File mẫu** | `cls_template_users.xlsx` | `TemplateV2.xlsx` |
| **Đầu ra** | Email, Mật khẩu, Họ và tên, Điện thoại | Tên tài khoản, Email, SĐT, Mật khẩu*, Giới tính*, Ngày sinh*, Đơn vị, Tổ chức |

*(để trống, không tự bịa)*

Chuyển chế độ sẽ tự động đổi file mẫu mặc định, đổi ô nhập liệu phù hợp, và
xoá lựa chọn tệp cũ để tránh nhầm lẫn.

## Quy tắc xử lý (áp dụng cho CẢ 2 chế độ)

1. **Email**: chuẩn theo định dạng hợp lệ, viết thường. Thiếu hoặc sai định
   dạng → **loại bỏ bản ghi** (không tự đoán/sửa email).
2. **Số điện thoại**: chuẩn theo định dạng Việt Nam (10 số, bắt đầu bằng 0).
   Không chuẩn hoá được → để trống (không loại cả bản ghi, chỉ email mới bắt buộc).
3. **Họ và tên**: viết hoa chữ cái đầu mỗi từ, bỏ ký tự đặc biệt/khoảng trắng thừa.
4. **Trùng email**: tự động loại bỏ (có thể tắt bằng ô tuỳ chọn).

**Riêng chế độ 2** có thêm 2 quy tắc chuẩn hoá:
- **Đơn vị**: `<cấp hành chính> <tên riêng> - <tỉnh/thành>` (vd `UBND phường
  Ninh Kiều` → `Phường Ninh Kiều - Cần Thơ`).
- **Tổ chức**: tên chính thức hiện nay của tỉnh/thành cấp 1, viết hoa toàn
  bộ. Tự nhận diện qua tên tỉnh cũ/mới trong dữ liệu, dùng ô "Tổ chức mặc
  định" làm phương án dự phòng. Có sẵn bảng tra cứu đầy đủ 34 tỉnh/thành sau
  sáp nhập 2025 (đã kiểm chứng qua Cổng TTĐT Chính phủ).

## Về độ chính xác

Tool **không bao giờ tự bịa dữ liệu thiếu**. Mọi dòng không đáp ứng đủ điều
kiện (thiếu/sai email, trùng email) đều bị loại khỏi kết quả và được liệt kê
**đầy đủ** trong tab **"Cần kiểm tra"** — kèm tệp nguồn/STT, toàn bộ dữ liệu
gốc, và lý do cụ thể — để bạn tự đối chiếu, sửa dữ liệu nguồn nếu cần, hoặc
xác nhận đó đúng là dữ liệu cần bỏ qua. Số liệu ở mục "Thống kê" và bảng
"Xem trước" luôn khớp chính xác 1-1 với dữ liệu thật, không làm tròn hay ước lượng.

Với dữ liệu KHÔNG đúng chuẩn ngay từ nguồn (vd email viết sai hoàn toàn, tên
đơn vị không có trong bảng tra cứu), tool sẽ không tự đoán mò mà đưa vào
"Cần kiểm tra" để con người quyết định — đây là cách duy nhất đảm bảo dữ
liệu ĐÃ XUẤT RA luôn chính xác, thay vì cố gắng "đoán" và có rủi ro sai.

## Luồng thao tác chung (cả 2 chế độ)

1. Chọn chế độ → chọn tệp nguồn → (chế độ 2: nhập thêm "Tổ chức mặc định").
2. Bấm **"Trích xuất & Xem trước"** (hoặc `Enter`) — xem thống kê, bảng "Xem
   trước" và tab "Cần kiểm tra". **Chưa ghi ra tệp** ở bước này.
3. Kiểm tra xong, bấm **"Xuất tệp kết quả..."** (hoặc `Ctrl+S`) — lúc này mới
   hiện hộp thoại đặt tên/chọn nơi lưu, và tệp mới thực sự được ghi. Nếu có
   dòng cần kiểm tra, tool tự ghi kèm 1 tệp báo cáo `..._can_kiem_tra.xlsx`.

**Phím tắt**: `Ctrl+O` chọn tệp · `Enter` trích xuất · `Ctrl+S` xuất tệp.

Có thể tìm kiếm (không phân biệt dấu) và bấm tiêu đề cột để sắp xếp trong cả
2 tab "Xem trước" và "Cần kiểm tra".

---

Chi tiết kỹ thuật đầy đủ (cách đọc từng loại tệp, bảng tra cứu 34 tỉnh/thành,
các trường hợp đặc biệt...) xem docstring đầu file `extract_contacts.py` và
`process_template_v2.py`.
