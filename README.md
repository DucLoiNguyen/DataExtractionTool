# Bộ công cụ tạo danh sách tài khoản

**1 giao diện duy nhất**, cho phép chọn 1 trong 2 chế độ trích xuất tuỳ theo
loại dữ liệu đầu vào bạn có.

## Cấu trúc thư mục

```
📁 (thư mục này)
├── gui.py                       🖥️  CHẠY FILE NÀY — giao diện duy nhất
├── extract_contacts.py          ⚙️  Logic dùng chung + chế độ 1
├── process_template_v2.py       ⚙️  Logic riêng cho chế độ 2
├── cls_template_users.xlsx      📄  File mẫu cho chế độ 1
├── TemplateV2.xlsx              📄  File mẫu cho chế độ 2
├── regression.py                🧪  Kiểm tra hồi quy (dành cho người bảo trì)
└── README.md                    📖  File này
```

**Quan trọng**: `gui.py`, 2 file logic và 2 file mẫu phải luôn nằm **cùng 1
thư mục**. Chỉ cần chạy `gui.py`.

## Cài đặt (chỉ cần làm 1 lần)

```bash
pip install openpyxl python-docx pdfplumber xlrd
```

Tuỳ chọn: cài [LibreOffice](https://www.libreoffice.org/) nếu cần đọc tệp Word
cũ `.doc` (Word 97-2003). Không có LibreOffice thì mở tệp bằng Word, chọn
"Save As" sang `.docx` rồi dùng tệp mới.

## Chạy

```bash
python3 gui.py
```

## Chọn chế độ

Chọn 1 trong 2 tab ở đầu trang:

| | Chế độ 1 | Chế độ 2 |
|---|---|---|
| **Dùng khi** | Chỉ cần Email / Họ tên / SĐT | Cần thêm Đơn vị và Tổ chức |
| **Đầu vào** | Nhiều tệp PDF, Word (.docx/.doc), Excel (.xlsx/.xls), CSV | 1 tệp PDF, Word (.docx/.doc), Excel (.xlsx/.xls) hoặc CSV |
| **File mẫu** | `cls_template_users.xlsx` | `TemplateV2.xlsx` |
| **Đầu ra** | Email, Mật khẩu, Họ và tên, Điện thoại | Tên tài khoản, Email, SĐT, Mật khẩu, Giới tính*, Ngày sinh*, Đơn vị, Tổ chức |

\* để trống — nguồn không có thông tin này nên tool không tự bịa.

Mỗi chế độ giữ riêng tệp đã chọn và kết quả xem trước; chuyển tab không làm
mất dữ liệu của tab kia. File mẫu mặc định tự đổi theo chế độ.

**Mật khẩu mặc định** (`Copenai@2026`) dùng chung cho cả 2 chế độ, có thể sửa
ở ô "Mật khẩu mặc định".

## Quy tắc xử lý (áp dụng cho CẢ 2 chế độ)

1. **Email**: viết thường, bỏ khoảng trắng/xuống dòng. Tự sửa **lỗi gõ phím rõ
   ràng**: dấu phẩy thay dấu chấm (`@gmail,com` → `@gmail.com`), 2 dấu chấm
   liền nhau, dấu chấm ngay trước/sau `@`. Email **thiếu hẳn** `@` hoặc thiếu
   tên miền (vd `nvgiang`) → **loại bỏ bản ghi**, không tự đoán.
2. **Số điện thoại**: chuẩn theo định dạng Việt Nam (10 số, bắt đầu bằng 0);
   tự xử lý `+84`/`84`, thiếu số 0 đầu. Ô có nhiều số → lấy số hợp lệ đầu
   tiên. Không chuẩn hoá được → để trống (vẫn giữ bản ghi, chỉ email mới bắt
   buộc).
3. **Họ và tên**: viết hoa chữ cái đầu mỗi từ, bỏ ký tự đặc biệt, cắt phần
   chức vụ/đơn vị bị dính vào tên.
4. **Trùng email**: mỗi email chỉ giữ 1 người (có thể tắt bằng ô tuỳ chọn).
   Khi trùng, giữ người có **tên khớp với email** hơn (vd `hvluyen@...` giữ
   "Hồ Văn Luyến"), vì nguồn hay dán nhầm email của người khác. Người bị loại
   hiện trong "Cần kiểm tra".

**Riêng chế độ 2** có thêm 2 quy tắc:
- **Đơn vị**: `<cấp hành chính> <tên riêng> - <tỉnh/thành>` (vd `UBND phường
  Ninh Kiều` → `Phường Ninh Kiều - Cần Thơ`). Không có cột đơn vị thì lấy từ
  tiêu đề văn bản (vd "Tên Cơ quan, đơn vị: ..." hoặc "Danh sách ... của ...").
- **Tổ chức**: tên chính thức hiện nay của tỉnh/thành cấp 1, viết hoa toàn
  bộ (vd `THÀNH PHỐ CẦN THƠ`). Tự quy tên tỉnh cũ về tỉnh mới theo bảng 34
  tỉnh/thành sau sáp nhập 2025. Ô "Tổ chức mặc định" dùng khi nguồn không nêu
  rõ; nếu ô này không phải tỉnh/thành (vd "Viện Năng lượng nguyên tử Việt
  Nam") thì dùng nguyên văn cho mọi dòng.

## Về độ chính xác

Tool **không bao giờ tự bịa dữ liệu thiếu**. Mọi dòng không đủ điều kiện
(thiếu/sai email, trùng email) đều bị loại khỏi kết quả và được liệt kê
**đầy đủ** trong tab **"Cần kiểm tra"** — kèm tệp nguồn, **STT gốc trong tệp**
(dòng không có STT hiện "sau STT N"), dữ liệu gốc và lý do — để bạn đối chiếu.

Tool còn tự phát hiện các dấu hiệu **mất dữ liệu âm thầm** và báo ở mục
"Thống kê" (⚠) cùng "Nhật ký xử lý":
- **STT bị nhảy số** (vd 22 → 24). Với PDF, tool tìm lại dòng bị thiếu trong
  văn bản gốc của trang và khôi phục (♻). Không tìm thấy thì chỉ cảnh báo để
  bạn mở tệp gốc kiểm tra (thường do nguồn đánh số sai).
- **Cả cột SĐT (hoặc Đơn vị) trống** trên mọi bản ghi — thường do tiêu đề cột
  lạ, tool không nhận ra cột.
- **Bảng có email nhưng không nhận ra tiêu đề cột** nên bị bỏ qua.

"0 dòng cần kiểm tra" chưa chắc đã đúng — hãy luôn đọc các cảnh báo này.

Tool cũng tự xử lý một số lỗi trình bày thường gặp trong PDF: bảng tách qua
nhiều trang, dòng tiêu đề nhóm (vd "Ủy ban nhân dân xã ...") không bị tính là
người, email bị cắt sang ô bên cạnh, cột bị lệch.

## Luồng thao tác chung (cả 2 chế độ)

1. Chọn chế độ → chọn tệp nguồn → (chế độ 2: nhập thêm "Tổ chức mặc định").
2. Bấm **"Trích xuất & Xem trước"** (hoặc `Enter`) — xem thống kê, cảnh báo,
   bảng "Xem trước" và tab "Cần kiểm tra". **Chưa ghi ra tệp** ở bước này.
3. Kiểm tra xong, bấm **"Xuất tệp kết quả..."** (hoặc `Ctrl+S`) — lúc này mới
   chọn nơi lưu và ghi tệp. Chế độ 2 tự ghi kèm tệp báo cáo
   `..._can_kiem_tra.xlsx` nếu có dòng bị loại.

**Phím tắt**: `Ctrl+O` chọn tệp · `Enter` trích xuất · `Ctrl+S` xuất tệp.

Có thể tìm kiếm (không phân biệt dấu) và bấm tiêu đề cột để sắp xếp trong cả
2 tab "Xem trước" và "Cần kiểm tra".

## Dòng lệnh (tuỳ chọn)

```bash
# Chế độ 1
python3 extract_contacts.py --input a.pdf b.xlsx --template cls_template_users.xlsx --output kq.xlsx

# Chế độ 2
python3 process_template_v2.py --input X.xlsx --template TemplateV2.xlsx \
    --output kq.xlsx --to-chuc "Cần Thơ" --password Copenai@2026
```

Chế độ 2 qua dòng lệnh để trống cột Mật khẩu nếu không truyền `--password`.
