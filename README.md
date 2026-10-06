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

Để đọc tệp Word cũ `.doc` (Word 97-2003), tool nhờ **Microsoft Word** trên máy
tự chuyển sang `.docx`; cần cài thêm:

```bash
pip install pywin32
```

Word chạy ẩn, không ảnh hưởng tới các cửa sổ Word bạn đang mở. Tệp `.xls` bị
hỏng mà thư viện xlrd không đọc được cũng tự chuyển qua Excel theo cách này.
Nếu máy không có Microsoft Office, tool dùng [LibreOffice](https://www.libreoffice.org/)
khi có cài; không có cả hai thì mở tệp bằng Word/Excel, chọn "Save As" sang
`.docx`/`.xlsx` rồi dùng tệp mới.

## Chạy

```bash
py gui.py
```

## Chọn chế độ

Chọn 1 trong 2 tab ở đầu trang:

| | Chế độ 1 | Chế độ 2 |
|---|---|---|
| **Dùng khi** | Chỉ cần Email / Họ tên / SĐT | Cần thêm Đơn vị và Tổ chức |
| **Đầu vào** | 1 tệp PDF, Word (.docx/.doc), Excel (.xlsx/.xls) hoặc CSV mỗi lần | Nhiều tệp PDF, Word (.docx/.doc), Excel (.xlsx/.xls), CSV cùng lúc |
| **File mẫu** | `cls_template_users.xlsx` | `TemplateV2.xlsx` |
| **Đầu ra** | Email, Mật khẩu, Họ và tên, Điện thoại | Tên tài khoản, Email, SĐT, Mật khẩu, Giới tính*, Ngày sinh*, Đơn vị, Tổ chức |

\* để trống — nguồn không có thông tin này nên tool không tự bịa.

Mỗi chế độ giữ riêng tệp đã chọn và kết quả xem trước; chuyển tab không làm
mất dữ liệu của tab kia. File mẫu mặc định tự đổi theo chế độ.

**Chế độ 2 xử lý nhiều tệp cùng lúc:**
- Mỗi tệp (hoặc mỗi sheet Excel, xem dưới) có ô **"Tổ chức mặc định"** riêng,
  vì các tệp trong một lần xử lý có thể thuộc các tỉnh khác nhau. Tệp mới thêm
  lấy giá trị đang có trong ô nhập. Muốn đổi thì chọn dòng, sửa ô rồi bấm
  **"Gán"**. Không chọn dòng nào thì "Gán" áp dụng cho tất cả.
- **Tệp Excel có nhiều sheet** (mỗi sheet có email) được **tách thành mỗi sheet
  một dòng**, để mỗi sheet có Tổ chức riêng (vd 1 tệp gộp 2 cơ quan ở 2 tỉnh).
  Sheet **trùng từ 50% email trở lên** với sheet trước (bản nháp, phụ lục) mặc
  định **"Bỏ qua"** (dòng tô xám); chọn dòng rồi bấm **"Dùng / Bỏ sheet"** (hoặc
  nhấp đúp) để đổi. Nhật ký ghi sheet nào đã bỏ và có bao nhiêu email mới bị bỏ
  theo. Chế độ 1 cũng tự bỏ sheet trùng và báo cảnh báo.
- Kết quả **gộp vào 1 tệp**. Tệp "Cần kiểm tra" có thêm cột **Tệp nguồn** và
  **Ghi chú**.
- **Lọc trùng email trên toàn bộ lô, kể cả khác tệp/sheet.** Người bị loại vào
  "Cần kiểm tra" kèm ghi chú trùng với **dòng nào, tệp nào** (vd "Trùng email
  với dòng 24 - Lừ Thị Chuyên (tệp SonLa.xlsx)"; dòng ghi theo số dòng trong
  Excel, kèm STT nếu có).
- Tệp nào đọc lỗi thì tool báo cảnh báo và vẫn xử lý tiếp các tệp còn lại.

**Mật khẩu mặc định** (`Copenai@2026`) dùng chung cho cả 2 chế độ, có thể sửa
ở ô "Mật khẩu mặc định".

## Quy tắc xử lý (áp dụng cho CẢ 2 chế độ)

1. **Email**: viết thường, bỏ khoảng trắng/xuống dòng. Tự sửa **lỗi gõ phím rõ
   ràng**: dấu phẩy thay dấu chấm (`@gmail,com` → `@gmail.com`), 2 dấu chấm
   liền nhau, dấu chấm ngay trước/sau `@`. Email **thiếu hẳn** `@` hoặc thiếu
   tên miền (vd `nvgiang`) → **loại bỏ bản ghi**, không tự đoán. Email có **dấu cách giữa
   phần tên** (vd `qthoang thongnhat@...`, có thể thiếu dấu chấm) hoặc **bị cắt giữa từ**
   do gõ chữ có dấu (`thươngpv@...`) cũng bị loại và hiện trong "Cần kiểm tra".
   Email bị cắt theo viền ô bảng PDF được tool tự khôi phục từ toạ độ chữ, kèm cảnh báo.
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
  Ninh Kiều` → `Phường Ninh Kiều - Cần Thơ`). Ô đơn vị chỉ ghi chức vụ ("Phó
  Giám đốc Sở") thì lấy cơ quan ban hành ở đầu văn bản (vd `Sở Nông nghiệp và
  Môi trường - Gia Lai`). Đơn vị cấp dưới xã/phường (phòng,
  ban, trung tâm, Đảng ủy, trạm y tế...) được gộp về chính xã/phường đó (vd
  `Phòng VHXH xã Tả Phìn` → `Xã Tả Phìn - Lào Cai`); "thị xã", "hợp tác xã",
  "Xã hội" không bị gộp. Không có cột đơn vị thì lấy từ
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
- **Sheet Excel bị bỏ qua** vì trùng với sheet trước, **họ tên bị ngắt thành 2
  dòng đã được nối lại** (PDF), hoặc **nhiều dòng bị ghép** vì cột Họ và tên
  trống.

"0 dòng cần kiểm tra" chưa chắc đã đúng — hãy luôn đọc các cảnh báo này.

Tool cũng tự xử lý một số lỗi trình bày thường gặp trong PDF: bảng tách qua
nhiều trang, dòng tiêu đề nhóm (vd "Ủy ban nhân dân xã ...") không bị tính là
người, email bị cắt sang ô bên cạnh, cột bị lệch.

## Luồng thao tác chung (cả 2 chế độ)

1. Chọn chế độ → chọn tệp nguồn (chế độ 2: có thể chọn nhiều tệp, mỗi tệp một
   "Tổ chức mặc định").
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
py extract_contacts.py --input a.pdf b.xlsx --template cls_template_users.xlsx --output kq.xlsx

# Chế độ 2 (1 hoặc nhiều tệp, gộp vào 1 tệp kết quả)
py process_template_v2.py --input X.xlsx Y.pdf --template TemplateV2.xlsx \
    --output kq.xlsx --to-chuc "Cần Thơ" --password Copenai@2026
```

Chế độ 2 qua dòng lệnh để trống cột Mật khẩu nếu không truyền `--password`.
Khi truyền nhiều tệp qua dòng lệnh, mọi tệp dùng chung `--to-chuc`. Muốn mỗi
tệp một Tổ chức riêng thì dùng giao diện. Dòng lệnh chế độ 1 vẫn nhận nhiều tệp
để tương thích với cách gọi cũ.
