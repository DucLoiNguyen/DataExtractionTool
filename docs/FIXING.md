# Yêu cầu sửa lỗi: TTYT Bác Ái + Tam Long (Excel) và Sở NN&MT Gia Lai (PDF)

Tài liệu này do Claude (bản web) viết sau khi phân tích 2 tệp đầu vào mới. Các lỗi được tìm và **cách sửa được kiểm chứng trên một bản code CŨ hơn** code trong repo (bản trước ngày 30/09/2026). Vì vậy:

1. Chép 2 tệp vào `test_data/` với tên ngắn: `TTYT_TamLong.xlsx` (gốc: `TTYT_KV_BAC_AI.xlsx`) và `GiaLai_SNNMT.pdf` (gốc: `11190_SNNMT-VP_30092026-signed_01.pdf`).
2. **Chạy thử trên code hiện tại trước**, đối chiếu với "Kết quả đúng" bên dưới. Lỗi nào code hiện tại đã xử lý đúng thì bỏ qua mục đó.
3. Sửa từng mục, chạy `python regression.py --detail` sau mỗi mục. Tên hàm/biến dưới đây theo bản cũ, có thể đã đổi.

## Kết quả đúng (đã kiểm chứng thủ công với tệp gốc)

| Tệp | Chế độ 2: Đọc / KQ / CKT | Đơn vị | Tổ chức |
|---|---|---|---|
| TTYT_TamLong.xlsx, sheet 1 (tên sheet "UBND PHƯỜNG TAM LONG" nhưng nội dung là TTYT Bác Ái) | 43 / 43 / 0 | `Trung Tâm Y Tế Khu Vực Bác Ái - Khánh Hòa` | `TỈNH KHÁNH HÒA` (Bác Ái thuộc Ninh Thuận cũ) |
| TTYT_TamLong.xlsx, sheet 2 ("ĐẢNG ỦY PHƯỜNG TAM LONG") | 36 / 36 / 0 | `Phường Tam Long - Hồ Chí Minh` | `THÀNH PHỐ HỒ CHÍ MINH` (email `@tphcm.gov.vn`) |
| GiaLai_SNNMT.pdf | 9 / 9 / 0 | `Sở Nông Nghiệp và Môi Trường - Gia Lai` | `TỈNH GIA LAI` |

Họ tên PDF phải đầy đủ: `Cao Thanh Thương`, `Nguyễn Thị Tố Trân`, `Trần Đình Chương`, `Hà Thị Thanh Hương`, `Nguyễn Văn Hoan`, `Nguyễn Thị Thế Vy`, `Trần Quốc Khánh`, `Đoàn Ngọc Có`, `Vũ Ngọc An`. Email `@gmaill.com` (sheet 1) phải được sửa về `@gmail.com` (quy tắc Gmail đã có). SĐT `'0913763278` (có dấu nháy đầu) phải ra `0913763278`.

## Lỗi 1 (nghiêm trọng): Chế độ 2 chỉ đọc 1 sheet Excel

**Triệu chứng:** Chế độ 2 ra 36/79 người, mất toàn bộ 43 người sheet 1 mà không có cảnh báo nào.

**Nguyên nhân:** `_read_tables_and_context` dùng `src_wb.active` (sheet đang mở khi lưu tệp, ở đây là sheet 2), và với `.xls` dùng `sheet_by_index(0)`. Chế độ 1 đã đọc mọi sheet.

**Cách sửa:** trả về 1 bảng cho **mỗi sheet** (`src_wb.worksheets`, `wb.sheets()` với xlrd). Nhưng **không được đọc mù quáng mọi sheet**: bỏ qua sheet có **≥ 50% email đã xuất hiện ở các sheet trước** (bản nháp, phụ lục), và ghi cảnh báo nêu tên sheet và số email mới bị bỏ. Đã kiểm chứng:
- Thái Nguyên `CV_163...xls` sheet "Trang_tính1" là bản nháp: 62/94 email trùng → bỏ (nếu không bỏ, CKT 2 → 97).
- Sơn La sheet 2 "Phụ lục II – DANH SÁCH ĐẦU MỐI TRIỂN KHAI" (người liên hệ, không phải danh sách đăng ký): 68/84 trùng → bỏ (nếu không bỏ, KQ 5053 → 5069, CKT 928 → 1047).
- TTYT_TamLong: 2 sheet không trùng nhau → đọc cả hai.

Với cách này, hồi quy cũ giữ nguyên 26/26. Cảnh báo cho Sơn La và Thái Nguyên là **cảnh báo mong đợi mới**, cần thêm vào `CLAUDE.md`. Chế độ 1 cũng nên áp dụng cùng quy tắc bỏ sheet trùng (chưa kiểm chứng ở Chế độ 1).

## Lỗi 2: Ghép ô tiêu đề Excel bằng dấu cách làm hỏng tên phường

**Triệu chứng:** Đơn vị sheet 2 ra `Phường Tam Long Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam - ...`.

**Nguyên nhân:** khi lấy văn bản các dòng phía trên bảng để tìm xã/phường ban hành (`_find_doc_commune`), các ô **cùng 1 dòng** được nối bằng dấu cách. Ô A1 `ỦY BAN NHÂN DÂN\nPHƯỜNG TAM LONG` và ô D1 `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` bị nối thành `PHƯỜNG TAM LONG CỘNG HÒA XÃ HỘI...`. Vì chữ viết hoa toàn bộ, tên phường bị kéo dài. Sau đó quy tắc "căn tên xã theo xã ban hành" (thêm cho Chiềng Ken) lại thay `Phường Tam Long` đúng trong ô bằng tên dài sai này.

**Cách sửa:** nối các ô bằng **xuống dòng**. Mỗi sheet phải có xã/phường ban hành **riêng** (tính lại cho từng bảng, không dùng chung cho cả tệp). Nên giới hạn thêm quy tắc căn tên xã: chỉ thay khi tên xã ban hành dài hơn tối đa 2 từ.

## Lỗi 3: 1 tệp chứa 2 tỉnh/thành, nhưng Tổ chức mặc định áp cho cả tệp

**Triệu chứng:** 36 người Phường Tam Long (email `@tphcm.gov.vn`) bị gán `TỈNH KHÁNH HÒA`.

**Người dùng đã chọn: tách tệp Excel nhiều sheet thành nhiều dòng trong danh sách tệp của Chế độ 2, mỗi sheet có Tổ chức mặc định riêng.** (Phương án chỉ cảnh báo theo tên miền email đã bị loại.)

### Phần xử lý (đã kiểm chứng trên bản cũ)

- Thêm `list_excel_sheets(path)` → danh sách sheet **có ít nhất 1 email**, theo thứ tự, mỗi phần tử gồm `{"sheet", "emails", "overlap_with_previous", "suggest_skip"}`. `suggest_skip = emails >= 3 and overlap >= 50%` (cùng quy tắc Lỗi 1). Tệp không phải Excel → `[]`. Kết quả mong đợi:
  - `TTYT_TamLong.xlsx` → 2 sheet: `UBND PHƯỜNG TAM LONG` (43), `ĐẢNG ỦY PHƯỜNG TAM LONG` (36), cả hai `suggest_skip=False`.
  - Thái Nguyên `.xls` → `Sheet1` (200), `Trang_tính1` (94, trùng 62, `suggest_skip=True`).
  - Sơn La → sheet chính (5091), `DANH SÁCH ĐẦU MỐI TRIỂN KHAI...` (84, trùng 68, `suggest_skip=True`).
  - VINATOM (`TK_AI_VNLNTVN.XLS`) → chỉ 1 sheet `"Người dùng "`. Các sheet danh mục khác không có email nên không liệt kê. **Giữ nguyên tên sheet, kể cả dấu cách thừa ở cuối.**
- `extract_v2(path, ..., sheet=None)`: có `sheet` → chỉ đọc đúng sheet đó (`wb[sheet]` / `book.sheet_by_name(sheet)`), không áp quy tắc bỏ sheet trùng. `sheet=None` → hành vi Lỗi 1 (đọc mọi sheet, bỏ sheet trùng), giữ cho CLI và tương thích ngược.
- `extract_v2_batch`: mỗi nguồn là `(đường dẫn, sheet hoặc None, Tổ chức mặc định)`. Cột "Tệp nguồn" ghi `tệp.xlsx [Sheet: tên]` khi có sheet. Lọc trùng email vẫn tính trên toàn bộ lô.

### Phần giao diện (Chế độ 2, `m2_tree`)

- Khi thêm 1 tệp Excel: gọi `list_excel_sheets`. **≥ 2 sheet** → chèn **mỗi sheet 1 dòng**, gồm Tệp, Sheet, Số email, Tổ chức mặc định, Trạng thái. **Chỉ 1 sheet** → 1 dòng như hiện nay (vẫn lưu tên sheet). Tệp PDF/Word giữ như cũ.
- Sheet `suggest_skip=True` vẫn hiện, nhưng Trạng thái = `Bỏ qua (trùng 62/94 email với sheet trước)` và **mặc định không được trích xuất**. Thêm nút (hoặc nhấp đúp) **"Dùng / Bỏ sheet"** để đổi trạng thái. Sheet có trạng thái "Bỏ qua" không được đưa vào lô.
- Tổ chức mặc định của các dòng sheet lấy theo giá trị đang nhập cho tệp. Người dùng sửa **từng dòng** bằng cơ chế sửa Tổ chức theo tệp đã có.
- Nhật ký ghi rõ sheet nào đã bỏ và số email mới bị bỏ theo, để người dùng chủ động bật lại nếu cần (vd 16 người mới trong phụ lục đầu mối Sơn La).

### Kiểm thử

- TTYT_TamLong: 2 dòng; đặt Khánh Hòa / Hồ Chí Minh → **79** tài khoản, đúng như bảng "Kết quả đúng" ở trên.
- Thái Nguyên, Sơn La: dòng sheet 2 mặc định "Bỏ qua" → kết quả **giữ nguyên như hồi quy** (153 / 5053). Bật sheet 2 lên → số dòng tăng, có trùng email, đúng như mong đợi.
- `regression.py`: thêm ca TTYT_TamLong dạng 2 nguồn theo sheet (mỗi nguồn 1 Tổ chức) và ca kiểm tra `list_excel_sheets` cho 3 tệp trên.
- Chế độ 1 không có Tổ chức nên giữ nguyên cách đọc mọi sheet. Nên thêm quy tắc bỏ sheet trùng (Lỗi 1) cho Chế độ 1.

## Lỗi 4: PDF Gia Lai, họ tên bị ngắt 2 dòng (Chế độ 2, đọc theo nội dung)

**Triệu chứng:** tên ra `Cao Thanh`, `Trần Quốc`... Có thêm 1 "người" ảo trong "Cần kiểm tra" (tên `Khánh`, đơn vị `Vy`, SĐT trùng STT 7).

**Nguyên nhân:** pdfplumber tách họ tên thành 2 dòng vật lý: `['1',...,'Cao Thanh',...,email,SĐT]` rồi `[...,'Thương',...]`. Dòng thứ 2 (1 ô văn bản, không email/SĐT) bị `_process_data_rows_content_based` hiểu là **dòng tiêu đề nhóm** nên thành đơn vị của người kế tiếp. Dòng `[...,'Khánh',...,'0914099910']` có SĐT trùng người trước nên bị coi là người mới.

**Cách sửa (đã kiểm chứng):**
- `_parse_row_by_content` trả thêm **vị trí cột** của ô tên (`name_idx`) và ô đơn vị.
- Trước khi kiểm tra tiêu đề nhóm: nếu dòng **không có email**, chỉ có **1 ô văn bản ngắn (≤ 3 từ)** nằm **đúng cột tên** của bản ghi vừa thêm, không bắt đầu bằng từ chỉ đơn vị, và các ô còn lại (nếu có) chỉ là **SĐT trùng SĐT bản ghi trước** → nối vào họ tên bản ghi trước (`normalize_name(ten + " " + phan_tiep)`), không tạo bản ghi mới.
- Tiêu đề nhóm chỉ nhận khi có dấu hiệu rõ: ô đầu là số La Mã, hoặc chữ viết hoa toàn bộ, hoặc bắt đầu bằng từ chỉ đơn vị. Nghĩa Đô ("I ĐẢNG ỦY") vẫn đúng.

## Lỗi 5: PDF Gia Lai, STT lệch cột làm tên bị hiểu sai

**Triệu chứng:** STT 8, 9 ra tên rỗng, đơn vị = `Đoàn Ngọc Có`.

**Nguyên nhân:** 2 dòng cuối bảng bị lệch 1 cột: `['', '8', '', '', 'Đoàn Ngọc Có', ...]`. Code lấy STT = `cells[0]` (rỗng), nên "8" thành ô văn bản đầu tiên = "họ tên".

**Cách sửa (đã kiểm chứng):** STT = **ô có nội dung đầu tiên** của dòng nếu nó là số (có thể có dấu chấm) hoặc số La Mã. Chỉ ô đó được đánh dấu đã dùng.

## Lỗi 6: Ô đơn vị chỉ ghi chức vụ, tên cơ quan nằm ở đầu văn bản

**Triệu chứng:** Đơn vị ra `Phó Giám đốc Sở - Gia Lai`.

**Cách sửa (đã kiểm chứng):** thêm `_find_doc_issuer(header_text)`, tìm **cơ quan ban hành** không phải xã/phường ở 8 dòng đầu văn bản. Lấy cụm từ **viết hoa toàn bộ ở đầu dòng**, bắt đầu bằng Sở / Ban / Trung tâm / Văn phòng / Bệnh viện / Viện / Chi cục / Cục / Trường / Bảo tàng... (PDF gộp 2 cột: `SỞ NÔNG NGHIỆP VÀ MÔI TRƯỜNG Độc lập - Tự do...` → chỉ lấy phần viết hoa), đổi sang `Sở Nông Nghiệp và Môi Trường`. Bỏ qua dòng `UBND TỈNH ...` (cơ quan cấp trên).

Trong `normalize_don_vi`: **chỉ khi** ô là chức vụ **thuần túy** (`_looks_like_job_title` đúng và `_strip_job_title` trả về rỗng), có cơ quan ban hành, và văn bản không phải của xã/phường → dùng cơ quan ban hành.

**Lưu ý:** lần thử đầu tôi cho bỏ cả chức vụ ở ô "chức vụ + đơn vị". Việc này đổi đơn vị ở VINATOM ("Giám đốc, Trung tâm Chiếu xạ Hà Nội") và Bến Cát ("Chủ tịch UBND phường"; người dùng đã chọn giữ nguyên). Phải giữ đúng điều kiện "chức vụ thuần túy". Hồi quy sau khi giới hạn: không đổi.

## Lỗi 7: Chế độ 1 với PDF Gia Lai ra 1 bản ghi rác (CHƯA kiểm chứng trên code hiện tại)

Bản cũ ra 1 bản ghi `thuongct@snnmt.gialai.gov.vntranntt`, tên rỗng. **Nguyên nhân** (theo bản cũ): bảng có 15 cột do ô gộp; tiêu đề "Họ và tên" ở cột 3 nhưng dữ liệu tên ở cột 4, email tiêu đề cột 9 nhưng 2 dòng cuối ở cột 10. Đọc theo vị trí → cột tên luôn rỗng → mọi dòng bị coi là "dòng bị ngắt" và ghép vào dòng đầu.

**Hướng sửa:** coi mỗi ô tiêu đề **bao phủ các cột tới ô tiêu đề kế tiếp** (cột 3 → 3..5); khi đọc 1 trường, lấy ô có nội dung đầu tiên trong vùng đó. Đồng thời áp dụng cách nối dòng tên bị ngắt như Lỗi 4. Code hiện tại đã có `_find_email_in_row`/`_find_phone_in_row`, nên có thể chỉ cần thêm phần vùng cột cho tên.

## Lỗi nhỏ cần để ý

- `unit_cache` đang dùng khóa là chuỗi đơn vị thô. Khi mỗi sheet có xã/cơ quan ban hành riêng, khóa phải gồm cả ngữ cảnh (`(raw_unit, doc_commune, doc_issuer)`), nếu không 2 sheet có cùng chuỗi đơn vị sẽ dùng nhầm kết quả của nhau.
- Chế độ 1, sheet 1: dòng nhóm `I | VIÊN CHỨC` từng ra "Thiếu email" (bản cũ). Theo `CLAUDE.md`, code hiện tại đã bỏ qua dòng nhóm, cần xác nhận lại.

## Vấn đề dữ liệu nguồn (báo người dùng, không sửa)

- Tên sheet không khớp nội dung (sheet "UBND PHƯỜNG TAM LONG" chứa TTYT Bác Ái). Một tệp gộp 2 đơn vị ở 2 tỉnh khác nhau, có thể gửi nhầm/gộp nhầm.
- `Nguyễn T.trang Vương` (TTYT Bác Ái): họ tên viết tắt, thứ tự lạ. Tool giữ nguyên.

## Sau khi sửa

Thêm 2 tệp vào `MODE1`/`MODE2` trong `regression.py` và bảng mục 7 của `CLAUDE.md`, ghi các cảnh báo mong đợi mới (bỏ sheet ở Sơn La và Thái Nguyên), ghi quyết định "tách sheet thành nhiều dòng, mỗi sheet 1 Tổ chức" vào mục quy tắc của `CLAUDE.md`, rồi commit.
