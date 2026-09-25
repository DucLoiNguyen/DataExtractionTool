# CLAUDE.md — Công cụ tạo danh sách tài khoản (Nền tảng AI công vụ)

Tài liệu này dành cho Claude (Claude Code / Claude local) đọc trước khi làm việc với dự án. Người dùng giao tiếp bằng **tiếng Việt**; luôn trả lời bằng tiếng Việt có dấu.

## 1. Dự án làm gì

Các cơ quan nhà nước (xã/phường, sở, viện, tỉnh/thành) gửi danh sách cán bộ đăng ký tài khoản "Nền tảng Trí tuệ nhân tạo hỗ trợ công vụ" dưới rất nhiều dạng: công văn PDF, Word (.docx/.doc), Excel (.xlsx/.xls), CSV. Mỗi đơn vị trình bày một kiểu, dữ liệu thường bẩn (email gõ sai, SĐT thừa số, trùng email, cột lệch, bảng bị cắt trang...).

Công cụ trích xuất **Họ tên / Email / SĐT** (và **Đơn vị / Tổ chức** ở chế độ 2), chuẩn hoá theo quy tắc cố định, rồi xuất ra file Excel theo 1 trong 2 file mẫu để nhập vào hệ thống tạo tài khoản.

**Yêu cầu cao nhất của người dùng: dữ liệu xuất ra phải chính xác.** Không bao giờ tự bịa dữ liệu. Dòng nào không xử lý chắc chắn được thì đưa vào tab/tệp **"Cần kiểm tra"** kèm lý do, không âm thầm bỏ qua.

## 2. Cấu trúc thư mục

```
gui.py                   # GUI Tkinter duy nhất — người dùng chạy file này
extract_contacts.py      # Logic cốt lõi + Chế độ 1 (dùng chung cho cả 2 chế độ)
process_template_v2.py   # Logic Chế độ 2 (import extract_contacts as core)
cls_template_users.xlsx  # File mẫu Chế độ 1
TemplateV2.xlsx          # File mẫu Chế độ 2
README.md                # Hướng dẫn cho người dùng cuối
regression.py            # Bảng hồi quy mục 7 + kiểm tra đơn vị (python regression.py --detail)
test_data/               # File gốc của người dùng, tên ngắn (TaPhin.pdf...) — KHÔNG commit (.gitignore)
```

Cả 5 file đầu phải nằm **cùng một thư mục**. `gui.py` tự tìm file mẫu trong thư mục của nó (`THIS_DIR`).

Môi trường máy người dùng (Windows): `git` không có trong PATH — dùng `C:\Users\Loind\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe`. Chưa cài LibreOffice (ca `.doc` không chạy được). File gốc nằm rải rác trong `~/Downloads` và `~/OneDrive - C-OPENAI/Premier Service Team - Tạo tài khoản GOV_260916/`.

Phụ thuộc: `pip install openpyxl python-docx pdfplumber xlrd`. Tuỳ chọn: **LibreOffice** (`soffice`) để đọc `.doc` (Word 97-2003); nếu thiếu, tool báo lỗi rõ ràng và hướng dẫn "Save As .docx".

Chạy:
```bash
python3 gui.py                                   # giao diện
python3 process_template_v2.py --input X.xlsx --template TemplateV2.xlsx \
    --output kq.xlsx --to-chuc "Cần Thơ" [--password Copenai@2026] [--no-dedupe]
python3 extract_contacts.py ...                  # CLI chế độ 1 (xem main())
```

## 3. Hai chế độ

| | Chế độ 1 | Chế độ 2 |
|---|---|---|
| Hàm chính | `core.extract_all(inputs, dedupe)` → `core.write_output(...)` | `v2.extract_v2(path, default_to_chuc, dedupe)` → `v2.write_v2_output(...)`, `v2.write_v2_issues(...)` |
| Đầu vào | Nhiều tệp PDF/Word/Excel/CSV | 1 tệp PDF/Word(.docx/.doc)/Excel(.xlsx/.xls)/CSV |
| File mẫu | `cls_template_users.xlsx`: Email, Mật khẩu, Họ và tên, Điện thoại | `TemplateV2.xlsx`: Tên tài khoản, Email, Số điện thoại, Mật khẩu, Giới tính, Ngày sinh, Đơn vị, Tổ chức |
| Cột để trống | — | Giới tính, Ngày sinh (không có trong nguồn → không bịa) |

Chế độ 2 là **phần mở rộng** của chế độ 1: áp dụng đủ quy tắc chuẩn hoá của chế độ 1 (gọi thẳng `core.normalize_email/phone/name`, `core.email_name_match_score`) rồi thêm Đơn vị/Tổ chức. **Sửa quy tắc email/SĐT/tên ở `extract_contacts.py` là tự động áp dụng cho cả 2 chế độ.**

GUI (`gui.py`): 2 tab chế độ ở đầu trang; toàn bộ trang cuộn được như một khối (header, 2 cột, nút Xuất tệp, Nhật ký đều nằm trong 1 Canvas). Mỗi chế độ có `ResultPanel` riêng với tab **"Xem trước (N)"** và **"Cần kiểm tra (N)"** (tìm kiếm không dấu, sắp xếp theo cột). Chỉ ghi file khi bấm "Xuất tệp kết quả..." (dùng chung `shared_export_button`). Chế độ 2 xuất kèm `<tên>_can_kiem_tra.xlsx` nếu có dòng bị loại. Phím tắt: Ctrl+O, Enter, Ctrl+S. Tông màu ấm (accent `#c15f3c`, nền `#f6f3ee`). Mọi chuỗi hiển thị là tiếng Việt có dấu.

## 4. Quy tắc nghiệp vụ (đã được người dùng xác nhận)

**Email** (`core.normalize_email`)
- Viết thường, bỏ mọi khoảng trắng/xuống dòng (PDF hay ngắt dòng giữa email).
- Tự sửa lỗi gõ phím rõ ràng (`_repair_email_typos`): `,` → `.`; `..` → `.`; bỏ `.` ngay trước/sau `@`. Ví dụ `ten..x@gmail,com` → `ten.x@gmail.com`.
- Sau khi sửa vẫn có "nhãn rỗng" (`_has_empty_email_label`) hoặc không có `@` → **loại** (lý do "Email sai định dạng"). Không bao giờ tự đoán/chèn thêm nội dung bị mất (ví dụ thiếu hẳn `@` hay thiếu tên miền).
- Không có email → loại (lý do "Thiếu email"). Email là trường bắt buộc duy nhất.
- Có thể có 2 cột email (vd "Thư điện tử công vụ" + "Thư điện tử Gmail"): dùng cột 2 (`email_alt`) khi cột 1 trống.

**Số điện thoại** (`core.normalize_phone`): 10 số, bắt đầu bằng `0`; xử lý `+84`/`84`, float từ Excel (`912345678.0`, kể cả khi đã bị ép thành chuỗi `"912345678.0"`), thiếu số 0 đầu (9 số). Ô có nhiều số → lấy số hợp lệ đầu tiên. Không chuẩn hoá được (vd 11 chữ số) → **để trống**, vẫn giữ bản ghi. **Luôn truyền giá trị thô (giữ kiểu float/int) vào hàm này, không ép `str()` trước** (từng làm mất SĐT).

**Họ tên** (`core.normalize_name`): viết hoa chữ cái đầu mỗi từ, bỏ ký tự đặc biệt, cắt phần chức vụ/đơn vị bị dính vào tên.

**Trùng email** (bật mặc định): mỗi email chỉ giữ 1 người. Khi trùng, giữ người có tên **khớp email** hơn (`core.email_name_match_score`: `hvluyen` ↔ "Hồ Văn Luyến" được 2 điểm); ngang điểm thì giữ người xuất hiện trước. Người bị loại vào "Cần kiểm tra" (lý do "Trùng email"). Lý do: nhiều file nguồn dán nhầm email của người khác vào dòng phía trên.

**Mật khẩu**: mặc định `Copenai@2026` (`core.DEFAULT_PASSWORD`) cho **cả 2 chế độ**; GUI có 1 ô nhập dùng chung. `write_v2_output(password=None)` để trống (giữ tương thích CLI cũ).

**Đơn vị** (Chế độ 2, `v2.normalize_don_vi`)
- Dạng `<cấp hành chính> <tên riêng> - <tỉnh/thành>`: `UBND phường Ninh Kiều` → `Phường Ninh Kiều - Cần Thơ`; `Sở Tư pháp` → `Sở Tư pháp - Cần Thơ`. Bỏ tiền tố UBND/HĐND trước Xã/Phường/Thị trấn.
- Mở rộng viết tắt: `TT` → Trung tâm, `TP.` → Thành phố, `CĐCĐ`, `QLDA`...
- **Luôn thêm hậu tố tỉnh/thành**, kể cả khi tên đã chứa tên tỉnh (ra `... Thành phố Cần Thơ - Cần Thơ`). Người dùng chưa yêu cầu bỏ phần lặp này; đừng tự đổi.
- Nếu giá trị chính là một tỉnh/thành → chỉ giữ tên riêng, không thêm hậu tố.
- Nếu "Tổ chức mặc định" **không phải** tỉnh/thành (vd "Viện Năng lượng nguyên tử Việt Nam") → không thêm hậu tố.
- Cột đơn vị ghi **chức vụ** thay vì tên đơn vị (vd Bến Cát: "Chủ tịch UBND phường") → **giữ nguyên dữ liệu gốc** (người dùng đã chọn như vậy), không suy từ đuôi email.

**Tổ chức** (Chế độ 2, `v2.determine_to_chuc`): tên chính thức **hiện nay** của tỉnh/thành cấp 1, VIẾT HOA (vd `THÀNH PHỐ CẦN THƠ`, `TỈNH LÀO CAI`). Dùng `PROVINCE_MERGE_MAP` (34 tỉnh/thành sau sáp nhập 2025, Nghị quyết 202/2025/QH15, hiệu lực 12/6/2025) để quy tên cũ về tên mới (Hậu Giang/Sóc Trăng → Cần Thơ; Kon Tum → Quảng Ngãi; Bắc Kạn → Thái Nguyên; Quảng Nam → Đà Nẵng...). Nguồn không nêu → dùng "Tổ chức mặc định" người dùng nhập. Nếu giá trị mặc định **không phải** tỉnh/thành → dùng nguyên văn viết hoa và **không** dò tên tỉnh trong văn bản (tránh "Trung tâm Chiếu xạ Hà Nội" bị nhận thành Hà Nội). Người dùng từng gọi nhầm "TỈNH CẦN THƠ" — tên đúng là "THÀNH PHỐ CẦN THƠ".

## 5. Kiến trúc đọc dữ liệu — các bẫy đã gặp

Đây là phần quan trọng nhất. Mỗi mục là một lỗi thật đã phát hiện qua file người dùng gửi.

**Nhận diện cột theo từ khoá** (`core.HEADER_KEYWORDS`, `v2.SOURCE_HEADER_KEYWORDS`; so khớp trên chuỗi đã bỏ dấu, viết thường). Từ khoá **email/SĐT** của Chế độ 2 được sinh tự động từ `core.HEADER_KEYWORDS` → chỉ cần thêm ở `extract_contacts.py`. Từ khoá **tên** giữ riêng: Chế độ 2 không được dùng từ chung `"ten"` (sẽ nhận nhầm "Tên đơn vị"). Các tiêu đề thực tế đã gặp: "mail công vụ", "Địa chỉ thư công vụ", "Thư điện tử công vụ", "Di động", "Số điện thoại sử dụng Zalo", "TT" (thay "STT"), "Tên người dùng", "Phòng ban/Bộ phận". Gặp file mới ra 0 bản ghi hoặc cột trống → kiểm tra tiêu đề cột trước tiên.

**Cột đơn vị** (`v2._find_source_columns`): quét theo **thứ tự từ khoá** (cụ thể trước: "đơn vị công tác" → "đơn vị" → "phòng ban"...), không theo thứ tự cột (Sơn La có "Tên đơn vị" gần như trống nằm bên trái cột "Cơ quan, đơn vị công tác" đầy đủ). "Chức vụ" chỉ là dự phòng. Hai cột liền kề **cùng tiêu đề** (tiêu đề gộp ô) → ghép giá trị qua `unit2` (Long Hòa: "Phòng VHXH" + "Xã Long Hòa").

**Dòng tiêu đề**: không mặc định là dòng 1; `_find_source_header_row` quét 15 dòng đầu, chọn dòng nhận diện được nhiều cột nhất.

**Đơn vị dự phòng khi không có cột đơn vị**: PDF/Word lấy từ "Tên Cơ quan, đơn vị: X" (`_extract_doc_header_unit`); Excel lấy từ tiêu đề phía trên bảng, phần sau chữ "của" (`_extract_title_unit_from_rows`, vd "Danh sách ... của Văn phòng UBND thành phố Đà Nẵng").

**Tiêu đề nhóm số La Mã** (I, II, III...): dòng nhóm (không email/SĐT, chỉ 1 ô tên có nội dung) → dùng tên nhóm làm đơn vị **chỉ khi** cột đơn vị trên dòng không giống tên đơn vị thật (`_looks_like_real_unit_name`: có khoảng trắng). VINATOM có cột đơn vị chứa username; Nghĩa Đô có cột đơn vị tốt → giữ.

**Excel**
- `.xls` đọc bằng xlrd, đã vá `handle_datemode` (`_patch_xlrd_tolerant_datemode`) cho file có VBA lỗi.
- Excel có nhiều sheet: Chế độ 2 chỉ đọc sheet đầu (`sheet_by_index(0)` / `wb.active`).
- **Không** gọi hàm ghép dòng bị ngắt (`merge_wrapped_continuation_rows`) cho Excel/CSV ở **cả 2 chế độ** (`extract_from_table_rows(..., merge_wrapped=False)`): dòng thiếu tên trong Excel là lỗi dữ liệu thật, ghép nhầm sẽ làm mất 1 người (Cần Thơ: từng mất 1 người ở cả 2 chế độ).

**PDF/Word trong Chế độ 2 — đọc theo NỘI DUNG, không theo vị trí cột**
- Các trang của cùng 1 PDF có thể khác bố cục cột (Nghĩa Đô: trang 2 email ở cột 4, trang 3 lệch 9 cột, trang 4-5 email ở cột 6). Vì vậy `_process_data_rows_content_based` + `_parse_row_by_content` nhận email/SĐT bằng regex ở bất kỳ ô nào; tên/đơn vị là 2 ô văn bản còn lại theo thứ tự.
- Dòng "mảnh vỡ" (chỉ có SĐT hoặc email, không tên) được ghép vào bản ghi ngay trước.
- `_flatten_doc_tables`: nối mọi bảng thành 1 danh sách; bắt đầu từ bảng **đầu tiên có tiêu đề hợp lệ** (file Word hay có bảng quốc hiệu/tiêu ngữ đứng trước); bỏ dòng **tiêu đề lặp lại** ở đầu mỗi trang.
- Không dùng `core.merge_multi_page_tables` cho Chế độ 2 (nó cắt danh sách khi số cột đổi).
- `.doc` chuyển sang `.docx` qua `core._convert_via_libreoffice(path, "docx")`, dọn thư mục tạm trong `finally`.

**STT bị nhảy số — cả 2 chế độ** (hàm dùng chung ở mục "2b. STT" của `extract_contacts.py`: `parse_stt`, `find_stt_gaps`, `find_stt_lines`, `split_recovered_stt_line`)
- **pdfplumber có thể bỏ sót hẳn 1 dòng** khi tách bảng (Tả Phìn mất STT 23, 40, 55, 79). PDF: dò STT nhảy số, tìm lại dòng trong văn bản thô (`page.extract_text()`), tách tên/đơn vị bằng `UNIT_KEYWORD_SPLIT_RE` rồi cắt chức vụ bằng `POSITION_SPLIT_RE`. Chế độ 1: `core._recover_missing_stt_rows`; Chế độ 2: `v2._recover_missing_stt_rows`.
- **Không tìm thấy dòng trong văn bản thô → chỉ ghi cảnh báo**, không tạo dòng nào (không bịa, cũng không tạo issue "Thiếu email" rỗng). Đồng Tháp: 724, 743, 915, 920, 2315 là **nguồn đánh số sai** (2315 gõ thành 2015), không phải mất dòng.
- STT "đã thấy" phải lấy từ **mọi dòng** của bảng, kể cả dòng sẽ bị ghép vào dòng trên và dòng tiêu đề nhóm. Đồng Tháp đánh STT cả cho dòng bị xuống hàng (910 chỉ chứa đuôi email `dongthap.gov.vn` của 909). Nếu chỉ lấy STT từ bản ghi, bước khôi phục sẽ tạo **người ảo**. Sơn La: STT 89 là dòng "Công an tỉnh — chưa có danh sách đăng ký".
- Excel/CSV/Word: không có văn bản thô để khôi phục → chỉ cảnh báo.

**Cảnh báo** (`stats["warnings"]`, `stats["recovered_stts"]` ở cả 2 chế độ; in qua `core.print_warnings`; GUI hiện dòng ♻/⚠ dưới pipeline Thống kê và tô màu trong Nhật ký): STT nhảy số không tìm lại được; cả tệp có ≥3 bản ghi nhưng không có SĐT nào (hoặc Đơn vị ở Chế độ 2); bảng có email nhưng không nhận ra tiêu đề (trước đây bị bỏ qua im lặng); tệp đọc lỗi.

**Chế độ 1** (`extract_contacts.py`): đọc bảng (`extract_from_table_rows`, gộp bảng nhiều trang bằng `merge_multi_page_tables`) + văn bản tự do (`extract_from_free_text`). Word dùng `_full_cell_text`/`_full_paragraph_text` để không sót email nằm trong hyperlink. Mỗi dòng bị loại được ghi vào `stats["issues"]` với `source`, `row_number`, `raw_*`, `reason`.
- `row_number` = **STT gốc** nếu bảng có cột STT (`is_stt_header`: STT/TT/Số TT); dòng không có STT hiện `"sau STT N"`. Trước đây là số tự đếm → lệch với tệp gốc (Hồ Quang Lớn STT 742 hiện thành 748).
- **Dòng tiêu đề nhóm** (`_is_group_header_row`): chỉ áp dụng khi bảng có cột STT; STT trống/La Mã, không email/SĐT, và (các ô khác ≤3 ký tự **hoặc** ô tên bắt đầu bằng loại cơ quan — `_GROUP_UNIT_START_RE`: Ủy ban, Sở, Trung tâm...) → bỏ qua, không tính là người. Đồng Tháp có 50 dòng như vậy, thường bị cắt sang 2 ô ("Ủy ban nhân dân Phường Lo" | "ng Thuận"). **Không** thêm "Đoàn" vào danh sách (là họ người: "Đoàn Văn Nỉ").
- **Email/SĐT theo nội dung khi lệch cột** (`_find_email_in_row`, `_find_phone_in_row`): cột Email không hợp lệ → tìm email hợp lệ ở ô khác, hoặc ghép 2 ô liền nhau khi ô sau là đuôi ngắn toàn ký tự email (`hoquanglon@gmail.c` + `om`). SĐT chỉ tìm sang ô khác khi có dấu hiệu lệch cột (tránh lấy nhầm số ở cột Ghi chú).
- Khi ghép dòng bị ngắt, **không nối ô STT** (trước ra "909910").

**Mã lý do** (`core.ISSUE_REASON_LABELS`): `missing_email` → "Thiếu email", `invalid_email_format` → "Email sai định dạng", `duplicate_email` → "Trùng email".

## 6. Cách làm việc bắt buộc

1. **Mỗi file mới người dùng gửi là một ca kiểm thử** cho cả 2 chế độ. Kiểm tra cấu trúc thô (sheet, dòng đầu, tiêu đề, số dòng thật), chạy thử, rồi **đối chiếu với file gốc**: số dòng đọc được có khớp số người thật không, STT có nhảy số không, cột nào trống bất thường (0 SĐT, đơn vị trống...). "0 lỗi" không có nghĩa là đúng — đã nhiều lần dữ liệu mất âm thầm.
2. Số dòng bị loại cao bất thường → xem vài dòng trong file gốc để phân biệt **lỗi dữ liệu nguồn** (báo người dùng, không sửa) và **lỗi tool** (vá).
3. Sau mọi thay đổi: `python3 -m py_compile` cả 3 file, chạy **bảng hồi quy ở mục 7**, và kiểm tra qua GUI thật khi đụng tới giao diện hoặc đường đi dữ liệu.
4. Quy tắc mơ hồ hoặc có nhiều cách hợp lý (vd cột đơn vị ghi chức vụ) → hỏi người dùng, đưa 2-3 lựa chọn cụ thể.
5. Khi báo cáo: nêu lỗi tìm được, cách sửa, số liệu trước/sau, và các lỗi dữ liệu nguồn cần người dùng xử lý (kèm STT, tên, giá trị cụ thể). Nếu bản vá làm thay đổi kết quả của file đã xử lý trước đó, nói rõ file nào và vì sao.
6. Quy ước code: chú thích trong code viết tiếng Việt **không dấu**, giải thích **lý do** (thường kèm ví dụ thực tế đã gặp); chuỗi hiển thị cho người dùng viết **có dấu**. Giữ tương thích ngược cho CLI và các hàm công khai.

## 7. Bảng hồi quy (số liệu hiện tại, đã xác minh)

Chạy: `python regression.py --detail` (đọc `test_data/`, file thiếu thì báo "BỎ QUA"; kèm kiểm tra đơn vị không cần file). Tên ngắn trong `test_data/` ở cột "test_data". Cột "Đọc" = số dòng đọc được, "KQ" = số tài khoản xuất ra, "CKT" = số dòng Cần kiểm tra.

Hiện chưa tìm thấy trên máy: `3_Khu_Di_tích_Phủ_chủ_tich.docx`, `CV_163_SXD_THAI_NGUYEN_...AI.xls`.

**Chế độ 1** (`core.extract_all([path], dedupe=True)`; Đọc = `stats["total_rows"]`)

| File | test_data | Đọc | KQ |
|---|---|---|---|
| CV_2096_UBNDX_...Nền_tảng_AI_công_vụ.pdf (Sơn Linh, Quảng Ngãi) | SonLinh.pdf | 25 | 25 |
| 3_Khu_Di_tích_Phủ_chủ_tich.docx | KhuDiTich.docx | 43 | 42 (số cũ, chưa chạy lại) |
| Danh_sa_üch__Éo_é_Çng_Tha_üp.pdf (Đồng Tháp) | DongThap.pdf | 2381 | 2241 |
| ĐT02_Đăng_ký_danh_sách_người_dùng_CLEX.pdf (Kon Đào) | KonDao.pdf | 10 | 10 |
| Tả_Phìn.pdf | TaPhin.pdf | 80 | 79 |
| Văn_phòng_UBND_thành_phố_Đà_Nẵng.xlsx | DaNang.xlsx | 82 | 82 |
| Xa_Long_Hoa_..._Vr_25_tháng_9_.xlsx | LongHoa.xlsx | 96 | 95 |
| DangKy_2109.xlsx (Cần Thơ) | CanTho.xlsx | 5000 | 4973 |

Thay đổi so với số cũ (25/09/2026): Tả Phìn 76/75 → 80/79 (khôi phục STT). Đồng Tháp 2431/2240 → 2381/2241 (bỏ 50 dòng tiêu đề nhóm; giữ được Hồ Quang Lớn — email bị cắt 2 ô, lệch cột). Cần Thơ chế độ 1 trước ra 4999 (ghép dòng Excel làm mất 1 người).

**Chế độ 2** (`v2.extract_v2(path, default_to_chuc=..., dedupe=True)`; Đọc = `stats["total_read"]`)

| File | test_data | Tổ chức mặc định | Đọc | KQ | CKT |
|---|---|---|---|---|---|
| DangKy_2109.xlsx | CanTho.xlsx | Cần Thơ | 5000 | 4973 | 27 |
| TK_AI_VNLNTVN.XLS | VINATOM.xls | Viện Năng lượng nguyên tử Việt Nam | 539 | 532 | 7 |
| CV_163_SXD_THAI_NGUYEN_...AI.xls | ThaiNguyen.xls | Thái Nguyên | 155 | 153 | 2 |
| ĐT02_Đăng_ký_danh_sách_người_dùng_CLEX.pdf | KonDao.pdf | Quảng Ngãi | 10 | 10 | 0 |
| Tả_Phìn.pdf | TaPhin.pdf | Lào Cai | 80 | 79 | 1 |
| Nghĩa_Đô.pdf | NghiaDo.pdf | Lào Cai | 70 | 69 | 1 |
| P_Bến_Cát.xlsx | BenCat.xlsx | Hồ Chí Minh | 23 | 22 | 1 |
| P_Bình_Hòa.docx | BinhHoa.docx | Hồ Chí Minh | 38 | 0 | 38 |
| P_Chánh_Hưng.docx | ChanhHung.docx | Hồ Chí Minh | 2 | 2 | 0 |
| P_Long_Hòa.doc (cần LibreOffice) | P_LongHoa.doc | Hồ Chí Minh | 97 | 0 | 97 |
| 2__TH_danh_sách_đăng_ký_..._kem_CV_.xlsx (Sơn La) | SonLa.xlsx | Sơn La | 5980 | 5089 | 891 |
| Văn_phòng_UBND_thành_phố_Đà_Nẵng.xlsx | DaNang.xlsx | Đà Nẵng | 82 | 82 | 0 |
| Xa_Long_Hoa_..._Vr_25_tháng_9_.xlsx | LongHoa.xlsx | Hồ Chí Minh | 96 | 95 | 1 |

Cảnh báo mong đợi: Đồng Tháp chế độ 1 "STT nhảy số 724, 743, 915, 920, 2315"; Sơn La chế độ 2 "STT nhảy số 85" (Văn phòng UBND tỉnh không được đánh số trong nguồn, dữ liệu vẫn đọc đủ). Các file khác không có cảnh báo.

Kiểm tra định tính (không chỉ đếm):
- Bình Hòa và Long Hòa (.doc) **không có email trong nguồn** → 0 là đúng.
- Trùng email phải giữ đúng chủ email: Bến Cát giữ "Lê Thị Hồng Hà" (`lthha`), Nghĩa Đô giữ "Trần Thế Anh" (`trantheanh2`), Long Hòa giữ "Hồ Văn Luyến" (`hvluyen`).
- Tả Phìn cả 2 chế độ phải có STT 23, 40, 55, 79 (khôi phục từ văn bản thô); chế độ 2 STT 40 có đơn vị "Ban kinh tế xã hội HĐND xã Tả Phìn".
- Đồng Tháp chế độ 1: có `hoquanglon@gmail.com` / SĐT `0919099499`.
- VINATOM: mọi dòng Tổ chức = "VIỆN NĂNG LƯỢNG NGUYÊN TỬ VIỆT NAM" (kể cả "Trung tâm Chiếu xạ Hà Nội").
- Đà Nẵng: 82/82 có SĐT (tiêu đề "Di động"), đơn vị "Văn phòng UBND thành phố Đà Nẵng - Đà Nẵng".

Khi thêm file mới vào bảng: chép vào `test_data/` với tên ngắn không dấu, thêm dòng vào `MODE1`/`MODE2` trong `regression.py` và vào bảng ở trên.

GUI kiểm tra được bằng script (dựng `gui.ContactExtractorGUI()`, gán `input_paths`/`mode2_input_path`, gọi `start_extract()`, đọc `active_panel.stats_text` và `log_text`). Máy người dùng có thể không cho chụp màn hình (`ImageGrab` báo "screen grab failed") → đọc nội dung widget thay thế.

## 8. Việc còn mở / ý tưởng đã nêu nhưng chưa làm

- Tự bỏ hậu tố tỉnh/thành khi tên đơn vị đã chứa tên đó (vd "...thành phố Đà Nẵng - Đà Nẵng"): đã đề xuất, người dùng chưa quyết.
- Kon Đào: đơn vị lấy từ cột "Phòng ban/Bộ phận" ("Phòng Kinh tế xã") thay vì tên đầy đủ ở đầu phiếu ("Phòng Kinh tế xã Kon Đào"); đã hỏi người dùng có muốn ưu tiên tên đầy đủ không, chưa có trả lời.
- Chế độ 1 vẫn đọc theo **vị trí cột** (chỉ tìm email/SĐT theo nội dung khi cột Email không hợp lệ). Chế độ 2 đọc PDF/Word hoàn toàn theo nội dung. Chế độ 1 còn bỏ qua bảng khi số cột đổi giữa các trang mà không có tiêu đề (`merge_multi_page_tables` coi là bảng mới) — hiện đã có cảnh báo nếu bảng đó chứa email.
- Bình Hòa chế độ 1 ra 0/0 không cảnh báo (nguồn không có cột email nên không nhận ra bảng, cũng không có "@").
- `PROVINCE_MERGE_MAP` là thông tin hành chính có thể thay đổi; cần kiểm tra lại nếu dùng cho dữ liệu ở thời điểm khác.
