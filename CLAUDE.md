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
regression.py            # Bảng hồi quy mục 7 + kiểm tra đơn vị: python regression.py --detail
test_data/               # File gốc của người dùng, tên ngắn (TaPhin.pdf...) — KHÔNG commit (.gitignore)
```

Môi trường máy người dùng (Windows): `git` không có trong PATH — dùng `C:\Users\Loind\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe`. File gốc nằm rải rác trong `~/Downloads` và `~/OneDrive - C-OPENAI/Premier Service Team - Tạo tài khoản GOV_260916/`. **Đã từng mất code do nhiều phiên làm việc ghi đè lên nhau** (phiên sau bắt đầu từ bản cũ): trước khi sửa, xem `git status`/`git log`; sau mỗi thay đổi đã kiểm thử, nhắc người dùng commit.

Cả 5 file đầu phải nằm **cùng một thư mục**. `gui.py` tự tìm file mẫu trong thư mục của nó (`THIS_DIR`).

Phụ thuộc: `pip install openpyxl python-docx pdfplumber xlrd`. Trên Windows thêm `pip install pywin32`.

**Chuyển đổi Office cũ** (`core.convert_office_file(path, "docx"|"xlsx")`): `.doc` luôn cần chuyển sang `.docx` (python-docx không đọc `.doc`); `.xls` chỉ chuyển khi `xlrd` thất bại. Thứ tự thử: (1) **Microsoft Word/Excel qua pywin32** (`_convert_via_ms_office`; chỉ Windows; `DispatchEx` mở phiên Word/Excel **riêng, chạy ẩn**; `CoInitialize` vì GUI chạy trong luồng nền; mở `ReadOnly` với mật khẩu giả để tệp có mật khẩu báo lỗi thay vì hiện hộp thoại làm treo; luôn `Close` + `Quit` trong `finally`; Word `FileFormat=16`, Excel `51`); (2) **LibreOffice** (`_convert_via_libreoffice`); (3) lỗi `ValueError` nêu lý do từng cách + hướng dẫn. **Người dùng không cài được LibreOffice**, nên nhánh pywin32 là nhánh chính trên máy họ. Đã chạy với Word thật trên máy người dùng (Office 16, pywin32 312): `P_LongHoa.doc`, `BaoTang.doc`, khoảng 4 giây/tệp, không để lại tiến trình Word/Excel hay thư mục tạm. Nếu gặp lỗi trên Windows, kiểm tra trước: đã cài pywin32 chưa, Word có mở tệp được bằng tay không, tệp có bị Windows chặn (Properties → Unblock) không.

Chạy:
```bash
python3 gui.py                                   # giao diện
python3 process_template_v2.py --input X.xlsx [Y.pdf ...] --template TemplateV2.xlsx \
    --output kq.xlsx --to-chuc "Cần Thơ" [--password Copenai@2026] [--no-dedupe]
    # nhieu tep: gop 1 ket qua, dung chung --to-chuc
python3 extract_contacts.py ...                  # CLI chế độ 1 (xem main())
```

## 3. Hai chế độ

| | Chế độ 1 | Chế độ 2 |
|---|---|---|
| Hàm chính | `core.extract_all(inputs, dedupe)` → `core.write_output(...)` | `v2.extract_v2(path, default_to_chuc, dedupe)` → `v2.write_v2_output(...)`, `v2.write_v2_issues(...)` |
| Đầu vào | GUI: **1 tệp mỗi lần** (hàm/CLI vẫn nhận danh sách, giữ tương thích) | GUI: **nhiều tệp**, mỗi tệp 1 Tổ chức mặc định → `v2.extract_v2_batch([(path, to_chuc), ...])` |
| File mẫu | `cls_template_users.xlsx`: Email, Mật khẩu, Họ và tên, Điện thoại | `TemplateV2.xlsx`: Tên tài khoản, Email, Số điện thoại, Mật khẩu, Giới tính, Ngày sinh, Đơn vị, Tổ chức |
| Cột để trống | — | Giới tính, Ngày sinh (không có trong nguồn → không bịa) |

Chế độ 2 là **phần mở rộng** của chế độ 1: áp dụng đủ quy tắc chuẩn hoá của chế độ 1 (gọi thẳng `core.normalize_email/phone/name`, `core.email_name_match_score`) rồi thêm Đơn vị/Tổ chức. **Sửa quy tắc email/SĐT/tên ở `extract_contacts.py` là tự động áp dụng cho cả 2 chế độ.**

GUI (`gui.py`): 2 tab chế độ ở đầu trang; toàn bộ trang cuộn được như một khối (header, 2 cột, nút Xuất tệp, Nhật ký đều nằm trong 1 Canvas). Mỗi chế độ có `ResultPanel` riêng với tab **"Xem trước (N)"** và **"Cần kiểm tra (N)"** (tìm kiếm không dấu, sắp xếp theo cột). Chỉ ghi file khi bấm "Xuất tệp kết quả..." (dùng chung `shared_export_button`). Chế độ 2 xuất kèm `<tên>_can_kiem_tra.xlsx` nếu có dòng bị loại.

**Chế độ 2 nhiều tệp** (người dùng chọn ngày 30/09/2026): danh sách tệp là Treeview `m2_tree` (cột Tệp + Tổ chức mặc định), dữ liệu ở `mode2_paths` + `mode2_to_chuc`. Tệp mới lấy giá trị ô `to_chuc_var`, nút "Gán" áp cho tệp đang chọn (không chọn = tất cả). `extract_v2_batch` gọi `extract_v2` cho từng tệp, **lọc trùng email chỉ trong từng tệp** (kết quả mỗi tệp phải giống hệt chạy riêng — có ca hồi quy), email lặp giữa các tệp chỉ cảnh báo; **gộp vào 1 tệp kết quả**; mỗi bản ghi/issue có khoá `source`; cảnh báo có tiền tố tên tệp; tệp đọc lỗi → cảnh báo và chạy tiếp (tất cả lỗi mới báo lỗi); `stats["files"]` thống kê từng tệp. Tab Xem trước/Cần kiểm tra và `write_v2_issues` có thêm cột "Tệp nguồn". Chế độ 1 trong GUI chỉ 1 tệp (`mode1_input_path`). Bố cục: cột trái rộng cố định 380px (`pack_propagate(False)`) nên chiều cao phải đặt theo nội dung (`_sync_left_height`), nếu không phần dưới (Mật khẩu, nút Trích xuất) bị che khi thẻ bên trái cao hơn cột phải; bảng dùng `_attach_scrollbars` (grid), không `pack side=left`, vì bảng rộng hơn khung sẽ đẩy thanh cuộn ra ngoài. Phím tắt: Ctrl+O, Enter, Ctrl+S. Tông màu ấm (accent `#c15f3c`, nền `#f6f3ee`). Mọi chuỗi hiển thị là tiếng Việt có dấu.

## 4. Quy tắc nghiệp vụ (đã được người dùng xác nhận)

**Email** (`core.normalize_email`)
- Viết thường, bỏ mọi khoảng trắng/xuống dòng (PDF hay ngắt dòng giữa email).
- Tự sửa lỗi gõ phím rõ ràng (`_repair_email_typos`): `,` → `.`; `..` → `.`; bỏ `.` ngay trước/sau `@`. Ví dụ `ten..x@gmail,com` → `ten.x@gmail.com`.
- `@@` → `@`. Chế độ 2 đọc PDF/Word theo nội dung: ô chứa `@` (không có khoảng trắng) được coi là ô email dù chưa đúng chuẩn, để bước sửa lỗi gõ xử lý (trước đây `@laocai,gov.vn`, `@@...` bị bỏ qua thành "thiếu email").
- **Tên miền bị cắt cụt** (`find_truncated_domain_emails`, chạy trên toàn file ở cả 2 chế độ): đuôi bị cắt (`laocai.go`, `sonla.gov`, `cantho.edu`) hoặc chỉ còn hậu tố (`@gov.vn`, `@edu.vn`) so với tên miền phổ biến trong file (≥ 3 lần, nhiều gấp ≥ 5) → loại, lý do "Email sai định dạng". Không tự điền phần thiếu. Chỉ xét 2 kiểu này để không loại nhầm tên miền hợp lệ như `daklak.gov.vn` khi file có nhiều `vubon.daklak.gov.vn`. Lưu ý: sửa `@.gov.vn` → `@gov.vn` (Sơn La, Lóng Phiêng) vẫn bị loại ở bước này vì mất tên đơn vị.
- Sửa tên miền Gmail gõ sai (`_fix_gmail_domain`, **luôn chạy**, kể cả khi email đúng cú pháp): `@gmailcom`, `@gmai.com`, `@gmial.com`, `@gamil.com`, `@gmail.con`, `@gmail.co`... → `@gmail.com`. Chỉ sửa Gmail, không đoán tên miền cơ quan. (Từng có 30 email ở Sơn La, 16 ở Đồng Tháp "hợp lệ" nhưng không nhận được thư.)
- Sau khi sửa vẫn có "nhãn rỗng" (`_has_empty_email_label`) hoặc không có `@` → **loại** (lý do "Email sai định dạng"). Không bao giờ tự đoán/chèn thêm nội dung bị mất (ví dụ thiếu hẳn `@` hay thiếu tên miền).
- Không có email → loại (lý do "Thiếu email"). Email là trường bắt buộc duy nhất.
- Có thể có 2 cột email (vd "Thư điện tử công vụ" + "Thư điện tử Gmail"): dùng cột 2 (`email_alt`) khi cột 1 trống.

**Số điện thoại** (`core.normalize_phone`): 10 số, bắt đầu bằng `0`; xử lý `+84`/`84`, float từ Excel (`912345678.0`), thiếu số 0 đầu (9 số). Ô có nhiều số → lấy số hợp lệ đầu tiên. Không chuẩn hoá được (vd 11 chữ số) → **để trống**, vẫn giữ bản ghi. **Luôn truyền giá trị thô (giữ kiểu float/int) vào hàm này, không ép `str()` trước** (từng làm mất SĐT).

**Họ tên** (`core.normalize_name`): viết hoa chữ cái đầu mỗi từ, bỏ ký tự đặc biệt, cắt phần chức vụ/đơn vị bị dính vào tên. Bỏ kính ngữ "Ông/Bà" ở đầu (`strip_honorific_by_email`) **chỉ khi** phần còn lại ≥ 3 chữ **và** khớp email hơn ("Bà Lý Thị Mụi" / `lythimui`). "Ông" cũng là họ thật ("Ông Quang Đông" / `dongoq`, "Ông Mai Xuân" / `ongmaixuan`) → không bỏ mù quáng.

**Trùng email** (bật mặc định): mỗi email chỉ giữ 1 người. Khi trùng, giữ người có tên **khớp email** hơn (`core.email_name_match_score`: `hvluyen` ↔ "Hồ Văn Luyến" được 2 điểm); ngang điểm thì giữ người xuất hiện trước. Người bị loại vào "Cần kiểm tra" (lý do "Trùng email"). Lý do: nhiều file nguồn dán nhầm email của người khác vào dòng phía trên.

**Mật khẩu**: mặc định `Copenai@2026` (`core.DEFAULT_PASSWORD`) cho **cả 2 chế độ**; GUI có 1 ô nhập dùng chung. `write_v2_output(password=None)` để trống (giữ tương thích CLI cũ).

**Đơn vị** (Chế độ 2, `v2.normalize_don_vi`)
- **Đơn vị cấp dưới xã/phường gộp thành chính xã/phường đó** (yêu cầu của người dùng): phòng ban, trường học, trạm y tế, đoàn thể, thôn/tổ dân phố... → `Xã/Phường/Thị trấn <Tên> - <tỉnh>`. Ví dụ `Phòng VH- XH Phường Nam Định` → `Phường Nam Định - Ninh Bình`; `Văn phòng Đảng ủy xã Tả Phìn` → `Xã Tả Phìn - Lào Cai`. Tên xã lấy từ chính ô đơn vị (`_find_commune_in_text`: lấy lần xuất hiện cuối có tên, nhưng bỏ qua từ khoá nằm **trong** tên vừa tìm — `Phường Xã Đàn` giữ nguyên; bỏ "xã hội", "xã viên"; bỏ khi từ đứng trước là `thị` (thị xã = cấp huyện cũ), `tác` (hợp tác xã), `các`/`cấp`; "XÃ/PHƯỜNG" viết hoa thì tên cũng phải viết hoa; **cả ô viết thường** (`phòng vh xã nam định`) thì nhận tên viết thường, tối đa 5 từ, dừng ở huyện/quận/tỉnh/thành phố; dừng ở `-`, `,`, `/`, `(`). Ô không ghi tên xã ("Văn phòng Đảng ủy", "Phòng kinh tế", "Hội nông dân Việt Nam xã") → dùng xã/phường của **cơ quan ban hành** ở đầu văn bản (`_find_doc_commune`: trang đầu PDF, bảng quốc hiệu + đoạn mở đầu Word, các dòng trên bảng Excel), trừ khi ô chứa dấu hiệu cấp cao hơn (`_ABOVE_COMMUNE_RE`: tỉnh, thành phố, Sở, Bộ, Cục) hoặc đơn vị không thuộc xã (`_NOT_COMMUNE_UNIT_RE`: huyện, quận, thị xã, bệnh viện, ngân hàng, chi nhánh, đại học, cao đẳng, kho bạc, BHXH, chi cục, điện lực, bưu điện, công ty...) — vd văn bản xã Nghĩa Đô có "Bệnh viện Đa khoa huyện Bảo Yên" thì giữ nguyên, không đoán là xã. Không có tên xã ở đâu cả (vd Bến Cát, chỉ có trong tên file/email) → giữ nguyên, không suy đoán.
- Đơn vị cấp xã trở lên (Sở, Viện, Ban QLDA...): dạng `<cấp hành chính> <tên riêng> - <tỉnh/thành>`: `UBND phường Ninh Kiều` → `Phường Ninh Kiều - Cần Thơ`; `Sở Tư pháp` → `Sở Tư pháp - Cần Thơ`. Bỏ tiền tố UBND/HĐND trước Xã/Phường/Thị trấn.
- Mở rộng viết tắt: `TT` → Trung tâm, `TP.` → Thành phố, `CĐCĐ`, `QLDA`...
- **Luôn thêm hậu tố tỉnh/thành**, kể cả khi tên đã chứa tên tỉnh (ra `... Thành phố Cần Thơ - Cần Thơ`). Người dùng chưa yêu cầu bỏ phần lặp này; đừng tự đổi.
- Nếu giá trị chính là một tỉnh/thành → chỉ giữ tên riêng, không thêm hậu tố.
- Nếu "Tổ chức mặc định" **không phải** tỉnh/thành (vd "Viện Năng lượng nguyên tử Việt Nam") → không thêm hậu tố.
- Cột đơn vị ghi **chức vụ** thay vì tên đơn vị (vd Bến Cát: "Chủ tịch UBND phường") → **giữ nguyên dữ liệu gốc** (người dùng đã chọn như vậy), không suy từ đuôi email.

**Tổ chức** (Chế độ 2, `v2.determine_to_chuc`): tên chính thức **hiện nay** của tỉnh/thành cấp 1, VIẾT HOA (vd `THÀNH PHỐ CẦN THƠ`, `TỈNH LÀO CAI`). Dùng `PROVINCE_MERGE_MAP` (34 tỉnh/thành sau sáp nhập 2025, Nghị quyết 202/2025/QH15, hiệu lực 12/6/2025) để quy tên cũ về tên mới (Hậu Giang/Sóc Trăng → Cần Thơ; Kon Tum → Quảng Ngãi; Bắc Kạn → Thái Nguyên; Quảng Nam → Đà Nẵng...). Nguồn không nêu → dùng "Tổ chức mặc định" người dùng nhập. Nếu giá trị mặc định **không phải** tỉnh/thành → dùng nguyên văn viết hoa và **không** dò tên tỉnh trong văn bản (tránh "Trung tâm Chiếu xạ Hà Nội" bị nhận thành Hà Nội). Nếu giá trị mặc định **là** tỉnh/thành → chỉ đổi sang tỉnh khác khi văn bản ghi rõ tiền tố "tỉnh X" / "thành phố X" / "TP. X" (`detect_to_chuc_from_text(..., require_prefix=True)`); chỉ nhắc tên thì giữ mặc định, vì tên tỉnh hay trùng tên xã/trường/tổ chức ("Xã Bình Thuận" ở Sơn La, "Đoàn TNCS Hồ Chí Minh"). So khớp tên tỉnh **giữ nguyên dấu** khi văn bản có dấu ("Hải Đường" ≠ "Hải Dương"), có tính cả 2 kiểu đặt dấu ("Hòa"/"Hoà"). Người dùng từng gọi nhầm "TỈNH CẦN THƠ" — tên đúng là "THÀNH PHỐ CẦN THƠ".

## 5. Kiến trúc đọc dữ liệu — các bẫy đã gặp

Đây là phần quan trọng nhất. Mỗi mục là một lỗi thật đã phát hiện qua file người dùng gửi.

**Nhận diện cột theo từ khoá** (`core.HEADER_KEYWORDS`, `v2.SOURCE_HEADER_KEYWORDS`; so khớp trên chuỗi đã bỏ dấu, viết thường). Từ khoá **email/SĐT** của Chế độ 2 được sinh tự động từ `core.HEADER_KEYWORDS` → chỉ thêm ở `extract_contacts.py`. Từ khoá **tên** giữ riêng: Chế độ 2 không được dùng từ chung `"ten"` (sẽ nhận nhầm "Tên đơn vị"). Các tiêu đề thực tế đã gặp: "mail công vụ", "Địa chỉ thư công vụ", "Thư điện tử công vụ", "Di động", "Số điện thoại sử dụng Zalo", "TT" (thay "STT"), "Tên người dùng", "Phòng ban/Bộ phận". Gặp file mới ra 0 bản ghi hoặc cột trống → kiểm tra tiêu đề cột trước tiên.

**Cột đơn vị** (`v2._find_source_columns`): quét theo **thứ tự từ khoá** (cụ thể trước: "đơn vị công tác" → "đơn vị" → "phòng ban"...), không theo thứ tự cột (Sơn La có "Tên đơn vị" gần như trống nằm bên trái cột "Cơ quan, đơn vị công tác" đầy đủ). "Chức vụ" chỉ là dự phòng. Hai cột liền kề **cùng tiêu đề** (tiêu đề gộp ô) → ghép giá trị qua `unit2` (Long Hòa: "Phòng VHXH" + "Xã Long Hòa").

**Dòng tiêu đề**: không mặc định là dòng 1; `_find_source_header_row` quét 15 dòng đầu, chọn dòng nhận diện được nhiều cột nhất.

**Bảng không có dòng tiêu đề** (vd `Support_excel.xlsx`: dữ liệu bắt đầu ngay dòng 1): `core.infer_columns_from_content` suy ra cột từ nội dung (email = cột nhiều `@`; SĐT = cột chuẩn hoá được thành SĐT; STT = cột số nhỏ; họ tên = cột văn bản dày, 2-6 chữ, ít trùng lặp; đơn vị = cột dày còn lại; `group` = cột thưa < 30%, thường là tên xã/phường chỉ ghi ở dòng đầu nhóm do gộp ô). Dùng cho cả 2 chế độ. Chế độ 2 điền tiếp `group` xuống dưới và ghép "phòng ban + xã/phường" (bỏ tiền tố UBND/HĐND; không ghép nếu phòng ban đã chứa tên đó). Khi có cột `group` thì tắt cơ chế "dòng tiêu đề nhóm La Mã" (nó sẽ nhầm người thiếu cả email lẫn SĐT thành tiêu đề nhóm).

**Cột khối thưa trong bảng CÓ tiêu đề** (`_detect_sparse_group_col`): cột chưa được gán (vd "GHI CHÚ" ở Đắk Lắk, "Tên đơn vị" ở Sơn La), < 30% dòng có giá trị, ≥ 80% giá trị giống tên đơn vị → dùng làm `group`, điền tiếp xuống dưới. Cách dùng tên khối: khối là **xã/phường** → luôn gộp vào xã/phường đó (ô từng dòng có thể sao chép nhầm, vd Sơn La ghi "UBND xã Tạ Khoa" cho 32 người thuộc khối Tà Xùa); khối khác → chỉ dùng khi ô đơn vị **trống hoặc chỉ ghi chức vụ** (`_looks_like_job_title`, so khớp có dấu để "Trưởng" ≠ "Trường"). Ô "chức vụ + đơn vị" → bỏ chức vụ (`_strip_job_title`: "Giám đốc Trung tâm TGPL Nhà nước" → "Trung tâm TGPL Nhà nước"). Ô đã là tên đơn vị thật → giữ nguyên, **không** ghép tên khối (tránh "Bệnh viện ... Sở Y tế").

**STT dạng số thực lệch** (Excel cộng dồn, vd `150.00000000000003`, Đắk Lắk `1.0000000000000848`): `v2._cell_display` và `core.parse_stt` làm tròn khi sai số < 1e-6 (thiếu ở `parse_stt` → cả cột STT bị coi là trống, báo "nhảy số" giả 193 STT).

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
- Nhận diện ô theo nội dung phải chịu được định dạng lệch: email có khoảng trắng sát `@` (`hamanhcuong3 @laocai.gov.vn`, Yên Thành: 69 người từng bị loại oan; xem `_squeeze_at`), SĐT có dấu chấm/gạch (`0399.517.520`), STT dạng `1.` (bỏ dấu chấm, nếu không cơ chế dò STT nhảy số sẽ không chạy).
- **Số trang** in ở đầu trang sau có thể bị một số công cụ đọc PDF nối vào SĐT dòng cuối trang trước (vd `0978850300` + trang "3" thành `09788503003`). pdfplumber tách bảng vẫn đúng 10 số. Đừng "sửa" SĐT dựa trên các chuỗi dài bất thường trong văn bản thô mà chưa kiểm tra vị trí ký tự (`page.extract_words()` có `x0/top`).

**STT bị nhảy số — cả 2 chế độ** (hàm dùng chung ở mục "2b. STT" của `extract_contacts.py`: `parse_stt`, `find_stt_gaps`, `find_stt_lines`, `split_recovered_stt_line`)
- **pdfplumber có thể bỏ sót hẳn 1 dòng** khi tách bảng (Tả Phìn mất STT 23, 40, 55, 79). PDF: dò STT nhảy số, tìm lại dòng trong văn bản thô (`page.extract_text()`), tách tên/đơn vị. Chế độ 1: `core._recover_missing_stt_rows`; Chế độ 2: `v2._recover_missing_stt_rows`.
- **Không tìm thấy dòng trong văn bản thô → chỉ ghi cảnh báo**, không tạo dòng nào (không bịa, không tạo issue "Thiếu email" rỗng). Đồng Tháp: 724, 743, 915, 920, 2315 là **nguồn đánh số sai** (2315 gõ thành 2015).
- STT "đã thấy" lấy từ **mọi dòng** của bảng, kể cả dòng sẽ bị ghép vào dòng trên và dòng tiêu đề nhóm. Đồng Tháp đánh STT cả cho dòng bị xuống hàng (910 chỉ chứa đuôi email `dongthap.gov.vn` của 909); chỉ lấy STT từ bản ghi sẽ tạo **người ảo**. Sơn La: STT 89 là dòng "Công an tỉnh — chưa có danh sách đăng ký".
- Excel/CSV/Word: không có văn bản thô → chỉ cảnh báo.

**Cảnh báo** (`stats["warnings"]`, `stats["recovered_stts"]` ở cả 2 chế độ; in qua `core.print_warnings`; GUI hiện dòng ♻/⚠ dưới pipeline Thống kê, tô màu trong Nhật ký): STT nhảy số không tìm lại được; tệp ≥3 bản ghi mà không có SĐT nào (Chế độ 2: hoặc không có Đơn vị nào); bảng có email nhưng không nhận ra tiêu đề (trước đây bị bỏ qua im lặng); tệp đọc lỗi.

- `.doc` chuyển sang `.docx` qua `core._convert_via_libreoffice(path, "docx")`, dọn thư mục tạm trong `finally`.

**Chế độ 1** (`extract_contacts.py`): đọc bảng (`extract_from_table_rows`, gộp bảng nhiều trang bằng `merge_multi_page_tables`) + văn bản tự do (`extract_from_free_text`). Word dùng `_full_cell_text`/`_full_paragraph_text` để không sót email nằm trong hyperlink. Mỗi dòng bị loại được ghi vào `stats["issues"]` với `source`, `row_number`, `raw_*`, `reason`.
- `row_number` = **STT gốc** nếu bảng có cột STT (`is_stt_header`: STT/TT/Số TT); dòng không có STT hiện `"sau STT N"`. Số tự đếm sẽ lệch với tệp gốc (Hồ Quang Lớn STT 742 từng hiện thành 748).
- **Dòng tiêu đề nhóm** (`_is_group_header_row`, chỉ khi bảng có cột STT và có dòng tiêu đề): STT trống/La Mã, không email/SĐT, và (các ô khác ≤3 ký tự **hoặc** ô tên bắt đầu bằng loại cơ quan — `_GROUP_UNIT_START_RE`) → bỏ qua. Đồng Tháp có 50 dòng như vậy, thường bị cắt sang 2 ô ("Ủy ban nhân dân Phường Lo" | "ng Thuận"). **Không** thêm "Đoàn" (là họ người: "Đoàn Văn Nỉ").
- **Email/SĐT theo nội dung khi lệch cột** (`_find_email_in_row`, `_find_phone_in_row`): cột Email không hợp lệ → tìm email hợp lệ ở ô khác, hoặc ghép 2 ô liền nhau khi ô sau là đuôi ngắn toàn ký tự email (`hoquanglon@gmail.c` + `om`). SĐT chỉ tìm sang ô khác khi có dấu hiệu lệch cột.
- Khi ghép dòng bị ngắt, **không nối ô STT** (trước ra "909910").
- Bảng không tiêu đề (`infer_columns_from_content`) không ghép dòng, không lọc tiêu đề nhóm (cột `group` đảm nhiệm).

**Mã lý do** (`core.ISSUE_REASON_LABELS`): `missing_email` → "Thiếu email", `invalid_email_format` → "Email sai định dạng", `duplicate_email` → "Trùng email".

## 6. Cách làm việc bắt buộc

1. **Mỗi file mới người dùng gửi là một ca kiểm thử** cho cả 2 chế độ. Kiểm tra cấu trúc thô (sheet, dòng đầu, tiêu đề, số dòng thật), chạy thử, rồi **đối chiếu với file gốc**: số dòng đọc được có khớp số người thật không, STT có nhảy số không, cột nào trống bất thường (0 SĐT, đơn vị trống...). "0 lỗi" không có nghĩa là đúng — đã nhiều lần dữ liệu mất âm thầm.
2. Số dòng bị loại cao bất thường → xem vài dòng trong file gốc để phân biệt **lỗi dữ liệu nguồn** (báo người dùng, không sửa) và **lỗi tool** (vá).
3. Sau mọi thay đổi: `python3 -m py_compile` cả 3 file, chạy **bảng hồi quy ở mục 7**, và kiểm tra qua GUI thật khi đụng tới giao diện hoặc đường đi dữ liệu.
4. Quy tắc mơ hồ hoặc có nhiều cách hợp lý (vd cột đơn vị ghi chức vụ) → hỏi người dùng, đưa 2-3 lựa chọn cụ thể.
5. Khi báo cáo: nêu lỗi tìm được, cách sửa, số liệu trước/sau, và các lỗi dữ liệu nguồn cần người dùng xử lý (kèm STT, tên, giá trị cụ thể). Nếu bản vá làm thay đổi kết quả của file đã xử lý trước đó, nói rõ file nào và vì sao.
6. Quy ước code: chú thích trong code viết tiếng Việt **không dấu**, giải thích **lý do** (thường kèm ví dụ thực tế đã gặp); chuỗi hiển thị cho người dùng viết **có dấu**. Giữ tương thích ngược cho CLI và các hàm công khai.

## 7. Bảng hồi quy (số liệu hiện tại, đã xác minh)

Chạy: `python regression.py --detail` (đọc `test_data/` với tên ngắn: SonLinh.pdf, DongThap.pdf, KonDao.pdf, TaPhin.pdf, DaNang.xlsx, LongHoa.xlsx, CanTho.xlsx, VINATOM.xls, NghiaDo.pdf, BenCat.xlsx, BinhHoa.docx, ChanhHung.docx, P_LongHoa.doc, BaoTang.doc, SonLa.xlsx, ChiengKen.pdf, YenThanh.pdf, DakLak.xlsx, NinhBinh.xlsx, ThaiNguyen.xls, KhuDiTich.docx; file thiếu thì báo "BO QUA"; kèm kiểm tra đơn vị không cần file). Khi thêm file mới: chép vào `test_data/` tên ngắn không dấu, thêm vào `MODE1`/`MODE2` trong `regression.py` và bảng dưới đây. Chưa tìm thấy trên máy (30/09/2026): Support_excel.xlsx (NinhBinh), CV_163 Thái Nguyên, Khu Di tích. Cột "Đọc" = số dòng đọc được, "KQ" = số tài khoản xuất ra, "CKT" = số dòng Cần kiểm tra.

**Chế độ 1** (`core.extract_all([path], dedupe=True)`; Đọc = `stats["total_rows"]`)

| File | Đọc | KQ |
|---|---|---|
| CV_2096_UBNDX_...Nền_tảng_AI_công_vụ.pdf (Sơn Linh, Quảng Ngãi) | 25 | 25 |
| 3_Khu_Di_tích_Phủ_chủ_tich.docx | 43 | 42 |
| Danh_sa_üch__Éo_é_Çng_Tha_üp.pdf (Đồng Tháp) | 2381 | 2239 |
| ĐT02_Đăng_ký_danh_sách_người_dùng_CLEX.pdf (Kon Đào) | 10 | 10 |
| Tả_Phìn.pdf | 80 | 78 |
| DangKy_2109.xlsx (Cần Thơ) | 5000 | 4968 |
| BẢO_TÀNG_MỸ_THUẬT_THÀNH_PHỐ_HỒ_CHÍ_MINH.doc | 10 | 10 |
| Văn_phòng_UBND_thành_phố_Đà_Nẵng.xlsx | 82 | 82 |
| Xa_Long_Hoa_..._Vr_25_tháng_9_.xlsx | 96 | 95 |
| Support_excel.xlsx (Ninh Bình, không có tiêu đề) | 5417 | 4327 |
| Chiềng_Ken.pdf | 71 | 67 |
| xã_Yên_Thành.pdf | 135 | 135 |
| Danh_sach_tao_tai_khoan_..._Dot_4.xlsx (Đắk Lắk) | 330 | 325 |

Chế độ 1 đổi ngày 30/09/2026 (khôi phục bản sửa bị ghi đè): Tả Phìn 76/74 → 80/78 (khôi phục STT 23/40/55/79); Đồng Tháp 2431/2238 → 2381/2239 (bỏ 50 dòng tiêu đề nhóm; giữ Hồ Quang Lớn — email bị cắt 2 ô, lệch cột); Cần Thơ 4999 → 5000 (Excel không ghép dòng). Chế độ 2 không đổi số nào.

Cảnh báo mong đợi (các file khác không có): Đồng Tháp chế độ 1 "STT nhảy số 724, 743, 915, 920, 2315"; Sơn La chế độ 2 "STT nhảy số 85" (Văn phòng UBND tỉnh không được đánh số, dữ liệu vẫn đủ). Tả Phìn cả 2 chế độ: khôi phục 23, 40, 55, 79.

**Chế độ 2** (`v2.extract_v2(path, default_to_chuc=..., dedupe=True)`; Đọc = `stats["total_read"]`)

| File | Tổ chức mặc định | Đọc | KQ | CKT |
|---|---|---|---|---|
| DangKy_2109.xlsx | Cần Thơ | 5000 | 4968 | 32 |
| TK_AI_VNLNTVN.XLS | Viện Năng lượng nguyên tử Việt Nam | 539 | 532 | 7 |
| CV_163_SXD_THAI_NGUYEN_...AI.xls | Thái Nguyên | 155 | 153 | 2 |
| ĐT02_Đăng_ký_danh_sách_người_dùng_CLEX.pdf | Quảng Ngãi | 10 | 10 | 0 |
| Tả_Phìn.pdf | Lào Cai | 80 | 78 | 2 |
| Nghĩa_Đô.pdf | Lào Cai | 70 | 69 | 1 |
| P_Bến_Cát.xlsx | Hồ Chí Minh | 23 | 22 | 1 |
| P_Bình_Hòa.docx | Hồ Chí Minh | 38 | 0 | 38 |
| P_Chánh_Hưng.docx | Hồ Chí Minh | 2 | 2 | 0 |
| P_Long_Hòa.doc (cần Word+pywin32 hoặc LibreOffice) | Hồ Chí Minh | 97 | 0 | 97 |
| BẢO_TÀNG_MỸ_THUẬT_THÀNH_PHỐ_HỒ_CHÍ_MINH.doc | Hồ Chí Minh | 10 | 10 | 0 |
| 2__TH_danh_sách_đăng_ký_..._kem_CV_.xlsx (Sơn La) | Sơn La | 5981 | 5053 | 928 |
| Văn_phòng_UBND_thành_phố_Đà_Nẵng.xlsx | Đà Nẵng | 82 | 82 | 0 |
| Xa_Long_Hoa_..._Vr_25_tháng_9_.xlsx | Hồ Chí Minh | 96 | 95 | 1 |
| Support_excel.xlsx | Ninh Bình | 5417 | 4327 | 1090 |
| Chiềng_Ken.pdf | Lào Cai | 71 | 67 | 4 |
| xã_Yên_Thành.pdf | Lào Cai | 135 | 135 | 0 |
| Danh_sach_tao_tai_khoan_..._Dot_4.xlsx | Đắk Lắk | 330 | 325 | 5 |

Kiểm tra định tính (không chỉ đếm):
- Bình Hòa và Long Hòa (.doc) **không có email trong nguồn** → 0 là đúng.
- Trùng email phải giữ đúng chủ email: Bến Cát giữ "Lê Thị Hồng Hà" (`lthha`), Nghĩa Đô giữ "Trần Thế Anh" (`trantheanh2`), Long Hòa giữ "Hồ Văn Luyến" (`hvluyen`).
- Tả Phìn cả 2 chế độ phải có STT 23, 40, 55, 79 (khôi phục từ văn bản thô).
- Đồng Tháp chế độ 1: có `hoquanglon@gmail.com` / SĐT `0919099499`.
- VINATOM: mọi dòng Tổ chức = "VIỆN NĂNG LƯỢNG NGUYÊN TỬ VIỆT NAM" (kể cả "Trung tâm Chiếu xạ Hà Nội").
- Đà Nẵng: 82/82 có SĐT (tiêu đề "Di động"), đơn vị "Văn phòng UBND thành phố Đà Nẵng - Đà Nẵng".
- Đơn vị sau khi gộp xã/phường: Tả Phìn, Nghĩa Đô, Kon Đào, Chánh Hưng, Long Hòa mỗi file chỉ còn **1** đơn vị (`Xã Tả Phìn - Lào Cai`, `Xã Nghĩa Đô - Lào Cai`, `Xã Kon Đào - Quảng Ngãi`, `Phường Chánh Hưng - Hồ Chí Minh`, `Xã Long Hòa - Hồ Chí Minh`); Chiềng Ken, Yên Thành cũng 1 đơn vị; Ninh Bình còn 33 đơn vị; Đắk Lắk còn 6 (Sở KHCN 143, 3 xã, Sở Y tế 31, và 1 người STT 107 ô đơn vị chỉ ghi "tỉnh Đắk Lắk" → "Đắk Lắk" — dữ liệu nguồn thiếu); Sơn La: 32 người email `.taxua@` thuộc `Xã Tà Xùa - Sơn La` (không phải Tạ Khoa); Bến Cát giữ 9 đơn vị gốc; Cần Thơ, VINATOM, Thái Nguyên, Sơn La, Đà Nẵng không đổi.
- **Mỗi file chế độ 2 chỉ có đúng 1 giá trị Tổ chức** (kiểm tra bằng `Counter(x["to_chuc"] for x in records)`). Từng lỗi: Ninh Bình có 134 người "Hải Đường" bị gán Hải Phòng, Sơn La có 25 người "Xã Bình Thuận" bị gán Lâm Đồng, Tả Phìn có 2 người "Đoàn TNCS Hồ Chí Minh" bị gán TP HCM.
- Ninh Bình (Support_excel): 117 dòng thiếu hẳn `@ninhbinh.gov.vn` (vd `kydh.phuongthientruong`) bị loại, đúng quy tắc hiện tại.

GUI kiểm tra được bằng script (dựng `gui.ContactExtractorGUI()`, gán `mode1_input_path.set(...)` / `add_mode2_paths([...], to_chuc=...)`, gọi `start_extract()`, đọc `active_panel.stats_text` và `log_text`). Luồng nền gọi `self.after(...)` nên script phải chạy trong `app.mainloop()` (dùng `app.after` để chờ và kiểm tra), không vòng lặp `update()` — Python 3.13 báo "main thread is not in main loop". Máy người dùng có thể không cho chụp màn hình (`ImageGrab` báo "screen grab failed") → đọc nội dung widget thay thế.

## 8. Việc còn mở / ý tưởng đã nêu nhưng chưa làm

- Tự bỏ hậu tố tỉnh/thành khi tên đơn vị đã chứa tên đó (vd "...thành phố Đà Nẵng - Đà Nẵng"): đã đề xuất, người dùng chưa quyết.
- Bảo tàng Mỹ thuật: đơn vị ra rất dài "Bảo tàng Mỹ thuật Thành phố Hồ Chí Minh, Sở Văn hóa và Thể thao Thành phố Hồ Chí Minh - Hồ Chí Minh". Đã hỏi có bỏ phần cơ quan chủ quản sau dấu phẩy không, chưa có trả lời.
- Gom logic dùng chung 2 chế độ: lọc trùng email đang viết 2 lần (`core.dedupe_by_email` và trong `v2.extract_v2`); bỏ "Ông/Bà", tên miền cắt cụt gọi riêng ở mỗi chế độ. Chế độ 1 chưa xuất tệp `_can_kiem_tra.xlsx`. Chế độ 1 vẫn đọc PDF theo vị trí cột (chỉ tìm email/SĐT theo nội dung khi cột Email không hợp lệ).
- Bình Hòa, P Long Hòa chế độ 1 ra 0/0 không cảnh báo (nguồn không có cột email, không có "@").
- Tự điền tên miền cho email thiếu/cắt cụt (vd `@laocai.go` → `@laocai.gov.vn`, `@gov.vn` → `@sonla.gov.vn`, thiếu hẳn `@ninhbinh.gov.vn`) theo tên miền phổ biến của file. Hiện đang loại. Người dùng chưa quyết.
- (cũ) Tự thêm tên miền cho email thiếu hẳn `@...` (suy từ các email cùng đơn vị, vd `@ninhbinh.gov.vn`): đã đề xuất như một tuỳ chọn bật/tắt, người dùng chưa quyết. Mặc định hiện tại là loại.
- Đắk Lắk: trước đây 5 người khối Sở KHCN ghi bộ phận chung chung ("Văn phòng Sở", "phòng HC-TH") được giữ nguyên; kết quả ngày 30/09/2026 không còn giá trị riêng lẻ nào (đã vào "Sở Khoa học và Công nghệ"). Câu hỏi "có gộp đơn vị cấp dưới Sở vào Sở không" vẫn chưa hỏi người dùng một cách tổng quát.
- `PROVINCE_MERGE_MAP` là thông tin hành chính có thể thay đổi; cần kiểm tra lại nếu dùng cho dữ liệu ở thời điểm khác.