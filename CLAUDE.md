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
| Đầu vào | GUI: **1 tệp mỗi lần** (hàm/CLI vẫn nhận danh sách, giữ tương thích) | GUI: **nhiều tệp, Excel tách theo sheet**, mỗi nguồn 1 Tổ chức mặc định → `v2.extract_v2_batch([(path, sheet, to_chuc), ...])` (tuple 2 phần tử `(path, to_chuc)` vẫn dùng được) |
| File mẫu | `cls_template_users.xlsx`: Email, Mật khẩu, Họ và tên, Điện thoại | `TemplateV2.xlsx`: Tên tài khoản, Email, Số điện thoại, Mật khẩu, Giới tính, Ngày sinh, Đơn vị, Tổ chức |
| Cột để trống | — | Giới tính, Ngày sinh (không có trong nguồn → không bịa) |

Chế độ 2 là **phần mở rộng** của chế độ 1: áp dụng đủ quy tắc chuẩn hoá của chế độ 1 (gọi thẳng `core.normalize_email/phone/name`, `core.email_name_match_score`) rồi thêm Đơn vị/Tổ chức. **Sửa quy tắc email/SĐT/tên ở `extract_contacts.py` là tự động áp dụng cho cả 2 chế độ.**

GUI (`gui.py`): 2 tab chế độ ở đầu trang; toàn bộ trang cuộn được như một khối (header, 2 cột, nút Xuất tệp, Nhật ký đều nằm trong 1 Canvas). Mỗi chế độ có `ResultPanel` riêng với tab **"Xem trước (N)"** và **"Cần kiểm tra (N)"** (tìm kiếm không dấu, sắp xếp theo cột). Chỉ ghi file khi bấm "Xuất tệp kết quả..." (dùng chung `shared_export_button`). Chế độ 2 xuất kèm `<tên>_can_kiem_tra.xlsx` nếu có dòng bị loại.

**Chế độ 2 nhiều tệp, tách sheet** (người dùng chọn ngày 30/09 và 02/10/2026): danh sách nguồn là Treeview `m2_tree` (cột Tệp › Sheet, Email, Tổ chức mặc định, Trạng thái), dữ liệu ở `mode2_sources` = list dict `{path, sheet, to_chuc, emails, skip, overlap}`. Tệp mới lấy giá trị ô `to_chuc_var`, nút "Gán" áp cho dòng đang chọn (không chọn = tất cả).
- **Tệp Excel có ≥ 2 sheet có email → mỗi sheet 1 dòng**, mỗi sheet 1 Tổ chức riêng (tệp TTYT Bác Ái/Tam Long chứa 2 cơ quan ở 2 tỉnh khác nhau). Sheet trùng ≥ 50% email với sheet trước mặc định **"Bỏ qua"** (tô xám); nút "Dùng / Bỏ sheet" hoặc nhấp đúp để đổi; Nhật ký ghi sheet nào đã bỏ và bao nhiêu email mới bị bỏ theo.
- `extract_v2_batch` gọi `extract_v2(..., dedupe=False, sheet=...)` cho từng nguồn rồi **lọc trùng email trên toàn lô, kể cả khác tệp/sheet** (`v2.dedupe_records`): giữ người có tên khớp email hơn, ngang điểm giữ người đứng trước; người bị loại vào Cần kiểm tra kèm `issue["note"]` = "Trùng email với dòng 24 (STT 5) - Tên (tệp X.xlsx)" (số dòng thật trong sheet Excel, vì nhiều nguồn không có STT — Sơn La). **Gộp vào 1 tệp kết quả**; mỗi bản ghi/issue có khoá `source` (nhãn `tệp.xlsx [Sheet: tên]` chỉ khi cùng 1 tệp có ≥ 2 nguồn); cảnh báo có tiền tố nguồn; nguồn đọc lỗi → cảnh báo và chạy tiếp (tất cả lỗi mới báo lỗi); `stats["files"]` thống kê từng nguồn. Tab Cần kiểm tra và `write_v2_issues` có thêm cột "Tệp nguồn" và "Ghi chú". Chế độ 1 trong GUI chỉ 1 tệp (`mode1_input_path`).
- Thêm tệp Excel gọi `core.list_excel_sheets` trên luồng giao diện (Sơn La ~1,4 giây). Bố cục: cột trái rộng cố định 380px (`pack_propagate(False)`) nên chiều cao phải đặt theo nội dung (`_sync_left_height`), nếu không phần dưới (Mật khẩu, nút Trích xuất) bị che khi thẻ bên trái cao hơn cột phải; bảng dùng `_attach_scrollbars` (grid), không `pack side=left`, vì bảng rộng hơn khung sẽ đẩy thanh cuộn ra ngoài. Phím tắt: Ctrl+O, Enter, Ctrl+S. Tông màu ấm (accent `#c15f3c`, nền `#f6f3ee`). Mọi chuỗi hiển thị là tiếng Việt có dấu.

## 4. Quy tắc nghiệp vụ (đã được người dùng xác nhận)

**Email** (`core.normalize_email`)
- Viết thường, bỏ mọi khoảng trắng/xuống dòng (PDF hay ngắt dòng giữa email).
- Tự sửa lỗi gõ phím rõ ràng (`_repair_email_typos`): `,` → `.`; `..` → `.`; bỏ `.` ngay trước/sau `@`. Ví dụ `ten..x@gmail,com` → `ten.x@gmail.com`.
- `@@` → `@`. Chế độ 2 đọc PDF/Word theo nội dung: ô chứa `@` (không có khoảng trắng) được coi là ô email dù chưa đúng chuẩn, để bước sửa lỗi gõ xử lý (trước đây `@laocai,gov.vn`, `@@...` bị bỏ qua thành "thiếu email").
- **Tên miền bị cắt cụt** (`find_truncated_domain_emails`, chạy trên toàn file ở cả 2 chế độ): đuôi bị cắt (`laocai.go`, `sonla.gov`, `cantho.edu`) hoặc chỉ còn hậu tố (`@gov.vn`, `@edu.vn`) so với tên miền phổ biến trong file (≥ 3 lần, nhiều gấp ≥ 5) → loại, lý do "Email sai định dạng". Không tự điền phần thiếu. Chỉ xét 2 kiểu này để không loại nhầm tên miền hợp lệ như `daklak.gov.vn` khi file có nhiều `vubon.daklak.gov.vn`. Lưu ý: sửa `@.gov.vn` → `@gov.vn` (Sơn La, Lóng Phiêng) vẫn bị loại ở bước này vì mất tên đơn vị.
- Sửa tên miền Gmail gõ sai (`_fix_gmail_domain`, **luôn chạy**, kể cả khi email đúng cú pháp): `@gmailcom`, `@gmai.com`, `@gmial.com`, `@gamil.com`, `@gmail.con`, `@gmail.co`... → `@gmail.com`. Chỉ sửa Gmail, không đoán tên miền cơ quan. (Từng có 30 email ở Sơn La, 16 ở Đồng Tháp "hợp lệ" nhưng không nhận được thư.)
- Sau khi sửa vẫn có "nhãn rỗng" (`_has_empty_email_label`) hoặc không có `@` → **loại** (lý do "Email sai định dạng"). Không bao giờ tự đoán/chèn thêm nội dung bị mất (ví dụ thiếu hẳn `@` hay thiếu tên miền).
- **Dấu cách giữa phần tên** (`_has_ambiguous_space_in_local`): `qthoang thongnhat@phutho.gov.vn`, `Nhen.BT daidong@...` (Phú Thọ) rất có thể thiếu dấu chấm → nối liền sẽ ra địa chỉ sai một cách âm thầm → **loại** ("Email sai định dạng"). Chỉ xét khi cả ô chỉ gồm ký tự email + khoảng trắng. Vẫn nối liền khi tên miền là Gmail (Gmail bỏ qua dấu chấm) hoặc phần sau dấu cách toàn chữ số (`hdkhang 1 @cantho.gov.vn` → `hdkhang1@`). Xuống dòng trong email (PDF ngắt dòng) và dấu cách sát `@`/`.` vẫn bỏ như cũ.
- **Email bị cắt giữa từ** (`_is_cut_mid_word`): người nhập gõ tiếng Việt có dấu trong email (`thươngpv.xaxuanhong@...`, `hoangminhhảo1992@gmail.com`); regex chỉ khớp phần đuôi (`ngpv.xaxuanhong@...`, `o1992@gmail.com`) → **loại** nếu ký tự ngay trước phần khớp là chữ/số (kể cả có dấu). `Email:abc@`, `(abc@` vẫn đúng. Đã soát từng dòng trên dữ liệu thật: mọi dòng đổi đều là email bị cắt thật. Đề xuất chưa làm: tự bỏ dấu (`thươngpv` → `thuongpv`) — nguy hiểm khi ký tự lạ do nhận dạng chữ (`ơlenganmau...`), chờ người dùng quyết.
- Không có email → loại (lý do "Thiếu email"). Email là trường bắt buộc duy nhất.
- Có thể có 2 cột email (vd "Thư điện tử công vụ" + "Thư điện tử Gmail"): dùng cột 2 (`email_alt`) khi cột 1 trống.

**Số điện thoại** (`core.normalize_phone`): 10 số, bắt đầu bằng `0`; xử lý `+84`/`84`, float từ Excel (`912345678.0`), thiếu số 0 đầu (9 số). Ô có nhiều số → lấy số hợp lệ đầu tiên. Không chuẩn hoá được (vd 11 chữ số) → **để trống**, vẫn giữ bản ghi. **Luôn truyền giá trị thô (giữ kiểu float/int) vào hàm này, không ép `str()` trước** (từng làm mất SĐT).

**Họ tên** (`core.normalize_name`): viết hoa chữ cái đầu mỗi từ, bỏ ký tự đặc biệt, cắt phần chức vụ/đơn vị bị dính vào tên. Bỏ kính ngữ "Ông/Bà" ở đầu (`strip_honorific_by_email`) **chỉ khi** phần còn lại ≥ 3 chữ **và** khớp email hơn ("Bà Lý Thị Mụi" / `lythimui`). "Ông" cũng là họ thật ("Ông Quang Đông" / `dongoq`, "Ông Mai Xuân" / `ongmaixuan`) → không bỏ mù quáng.

**Trùng email** (bật mặc định): mỗi email chỉ giữ 1 người. Khi trùng, giữ người có tên **khớp email** hơn (`core.email_name_match_score`: `hvluyen` ↔ "Hồ Văn Luyến" được 2 điểm); ngang điểm thì giữ người xuất hiện trước. Người bị loại vào "Cần kiểm tra" (lý do "Trùng email"). Lý do: nhiều file nguồn dán nhầm email của người khác vào dòng phía trên.

**Mật khẩu**: mặc định `Copenai@2026` (`core.DEFAULT_PASSWORD`) cho **cả 2 chế độ**; GUI có 1 ô nhập dùng chung. `write_v2_output(password=None)` để trống (giữ tương thích CLI cũ).

**Đơn vị** (Chế độ 2, `v2.normalize_don_vi`)
- **Đơn vị cấp dưới xã/phường gộp thành chính xã/phường đó** (yêu cầu của người dùng): phòng ban, trường học, trạm y tế, đoàn thể, thôn/tổ dân phố... → `Xã/Phường/Thị trấn <Tên> - <tỉnh>`. Ví dụ `Phòng VH- XH Phường Nam Định` → `Phường Nam Định - Ninh Bình`; `Văn phòng Đảng ủy xã Tả Phìn` → `Xã Tả Phìn - Lào Cai`. Tên xã lấy từ chính ô đơn vị (`_find_commune_in_text`: lấy lần xuất hiện cuối có tên, nhưng bỏ qua từ khoá nằm **trong** tên vừa tìm — `Phường Xã Đàn` giữ nguyên; bỏ "xã hội", "xã viên"; bỏ khi từ đứng trước là `thị` (thị xã = cấp huyện cũ), `tác` (hợp tác xã), `các`/`cấp`; "XÃ/PHƯỜNG" viết hoa thì tên cũng phải viết hoa; **cả ô viết thường** (`phòng vh xã nam định`) thì nhận tên viết thường, tối đa 5 từ, dừng ở huyện/quận/tỉnh/thành phố; dừng ở `-`, `,`, `/`, `(`). Ô không ghi tên xã ("Văn phòng Đảng ủy", "Phòng kinh tế", "Hội nông dân Việt Nam xã") → dùng xã/phường của **cơ quan ban hành** ở đầu văn bản (`_find_doc_commune`: trang đầu PDF, bảng quốc hiệu + đoạn mở đầu Word, các dòng trên bảng Excel), trừ khi ô chứa dấu hiệu cấp cao hơn (`_ABOVE_COMMUNE_RE`: tỉnh, thành phố, Sở, Bộ, Cục; **"Cán bộ" không phải "Bộ"**) hoặc đơn vị không thuộc xã (`_NOT_COMMUNE_UNIT_RE`: huyện, quận, thị xã, bệnh viện, ngân hàng, chi nhánh, đại học, cao đẳng, kho bạc, BHXH, chi cục, điện lực, bưu điện, công ty...) — vd văn bản xã Nghĩa Đô có "Bệnh viện Đa khoa huyện Bảo Yên" thì giữ nguyên, không đoán là xã. Không có tên xã ở đâu cả (vd Bến Cát, chỉ có trong tên file/email) → giữ nguyên, không suy đoán.
- Đơn vị cấp xã trở lên (Sở, Viện, Ban QLDA...): dạng `<cấp hành chính> <tên riêng> - <tỉnh/thành>`: `UBND phường Ninh Kiều` → `Phường Ninh Kiều - Cần Thơ`; `Sở Tư pháp` → `Sở Tư pháp - Cần Thơ`. Bỏ tiền tố UBND/HĐND trước Xã/Phường/Thị trấn.
- Mở rộng viết tắt: `TT` → Trung tâm, `TP.` → Thành phố, `CĐCĐ`, `QLDA`...
- **Luôn thêm hậu tố tỉnh/thành**, kể cả khi tên đã chứa tên tỉnh (ra `... Thành phố Cần Thơ - Cần Thơ`). Người dùng chưa yêu cầu bỏ phần lặp này; đừng tự đổi.
- Nếu giá trị chính là một tỉnh/thành → chỉ giữ tên riêng, không thêm hậu tố.
- Nếu "Tổ chức mặc định" **không phải** tỉnh/thành (vd "Viện Năng lượng nguyên tử Việt Nam") → không thêm hậu tố.
- Cột đơn vị ghi **chức vụ** thay vì tên đơn vị (vd Bến Cát: "Chủ tịch UBND phường") → **giữ nguyên dữ liệu gốc** (người dùng đã chọn như vậy), không suy từ đuôi email.
- **Ngoại lệ (người dùng đồng ý 02/10/2026): ô chỉ ghi chức vụ THUẦN TÚY** (`_looks_like_job_title` đúng **và** `_strip_job_title` rỗng: "Giám đốc sở", "Phó Giám đốc Sở") **mà văn bản có cơ quan ban hành rõ ở đầu và không phải của xã/phường** → đơn vị = cơ quan ban hành (`_find_doc_issuer`: cụm chữ **viết hoa** ở đầu dòng bắt đầu bằng Sở/Ban/Trung tâm/Văn phòng/Bệnh viện/Viện/Chi cục/Cục/Trường/Bảo tàng/Thanh tra, dừng ở từ có chữ thường hoặc "CỘNG HÒA XÃ HỘI"; bỏ dòng "UBND TỈNH..."; cách viết lấy theo thân văn bản nếu có, vd "Sở Nông nghiệp và Môi trường", không thì viết hoa chữ đầu mỗi từ). Gia Lai: `Sở Nông nghiệp và Môi trường - Gia Lai`. Ô "chức vụ + đơn vị" (VINATOM "Giám đốc, Trung tâm Chiếu xạ Hà Nội") và "Chủ tịch UBND phường" vẫn giữ nguyên. Văn bản của xã/phường: không áp dụng.

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
- **Excel nhiều sheet — cả 2 chế độ** (`core.analyze_sheets`, `SHEET_SKIP_MIN_EMAILS=3`, `SHEET_SKIP_OVERLAP=0.5`): trước đây Chế độ 2 chỉ đọc sheet đang mở khi lưu (`wb.active`, mất 43 người của TTYT Bác Ái, không cảnh báo), Chế độ 1 đọc mù quáng mọi sheet (kéo 16 người "đầu mối" của Sơn La vào kết quả). Quy tắc: đọc sheet **có email**; **bỏ sheet có ≥ 3 email mà ≥ 50% đã xuất hiện ở các sheet trước** (bản nháp/phụ lục: Thái Nguyên "Trang_tính1" 62/94, Sơn La phụ lục đầu mối 68/85) kèm cảnh báo nêu số email mới bị bỏ; sheet không có email (danh mục VINATOM) không đọc; không sheet nào có email → đọc sheet đầu như cũ. Chế độ 2 gọi với `sheet="tên"` thì **chỉ** đọc sheet đó, không áp quy tắc bỏ trùng (người dùng đã chủ động chọn). Hàm dùng chung: `core.read_excel_sheets` (giữ nguyên tên sheet, kể cả dấu cách thừa như `"Người dùng "`), `core.list_excel_sheets`. **STT nhảy số tính riêng từng sheet.** Xã/phường/cơ quan ban hành tính lại **cho từng bảng** (từng sheet), `unit_cache` khoá theo `(đơn vị thô, xã ban hành, cơ quan ban hành)`.
- **Dòng phía trên bảng Excel** (tìm xã/phường ban hành): nối các ô của cùng 1 dòng bằng **xuống dòng**, không bằng dấu cách ("PHƯỜNG TAM LONG" ở A1 + "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" ở D1 từng ra `Phường Tam Long Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam`). `_find_commune_in_text` ở chế độ viết hoa dừng trước "CỘNG HÒA XÃ HỘI" và "ĐỘC LẬP" ("Xã Cộng Hòa" là tên xã thật nên không dừng ở "CỘNG HÒA" trơn). Quy tắc căn tên xã theo xã ban hành chỉ thay khi dài hơn tối đa 2 từ.
- **Không** gọi hàm ghép dòng bị ngắt (`merge_wrapped_continuation_rows`) cho Excel/CSV ở **cả 2 chế độ** (`extract_from_table_rows(..., merge_wrapped=False)`): dòng thiếu tên trong Excel là lỗi dữ liệu thật, ghép nhầm sẽ làm mất 1 người (Cần Thơ: từng mất 1 người ở cả 2 chế độ).

**PDF/Word trong Chế độ 2 — đọc theo NỘI DUNG, không theo vị trí cột**
- Các trang của cùng 1 PDF có thể khác bố cục cột (Nghĩa Đô: trang 2 email ở cột 4, trang 3 lệch 9 cột, trang 4-5 email ở cột 6). Vì vậy `_process_data_rows_content_based` + `_parse_row_by_content` nhận email/SĐT bằng regex ở bất kỳ ô nào; tên/đơn vị là 2 ô văn bản còn lại theo thứ tự.
- Dòng "mảnh vỡ" (chỉ có SĐT hoặc email, không tên) được ghép vào bản ghi ngay trước.
- **Họ tên bị ngắt 2 dòng vật lý** (Gia Lai: `Cao Thanh` + dòng sau `Thương`, nhiều khi kèm SĐT lặp lại): dòng chỉ có **đúng 1 từ chữ cái** ở **đúng cột tên** của người vừa thêm, không email/STT, ô khác rỗng (hoặc chỉ là SĐT trùng SĐT người trước), tên trước ≤ 3 từ, cùng kiểu viết hoa, không phải từ chỉ cơ quan → nối vào họ tên (`_is_name_continuation`; Chế độ 1: `core.merge_name_fragment_rows`). Từng bị hiểu là tiêu đề nhóm (thành đơn vị của người kế tiếp) hoặc người ảo (tên `Khánh`, đơn vị `Vy`). Luôn kèm **cảnh báo** "Đã nối N họ tên bị ngắt thành 2 dòng (STT ...)" để đối chiếu.
- **Email bị cắt theo viền ô bảng PDF** (Phú Thọ Nguyệt Đức, 9 người): pdfplumber nhận nhầm viền cột email chỉ rộng ~31pt, ô email chỉ còn mảnh đầu (`Hainv`, `Huypq`, `dthuong21`), phần còn lại rơi ngoài bảng, văn bản thô xen kẽ các cột nên không nối được. `core.repair_cut_email_cells` (gọi ở cả `core.read_pdf_file` và `v2._read_tables_and_context`, **trước** khi phân tích dòng): dòng không có ô nào chứa `@` mà có mảnh (cột ≥ 2, 1 từ chỉ ký tự email, có chữ cái) → tìm từ trên trang bắt đầu bằng mảnh (`extract_words`, toạ độ), nối các từ **cùng cột** (|Δx0| < 3) ở dòng ngay dưới (cách < 16pt), bỏ SĐT của chính dòng bị dính vào (`phutho.0387875219`), dừng khi đủ tên miền hoặc gặp email của dòng sau; **chỉ nhận** khi tên miền == tên miền phổ biến nhất của tệp (≥ 3 lần) và phần trước `@` bắt đầu bằng mảnh. Luôn kèm cảnh báo "Khôi phục N email bị cắt theo viền ô bảng PDF (STT ...)".
- **Tiêu đề cột bị ngắt dòng** (`Số điện\nthoại sử\ndụng Zalo`): `_strip_accents_lower` gộp mọi khoảng trắng/xuống dòng thành 1 dấu cách, nếu không cả cột SĐT bị bỏ (Chế độ 1 Nguyệt Đức từng ra 0 SĐT).
- **STT lệch cột**: nếu ô đầu rỗng, STT = ô có nội dung đầu tiên **nếu là số ≤ 4 chữ số hoặc số La Mã** (không lấy SĐT/tên; Gia Lai STT 8, 9 từng ra tên rỗng).
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
- Khi ghép dòng bị ngắt, **không nối ô STT** (trước ra "909910"), và **không nối khi cả 2 dòng đều có email** (2 người khác nhau; từng ra email rác `thuongct@...vntranntt`).
- **Ô tiêu đề gộp ô** (`core.collapse_header_spans`): pdfplumber trả tiêu đề "Họ và tên" ở cột 3, cột 4-5 là `None`, dữ liệu thật ở cột 4 (Gia Lai: bảng 15 cột, email cột 9 hoặc 10). Mỗi ô tiêu đề bao phủ các cột tới ô tiêu đề kế (tối đa 3); **chỉ thu gọn khi có bằng chứng** ≥ 2 dòng có ô cột tiêu đề rỗng mà ô khác trong vùng có giá trị. Trước đây cột Tên rỗng ở mọi dòng → mọi dòng bị coi là "dòng bị ngắt" và ghép hết vào dòng đầu: **9 người thành 1 bản ghi sai, không cảnh báo**. Nay còn có cảnh báo khi ≥ 30% dòng bị ghép vì cột tên trống.
- Họ tên bị ngắt 2 dòng: xem mục PDF/Word ở trên, Chế độ 1 dùng `merge_name_fragment_rows` trước `merge_wrapped_continuation_rows` (chỉ PDF/Word, không bảng suy cột).
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

Chạy: `python regression.py --detail` (đọc `test_data/` với tên ngắn: SonLinh.pdf, DongThap.pdf, KonDao.pdf, TaPhin.pdf, DaNang.xlsx, LongHoa.xlsx, CanTho.xlsx, VINATOM.xls, NghiaDo.pdf, BenCat.xlsx, BinhHoa.docx, ChanhHung.docx, P_LongHoa.doc, BaoTang.doc, SonLa.xlsx, ChiengKen.pdf, YenThanh.pdf, DakLak.xlsx, GiaLai_SNNMT.pdf, TTYT_TamLong.xlsx, NinhBinh.xlsx, ThaiNguyen.xls, KhuDiTich.docx; file thiếu thì báo "BO QUA"; kèm kiểm tra đơn vị không cần file). Khi thêm file mới: chép vào `test_data/` tên ngắn không dấu, thêm vào `MODE1`/`MODE2` trong `regression.py` và bảng dưới đây. Chưa tìm thấy trên máy (30/09/2026): Support_excel.xlsx (NinhBinh), CV_163 Thái Nguyên, Khu Di tích. Cột "Đọc" = số dòng đọc được, "KQ" = số tài khoản xuất ra, "CKT" = số dòng Cần kiểm tra.

**Chế độ 1** (`core.extract_all([path], dedupe=True)`; Đọc = `stats["total_rows"]`)

| File | Đọc | KQ |
|---|---|---|
| CV_2096_UBNDX_...Nền_tảng_AI_công_vụ.pdf (Sơn Linh, Quảng Ngãi) | 25 | 25 |
| 3_Khu_Di_tích_Phủ_chủ_tich.docx | 43 | 42 |
| Danh_sa_üch__Éo_é_Çng_Tha_üp.pdf (Đồng Tháp) | 2381 | 2229 |
| ĐT02_Đăng_ký_danh_sách_người_dùng_CLEX.pdf (Kon Đào) | 10 | 10 |
| Tả_Phìn.pdf | 80 | 78 |
| DangKy_2109.xlsx (Cần Thơ) | 5000 | 4964 |
| BẢO_TÀNG_MỸ_THUẬT_THÀNH_PHỐ_HỒ_CHÍ_MINH.doc | 10 | 10 |
| Văn_phòng_UBND_thành_phố_Đà_Nẵng.xlsx | 82 | 82 |
| Xa_Long_Hoa_..._Vr_25_tháng_9_.xlsx | 96 | 95 |
| Support_excel.xlsx (Ninh Bình, không có tiêu đề) | 5417 | 4327 |
| Chiềng_Ken.pdf | 71 | 67 |
| xã_Yên_Thành.pdf | 135 | 135 |
| Danh_sach_tao_tai_khoan_..._Dot_4.xlsx (Đắk Lắk) | 330 | 324 |
| 11190_SNNMT-VP_30092026-signed_01.pdf (Sở NN&MT Gia Lai) | 9 | 9 |
| TTYT_KV_BAC_AI.xlsx (2 sheet: TTYT Bác Ái 43 + Đảng ủy P. Tam Long 36) | 79 | 79 |
| 2__TH_danh_sách_đăng_ký_..._kem_CV_.xlsx (Sơn La, phụ lục đầu mối bị bỏ) | 5981 | 5043 |
| cv_cc_danh_sach_ccvc_dk_tao_tk_..._ptn...pdf (Phú Thọ, P. Thống Nhất) | 69 | 67 |
| 261005_ra_soat_cung_cap_danh_sach_ccvc_...pdf (Phú Thọ, xã Nguyệt Đức) | 54 | 53 |
| danh_sach_dang_ky_tai_khoan_ai_xa_dai_dong-5.xlsx (Phú Thọ, xã Đại Đồng) | 172 | 169 |

Chế độ 1 đổi ngày 06/10/2026 (quy tắc email Phú Thọ): Đồng Tháp 2239 → 2229, Cần Thơ 4968 → 4964, Đắk Lắk 325 → 324, Sơn La 5053 → 5043 — toàn bộ là email bị cắt giữa từ (đã soát từng dòng); các file khác không đổi.

Chế độ 1 đổi ngày 02/10/2026: Gia Lai 1/1 (email rác) → 9/9; Sơn La 6116/5069 → 5981/5053 (bỏ sheet phụ lục đầu mối trùng 80%; 16 người liên hệ không còn bị đưa vào danh sách tài khoản); các file cũ không đổi số nào.

Chế độ 1 đổi ngày 30/09/2026 (khôi phục bản sửa bị ghi đè): Tả Phìn 76/74 → 80/78 (khôi phục STT 23/40/55/79); Đồng Tháp 2431/2238 → 2381/2239 (bỏ 50 dòng tiêu đề nhóm; giữ Hồ Quang Lớn — email bị cắt 2 ô, lệch cột); Cần Thơ 4999 → 5000 (Excel không ghép dòng). Chế độ 2 không đổi số nào.

Cảnh báo mong đợi (các file khác không có): Đồng Tháp chế độ 1 "STT nhảy số 724, 743, 915, 920, 2315"; Sơn La cả 2 chế độ "STT nhảy số 85" (Văn phòng UBND tỉnh không được đánh số, dữ liệu vẫn đủ) và "Bỏ qua sheet 'DANH SÁCH ĐẦU MỐI TRIỂN KHAI NỀ': trùng 68/85 email (17 email mới không được đọc)"; Gia Lai cả 2 chế độ "Đã nối 7 họ tên bị ngắt thành 2 dòng (STT 1-7)". Tả Phìn cả 2 chế độ: khôi phục 23, 40, 55, 79. Thái Nguyên (chưa có tệp): kỳ vọng cảnh báo bỏ sheet "Trang_tính1" (62/94 trùng).

**Chế độ 2** (`v2.extract_v2(path, default_to_chuc=..., dedupe=True)`; Đọc = `stats["total_read"]`)

| File | Tổ chức mặc định | Đọc | KQ | CKT |
|---|---|---|---|---|
| DangKy_2109.xlsx | Cần Thơ | 5000 | 4964 | 36 |
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
| 2__TH_danh_sách_đăng_ký_..._kem_CV_.xlsx (Sơn La) | Sơn La | 5981 | 5043 | 938 |
| Văn_phòng_UBND_thành_phố_Đà_Nẵng.xlsx | Đà Nẵng | 82 | 82 | 0 |
| Xa_Long_Hoa_..._Vr_25_tháng_9_.xlsx | Hồ Chí Minh | 96 | 95 | 1 |
| Support_excel.xlsx | Ninh Bình | 5417 | 4327 | 1090 |
| Chiềng_Ken.pdf | Lào Cai | 71 | 67 | 4 |
| xã_Yên_Thành.pdf | Lào Cai | 135 | 135 | 0 |
| Danh_sach_tao_tai_khoan_..._Dot_4.xlsx | Đắk Lắk | 330 | 324 | 6 |
| 11190_SNNMT-VP_30092026-signed_01.pdf | Gia Lai | 9 | 9 | 0 |
| TTYT_KV_BAC_AI.xlsx (`sheet=None`, đọc cả 2 sheet với 1 Tổ chức) | Khánh Hòa | 79 | 79 | 0 |
| Phú Thọ (P. Thống Nhất) | Phú Thọ | 69 | 67 | 2 |
| Phú Thọ (xã Nguyệt Đức) | Phú Thọ | 54 | 53 | 1 |
| Phú Thọ (xã Đại Đồng) | Phú Thọ | 172 | 169 | 3 |

Kiểm tra định tính (không chỉ đếm):
- **Phú Thọ:** mỗi tệp 1 đơn vị (`Phường Thống Nhất`/`Xã Nguyệt Đức`/`Xã Đại Đồng - Phú Thọ`), Tổ chức `TỈNH PHÚ THỌ`. Cần kiểm tra: Thống Nhất STT 36 (email có dấu cách), STT 65 (trùng STT 29); Nguyệt Đức STT 33 (Trần Thị Huệ trùng Trần Thị Huế STT 18); Đại Đồng STT 21 (email có dấu cách), 16 (trùng 15), 53 (Bùi Văn Sợn trùng Bùi Văn Sòn STT 42). Nguyệt Đức: 9 email khôi phục STT 29, 30, 31, 32, 34, 35, 37, 39, 40.
- Đại Đồng cảnh báo mong đợi "STT nhảy số (170, 172, 174)" — nguồn bỏ số thật. STT 158 Phạm Thúy Ngà cột SĐT ghi "Nghỉ hưu" (SĐT để trống, vẫn tạo tài khoản) — hỏi người dùng. Thống Nhất STT 53/54 cùng SĐT `0382670103`.
- **Gia Lai:** 9 họ tên đầy đủ (`Cao Thanh Thương`, `Nguyễn Thị Tố Trân`, `Trần Đình Chương`, `Hà Thị Thanh Hương`, `Nguyễn Văn Hoan`, `Nguyễn Thị Thế Vy`, `Trần Quốc Khánh`, `Đoàn Ngọc Có`, `Vũ Ngọc An`), đơn vị `Sở Nông nghiệp và Môi trường - Gia Lai`, `TỈNH GIA LAI`.
- **TTYT Bác Ái / Tam Long, theo từng sheet** (`extract_v2_batch` với `(path, sheet, to_chuc)`): sheet 1 → `Trung Tâm Y Tế Khu Vực Bác Ái - Khánh Hòa` / `TỈNH KHÁNH HÒA` (43 người, Bác Ái thuộc Ninh Thuận cũ); sheet 2 → `Phường Tam Long - Hồ Chí Minh` / `THÀNH PHỐ HỒ CHÍ MINH` (36 người, email `@tphcm.gov.vn`). Tổng 79, 0 cần kiểm tra. Với `sheet=None` và 1 Tổ chức mặc định thì cả 79 người cùng 1 tỉnh (đúng chỉ khi tệp thuộc 1 tỉnh).
- `list_excel_sheets`: TTYT → 2 sheet (43, 36 email, không bỏ); Sơn La → 2 sheet, sheet 2 `suggest_skip`; VINATOM → chỉ `"Người dùng "`.
- Lọc trùng khác tệp: chạy `TaPhin.pdf` 2 lần trong 1 lô → 78 người, 78 dòng trùng, mỗi dòng có ghi chú trỏ tới STT và tệp.
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
- **Họ tên có dấu chấm giữa chữ** (TTYT Bác Ái: `Nguyễn T.trang Vương`) hiện ra `Nguyễn Ttrang Vương` vì bước bỏ ký tự đặc biệt xoá dấu chấm làm dính 2 từ. Đề xuất đổi dấu chấm giữa chữ thành khoảng trắng (`Nguyễn T Trang Vương`); chưa làm, người dùng chưa quyết (đổi ở `core.normalize_name` → ảnh hưởng cả 2 chế độ, cần chạy hồi quy).
- Chế độ 1 chưa có ghi chú "trùng với dòng nào" (chỉ Chế độ 2) và vẫn đọc PDF theo vị trí cột (chỉ gom ô tiêu đề gộp khi có bằng chứng, xem `collapse_header_spans`).
- Bỏ sheet trùng: 17 email "mới" trong phụ lục đầu mối Sơn La (16 người sau khi chuẩn hoá) không được đọc; nếu muốn xem họ, bật lại sheet trong GUI (họ sẽ vào kết quả cùng 89 dòng trùng ở Cần kiểm tra).
- Các dòng `Phó Giám đốc Sở` ở bảng phụ 1 cột cuối trang PDF (Gia Lai) vẫn bị coi là tiêu đề nhóm (vô hại vì đứng cuối); nếu gặp PDF có bảng 1 cột xen giữa danh sách thì kiểm tra lại.