# Công cụ tạo danh sách tài khoản

Các cơ quan nhà nước gửi danh sách cán bộ đăng ký tài khoản "Nền tảng Trí tuệ nhân tạo hỗ trợ công vụ"
dưới rất nhiều dạng: công văn PDF, Word (`.docx`/`.doc`), Excel (`.xlsx`/`.xls`), CSV. Công cụ này
**trích xuất Họ tên / Email / SĐT** (và **Đơn vị / Tổ chức** ở chế độ 2), chuẩn hoá theo quy tắc cố định rồi
xuất ra tệp Excel theo file mẫu để nhập vào hệ thống tạo tài khoản. Dòng nào không xử lý chắc chắn được
thì đưa vào danh sách **"Cần kiểm tra"** kèm lý do; công cụ không tự bịa dữ liệu.

Có **3 cách dùng**, cùng một bộ máy trích xuất (sửa quy tắc ở `core.py`/`mode2.py` là áp dụng cho cả 3):

| Cách dùng | Dành cho | Cách chạy |
|---|---|---|
| **Giao diện Tkinter** (máy tính để bàn) | Người dùng cuối trên Windows | `py gui.py` |
| **Web** (backend FastAPI + frontend React) | Dùng qua trình duyệt, và để luyện triển khai | xem [Chạy bản web](#cách-2-bản-web) |
| **Dòng lệnh** | Chạy hàng loạt, viết script | `py -m dataextraction.core ...` / `py -m dataextraction.mode2 ...` |

## Cấu trúc thư mục

```
(thư mục gốc)
├── gui.py                       Chạy giao diện Tkinter (không cần cài package)
├── pyproject.toml               Khai báo package "dataextraction", thư viện, tuỳ chọn cài (web / dev / windows)
├── src/dataextraction/
│   ├── core.py                  Logic dùng chung + chế độ 1
│   ├── mode2.py                 Logic riêng của chế độ 2 (dùng lại core.py)
│   ├── gui.py                   Mã giao diện Tkinter
│   ├── web/app.py               API web (FastAPI); KHÔNG phục vụ giao diện
│   └── templates/               2 file mẫu Excel (chế độ 1 và chế độ 2)
├── frontend/                    Giao diện web: Vite + React + TypeScript + Tailwind CSS
├── DESIGN.md                    Quy chuẩn thiết kế giao diện web (màu, chữ, bố cục, chuyển động)
├── tests/                       Test API web (pytest)
├── scripts/regression.py        Bảng hồi quy trên tệp thật (dành cho người bảo trì)
├── docs/                        web.md (kiến trúc web, API, biến môi trường), FIXING.md
├── legacy/                      Giao diện Tkinter bản cũ, chỉ để tham khảo, không còn chạy được
├── CLAUDE.md                    Ghi chú kỹ thuật chi tiết và các quy tắc nghiệp vụ đã chốt (cho người bảo trì / Claude)
└── test_data/                   Tệp thật của người dùng để chạy hồi quy; KHÔNG commit (có thể không tồn tại)
```

**Nhánh Git:** toàn bộ mã (Tkinter, API web, frontend, tài liệu) nằm trên nhánh **`main`**.
Các nhánh `claude/...` là nhánh làm việc tạm cho từng đợt thay đổi; muốn dùng bản mới nhất thì dùng `main`.

## Cài đặt trên Windows (làm 1 lần)

Cần **Python 3.10 trở lên** (đã thử trên 3.13) và, nếu chạy bản web, **Node.js 20.19+ hoặc 22.12+** (kèm `npm`).
Mở PowerShell trong thư mục dự án (ví dụ `D:\DataExtractionTool`) rồi chạy:

```powershell
git checkout main
git pull
py -m pip install -e ".[windows]"      # thư viện cốt lõi + pywin32 (để đọc .doc/.xls cũ qua Microsoft Office)
```

Máy bạn không có `git` trong PATH thì dùng GitHub Desktop (Fetch origin, chọn nhánh `main`, Pull origin) hoặc
chạy `git.exe` bằng đường dẫn đầy đủ của GitHub Desktop.

Các tuỳ chọn cài đặt (`pip install -e ".[tên]"`):

| Tuỳ chọn | Cài thêm | Khi nào cần |
|---|---|---|
| (không có) | openpyxl, python-docx, pdfplumber, xlrd | Chỉ chạy giao diện Tkinter / dòng lệnh |
| `windows` | pywin32 (chỉ trên Windows) | Đọc `.doc` (Word 97-2003) và `.xls` hỏng qua Microsoft Word/Excel |
| `web` | fastapi, uvicorn, python-multipart | Chạy API của bản web |
| `dev` | gồm `web` + pytest, httpx, ruff | Phát triển, chạy test |

Có thể ghép: `py -m pip install -e ".[windows,web]"`. Nếu chỉ chạy `gui.py`, cài nhanh bằng
`py -m pip install openpyxl python-docx pdfplumber xlrd` cũng được.

**Đọc tệp `.doc`/`.xls` cũ:** công cụ nhờ **Microsoft Word/Excel** trên máy tự chuyển sang `.docx`/`.xlsx` (Word chạy ẩn,
không ảnh hưởng các cửa sổ Word bạn đang mở; cần `pywin32`). Không có Microsoft Office thì dùng
[LibreOffice](https://www.libreoffice.org/) nếu có cài; không có cả hai thì mở tệp bằng Word/Excel, chọn "Save As"
sang `.docx`/`.xlsx` rồi dùng tệp mới. Trên Linux/macOS chỉ có LibreOffice.

## Cách 1: Giao diện Tkinter

```powershell
py gui.py
```

Chọn chế độ ở tab đầu trang, chọn tệp, bấm **Trích xuất & Xem trước**, kiểm tra kết quả, rồi **Xuất tệp kết quả...**
(chi tiết thao tác ở mục [Luồng thao tác trên giao diện Tkinter](#luồng-thao-tác-trên-giao-diện-tkinter-cả-2-chế-độ)).
Cách khác: `py -m dataextraction` (sau khi `pip install -e .`).

## Cách 2: Bản web

Bản web gồm **2 phần chạy riêng**: backend (API Python, cổng 8000) và frontend (React, cổng 5173 khi phát triển).
Cần mở **2 cửa sổ PowerShell**, **chạy backend trước**:

```powershell
# Cửa sổ 1 - backend (ở thư mục gốc dự án)
py -m pip install -e ".[web]"
py -m dataextraction.web                 # http://127.0.0.1:8000  (kiểm tra: /api/health trả {"ok": true})

# Cửa sổ 2 - frontend
cd frontend
npm install
npm run dev                              # http://localhost:5173  <- mở địa chỉ này trong trình duyệt
```

Frontend tự chuyển các yêu cầu `/api/...` sang backend. Nếu backend chạy ở nơi khác:
`$env:VITE_API_PROXY="http://host:cổng"; npm run dev`.

**Lỗi thường gặp**

| Thông báo | Nguyên nhân và cách xử lý |
|---|---|
| `connect ECONNREFUSED 127.0.0.1:8000` (cửa sổ `npm run dev`) | Backend chưa chạy. Mở cửa sổ khác chạy `py -m dataextraction.web` trước, rồi tải lại trang. |
| `No module named 'dataextraction'` | Chưa cài package vào đúng bản Python, hoặc chưa ở nhánh `main` (chưa có thư mục `src`). Chạy `git checkout main`, `git pull`, rồi `py -m pip install -e ".[web]"` trong thư mục gốc. |
| `npm` không nhận ra lệnh | Chưa cài Node.js (bản LTS), hoặc cần mở lại PowerShell sau khi cài. |
| Trang trắng hoặc báo "Lỗi 404/500" | Mở `http://127.0.0.1:8000/api/health` kiểm tra backend; xem log ở cửa sổ backend. |

Build frontend ra tệp tĩnh: `cd frontend; npm run build` (kết quả trong `frontend/dist/`). Backend không tự phục vụ giao diện, nên
khi triển khai thật cần một lớp phía trước phục vụ `dist/` và chuyển `/api/*` sang backend.
Biến môi trường (`DATAEXTRACTION_MAX_UPLOAD_MB`, `DATAEXTRACTION_TTL_SECONDS`, `DATAEXTRACTION_WORK_DIR`), danh sách API và các giới hạn
(chỉ chạy được 1 bản backend, chưa có đăng nhập): xem [`docs/web.md`](docs/web.md). Giao diện và thiết kế: [`frontend/README.md`](frontend/README.md), [`DESIGN.md`](DESIGN.md).

## Cách 3: Dòng lệnh

```powershell
# Chế độ 1 (nhiều tệp được gộp vào 1 kết quả)
py -m dataextraction.core --input a.pdf b.xlsx --output kq.xlsx [--no-dedupe] [--password MẬT_KHẨU]

# Chế độ 2 (1 hoặc nhiều tệp, gộp vào 1 tệp kết quả, dùng chung --to-chuc)
py -m dataextraction.mode2 --input X.xlsx Y.pdf --output kq.xlsx --to-chuc "Cần Thơ" [--password Copenai@2026] [--issues-output kq_can_kiem_tra.xlsx]
```

`--template` mặc định là file mẫu đi kèm trong `src/dataextraction/templates/`. Chế độ 2 qua dòng lệnh để trống cột Mật khẩu
nếu không truyền `--password`. Khi truyền nhiều tệp, mọi tệp dùng chung `--to-chuc`; muốn mỗi tệp (hoặc mỗi sheet) một Tổ chức
riêng thì dùng giao diện Tkinter hoặc web.

## Hai chế độ trích xuất

Giao diện Tkinter và giao diện web đều có 2 tab chế độ ở đầu trang:

| | Chế độ 1 | Chế độ 2 |
|---|---|---|
| **Dùng khi** | Chỉ cần Email / Họ tên / SĐT | Cần thêm Đơn vị và Tổ chức |
| **Đầu vào** | 1 tệp PDF, Word (.docx/.doc), Excel (.xlsx/.xls) hoặc CSV mỗi lần | Nhiều tệp PDF, Word (.docx/.doc), Excel (.xlsx/.xls), CSV cùng lúc |
| **File mẫu** (trong `src/dataextraction/templates/`) | `cls_template_users.xlsx` | `TemplateV2.xlsx` |
| **Đầu ra** | Email, Mật khẩu, Họ và tên, Điện thoại | Tên tài khoản, Email, SĐT, Mật khẩu, Giới tính*, Ngày sinh*, Đơn vị, Tổ chức |

\* để trống — nguồn không có thông tin này nên tool không tự bịa.

Trên giao diện Tkinter, mỗi chế độ giữ riêng tệp đã chọn và kết quả xem trước; chuyển tab không làm
mất dữ liệu của tab kia (giao diện web thì xoá danh sách tệp khi đổi chế độ). File mẫu mặc định tự đổi theo chế độ.

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
   Tên miền gõ sai rõ ràng so với tên miền phổ biến trong cùng tệp (vd `@tphcm.gov.com`,
   `@laocai.goc.vn`, `@phuttho.edu.vn`) được **tự sửa** kèm cảnh báo nêu STT.
   Email có dấu cách trong **tên miền** (vd `...@sonla gov.vn`, `...@p htho.gov.vn`) cũng bị loại thay vì
   nối liền thành tên miền sai.
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
- **Đơn vị**: ô đơn vị ghi kèm Sở chủ quản ở cuối (vd `Phòng X - Sở Nông nghiệp và Môi trường thành phố
  Hồ Chí Minh`) được gộp về Sở (`Sở Nông nghiệp và Môi trường - Hồ Chí Minh`).
  Mặc định: `<cấp hành chính> <tên riêng> - <tỉnh/thành>` (vd `UBND phường
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
  Tool còn cảnh báo khi "Tổ chức mặc định" khác tỉnh gợi ý bởi tên miền email phổ biến (vd email
  `@laocai.gov.vn` nhưng mặc định `Hồ Chí Minh`), và điền sẵn gợi ý vào ô này khi để trống (Excel/Word).

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
- **Dòng khôi phục từ văn bản gốc PDF** (kèm họ tên đã lấy được để bạn soát nhanh).
- **Sheet Excel bị bỏ qua** vì trùng với sheet trước, **họ tên bị ngắt thành 2
  dòng đã được nối lại** (PDF), hoặc **nhiều dòng bị ghép** vì cột Họ và tên
  trống.

"0 dòng cần kiểm tra" chưa chắc đã đúng — hãy luôn đọc các cảnh báo này.

Tool cũng tự xử lý một số lỗi trình bày thường gặp trong PDF: bảng tách qua
nhiều trang, dòng tiêu đề nhóm (vd "Ủy ban nhân dân xã ...") không bị tính là
người, email bị cắt sang ô bên cạnh, cột bị lệch.

## Luồng thao tác trên giao diện Tkinter (cả 2 chế độ)

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

## Phát triển và kiểm thử

```powershell
py -m pip install -e ".[dev]"
py -m pytest -q                          # test API web (gọi API thật, không mock logic trích xuất)
py -m ruff check .                       # kiểm tra mã Python mới (web, tests)

cd frontend
npm test                                 # Vitest
npm run lint                             # ESLint
npm run build                            # tsc + vite build
```

Sau khi sửa quy tắc trích xuất ở `core.py`/`mode2.py`, chạy thêm bảng hồi quy trên tệp thật:
`py scripts/regression.py --detail` (đọc `test_data/`; tệp nào thiếu thì báo "BO QUA"). Số liệu chuẩn và các bẫy dữ liệu đã gặp
nằm trong [`CLAUDE.md`](CLAUDE.md).

## Triển khai

Repo cố ý **không** có Dockerfile, docker compose, cấu hình Nginx, CI hay manifest Kubernetes: phần đóng gói và triển khai do chủ dự án
tự làm. Mã được viết để triển khai thuận tay: backend cấu hình bằng biến môi trường, có `/api/health` làm healthcheck, và frontend
build ra tệp tĩnh. Chi tiết và các giới hạn cần biết: [`docs/web.md`](docs/web.md).

## Tài liệu khác

| Tệp | Nội dung |
|---|---|
| [`docs/web.md`](docs/web.md) | Kiến trúc web, API, biến môi trường, hiệu năng, giới hạn |
| [`frontend/README.md`](frontend/README.md) | Chạy/build frontend, cấu trúc mã giao diện |
| [`DESIGN.md`](DESIGN.md) | Quy chuẩn thiết kế giao diện web |
| [`CLAUDE.md`](CLAUDE.md) | Ghi chú kỹ thuật chi tiết, quy tắc nghiệp vụ đã chốt, bảng hồi quy |
| [`legacy/README.md`](legacy/README.md) | Giao diện Tkinter bản cũ (không còn chạy được) |
