# DESIGN.md — Giao diện web "Tạo danh sách tài khoản"

Nguồn chuẩn cho giao diện web (`frontend/`). Khi mã và tệp này khác nhau thì sửa cho khớp tệp này, hoặc cập nhật tệp này cùng lúc. Thiết kế dựa trên phong cách dịu, ấm, tương phản nhẹ ("soft") và giữ tông của bản Tkinter hiện có (nền kem, coral, bề mặt tối cho Nhật ký). Đây là mô tả tự viết cho dự án này, không chép từ thương hiệu nào.

## 1. Tinh thần
- Công cụ làm việc nội bộ, người dùng xử lý nhiều tệp liên tục: **rõ ràng, dày thông tin, êm mắt**. Không phải trang giới thiệu.
- Núm điều chỉnh (đã thống nhất với người dùng): độ phá cách bố cục **thấp** (1 cột chính, căn lề đều), độ dày thông tin **cao** (bảng gọn), chuyển động **vừa phải (4/10)**.
- Dữ liệu chính xác là ưu tiên số 1: lỗi, cảnh báo, dòng cần kiểm tra luôn dễ thấy và không bao giờ chỉ dựa vào màu.

## 2. Màu (token CSS, định nghĩa ở `frontend/src/index.css`)
| Token | Sáng | Tối | Dùng cho |
|---|---|---|---|
| `--bg` | `#faf9f5` | `#1c1b18` | Nền trang |
| `--surface` | `#ffffff` | `#262420` | Thẻ, bảng |
| `--surface-2` | `#f4f1e8` | `#2e2b26` | Hàng tiêu đề bảng, vùng nhấn nhẹ |
| `--border` | `#e6dfd0` | `#3a372f` | Viền, đường kẻ |
| `--text` | `#29261b` | `#ece6d8` | Chữ chính |
| `--muted` | `#6f6858` | `#a39b88` | Chữ phụ, nhãn |
| `--accent` | `#cc785c` | `#d98a6e` | Nút chính, tab đang chọn, số kết quả |
| `--accent-strong` | `#b4634a` | `#e29b81` | Hover của nút chính |
| `--on-accent` | `#181715` | `#181715` | Chữ/biểu tượng trên nền coral (chữ trắng trên `#cc785c` chỉ đạt 3.3:1, không đủ) |
| `--ink` | `#181715` | `#0f0e0d` | Bề mặt tối: nút "Tải tệp", Nhật ký |
| `--warn` / `--warn-bg` | `#8a5a0c` / `#fbf1dc` | `#e3b866` / `#3a2f14` | Cảnh báo |
| `--danger` / `--danger-bg` | `#a8402b` / `#f9e9e4` | `#f0937c` / `#3b1f18` | Lỗi |
- Coral **chỉ** cho: nút chính, tab/hàng đang chọn, con số kết quả. Không tô nền lớn bằng coral.
- Hệ thống **không dùng xanh lá** (kể cả trạng thái thành công): thành công = chữ thường + biểu tượng dấu tích màu `--text`.
- Chữ trên nền màu phải đạt tương phản ≥ 4.5:1 ở cả hai giao diện sáng/tối.

## 3. Chữ
- Thân: `"Segoe UI", system-ui, sans-serif`, 15px / 1.5. Bảng: 14px.
- Tiêu đề trang và con số kết quả lớn: serif (`Georgia, "Times New Roman", serif`), weight thường (400). Tiêu đề thẻ: 15px weight 600.
- Mã/email/SĐT trong bảng: `"Cascadia Mono", Consolas, ui-monospace, monospace`, 13px.
- Mọi chuỗi hiển thị là **tiếng Việt có dấu**. Không emoji trong nút/thẻ/tab; dùng biểu tượng nét mảnh (lucide-react, nét 1.75).

## 4. Bố cục
- Thanh nội dung tối đa 1280px, lề 16px (mobile) / 24px (≥ 768px). **Hai khung** (≥ 1024px): trái 380px cố định (dính khi cuộn) gồm chế độ, vùng thả tệp + danh sách nguồn, Tuỳ chọn (thu gọn), nút Trích xuất; phải là kết quả (trống: 3 bước hướng dẫn; đang chạy: thanh tiến độ + số giây; xong: nút tải, dải thống kê, bảng). Dưới 1024px xếp dọc, kết quả nằm dưới và tự cuộn tới khi xong.
- Thẻ: nền `--surface`, viền 1px `--border`, bo góc 12px, đệm 16px, không đổ bóng (chỉ thẻ nổi như menu mới có bóng nhẹ).
- Khoảng cách theo bậc 4/8/12/16/24. Giữa các thẻ 16px.
- Bảng dữ liệu: hàng cao 36px, tiêu đề dính trên cùng, **ảo hoá hàng** (chỉ vẽ hàng đang thấy), cuộn trong khung cao tối đa 520px. Không phân trang.
- Mỗi nguồn là một thẻ xếp dọc (tên tệp › sheet, số email, Bỏ/Dùng sheet, Xoá, ô Tổ chức). Không cuộn ngang toàn trang; tên tệp dài phải xuống dòng.

## 5. Thành phần
- **Nút chính** (Trích xuất): nền `--accent`, chữ `--on-accent`, bo 8px, cao 36px. Hover `--accent-strong`. Khi đang chạy: biểu tượng xoay + chữ "Đang trích xuất…", không cho bấm lại.
- **Nút phụ**: nền `--surface`, viền `--border`. **Nút tải tệp**: nền `--ink`, chữ sáng.
- **Tab chế độ**: hai nút dạng phân đoạn, mục đang chọn nền `--accent`. Có thể điều hướng bằng bàn phím (mũi tên).
- **Ô nhập**: nền `--bg`, viền `--border`, bo 6px; focus: viền `--accent` + vòng 2px cùng màu 30%.
- **Vùng thả tệp**: viền đứt 2px `--border`; khi kéo tệp vào: viền `--accent`, nền accent 8%. Bấm hoặc Enter để mở hộp chọn tệp.
- **Dòng nguồn**: tên tệp › sheet, số email, ô Tổ chức mặc định, nút Bỏ/Dùng sheet, nút Xoá. Sheet bị bỏ: chữ `--muted` + gạch ngang + nhãn "Bỏ qua: trùng N/M email với sheet trước" (không chỉ dựa vào màu).
- **Dải thống kê**: `N dòng đọc được → −a thiếu email → −b sai định dạng → −c trùng email → **K** dòng kết quả`; K cỡ serif 28px màu `--accent`.
- **Cảnh báo**: nền `--warn-bg`, chữ `--warn`, biểu tượng tam giác cảnh báo, bo 8px. **Lỗi**: tương tự với `--danger`.

## 6. Chuyển động (mức 4/10)
- Thời lượng 150–250ms, easing `ease-out`. Chỉ dùng cho: thẻ kết quả hiện dần + trượt lên 8px; chuyển tab Xem trước/Cần kiểm tra mờ dần; dòng nguồn mới vào/ra; con số kết quả đếm lên; nút bấm lún nhẹ (scale 0.98).
- **Không** hiệu ứng trong bảng dữ liệu, không hiệu ứng cuộn, không chuyển động lặp vô hạn ngoài biểu tượng đang tải.
- Tôn trọng `prefers-reduced-motion`: tắt mọi chuyển động trừ biểu tượng đang tải (đổi thành tĩnh).

## 7. Trợ năng
- Mọi điều khiển dùng được bằng bàn phím, vòng focus nhìn thấy rõ. Nhãn `aria-label` tiếng Việt cho nút chỉ có biểu tượng.
- Trạng thái đang chạy và kết quả có `aria-live="polite"`.
- Không truyền đạt thông tin chỉ bằng màu (xem phần 2, 5).

## 8. Nên / Không nên
- Nên: giữ con số và nhãn thật (dòng bị loại, lý do, STT) ở ngay chỗ nhìn thấy; hiện đủ tên tệp nguồn.
- Nên: giữ nguyên thuật ngữ của công cụ: "Cần kiểm tra", "Tổ chức mặc định", "Lọc trùng email".
- Không: gradient, kính mờ, đổ bóng nặng, ảnh nền, chữ in hoa toàn bộ trong thân trang, xanh lá.
- Không: ẩn lỗi sau hộp thoại; lỗi và cảnh báo nằm trong trang.
