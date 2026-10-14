# legacy

`gui_notion.py` và `guiv2.py` là các bản giao diện Tkinter cũ (nhập `extract_contacts`/`process_template_v2` theo thư mục phẳng). Giữ lại để tham khảo thiết kế; sau khi chuyển sang package chúng không còn chạy được nếu không sửa import thành `from dataextraction import core` / `from dataextraction import mode2 as v2`.
