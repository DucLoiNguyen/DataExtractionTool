#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui.py
======
Giao diện đồ hoạ DUY NHẤT (Tkinter) cho tool tạo danh sách tài khoản, hỗ
trợ 2 CHẾ ĐỘ trích xuất (chọn bằng nút chọn ở đầu trang):

  - CHẾ ĐỘ 1: Trích xuất từ tài liệu (PDF/Word/Excel/CSV) — dùng khi có
    công văn/danh sách tự do, cần trích Email/Họ tên/SĐT. Xuất theo file
    mẫu "cls_template_users.xlsx".
  - CHẾ ĐỘ 2: Chuẩn hoá danh sách đăng ký (Excel/PDF/Word có cột hoặc ngữ
    cảnh Đơn vị công
    tác) — dùng khi đã có bảng Họ và tên/Đơn vị công tác/Email/SĐT, cần
    chuẩn hoá thêm Đơn vị/Tổ chức. Xuất theo file mẫu "TemplateV2.xlsx".

Cả 2 chế độ đều:
  - Áp dụng ĐẦY ĐỦ 3 quy tắc chuẩn hoá gốc (email đúng định dạng + viết
    thường, số điện thoại chuẩn VN, họ tên viết hoa đúng cách) — không bao
    giờ tự bịa dữ liệu thiếu.
  - Có tab "Cần kiểm tra" liệt kê ĐẦY ĐỦ các dòng bị loại kèm dữ liệu gốc
    và lý do cụ thể (thiếu email / email sai định dạng / trùng email) —
    không chỉ báo 1 con số tổng.
  - CHƯA ghi ra tệp cho đến khi bấm "Xuất tệp kết quả..." và chọn nơi lưu.

extract_contacts.py và process_template_v2.py phải nằm CÙNG THƯ MỤC với
file này.

Phím tắt: Ctrl+O thêm tệp, Enter trích xuất, Ctrl+S xuất tệp kết quả.

Chạy:
    python3 gui.py
"""

import os
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
from tkinter import filedialog, messagebox, ttk

import extract_contacts as core
import process_template_v2 as v2

THIS_DIR = os.path.dirname(os.path.abspath(__file__))

MODE1_INPUT_FILETYPES = [
    ("Tất cả tệp hỗ trợ", "*.pdf *.docx *.xlsx *.xls *.csv"),
    ("PDF", "*.pdf"), ("Word", "*.docx"), ("Excel", "*.xlsx *.xls"),
    ("CSV", "*.csv"), ("Tất cả tệp", "*.*"),
]
MODE2_INPUT_FILETYPES = [
    ("Tất cả tệp hỗ trợ", "*.xlsx *.xls *.pdf *.docx *.csv"),
    ("Excel", "*.xlsx *.xls"), ("PDF", "*.pdf"), ("Word", "*.docx"),
    ("CSV", "*.csv"), ("Tất cả tệp", "*.*"),
]
EXCEL_FILETYPES = [("Excel", "*.xlsx"), ("Tất cả tệp", "*.*")]

PREVIEW_LIMIT = 200

FILE_ICONS = {".pdf": "\U0001F4D5", ".docx": "\U0001F4D8", ".xlsx": "\U0001F4D7",
              ".xls": "\U0001F4D7", ".csv": "\U0001F4C4"}

# ----- Cot cho CHE DO 1 -----
M1_PREVIEW_COLUMNS = ("email", "password", "name", "phone")
M1_PREVIEW_HEADINGS = {"email": "Email", "password": "Mật khẩu", "name": "Họ và tên", "phone": "Điện thoại"}
M1_SORTABLE = {"email": "email", "name": "name", "phone": "phone"}

M1_ISSUE_COLUMNS = ("source", "row_number", "raw_email", "raw_name", "raw_phone", "reason")
M1_ISSUE_HEADINGS = {
    "source": "Tệp nguồn", "row_number": "STT", "raw_email": "Email (gốc)",
    "raw_name": "Họ và tên (gốc)", "raw_phone": "Điện thoại (gốc)", "reason": "Lý do",
}
M1_ISSUE_SORTABLE = {"source": "source", "row_number": "row_number", "reason": "reason"}

# ----- Cot cho CHE DO 2 -----
M2_PREVIEW_COLUMNS = ("name", "email", "phone", "don_vi", "to_chuc")
M2_PREVIEW_HEADINGS = {"name": "Tên tài khoản", "email": "Email", "phone": "SĐT",
                        "don_vi": "Đơn vị", "to_chuc": "Tổ chức"}
M2_SORTABLE = {"name": "name", "email": "email", "phone": "phone", "don_vi": "don_vi", "to_chuc": "to_chuc"}

M2_ISSUE_COLUMNS = ("stt", "raw_name", "raw_email", "raw_phone", "raw_unit", "reason")
M2_ISSUE_HEADINGS = {
    "stt": "STT", "raw_name": "Họ và tên (gốc)", "raw_email": "Email (gốc)",
    "raw_phone": "Điện thoại (gốc)", "raw_unit": "Đơn vị công tác (gốc)", "reason": "Lý do",
}
M2_ISSUE_SORTABLE = {"stt": "stt", "raw_name": "raw_name", "reason": "reason"}

# ---------------------------------------------------------------------------
# BẢNG MÀU / THIẾT KẾ (tông ấm)
# ---------------------------------------------------------------------------
COLOR_BG = "#f6f3ee"
COLOR_CARD = "#fffefb"
COLOR_BORDER = "#e8e2d6"
COLOR_HEADER_BG = "#efe8dc"
COLOR_ACCENT = "#c15f3c"
COLOR_ACCENT_DARK = "#a34f31"
COLOR_ACCENT_SOFT = "#f7e9e3"
COLOR_SUCCESS = "#4b7f52"
COLOR_SUCCESS_BG = "#eaf3e8"
COLOR_WARN = "#b3781f"
COLOR_WARN_BG = "#faf1e0"
COLOR_DANGER = "#b1503a"
COLOR_DANGER_BG = "#f8ece7"
COLOR_TEXT = "#3d3929"
COLOR_TEXT_MUTED = "#7a7263"
COLOR_CONSOLE_BG = "#2b2620"
COLOR_CONSOLE_FG = "#e8e2d6"
COLOR_CONSOLE_OK = "#8fbf8f"


def _pick_font(preferred, fallback="TkDefaultFont"):
    try:
        families = set(tkfont.families())
    except Exception:
        return fallback
    for name in preferred:
        if name in families:
            return name
    return fallback


def _search_key(text):
    if not text:
        return ""
    text = str(text).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.lower()


def _fmt(n):
    return f"{n:,}".replace(",", ".")


class ResultPanel:
    """
    Goi 1 'panel ket qua' hoan chinh cho 1 CHE DO: dong thong ke dang
    pipeline + Notebook 2 tab (Xem truoc / Can kiem tra, moi tab co tim
    kiem + sap xep) + nut Xuat tep. Dung CHUNG cau truc cho ca 2 che do,
    chi khac BO COT duoc truyen vao luc khoi tao.

    Day la 1 class thuan UI - KHONG chua logic xu ly du lieu (logic van o
    extract_contacts.py / process_template_v2.py), chi hien thi va cho
    phep loc/sap xep PHAN HIEN THI (khong lam thay doi du lieu goc).
    """

    def __init__(self, app, parent, key_field, preview_columns, preview_headings, preview_sortable,
                 issue_columns, issue_headings, issue_sortable, row_to_values, issue_row_to_values,
                 pipeline_labels):
        self.app = app
        self.key_field = key_field  # field dung de tim kiem chinh (khong dung truc tiep, xem row_to_values)
        self.preview_columns = preview_columns
        self.preview_sortable = preview_sortable
        self.issue_columns = issue_columns
        self.issue_sortable = issue_sortable
        self.row_to_values = row_to_values            # ham: record -> tuple gia tri hien thi
        self.issue_row_to_values = issue_row_to_values  # ham: issue -> tuple gia tri hien thi
        self.pipeline_labels = pipeline_labels

        self.records = None   # danh sach ban ghi hop le, THU TU GOC (khong doi khi loc/sap xep)
        self.stats = None
        self.preview_sort_field = None
        self.preview_sort_reverse = False
        self.issue_sort_field = None
        self.issue_sort_reverse = False

        self.frame = tk.Frame(parent, bg=COLOR_BG)
        self._build(preview_headings, issue_headings)

    # ------------------------------------------------------------------
    def _build(self, preview_headings, issue_headings):
        f = self.app.base_font_family

        # ---- Thong ke dang pipeline ----
        body3 = self.app._card(self.frame, "\U0001F4CA", "3. Thống kê")
        self.stats_text = tk.Text(
            body3, height=3, wrap="word", relief="flat", bg=COLOR_CARD,
            font=(f, 11), state="disabled", cursor="arrow", highlightthickness=0, padx=0, pady=0,
        )
        self.stats_text.pack(fill="x")
        self.stats_text.tag_configure("neutral", foreground=COLOR_TEXT, font=(f, 11))
        self.stats_text.tag_configure("warn", foreground=COLOR_WARN, font=(f, 11))
        self.stats_text.tag_configure("danger", foreground=COLOR_DANGER, font=(f, 11))
        self.stats_text.tag_configure("arrow", foreground=COLOR_TEXT_MUTED, font=(f, 11))
        self.stats_text.tag_configure("final", foreground=COLOR_SUCCESS, font=(f, 15, "bold"))
        self.set_stats_placeholder()

        # ---- The chua Notebook (Xem truoc / Can kiem tra) ----
        body4 = self.app._card(self.frame, "\U0001F441", "4. Kết quả chi tiết")
        self.notebook = ttk.Notebook(body4)
        self.notebook.pack(fill="both", expand=True)

        preview_tab = tk.Frame(self.notebook, bg=COLOR_CARD)
        issues_tab = tk.Frame(self.notebook, bg=COLOR_CARD)
        self.notebook.add(preview_tab, text="Xem trước (0)")
        self.notebook.add(issues_tab, text="Cần kiểm tra (0)")

        # --- tab Xem truoc ---
        pbody = ttk.Frame(preview_tab, style="Card.TFrame")
        pbody.pack(fill="both", expand=True, padx=4, pady=10)

        prow = ttk.Frame(pbody, style="Card.TFrame")
        prow.pack(fill="x", pady=(0, 8))
        ttk.Label(prow, text="\U0001F50D", style="Card.TLabel").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.refresh_preview())
        ttk.Entry(prow, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Label(prow, text="Tìm kiếm \u00b7 bấm tiêu đề cột để sắp xếp",
                  style="Muted.TLabel").pack(side="left")

        tree_wrap = tk.Frame(pbody, bg=COLOR_BORDER)
        tree_wrap.pack(fill="both", expand=True)
        tree_inner = tk.Frame(tree_wrap, bg=COLOR_CARD)
        tree_inner.pack(fill="both", expand=True, padx=1, pady=1)

        self.preview_tree = ttk.Treeview(tree_inner, columns=self.preview_columns, show="headings", height=11)
        for col in self.preview_columns:
            if col in self.preview_sortable:
                self.preview_tree.heading(col, text=preview_headings[col],
                                           command=lambda c=col: self._sort_preview(c))
            else:
                self.preview_tree.heading(col, text=preview_headings[col])
            self.preview_tree.column(col, width=150, anchor="w")
        self.preview_tree.pack(side="left", fill="both", expand=True)
        self.preview_tree.tag_configure("odd", background="#ffffff")
        self.preview_tree.tag_configure("even", background="#f7f2e9")
        psb = ttk.Scrollbar(tree_inner, orient="vertical", command=self.preview_tree.yview)
        psb.pack(side="left", fill="y")
        self.preview_tree.config(yscrollcommand=psb.set)

        self.preview_note_var = tk.StringVar(value='Chưa có dữ liệu. Hãy bấm "Trích xuất & Xem trước".')
        ttk.Label(pbody, textvariable=self.preview_note_var, style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

        # --- tab Can kiem tra ---
        ibody = ttk.Frame(issues_tab, style="Card.TFrame")
        ibody.pack(fill="both", expand=True, padx=4, pady=10)

        irow = ttk.Frame(ibody, style="Card.TFrame")
        irow.pack(fill="x", pady=(0, 8))
        ttk.Label(irow, text="\U0001F50D", style="Card.TLabel").pack(side="left")
        self.issue_search_var = tk.StringVar()
        self.issue_search_var.trace_add("write", lambda *a: self.refresh_issues())
        ttk.Entry(irow, textvariable=self.issue_search_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Label(irow, text="Tìm kiếm \u00b7 bấm tiêu đề cột để sắp xếp",
                  style="Muted.TLabel").pack(side="left")

        itree_wrap = tk.Frame(ibody, bg=COLOR_BORDER)
        itree_wrap.pack(fill="both", expand=True)
        itree_inner = tk.Frame(itree_wrap, bg=COLOR_CARD)
        itree_inner.pack(fill="both", expand=True, padx=1, pady=1)

        self.issues_tree = ttk.Treeview(itree_inner, columns=self.issue_columns, show="headings", height=11)
        widths = {"source": 190, "row_number": 55, "stt": 55}
        for col in self.issue_columns:
            if col in self.issue_sortable:
                self.issues_tree.heading(col, text=issue_headings[col],
                                          command=lambda c=col: self._sort_issue(c))
            else:
                self.issues_tree.heading(col, text=issue_headings[col])
            anchor = "center" if col in ("row_number", "stt") else "w"
            self.issues_tree.column(col, width=widths.get(col, 140), anchor=anchor)
        self.issues_tree.pack(side="left", fill="both", expand=True)
        self.issues_tree.tag_configure("odd", background="#ffffff")
        self.issues_tree.tag_configure("even", background="#f7f2e9")
        self.issues_tree.tag_configure("reason_missing", foreground=COLOR_WARN)
        self.issues_tree.tag_configure("reason_invalid", foreground=COLOR_WARN)
        self.issues_tree.tag_configure("reason_duplicate", foreground=COLOR_DANGER)
        isb = ttk.Scrollbar(itree_inner, orient="vertical", command=self.issues_tree.yview)
        isb.pack(side="left", fill="y")
        self.issues_tree.config(yscrollcommand=isb.set)

        self.issues_note_var = tk.StringVar(
            value="Chưa có dữ liệu. Các dòng bị loại (thiếu email, sai định dạng, trùng email) sẽ hiện ở đây."
        )
        ttk.Label(ibody, textvariable=self.issues_note_var, style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

        # ---- Nut xuat tep: KHONG con nam trong panel nua (xem ghi chu o
        # class ContactExtractorGUI._build_widgets) - moi che do dung
        # CHUNG 1 nut Xuat tep DUY NHAT, dat CO DINH o day man hinh, de
        # luon nhin thay du man hinh nho/DPI cao den dau.

    # ------------------------------------------------------------------
    def pack(self):
        self.frame.pack(fill="both", expand=True)

    def pack_forget(self):
        self.frame.pack_forget()

    # ------------------------------------------------------------------
    def set_stats_placeholder(self, text=None):
        self.stats_text.config(state="normal")
        self.stats_text.delete("1.0", tk.END)
        self.stats_text.insert(
            tk.END, text or 'Chưa có dữ liệu. Hãy bấm "Trích xuất & Xem trước" (hoặc nhấn Enter).', "arrow"
        )
        self.stats_text.config(state="disabled")

    def set_stats_pipeline(self, segments):
        """segments: list (text, tag) - xem cac ham _mode1_pipeline_segments /
        _mode2_pipeline_segments trong app de biet cach dung cu the."""
        self.stats_text.config(state="normal")
        self.stats_text.delete("1.0", tk.END)
        for text, tag in segments:
            self.stats_text.insert(tk.END, text, tag)
        self.stats_text.config(state="disabled")

    def reset(self):
        self.records = None
        self.stats = None
        self.search_var.set("")
        self.issue_search_var.set("")
        self.preview_sort_field = None
        self.preview_sort_reverse = False
        self.issue_sort_field = None
        self.issue_sort_reverse = False
        self._update_preview_arrows()
        self._update_issue_arrows()
        self.clear_preview()
        self.clear_issues()
        self.set_stats_placeholder()
        self.app._sync_export_button()

    def clear_preview(self):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.preview_note_var.set('Chưa có dữ liệu. Hãy bấm "Trích xuất & Xem trước".')
        self.notebook.tab(0, text="Xem trước (0)")

    def clear_issues(self):
        for item in self.issues_tree.get_children():
            self.issues_tree.delete(item)
        self.issues_note_var.set(
            "Chưa có dữ liệu. Các dòng bị loại (thiếu email, sai định dạng, trùng email) sẽ hiện ở đây."
        )
        self.notebook.tab(1, text="Cần kiểm tra (0)")

    # ------------------------------------------------------------------
    def load_results(self, records, stats):
        self.records = records
        self.stats = stats
        self.refresh_preview()
        self.refresh_issues()
        self.app._sync_export_button()

    # ------------------------------------------------------------------
    def _update_preview_arrows(self):
        pass  # (giu don gian - khong ve lai tieu de moi lan, tranh phuc tap khong can thiet)

    def _sort_preview(self, col):
        field = self.preview_sortable.get(col)
        if not field:
            return
        if self.preview_sort_field == field:
            self.preview_sort_reverse = not self.preview_sort_reverse
        else:
            self.preview_sort_field = field
            self.preview_sort_reverse = False
        self.refresh_preview()

    def _sort_issue(self, col):
        field = self.issue_sortable.get(col)
        if not field:
            return
        if self.issue_sort_field == field:
            self.issue_sort_reverse = not self.issue_sort_reverse
        else:
            self.issue_sort_field = field
            self.issue_sort_reverse = False
        self.refresh_issues()

    def _update_issue_arrows(self):
        pass  # (giu don gian - khong ve lai tieu de moi lan, tranh phuc tap khong can thiet)

    # ------------------------------------------------------------------
    def refresh_preview(self):
        if self.records is None:
            return
        records = self.records
        query = self.search_var.get().strip()
        if query:
            q = _search_key(query)
            records = [r for r in records if any(q in _search_key(v) for v in self.row_to_values(r))]
        if self.preview_sort_field:
            key = self.preview_sort_field
            records = sorted(records, key=lambda r: _search_key(r.get(key, "")),
                              reverse=self.preview_sort_reverse)
        self._render_preview(records, total_all=len(self.records), query=query)
        self.notebook.tab(0, text=f"Xem trước ({_fmt(len(self.records))})")

    def _render_preview(self, records, total_all, query):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        shown = records[:PREVIEW_LIMIT]
        for i, rec in enumerate(shown):
            tag = "even" if i % 2 == 0 else "odd"
            self.preview_tree.insert("", tk.END, values=self.row_to_values(rec), tags=(tag,))

        total_filtered = len(records)
        if total_all == 0:
            self.preview_note_var.set("Không có dữ liệu hợp lệ để hiển thị.")
        elif query:
            if total_filtered == 0:
                self.preview_note_var.set(f'Không tìm thấy dòng nào khớp với "{query}" (trên {_fmt(total_all)} dòng).')
            else:
                self.preview_note_var.set(
                    f'Đang hiển thị {min(total_filtered, PREVIEW_LIMIT)}/{_fmt(total_filtered)} dòng khớp với '
                    f'"{query}" (trên tổng {_fmt(total_all)} dòng).'
                )
        elif total_filtered > PREVIEW_LIMIT:
            self.preview_note_var.set(f"Đang hiển thị {PREVIEW_LIMIT}/{_fmt(total_filtered)} dòng đầu tiên.")
        else:
            self.preview_note_var.set(f"Đang hiển thị toàn bộ {_fmt(total_filtered)} dòng.")

    # ------------------------------------------------------------------
    def refresh_issues(self):
        if self.stats is None:
            return
        issues = self.stats.get("issues", [])
        query = self.issue_search_var.get().strip()
        filtered = issues
        if query:
            q = _search_key(query)

            def matches(issue):
                reason_label = core.ISSUE_REASON_LABELS.get(issue.get("reason"), issue.get("reason", ""))
                vals = list(self.issue_row_to_values(issue))
                vals.append(reason_label)
                return any(q in _search_key(v) for v in vals)

            filtered = [i for i in issues if matches(i)]

        if self.issue_sort_field:
            key = self.issue_sort_field
            filtered = sorted(filtered, key=lambda i: _search_key(str(i.get(key, ""))),
                               reverse=self.issue_sort_reverse)

        self._render_issues(filtered, total_all=len(issues), query=query)
        self.notebook.tab(1, text=f"Cần kiểm tra ({_fmt(len(issues))})")

    def _render_issues(self, issues, total_all, query):
        for item in self.issues_tree.get_children():
            self.issues_tree.delete(item)
        reason_tag = {"missing_email": "reason_missing", "invalid_email_format": "reason_invalid",
                      "duplicate_email": "reason_duplicate"}
        shown = issues[:PREVIEW_LIMIT]
        for issue in shown:
            reason_label = core.ISSUE_REASON_LABELS.get(issue.get("reason"), issue.get("reason", ""))
            values = list(self.issue_row_to_values(issue)) + [reason_label]
            self.issues_tree.insert("", tk.END, values=values, tags=(reason_tag.get(issue.get("reason"), ""),))

        total_filtered = len(issues)
        if total_all == 0:
            self.issues_note_var.set("Không có dòng nào cần kiểm tra - toàn bộ dữ liệu đều hợp lệ.")
        elif query:
            if total_filtered == 0:
                self.issues_note_var.set(f'Không tìm thấy dòng nào khớp với "{query}" (trên {_fmt(total_all)} dòng).')
            else:
                self.issues_note_var.set(
                    f'Đang hiển thị {min(total_filtered, PREVIEW_LIMIT)}/{_fmt(total_filtered)} dòng khớp với '
                    f'"{query}" (trên tổng {_fmt(total_all)} dòng cần kiểm tra).'
                )
        elif total_filtered > PREVIEW_LIMIT:
            self.issues_note_var.set(f"Đang hiển thị {PREVIEW_LIMIT}/{_fmt(total_filtered)} dòng đầu tiên cần kiểm tra.")
        else:
            self.issues_note_var.set(f"Có {_fmt(total_filtered)} dòng cần kiểm tra.")


class ContactExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tạo danh sách tài khoản \u2192 Excel")
        self.geometry("1240x800")
        self.minsize(1060, 660)
        self.configure(bg=COLOR_BG)

        self.mode_var = tk.StringVar(value="mode1")
        self._ui_ready = False  # tranh <<NotebookTabChanged>> kich hoat qua som luc dang dung giao dien

        # Che do 1
        self.input_paths = []
        # Che do 2
        self.mode2_input_path = tk.StringVar()
        self.to_chuc_var = tk.StringVar()

        self.base_font_family = _pick_font(["Segoe UI", "Helvetica Neue", "Helvetica", "Arial"])
        self.mono_font_family = _pick_font(["Cascadia Mono", "Consolas", "Menlo", "Courier New"])

        self._setup_style()
        self._build_widgets()
        self._setup_shortcuts()
        self._ui_ready = True
        self._apply_mode()

    # ------------------------------------------------------------------
    # STYLE
    # ------------------------------------------------------------------
    def _setup_style(self):
        f = self.base_font_family
        self.option_add("*Font", (f, 10))
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=(f, 10))
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_CARD)
        style.configure("Header.TFrame", background=COLOR_HEADER_BG)
        style.configure("CardTitle.TLabel", background=COLOR_CARD, foreground=COLOR_TEXT, font=(f, 11, "bold"))
        style.configure("Card.TLabel", background=COLOR_CARD, foreground=COLOR_TEXT, font=(f, 10))
        style.configure("Muted.TLabel", background=COLOR_CARD, foreground=COLOR_TEXT_MUTED, font=(f, 9))
        style.configure("MutedOnBg.TLabel", background=COLOR_BG, foreground=COLOR_TEXT_MUTED, font=(f, 9))
        style.configure("AppTitle.TLabel", background=COLOR_HEADER_BG, foreground=COLOR_TEXT, font=(f, 16, "bold"))
        style.configure("AppSubtitle.TLabel", background=COLOR_HEADER_BG, foreground=COLOR_TEXT_MUTED, font=(f, 10))
        style.configure("Shortcut.TLabel", background=COLOR_HEADER_BG, foreground=COLOR_TEXT_MUTED, font=(f, 9))

        style.configure("TButton", font=(f, 10), padding=(10, 6))
        style.configure("Secondary.TButton", font=(f, 10), padding=(10, 6), background="#f0ebe0", foreground=COLOR_TEXT)
        style.map("Secondary.TButton", background=[("active", "#e6ddcc"), ("disabled", "#f2ede2")],
                  foreground=[("disabled", "#a89f8d")])

        style.configure("Accent.TButton", font=(f, 11, "bold"), padding=(14, 9),
                         background=COLOR_ACCENT, foreground="#ffffff")
        style.map("Accent.TButton", background=[("active", COLOR_ACCENT_DARK), ("disabled", "#e3c4b5")],
                  foreground=[("disabled", "#f7ece6")])

        style.configure("Export.TButton", font=(f, 11, "bold"), padding=(14, 9),
                         background=COLOR_SUCCESS, foreground="#ffffff")
        style.map("Export.TButton", background=[("active", "#3a6640"), ("disabled", "#d6e5d3")],
                  foreground=[("disabled", "#eef5ee")])

        style.configure("TEntry", padding=6, fieldbackground="#ffffff")
        style.configure("Card.TCheckbutton", background=COLOR_CARD, foreground=COLOR_TEXT, font=(f, 10))
        style.configure("Card.TRadiobutton", background=COLOR_CARD, foreground=COLOR_TEXT, font=(f, 10, "bold"))

        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#ece5d8",
                         background=COLOR_ACCENT, thickness=8, borderwidth=0)

        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff",
                         foreground=COLOR_TEXT, rowheight=25, font=(f, 10), borderwidth=0)
        style.configure("Treeview.Heading", background=COLOR_ACCENT_SOFT, foreground=COLOR_ACCENT_DARK,
                         font=(f, 10, "bold"), padding=(6, 6), relief="flat")
        style.map("Treeview.Heading", background=[("active", COLOR_ACCENT_SOFT)])
        style.map("Treeview", background=[("selected", COLOR_ACCENT)], foreground=[("selected", "#ffffff")])

        style.configure("TNotebook", background=COLOR_CARD, borderwidth=0)
        style.configure("TNotebook.Tab", background="#f0ebe0", foreground=COLOR_TEXT, font=(f, 10), padding=(14, 7))
        style.map("TNotebook.Tab", background=[("selected", COLOR_CARD)],
                  foreground=[("selected", COLOR_ACCENT_DARK)], font=[("selected", (f, 10, "bold"))])

        style.configure("Vertical.TScrollbar", background="#e3dccc", troughcolor=COLOR_BG,
                         arrowsize=12, borderwidth=0)

    # ------------------------------------------------------------------
    def _card(self, parent, icon, title, subtitle=None):
        outer = tk.Frame(parent, bg=COLOR_BORDER)
        outer.pack(fill="x", pady=(0, 16))
        inner = tk.Frame(outer, bg=COLOR_CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        header = ttk.Frame(inner, style="Card.TFrame")
        header.pack(fill="x", padx=18, pady=(14, 6))
        ttk.Label(header, text=f"{icon}  {title}", style="CardTitle.TLabel").pack(side="left")
        if subtitle:
            ttk.Label(header, text=subtitle, style="Muted.TLabel").pack(side="right")
        body = ttk.Frame(inner, style="Card.TFrame")
        body.pack(fill="both", expand=True, padx=18, pady=(4, 16))
        return body

    # ------------------------------------------------------------------
    # GIAO DIEN
    # ------------------------------------------------------------------
    def _build_widgets(self):
        # ===== 1 VUNG CUON DUY NHAT CHO TOAN BO TRANG =====
        # Header, tab chon che do, 2 cot noi dung, nut Xuat tep, Nhat ky -
        # TAT CA nam trong CUNG 1 Canvas cuon duoc, khong con khu vuc co
        # dinh rieng o tren/duoi nua - cuon nhu 1 trang web thong thuong.
        canvas = tk.Canvas(self, bg=COLOR_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        main = tk.Frame(canvas, bg=COLOR_BG)
        main_window = canvas.create_window((0, 0), window=main, anchor="nw")

        def _on_main_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(main_window, width=event.width)

        main.bind("<Configure>", _on_main_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            delta = event.delta
            if delta == 0:
                return
            canvas.yview_scroll(int(-1 * (delta / 120)), "units")

        def _on_mousewheel_linux(event):
            canvas.yview_scroll(-1 if event.num == 4 else 1, "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)     # Windows/macOS
        canvas.bind_all("<Button-4>", _on_mousewheel_linux)  # Linux scroll up
        canvas.bind_all("<Button-5>", _on_mousewheel_linux)  # Linux scroll down

        # ===== HEADER (nam trong vung cuon) =====
        header = tk.Frame(main, bg=COLOR_HEADER_BG)
        header.pack(fill="x")
        header_inner = ttk.Frame(header, style="Header.TFrame")
        header_inner.pack(fill="x", padx=24, pady=16)
        ttk.Label(header_inner, text="\U0001F4C7  Tạo danh sách tài khoản",
                  style="AppTitle.TLabel").pack(anchor="w")
        row = ttk.Frame(header_inner, style="Header.TFrame")
        row.pack(fill="x", pady=(2, 0))
        ttk.Label(row, text="Trích xuất và chuẩn hoá Email / Họ tên / SĐT, xuất theo file mẫu Excel",
                  style="AppSubtitle.TLabel").pack(side="left")
        ttk.Label(row, text="  \u2022  Ctrl+O thêm tệp \u00b7 Enter trích xuất \u00b7 Ctrl+S xuất tệp",
                  style="Shortcut.TLabel").pack(side="left")

        # ===== TAB CHON CHE DO (thay cho 2 nut chon truoc day) =====
        mode_wrap = tk.Frame(main, bg=COLOR_BG)
        mode_wrap.pack(fill="x", padx=20, pady=(14, 0))
        self.mode_notebook = ttk.Notebook(mode_wrap)
        self.mode_notebook.pack(fill="x")

        tab1 = tk.Frame(self.mode_notebook, bg=COLOR_CARD)
        tab2 = tk.Frame(self.mode_notebook, bg=COLOR_CARD)
        self.mode_notebook.add(tab1, text="\U0001F4D5  Trích xuất từ tài liệu (PDF/Word/Excel/CSV)")
        self.mode_notebook.add(tab2, text="\U0001F4CB  Chuẩn hoá danh sách đăng ký (có cột Đơn vị công tác)")
        ttk.Label(tab1, text="Công văn/danh sách tự do \u2192 file mẫu cls_template_users.xlsx",
                  style="Muted.TLabel").pack(anchor="w", padx=14, pady=10)
        ttk.Label(tab2, text="Bảng Họ tên/Đơn vị/Email/SĐT có sẵn \u2192 file mẫu TemplateV2.xlsx",
                  style="Muted.TLabel").pack(anchor="w", padx=14, pady=10)
        self.mode_notebook.bind("<<NotebookTabChanged>>", self._on_mode_tab_changed)

        # ===== 2 COT =====
        content = tk.Frame(main, bg=COLOR_BG)
        content.pack(fill="both", expand=True, padx=20, pady=16)

        left_col = tk.Frame(content, bg=COLOR_BG, width=380)
        left_col.pack(side="left", fill="y", padx=(0, 8))
        left_col.pack_propagate(False)

        right_col = tk.Frame(content, bg=COLOR_BG)
        right_col.pack(side="left", fill="both", expand=True, padx=(8, 0))

        self._build_left_column(left_col)
        self._build_right_column(right_col)

        # ===== NUT XUAT TEP (dung chung cho ca 2 che do, nam trong vung
        # cuon nhu phan con lai - khong con co dinh rieng nua) =====
        export_wrap = tk.Frame(main, bg=COLOR_BG)
        export_wrap.pack(fill="x", padx=20, pady=(0, 4))
        self.shared_export_button = ttk.Button(
            export_wrap, text="\U0001F4BE  Xuất tệp kết quả...", style="Export.TButton",
            command=self.export_result, state="disabled",
        )
        self.shared_export_button.pack(side="left")
        ttk.Label(export_wrap, text="Chỉ ghi ra tệp khi bạn bấm nút này và chọn nơi lưu",
                  style="MutedOnBg.TLabel").pack(side="left", padx=(12, 0))

        # ===== NHAT KY (cung nam trong vung cuon) =====
        log_wrap = tk.Frame(main, bg=COLOR_BG)
        log_wrap.pack(fill="x", padx=20, pady=(12, 20))
        log_body = self._card(log_wrap, "\U0001F5A5", "Nhật ký xử lý")
        console_wrap = tk.Frame(log_body, bg=COLOR_CONSOLE_BG)
        console_wrap.pack(fill="both", expand=True)
        self.log_text = tk.Text(
            console_wrap, height=6, state="disabled", wrap="word",
            bg=COLOR_CONSOLE_BG, fg=COLOR_CONSOLE_FG, insertbackground=COLOR_CONSOLE_FG,
            relief="flat", font=(self.mono_font_family, 9), padx=10, pady=8,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("ok", foreground=COLOR_CONSOLE_OK)
        self.log_text.tag_configure("err", foreground="#e08a72")

    def _on_mode_tab_changed(self, event=None):
        if not self._ui_ready:
            return
        selected = self.mode_notebook.select()
        if not selected:
            return
        try:
            idx = self.mode_notebook.index(selected)
        except tk.TclError:
            return
        self.mode_var.set("mode1" if idx == 0 else "mode2")
        self._apply_mode()

    def _sync_export_button(self):
        """Bat/tat nut Xuat tep DUNG CHUNG dua tren panel dang duoc chon co
        du lieu hay khong. Goi moi khi doi che do hoac sau khi trich xuat."""
        has_records = bool(self.active_panel.records)
        self.shared_export_button.config(state="normal" if has_records else "disabled")

    # ------------------------------------------------------------------
    def _build_left_column(self, parent):
        # ===== THE 1: CHON TEP NGUON (noi dung doi theo che do) =====
        self.file_card_body = self._card(parent, "\U0001F4C2", "1. Chọn tệp nguồn")

        # -- noi dung cho CHE DO 1: danh sach nhieu tep --
        self.m1_file_frame = tk.Frame(self.file_card_body, bg=COLOR_CARD)
        list_wrap = tk.Frame(self.m1_file_frame, bg=COLOR_BORDER)
        list_wrap.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(
            list_wrap, height=6, selectmode=tk.EXTENDED, activestyle="none",
            bg="#ffffff", fg=COLOR_TEXT, selectbackground=COLOR_ACCENT, selectforeground="#ffffff",
            relief="flat", highlightthickness=0, font=(self.base_font_family, 10),
        )
        self.listbox.pack(fill="both", expand=True, padx=1, pady=1)
        btns_input = ttk.Frame(self.m1_file_frame, style="Card.TFrame")
        btns_input.pack(fill="x", pady=(8, 0))
        for i in range(3):
            btns_input.grid_columnconfigure(i, weight=1)
        ttk.Button(btns_input, text="\U0001F4C1 Thêm", style="Secondary.TButton",
                   command=self.add_files).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(btns_input, text="\u2716 Xoá chọn", style="Secondary.TButton",
                   command=self.remove_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns_input, text="\U0001F9F9 Xoá hết", style="Secondary.TButton",
                   command=self.clear_files).grid(row=0, column=2, sticky="ew", padx=(4, 0))
        self.file_count_var = tk.StringVar(value="Chưa chọn tệp nào.")
        ttk.Label(self.m1_file_frame, textvariable=self.file_count_var, style="Muted.TLabel").pack(
            anchor="w", pady=(8, 0))

        # -- noi dung cho CHE DO 2: 1 tep duy nhat --
        self.m2_file_frame = tk.Frame(self.file_card_body, bg=COLOR_CARD)
        ttk.Label(self.m2_file_frame, text="Tệp danh sách đăng ký (Excel/PDF/Word/CSV):",
                  style="Card.TLabel").pack(anchor="w")
        row_m2 = ttk.Frame(self.m2_file_frame, style="Card.TFrame")
        row_m2.pack(fill="x", pady=(4, 0))
        ttk.Entry(row_m2, textvariable=self.mode2_input_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row_m2, text="...", width=3, style="Secondary.TButton",
                   command=self.choose_mode2_file).pack(side="left", padx=(6, 0))
        ttk.Label(self.m2_file_frame,
                  text="Cần nhận diện được cột Họ và tên/Email (SĐT, Đơn vị công tác nếu có) — "
                       "chấp nhận cả bảng Excel lẫn tài liệu PDF/Word dạng danh sách",
                  style="Muted.TLabel", wraplength=320).pack(anchor="w", pady=(8, 0))

        # ===== THE 2: FILE MAU + TUY CHON =====
        body2 = self._card(parent, "\u2699", "2. File mẫu & Tuỳ chọn")
        ttk.Label(body2, text="File mẫu Excel:", style="Card.TLabel").pack(anchor="w")
        row_tpl = ttk.Frame(body2, style="Card.TFrame")
        row_tpl.pack(fill="x", pady=(4, 10))
        self.template_var = tk.StringVar()
        ttk.Entry(row_tpl, textvariable=self.template_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row_tpl, text="...", width=3, style="Secondary.TButton",
                   command=self.choose_template).pack(side="left", padx=(6, 0))

        # -- rieng CHE DO 1: mat khau mac dinh --
        self.m1_password_frame = tk.Frame(body2, bg=COLOR_CARD)
        ttk.Label(self.m1_password_frame, text="Mật khẩu mặc định:", style="Card.TLabel").pack(anchor="w")
        self.password_var = tk.StringVar(value=core.DEFAULT_PASSWORD)
        ttk.Entry(self.m1_password_frame, textvariable=self.password_var).pack(fill="x", pady=(4, 10))

        # -- rieng CHE DO 2: to chuc mac dinh --
        self.m2_to_chuc_frame = tk.Frame(body2, bg=COLOR_CARD)
        ttk.Label(self.m2_to_chuc_frame, text="Tổ chức mặc định (tỉnh/thành, không tiền tố):",
                  style="Card.TLabel").pack(anchor="w")
        ttk.Entry(self.m2_to_chuc_frame, textvariable=self.to_chuc_var).pack(fill="x", pady=(4, 4))
        ttk.Label(self.m2_to_chuc_frame,
                  text="Dùng khi dữ liệu nguồn không tự nêu rõ tỉnh/thành (vd \"Sở Tư pháp\"). Ví dụ: Cần Thơ",
                  style="Muted.TLabel", wraplength=320).pack(anchor="w", pady=(0, 10))

        self.dedupe_var = tk.BooleanVar(value=True)
        self.dedupe_check = ttk.Checkbutton(body2, text="Tự động loại bỏ bản ghi trùng email",
                                             variable=self.dedupe_var, style="Card.TCheckbutton")
        self.dedupe_check.pack(anchor="w")

        # ===== NUT TRICH XUAT =====
        run_wrap = tk.Frame(parent, bg=COLOR_BG)
        run_wrap.pack(fill="x", pady=(0, 12))
        self.run_button = ttk.Button(run_wrap, text="\u25B6  Trích xuất & Xem trước",
                                      style="Accent.TButton", command=self.start_extract)
        self.run_button.pack(fill="x")
        self.progress = ttk.Progressbar(run_wrap, mode="indeterminate", style="Accent.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(8, 4))
        self.status_var = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(run_wrap, textvariable=self.status_var, style="MutedOnBg.TLabel").pack(anchor="w")

    # ------------------------------------------------------------------
    def _build_right_column(self, parent):
        self.result_container = tk.Frame(parent, bg=COLOR_BG)
        self.result_container.pack(fill="both", expand=True)

        def m1_row_to_values(rec):
            password = self.password_var.get() or core.DEFAULT_PASSWORD
            return (rec["email"], password, rec["name"], rec["phone"])

        def m1_issue_row_to_values(issue):
            return (issue.get("source", ""), issue.get("row_number") if issue.get("row_number") is not None else "",
                    issue.get("raw_email", ""), issue.get("raw_name", ""), issue.get("raw_phone", ""))

        self.panel_m1 = ResultPanel(
            self, self.result_container, "email",
            M1_PREVIEW_COLUMNS, M1_PREVIEW_HEADINGS, M1_SORTABLE,
            M1_ISSUE_COLUMNS, M1_ISSUE_HEADINGS, M1_ISSUE_SORTABLE,
            m1_row_to_values, m1_issue_row_to_values, None,
        )

        def m2_row_to_values(rec):
            return (rec["name"], rec["email"], rec["phone"], rec["don_vi"], rec["to_chuc"])

        def m2_issue_row_to_values(issue):
            return (issue.get("stt", ""), issue.get("raw_name", ""), issue.get("raw_email", ""),
                    issue.get("raw_phone", ""), issue.get("raw_unit", ""))

        self.panel_m2 = ResultPanel(
            self, self.result_container, "email",
            M2_PREVIEW_COLUMNS, M2_PREVIEW_HEADINGS, M2_SORTABLE,
            M2_ISSUE_COLUMNS, M2_ISSUE_HEADINGS, M2_ISSUE_SORTABLE,
            m2_row_to_values, m2_issue_row_to_values, None,
        )

    # ------------------------------------------------------------------
    # CHUYEN DOI CHE DO
    # ------------------------------------------------------------------
    def _apply_mode(self):
        mode = self.mode_var.get()
        # An het truoc, roi hien dung phan
        self.m1_file_frame.pack_forget()
        self.m2_file_frame.pack_forget()
        self.m1_password_frame.pack_forget()
        self.m2_to_chuc_frame.pack_forget()
        self.panel_m1.pack_forget()
        self.panel_m2.pack_forget()

        if mode == "mode1":
            self.m1_file_frame.pack(fill="both", expand=True)
            self.m1_password_frame.pack(fill="x")
            self.panel_m1.pack()
            default_tpl = os.path.join(THIS_DIR, "cls_template_users.xlsx")
        else:
            self.m2_file_frame.pack(fill="both", expand=True)
            self.m2_to_chuc_frame.pack(fill="x")
            self.panel_m2.pack()
            default_tpl = os.path.join(THIS_DIR, "TemplateV2.xlsx")

        # Dam bao checkbox "loai trung email" LUON nam SAU cung trong the 2,
        # bat ke frame nao (mat khau / to chuc) vua duoc pack o tren.
        self.dedupe_check.pack_forget()
        self.dedupe_check.pack(anchor="w")

        if os.path.exists(default_tpl):
            self.template_var.set(default_tpl)
        else:
            self.template_var.set("")

        self._sync_export_button()

    @property
    def active_panel(self):
        return self.panel_m1 if self.mode_var.get() == "mode1" else self.panel_m2

    # ------------------------------------------------------------------
    # PHIM TAT
    # ------------------------------------------------------------------
    def _setup_shortcuts(self):
        self.bind_all("<Control-o>", lambda e: self.add_files() if self.mode_var.get() == "mode1"
                      else self.choose_mode2_file())
        self.bind_all("<Control-O>", lambda e: self.add_files() if self.mode_var.get() == "mode1"
                      else self.choose_mode2_file())
        self.bind_all("<Return>", self._shortcut_extract)
        self.bind_all("<Control-s>", self._shortcut_export)
        self.bind_all("<Control-S>", self._shortcut_export)

    def _shortcut_extract(self, event=None):
        if str(self.run_button["state"]) != "disabled":
            self.start_extract()

    def _shortcut_export(self, event=None):
        if str(self.shared_export_button["state"]) != "disabled":
            self.export_result()

    # ------------------------------------------------------------------
    # CHE DO 1: FILE NGUON (nhieu tep)
    # ------------------------------------------------------------------
    def add_files(self):
        paths = filedialog.askopenfilenames(title="Chọn tệp nguồn", filetypes=MODE1_INPUT_FILETYPES)
        for p in paths:
            if p not in self.input_paths:
                self.input_paths.append(p)
                icon = FILE_ICONS.get(os.path.splitext(p)[1].lower(), "\U0001F4C4")
                self.listbox.insert(tk.END, f"{icon}  {os.path.basename(p)}")
        self._update_file_count()

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.input_paths[idx]
        self._update_file_count()

    def clear_files(self):
        self.listbox.delete(0, tk.END)
        self.input_paths = []
        self._update_file_count()

    def _update_file_count(self):
        n = len(self.input_paths)
        if n == 0:
            self.file_count_var.set("Chưa chọn tệp nào.")
        elif n == 1:
            self.file_count_var.set("Đã chọn 1 tệp.")
        else:
            self.file_count_var.set(f"Đã chọn {n} tệp.")

    # ------------------------------------------------------------------
    # CHE DO 2: FILE NGUON (1 tep)
    # ------------------------------------------------------------------
    def choose_mode2_file(self):
        path = filedialog.askopenfilename(title="Chọn tệp danh sách đăng ký",
                                           filetypes=MODE2_INPUT_FILETYPES)
        if path:
            self.mode2_input_path.set(path)

    # ------------------------------------------------------------------
    def choose_template(self):
        path = filedialog.askopenfilename(title="Chọn file mẫu Excel", filetypes=EXCEL_FILETYPES)
        if path:
            self.template_var.set(path)

    # ------------------------------------------------------------------
    # LOG
    # ------------------------------------------------------------------
    def log(self, message):
        key = _search_key(message)
        tag = None
        if "loi" in key or "!!" in message:
            tag = "err"
        elif "da ghi" in key or "->" in message or "hoan tat" in key:
            tag = "ok"
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")

    # ------------------------------------------------------------------
    # PIPELINE STATS - CHE DO 1
    # ------------------------------------------------------------------
    def _mode1_pipeline_segments(self, stats):
        return [
            (f"{_fmt(stats['total_rows'])} dòng đọc được", "neutral"),
            ("  \u2192  ", "arrow"),
            (f"-{_fmt(stats['missing_email'])} thiếu email", "warn"),
            ("  \u2192  ", "arrow"),
            (f"-{_fmt(stats['invalid_email_format'])} sai định dạng", "warn"),
            ("  \u2192  ", "arrow"),
            (f"-{_fmt(stats['duplicates_removed'])} trùng email", "danger"),
            ("  \u2192  ", "arrow"),
            (f"{_fmt(stats['final_count'])} dòng kết quả", "final"),
        ]

    def _mode2_pipeline_segments(self, stats):
        missing_invalid = sum(1 for i in stats["issues"]
                               if i["reason"] in ("missing_email", "invalid_email_format"))
        return [
            (f"{_fmt(stats['total_read'])} dòng đọc được", "neutral"),
            ("  \u2192  ", "arrow"),
            (f"-{_fmt(missing_invalid)} thiếu/sai email", "warn"),
            ("  \u2192  ", "arrow"),
            (f"-{_fmt(stats['duplicates_removed'])} trùng email", "danger"),
            ("  \u2192  ", "arrow"),
            (f"{_fmt(stats['final_count'])} dòng kết quả", "final"),
        ]

    # ------------------------------------------------------------------
    # BUOC 1: TRICH XUAT
    # ------------------------------------------------------------------
    def start_extract(self):
        mode = self.mode_var.get()
        if mode == "mode1":
            if not self.input_paths:
                messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn ít nhất 1 tệp nguồn.")
                return
        else:
            if not self.mode2_input_path.get().strip():
                messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn tệp danh sách đăng ký.")
                return

        self.run_button.config(state="disabled")
        self.status_var.set("Đang xử lý...")
        self.progress.start(12)
        self.clear_log()
        self.active_panel.reset()

        if mode == "mode1":
            thread = threading.Thread(
                target=self._extract_worker_mode1,
                args=(list(self.input_paths), self.dedupe_var.get()), daemon=True,
            )
        else:
            thread = threading.Thread(
                target=self._extract_worker_mode2,
                args=(self.mode2_input_path.get().strip(), self.to_chuc_var.get().strip(), self.dedupe_var.get()),
                daemon=True,
            )
        thread.start()

    def _extract_worker_mode1(self, inputs, dedupe):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = _StreamToLog(self)
        try:
            records, stats = core.extract_all(inputs, dedupe=dedupe, verbose=True)
            self.after(0, self._on_extract_success_mode1, records, stats)
        except Exception as e:
            self.after(0, self._on_extract_error, str(e))
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr

    def _extract_worker_mode2(self, input_path, to_chuc, dedupe):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = _StreamToLog(self)
        try:
            records, stats = v2.extract_v2(input_path, default_to_chuc=to_chuc or None,
                                            dedupe=dedupe, verbose=True)
            self.after(0, self._on_extract_success_mode2, records, stats)
        except Exception as e:
            self.after(0, self._on_extract_error, str(e))
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr

    def _on_extract_success_mode1(self, records, stats):
        self.progress.stop()
        self.run_button.config(state="normal")
        self.status_var.set("Hoàn tất trích xuất.")
        self.panel_m1.set_stats_pipeline(self._mode1_pipeline_segments(stats))
        self.panel_m1.load_results(records, stats)
        if not records:
            messagebox.showinfo("Không có dữ liệu hợp lệ",
                                 "Không trích xuất được bản ghi nào có email hợp lệ từ các tệp đã chọn.\n"
                                 'Xem tab "Cần kiểm tra" để biết chi tiết từng dòng bị loại.')

    def _on_extract_success_mode2(self, records, stats):
        self.progress.stop()
        self.run_button.config(state="normal")
        self.status_var.set("Hoàn tất trích xuất.")
        self.panel_m2.set_stats_pipeline(self._mode2_pipeline_segments(stats))
        self.panel_m2.load_results(records, stats)
        if stats.get("to_chuc_missing"):
            self.log(f"CẢNH BÁO: {stats['to_chuc_missing']} dòng không xác định được 'Tổ chức' "
                     f"(để trống) - kiểm tra lại ô 'Tổ chức mặc định' hoặc dữ liệu nguồn.")
        if not records:
            messagebox.showinfo("Không có dữ liệu hợp lệ",
                                 "Không trích xuất được bản ghi nào có email hợp lệ từ tệp đã chọn.\n"
                                 'Xem tab "Cần kiểm tra" để biết chi tiết từng dòng bị loại.')

    def _on_extract_error(self, error_message):
        self.progress.stop()
        self.run_button.config(state="normal")
        self.status_var.set("Có lỗi xảy ra.")
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi trong quá trình trích xuất:\n{error_message}")

    # ------------------------------------------------------------------
    # BUOC 2: XUAT TEP KET QUA
    # ------------------------------------------------------------------
    def export_result(self):
        panel = self.active_panel
        if not panel.records:
            messagebox.showwarning("Chưa có dữ liệu", 'Vui lòng bấm "Trích xuất & Xem trước" trước.')
            return

        template_path = self.template_var.get().strip()
        if not template_path or not os.path.exists(template_path):
            messagebox.showwarning("Thiếu file mẫu", "Vui lòng chọn file mẫu Excel hợp lệ.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Đặt tên và chọn nơi lưu tệp kết quả", defaultextension=".xlsx",
            filetypes=EXCEL_FILETYPES, initialfile="ket_qua.xlsx",
        )
        if not output_path:
            return

        self.shared_export_button.config(state="disabled")
        self.status_var.set("Đang xuất tệp...")
        self.progress.start(12)

        mode = self.mode_var.get()
        if mode == "mode1":
            password = self.password_var.get() or core.DEFAULT_PASSWORD
            thread = threading.Thread(
                target=self._export_worker_mode1,
                args=(panel.records, template_path, output_path, password), daemon=True,
            )
        else:
            thread = threading.Thread(
                target=self._export_worker_mode2,
                args=(panel, panel.records, panel.stats.get("issues", []), template_path, output_path),
                daemon=True,
            )
        thread.start()

    def _export_worker_mode1(self, records, template_path, output_path, password):
        try:
            core.write_output(records, template_path, output_path, password=password)
            self.after(0, self._on_export_success, len(records), output_path)
        except Exception as e:
            self.after(0, self._on_export_error, str(e))

    def _export_worker_mode2(self, panel, records, issues, template_path, output_path):
        try:
            v2.write_v2_output(records, template_path, output_path)
            issues_path = None
            if issues:
                base, _ext = os.path.splitext(output_path)
                issues_path = f"{base}_can_kiem_tra.xlsx"
                v2.write_v2_issues(issues, issues_path)
            self.after(0, self._on_export_success, len(records), output_path, issues_path)
        except Exception as e:
            self.after(0, self._on_export_error, str(e))

    def _on_export_success(self, count, output_path, issues_path=None):
        self.progress.stop()
        self.shared_export_button.config(state="normal")
        self.status_var.set("Đã xuất tệp thành công.")
        msg = f"Đã xuất {count} bản ghi ra tệp:\n{output_path}"
        if issues_path:
            msg += f"\n\nĐã ghi kèm báo cáo các dòng bị loại:\n{issues_path}"
        messagebox.showinfo("Hoàn tất", msg)

    def _on_export_error(self, error_message):
        self.progress.stop()
        self.shared_export_button.config(state="normal")
        self.status_var.set("Có lỗi xảy ra.")
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi xuất tệp:\n{error_message}")


class _StreamToLog:
    def __init__(self, gui):
        self.gui = gui

    def write(self, text):
        text = text.rstrip("\n")
        if text:
            self.gui.after(0, self.gui.log, text)

    def flush(self):
        pass


def main():
    app = ContactExtractorGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
