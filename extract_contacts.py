#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_contacts.py
====================
Trich xuat du lieu Email / Ho va ten / So dien thoai tu cac file nguon
(PDF, Word .docx, Excel .xlsx/.xls, CSV) va xuat ket qua ra file Excel
theo dung template co san (giu nguyen tieu de, dinh dang cot).

QUY TAC XU LY (theo yeu cau):
  1) Email:
       - Chuan hoa ve dang hop le, viet thuong toan bo.
       - Neu ban ghi KHONG co email -> BO QUA, khong dua vao file ket qua.
  2) So dien thoai:
       - Chuan ve dang so Viet Nam: 10 so, bat dau bang so 0,
         khong ky tu dac biet, viet lien khong khoang trang.
       - Tu dong xu ly cac dang nhap: +84..., 84..., co dau cach/gach ngang/dau cham...
       - Neu khong the chuan hoa duoc thanh so hop le -> de trong o do (van giu ban ghi,
         vi chi co quy tac bo qua danh cho truong email).
  3) Ho va ten:
       - Viet hoa chu cai dau moi tu.
       - Loai bo ky tu dac biet (chi giu chu cai - ke ca co dau tieng Viet - va khoang trang).
  4) Loc bo cac truong "Chuc vu", "Don vi" neu co trong du lieu nguon (khong
     dua 2 truong nay vao file ket qua, va khong de chung lam nhieu du lieu
     Ho ten khi trich xuat tu van ban tu do).

THONG KE (tra ve tu ham extract_all() / run()):
  - total_extracted     : tong so dong du lieu hop le (co email) doc duoc, TRUOC khi loc trung
  - duplicates_removed  : so dong bi trung email da bi loai bo
  - final_count         : so dong con lai sau khi loc trung (= total_extracted - duplicates_removed)

CACH DUNG (dong lenh):
  python3 extract_contacts.py \
      --input file1.pdf file2.docx file3.xlsx \
      --template cls_template_users.xlsx \
      --output ket_qua.xlsx

  --input     : mot hoac nhieu file nguon (PDF / DOCX / XLSX / XLS / CSV)
  --template  : file Excel mau (template) can theo dung dinh dang cot
  --output    : duong dan file Excel ket qua se duoc tao ra

Co the import va goi ham `run(inputs, template_path, output_path)` truc tiep
tu code Python khac neu khong muon dung dong lenh, hoac `extract_all(inputs)`
neu chi can trich xuat + thong ke ma chua muon ghi file (dung cho xem truoc).
"""

import argparse
import csv
import os
import re
import sys
import unicodedata
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. REGEX / TU KHOA NHAN DIEN
# ---------------------------------------------------------------------------

EMAIL_REGEX = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# Chap nhan cac dang so co the co +84 / 84 / 0. Dung ANCHOR chinh xac (khong
# phai khoang mo {8,10}) de KHONG "an lan" sang so dien thoai KE BEN khi 1 o /
# 1 dong chua NHIEU so cach nhau boi dau cach (vd "0912345678 0987654321"):
#   - (?<!\d)      : khong duoc dung ngay sau 1 chu so khac (tranh bat giua so)
#   - prefix        : +84 / 84 / 0
#   - dung 9 chu so tiep theo (moi so co the co 0-1 ky tu phan cach truoc no:
#     khoang trang, dau cham, dau gach ngang, ngoac don)
#   - (?!\d)        : ky tu ngay sau KHONG duoc la chu so, tranh bat lan sang
#     chu so dau tien cua mot so khac ngay ke ben
PHONE_CANDIDATE_REGEX = re.compile(
    r"(?<!\d)(?:\+84|84|0)(?:[\s.\-()]?\d){9}(?!\d)"
)

# Nhan dien nhan truong trong van ban tu do (PDF / DOCX dang van ban).
# LUU Y: cac pattern nay duoc ap dung tren BAN DA BO DAU (xem
# `_deaccent_keep_len`), vi van ban goc thuong co dau tieng Viet (vd
# "Họ và tên:", "Điện thoại:") ma pattern ASCII thuan se khong khop truc tiep.
LABEL_PATTERNS = {
    "email": re.compile(r"(?:email|e-mail|mail)\s*[:\-]\s*(.+)", re.IGNORECASE),
    "phone": re.compile(
        r"(?:so\s*dien\s*thoai|sdt|dien\s*thoai|phone|hotline|tel)\s*[:\-]\s*(.+)",
        re.IGNORECASE,
    ),
    "name": re.compile(
        r"(?:ho\s*va\s*ten|ho\s*ten|ho\s*va\s*ten\s*dem|full\s*name|ten\s*(?:tai\s*khoan)?|name)\s*[:\-]\s*(.+)",
        re.IGNORECASE,
    ),
}

# Danh sach tu khoa cua cac truong CAN LOC BO khoi ket qua (Chuc vu / Don vi
# va cac bien the thuong gap). Dung CHUNG cho ca 3 cho: header cot dang bang,
# dong van ban tu do, va PHAN CAT bo phan "dinh kem" trong CUNG 1 chuoi ten
# (vd 1 o Excel/PDF ghi gop "Nguyen Van A - Truong phong - Phong Ky thuat").
NOISE_KEYWORDS = [
    "chuc vu", "chuc danh", "vi tri cong tac", "vi tri",
    "don vi cong tac", "don vi", "phong ban", "bo phan",
    "co quan", "noi cong tac", "chi nhanh",
]

# Khop bat ky tu khoa nao o tren, XUAT HIEN O DAU TRONG CHUOI/DONG (khong bat
# buoc co dau ":" hay "-" ngay sau, vi nhieu file khong co dau cau ro rang).
# Sap xep tu khoa dai truoc de uu tien khop cum dai hon (vd "vi tri cong tac"
# truoc "vi tri").
NOISE_INLINE_REGEX = re.compile(
    "|".join(re.escape(kw) for kw in sorted(NOISE_KEYWORDS, key=len, reverse=True))
)

# Nhan cua cac truong CAN LOC BO, dang "nhan: gia tri" trong van ban tu do
# (dung de loai dong nay hoan toan khoi vong lap doan van khi suy doan ten).
NOISE_LABEL_REGEX = NOISE_INLINE_REGEX

# Tu khoa nhan dien header cot khi doc du lieu dang bang (Excel/CSV/table Word/PDF)
HEADER_KEYWORDS = {
    "email": [
        "email", "e-mail", "mail",
        "thu dien tu", "hop thu dien tu", "dia chi email", "dia chi mail",
    ],
    "phone": ["dien thoai", "sdt", "so dien thoai", "phone", "hotline", "tel"],
    "name": ["ho va ten", "ho ten", "ten tai khoan", "ten", "full name", "name", "hoten"],
    "password": ["mat khau", "password", "pass", "mk"],
}

# Tu khoa header can LOAI BO hoan toan (khong dua vao file ket qua), du co
# the trung mot phan voi tu khoa o tren. Duoc kiem tra TRUOC HEADER_KEYWORDS.
NOISE_HEADER_KEYWORDS = NOISE_KEYWORDS

# Mat khau mac dinh dien vao cot "Mat khau" cua file ket qua (neu template co cot nay)
DEFAULT_PASSWORD = "Copenai@2026"


def _strip_accents_lower(text: str) -> str:
    """Bo dau + ve chu thuong, dung de so khop tu khoa header khong phan biet dau."""
    if text is None:
        return ""
    text = str(text).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.lower().strip()


def _deaccent_keep_len(text: str) -> str:
    """
    Bo dau tieng Viet + ve chu thuong nhung GIU NGUYEN DO DAI CHUOI (moi ky
    tu goc -> dung 1 ky tu ket qua). Nho vay vi tri (index) tim duoc tren
    chuoi da bo dau van dung de cat chuoi GOC (con dau), phuc vu viec trich
    xuat gia tri (vd ten nguoi) khong bi mat dau khi tim theo nhan khong dau.
    """
    if text is None:
        return ""
    out_chars = []
    for ch in str(text):
        base = ch.replace("đ", "d").replace("Đ", "D")
        nfd = unicodedata.normalize("NFD", base)
        stripped = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
        out_chars.append((stripped or base)[:1].lower() if stripped else base.lower())
    return "".join(out_chars)


def is_noise_line(line: str) -> bool:
    """True neu dong van ban thuoc truong 'Chuc vu' / 'Don vi' / bien the
    (can loc bo), du co dau ':' '-' hay khong."""
    return bool(NOISE_INLINE_REGEX.search(_deaccent_keep_len(line)))


def strip_inline_noise_from_name(raw_name: str) -> str:
    """
    Xu ly truong hop 1 CHUOI TEN bi ghep chung voi Chuc vu/Don vi (thuong gap
    khi PDF/Excel gop nhieu cot vao 1 o, vd:
    "Nguyen Van A - Truong phong - Phong Ky thuat"). Cat bo tu vi tri xuat
    hien tu khoa nhieu dau tien, chi giu lai phan TEN o truoc do.
    """
    if not raw_name:
        return raw_name
    deaccented = _deaccent_keep_len(raw_name)
    match = NOISE_INLINE_REGEX.search(deaccented)
    if not match:
        return raw_name
    if match.start() == 0:
        # Toan bo chuoi la Chuc vu/Don vi, khong co phan ten hop le phia truoc
        return ""
    return raw_name[: match.start()].strip(" \t-|/,;:")


def match_header_field(header_text: str):
    """Tra ve 'email' / 'phone' / 'name' / 'password' neu header khop tu
    khoa; tra ve None neu khong khop HOAC khop tu khoa "nhieu" (Chuc vu,
    Don vi...) can loai bo khoi ket qua."""
    norm = _strip_accents_lower(header_text)
    if not norm:
        return None
    for kw in NOISE_HEADER_KEYWORDS:
        if kw in norm:
            return None  # cot nhieu (Chuc vu / Don vi...) -> khong lay
    for field, keywords in HEADER_KEYWORDS.items():
        for kw in keywords:
            if kw in norm:
                return field
    return None


# ---------------------------------------------------------------------------
# 2. CHUAN HOA DU LIEU (3 QUY TAC)
# ---------------------------------------------------------------------------

def normalize_email(raw):
    """Tra ve email hop le, viet thuong; None neu khong hop le / rong.
    Bo TOAN BO khoang trang/xuong dong truoc khi tim, vi PDF/Word co the
    ngat dong GIUA 1 dia chi email dai (vd trong 1 o bang bi wrap chu),
    lam email bi tach thanh 2 dong va regex email se bat thieu neu khong
    noi lai truoc."""
    if not raw:
        return None
    raw = re.sub(r"\s+", "", str(raw))
    match = EMAIL_REGEX.search(raw)
    if not match:
        return None
    return match.group(0).lower()


def normalize_phone(raw):
    """
    Chuan hoa ve dang so dien thoai Viet Nam: 10 so, bat dau bang '0',
    khong ky tu dac biet, khong khoang trang. Tra ve '' neu khong chuan hoa duoc.

    QUAN TRONG: neu o du lieu chua NHIEU so dien thoai (vd "0912345678 /
    0987654321") hoac co chu/ky hieu xen giua, ham se TIM va lay SO DAU TIEN
    khop dinh dang hop le (qua PHONE_CANDIDATE_REGEX) truoc khi lam sach,
    thay vi gop toan bo ky tu con lai thanh 1 chuoi so dai sai (nguyen nhan
    gay mat/sai du lieu o cac phien ban truoc).
    """
    if raw is None or raw == "":
        return ""

    # Excel doc so nguyen co the tra ve float (vd 912345678.0) -> quy ve int
    if isinstance(raw, float):
        raw = str(int(raw)) if raw.is_integer() else str(raw)
    else:
        raw = str(raw)
    raw = raw.strip()
    if not raw:
        return ""

    def _finalize(cleaned: str):
        if cleaned.startswith("+84"):
            digits = "0" + cleaned[3:]
        elif cleaned.startswith("84") and len(cleaned) >= 10:
            digits = "0" + cleaned[2:]
        elif cleaned.startswith("0"):
            digits = cleaned
        else:
            # Khong co ma vung ro rang; neu du 9 so (thieu so 0 dau) thi them vao
            digits = "0" + cleaned if len(cleaned) == 9 else cleaned
        digits = re.sub(r"\D", "", digits)
        if len(digits) == 10 and digits.startswith("0"):
            return digits
        return None

    # Buoc 1 (uu tien): tim so hop le DAU TIEN bang regex co "neo" chinh xac,
    # tranh bi "an lan" tu so ben canh khi 1 o co nhieu so.
    candidate_match = PHONE_CANDIDATE_REGEX.search(raw)
    if candidate_match:
        cleaned = re.sub(r"[^\d+]", "", candidate_match.group(0))
        result = _finalize(cleaned)
        if result:
            return result

    # Buoc 2 (du phong): khong tim thay theo regex co neo (vd thieu so 0 dau
    # nen khong khop prefix) -> thu lam sach TOAN BO chuoi nhu truoc.
    cleaned = re.sub(r"[^\d+]", "", raw)
    if not cleaned:
        return ""
    result = _finalize(cleaned)
    return result or ""


def normalize_name(raw):
    """Viet hoa chu cai dau moi tu, bo ky tu dac biet, giu chu + khoang trang.
    Cung tu dong cat bo phan Chuc vu/Don vi neu bi ghep chung trong CUNG 1
    chuoi ten (vd "Nguyen Van A - Truong phong - Phong Ky thuat")."""
    if not raw:
        return ""
    raw = str(raw).strip()
    raw = strip_inline_noise_from_name(raw)
    if not raw:
        return ""
    # Bo ky tu dac biet: chi giu chu cai (Unicode, gom co dau tieng Viet) va khoang trang
    cleaned = "".join(ch for ch in raw if ch.isalpha() or ch.isspace())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return ""
    words = cleaned.split(" ")
    capitalized = [w[0].upper() + w[1:].lower() if w else w for w in words]
    return " ".join(capitalized)


# Nhan hien thi (co dau, de doc) cho tung ly do bi loai - dung chung cho CLI
# va giao dien (tab "Can kiem tra").
ISSUE_REASON_LABELS = {
    "missing_email": "Thiếu email",
    "invalid_email_format": "Email sai định dạng",
    "duplicate_email": "Trùng email",
}


def build_clean_record(raw_email, raw_name, raw_phone, row_stats=None, row_context=None):
    """Ap dung 3 quy tac; tra ve dict hoac None neu bi bo qua (thieu/sai email).
    Neu duoc truyen row_stats (dict dem so lieu, dung cho thong ke), se cong
    don: tong so dong da xu ly, so dong thieu email, so dong email sai dinh
    dang, so dong hop le. Neu duoc truyen row_context (dict {'source':...,
    'row_number':...} de biet dong nay o dau), khi bi loai se GHI LAI vao
    row_stats['issues'] du lieu goc + ly do, phuc vu tab "Can kiem tra" tren
    giao dien (nguoi dung can biet CHINH XAC dong nao, o tep nao, vi sao bi
    loai - chi bao 1 con so tong khong du de tu sua du lieu nguon)."""
    if row_stats is not None:
        row_stats["total_rows"] = row_stats.get("total_rows", 0) + 1

    raw_email_str = str(raw_email).strip() if raw_email else ""
    raw_name_str = str(raw_name).strip() if raw_name else ""
    raw_phone_str = str(raw_phone).strip() if raw_phone else ""
    email = normalize_email(raw_email)

    if not email:
        reason = "invalid_email_format" if raw_email_str else "missing_email"
        if row_stats is not None:
            row_stats[reason] = row_stats.get(reason, 0) + 1
            if row_context is not None:
                row_stats.setdefault("issues", []).append({
                    "source": row_context.get("source", ""),
                    "row_number": row_context.get("row_number"),
                    "raw_email": raw_email_str,
                    "raw_name": raw_name_str,
                    "raw_phone": raw_phone_str,
                    "reason": reason,
                })
        return None  # Quy tac 1: khong co email hop le -> bo qua ban ghi

    if row_stats is not None:
        row_stats["valid_rows"] = row_stats.get("valid_rows", 0) + 1

    record = {
        "email": email,
        "name": normalize_name(raw_name),
        "phone": normalize_phone(raw_phone),
    }
    if row_context is not None:
        record["_context"] = dict(row_context)
    return record


# ---------------------------------------------------------------------------
# 3. DOC DU LIEU DANG BANG (Excel / CSV / bang trong Word / PDF)
# ---------------------------------------------------------------------------

def merge_wrapped_continuation_rows(rows, name_col_idx):
    """
    Xu ly truong hop 1 O BI NGAT DONG giua chung (thuong gap o cot Email/
    Chuc vu khi noi dung dai, hoac khi dong roi vao ranh gioi trang PDF),
    khien pdfplumber tach thanh 1 DONG BANG RIENG chi co du lieu o 1-2 cot
    (vd phan con lai cua 1 email bi cat: "...@" roi "dongthap.gov.vn" o
    dong ke tiep). Neu khong xu ly: nguoi TRUOC do bi email SAI/CUT, va
    phan bi ngat co the bi hieu nham la 1 NGUOI MOI (ten rong, thieu du
    lieu that su).

    Quy tac: MOI NGUOI THAT SU DEU CO TEN. Vi vay 1 dong duoc coi la PHAN
    TIEP CUA DONG TRUOC neu cot TEN cua no RONG (nhung dong đó vẫn co it
    nhat 1 o khac co du lieu). Khi do, NOI TRUC TIEP (khong them khoang
    trang - vi thuong la ngat GIUA 1 tu/dia chi) noi dung cac o cua dong
    nay vao CUOI o tuong ung cua dong TRUOC, roi loai hoan toan dong nay.
    """
    if not rows or name_col_idx is None:
        return rows

    merged = []
    for row in rows:
        row = list(row)
        name_val = row[name_col_idx] if name_col_idx < len(row) else None
        name_is_empty = not (name_val and str(name_val).strip())
        has_any_data = any((c and str(c).strip()) for c in row)

        if merged and name_is_empty and has_any_data:
            prev = merged[-1]
            for idx in range(min(len(prev), len(row))):
                cell = row[idx]
                if cell and str(cell).strip():
                    prev[idx] = (str(prev[idx]).strip() if prev[idx] else "") + str(cell).strip()
        elif has_any_data or not merged:
            merged.append(row)
        # dong hoan toan rong (khong co du lieu o bat ky cot nao) -> bo qua

    return merged


def find_header_row(rows, max_scan=15):
    """
    Tim dong THUC SU la tieu de cot trong 15 dong dau tien, thay vi mac dinh
    dong 0 la tieu de. CAN THIET vi nhieu file (dac biet PDF cong van) co
    VAI DONG TIEU DE VAN BAN (ten van ban, so cong van, dong trong...) NAM
    TRUOC dong tieu de cot thuc su - neu cu dung dong 0 se khong nhan dien
    duoc cot nao (dong tieu de van ban khong khop tu khoa nao) va MAT TRANG
    toan bo du lieu phia sau.

    Tra ve (index, col_map) cua dong dau tien (trong 15 dong quet) ma tu do
    xac dinh duoc CA cot email. Neu khong tim thay, tra ve (None, None).
    """
    for i, row in enumerate(rows[:max_scan]):
        col_map = {}
        for idx, cell in enumerate(row):
            field = match_header_field(cell)
            if field and field not in col_map:
                col_map[field] = idx
        if "email" in col_map:
            return i, col_map
    return None, None


def extract_from_table_rows(rows, row_stats=None, source_label=""):
    """
    rows: list cac list o. Tu dong TIM dong tieu de that su (xem
    find_header_row - co the khong phai dong 0 neu co vai dong tieu de van
    ban phia truoc) roi xac dinh cot nao la email / ten / dien thoai dua
    tren HEADER_KEYWORDS. Neu khong tim thay dong tieu de phu hop trong 15
    dong dau, tra ve [].

    source_label: nhan mo ta nguon du lieu (vd "ten_file.xlsx" hoac
    "ten_file.xlsx [Sheet: Thang1]") - duoc ghi kem theo moi dong bi loai
    de nguoi dung biet CHINH XAC no thuoc tep/sheet nao (xem tab "Can kiem
    tra" tren giao dien).
    """
    if not rows:
        return []

    header_idx, col_map = find_header_row(rows)
    if header_idx is None:
        return []  # khong nhan dien duoc cot email trong pham vi quet -> bo qua

    data_rows = rows[header_idx + 1:]
    # Gop cac dong bi "gay" do PDF ngat dong giua o (xem
    # merge_wrapped_continuation_rows) TRUOC KHI xu ly tung dong, de khong
    # mat/sai du lieu (vd email bi cat lam doi) va khong tao ra "nguoi ao"
    # tu phan bi ngat.
    data_rows = merge_wrapped_continuation_rows(data_rows, col_map.get("name"))

    records = []
    row_number = 0
    for row in data_rows:
        if row is None or all(c in (None, "") for c in row):
            continue
        row_number += 1

        def get(field):
            idx = col_map.get(field)
            if idx is None or idx >= len(row):
                return ""
            return row[idx] if row[idx] is not None else ""

        row_context = {"source": source_label, "row_number": row_number}
        rec = build_clean_record(get("email"), get("name"), get("phone"),
                                  row_stats=row_stats, row_context=row_context)
        if rec:
            records.append(rec)
    return records


def _row_matches_known_header(row, header_row):
    """
    So sanh 1 dong voi dong TIEU DE DA BIET (cua bang truoc do) de phat hien
    truong hop bang tiep theo LAP LAI dong tieu de (thay vi doan luon vao du
    lieu). CHI dung de phat hien header LAP LAI (giong het/gan giong header
    cu) - KHONG dung cach do "co chua tu khoa header hay khong" nhu truoc,
    vi cach do de bi NHAM khi 1 O DU LIEU (vd dia chi email that su chua
    chu "mail" trong "gmail.com") vo tinh khop tu khoa header, khien dong
    DU LIEU dau tien cua bang sau bi hieu nham la "tieu de moi" va lam mat
    ca cot Ten/SDT cua toan bo phan con lai.
    """
    if len(row) != len(header_row):
        return False
    matches = 0
    for a, b in zip(row, header_row):
        if _strip_accents_lower(a) == _strip_accents_lower(b):
            matches += 1
    # Da so cac cot giong het header cu -> chac chan la dong tieu de lap lai
    return matches >= max(1, (len(row) + 1) // 2)


def merge_multi_page_tables(tables):
    """
    Gop cac bang lien tiep (PDF nhieu trang hoac Word bi tach thanh nhieu
    bang) co CUNG SO COT thanh 1 bang logic duy nhat. Neu khong gop, dong DU
    LIEU dau tien cua bang tiep theo se bi nham thanh "header" -> mat du
    lieu cua ca bang do.

    Quy tac:
      - Neu bang tiep theo co dong dau LAP LAI dong tieu de cua bang truoc
        (giong het/gan giong) -> bo dong tieu de lap do, noi phan CON LAI
        lam du lieu.
      - Neu KHONG giong header cu (tuc la dong dau da la DU LIEU) -> noi
        TOAN BO cac dong (ke ca dong dau) lam du lieu.
      - Neu so cot khac -> coi la bang logic MOI, dong dau cua no la header
        moi.
    """
    if not tables:
        return []

    logical_tables = []
    current = None
    current_header = None

    for table in tables:
        if not table:
            continue
        if current is None:
            current = [list(r) for r in table]
            # Lay dong tieu de THUC SU (co the khong phai dong 0, vd co vai
            # dong tieu de van ban phia truoc) de so sanh chinh xac voi cac
            # bang tiep theo khi phat hien header lap lai.
            idx_found, _ = find_header_row(table)
            current_header = table[idx_found] if idx_found is not None else table[0]
            continue

        same_width = len(table[0]) == len(current[0])
        if not same_width:
            logical_tables.append(current)
            current = [list(r) for r in table]
            idx_found, _ = find_header_row(table)
            current_header = table[idx_found] if idx_found is not None else table[0]
            continue

        if current_header is not None and _row_matches_known_header(table[0], current_header):
            current.extend(list(r) for r in table[1:])  # bo dong tieu de lap lai
        else:
            current.extend(list(r) for r in table)  # toan bo la du lieu

    if current is not None:
        logical_tables.append(current)

    return logical_tables


_XLRD_PATCHED = False


def _patch_xlrd_tolerant_datemode():
    """
    Va 1 loi rat hay gap trong xlrd voi cac file .xls da qua nhieu lan chinh
    sua / co nhung macro (VBA project): ban ghi DATEMODE trong file bi thieu
    du lieu (0 byte thay vi 2 byte), khien xlrd.book.Book.handle_datemode()
    crash voi struct.error ngay tu buoc doc file, du toan bo du lieu con lai
    van hoan toan doc duoc binh thuong. Ban vay chi don gian: neu thieu du
    lieu, mac dinh datemode=0 (he ngay 1900 - pho bien nhat) thay vi crash,
    giup doc duoc file ma KHONG CAN cai them phan mem nao (vd LibreOffice).
    Chi vay 1 lan duy nhat (idempotent).
    """
    global _XLRD_PATCHED
    if _XLRD_PATCHED:
        return
    import xlrd.book as _xlrd_book

    _original_handle_datemode = _xlrd_book.Book.handle_datemode

    def _tolerant_handle_datemode(self, data):
        if len(data) < 2:
            self.datemode = 0
            return
        return _original_handle_datemode(self, data)

    _xlrd_book.Book.handle_datemode = _tolerant_handle_datemode
    _XLRD_PATCHED = True


def _rows_from_xlrd_sheet(sheet):
    """Chuyen 1 sheet xlrd thanh list-of-lists giong dinh dang openpyxl dung,
    dam bao gia tri kieu float/int duoc giu nguyen (khong ep ve string) de
    cac ham chuan hoa (normalize_phone...) xu ly dung nhu voi openpyxl."""
    rows = []
    for r in range(sheet.nrows):
        rows.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
    return rows


def read_xls_file(path, row_stats=None, source_label=None):
    """
    Doc file Excel 97-2003 (.xls - dinh dang nhi phan cu, KHAC voi .xlsx).
    openpyxl KHONG doc duoc dinh dang nay, nen dung xlrd (co va loi datemode
    thuong gap - xem _patch_xlrd_tolerant_datemode). Neu van khong doc duoc
    (truong hop file hong nang hon), tu dong du phong sang LibreOffice NEU
    may co cai san; khong bat buoc phai co LibreOffice cho truong hop thong
    thuong.
    """
    import xlrd

    _patch_xlrd_tolerant_datemode()
    base_label = source_label or os.path.basename(path)

    all_records = []
    try:
        wb = xlrd.open_workbook(path)
        for sheet in wb.sheets():
            rows = _rows_from_xlrd_sheet(sheet)
            while rows and all(c in (None, "") for c in rows[0]):
                rows.pop(0)
            sheet_label = f"{base_label} [Sheet: {sheet.name}]"
            all_records.extend(extract_from_table_rows(rows, row_stats=row_stats, source_label=sheet_label))
        return all_records
    except Exception as xlrd_error:
        converted_path = _convert_xls_via_libreoffice(path)
        if converted_path is None:
            raise ValueError(
                f"Không thể đọc tệp .xls này bằng xlrd ({xlrd_error}), và không tìm thấy "
                "LibreOffice (soffice) trên máy để dự phòng. Cách khắc phục: mở tệp bằng "
                "Excel và 'Save As' sang định dạng .xlsx rồi thử lại, hoặc cài LibreOffice "
                "(https://www.libreoffice.org/) để tool tự động chuyển đổi."
            ) from xlrd_error
        try:
            return read_excel_file(converted_path, row_stats=row_stats, source_label=base_label)
        finally:
            try:
                os.remove(converted_path)
                os.rmdir(os.path.dirname(converted_path))
            except OSError:
                pass


def _convert_via_libreoffice(path, target_format):
    """Goi LibreOffice (soffice) o che do nen de chuyen 1 file sang dinh
    dang `target_format` (vd 'xlsx', 'docx') trong 1 thu muc tam. Tra ve
    duong dan file da chuyen, hoac None neu khong tim thay LibreOffice /
    chuyen doi that bai. Dung chung cho ca .xls->.xlsx va .doc->.docx (2
    dinh dang Office cu deu can LibreOffice vi thu vien Python thuan
    (openpyxl/python-docx) chi doc duoc dinh dang moi hon)."""
    import shutil
    import subprocess
    import tempfile

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None

    out_dir = tempfile.mkdtemp(prefix="office_convert_")
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", target_format, "--outdir", out_dir, path],
            check=True, capture_output=True, timeout=120,
        )
    except Exception:
        try:
            os.rmdir(out_dir)
        except OSError:
            pass
        return None

    base_name = os.path.splitext(os.path.basename(path))[0]
    converted = os.path.join(out_dir, base_name + "." + target_format)
    return converted if os.path.exists(converted) else None


def _convert_xls_via_libreoffice(path):
    """Rieng cho .xls->.xlsx - xem _convert_via_libreoffice."""
    return _convert_via_libreoffice(path, "xlsx")


def read_excel_file(path, row_stats=None, source_label=None):
    base_label = source_label or os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    all_records = []
    for ws in wb.worksheets:
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
        # Bo cac dong trong hoan toan o dau
        while rows and all(c in (None, "") for c in rows[0]):
            rows.pop(0)
        sheet_label = f"{base_label} [Sheet: {ws.title}]" if len(wb.worksheets) > 1 else base_label
        all_records.extend(extract_from_table_rows(rows, row_stats=row_stats, source_label=sheet_label))
    return all_records


def read_csv_file(path, row_stats=None, source_label=None):
    label = source_label or os.path.basename(path)
    with open(path, newline="", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        rows = list(reader)
    return extract_from_table_rows(rows, row_stats=row_stats, source_label=label)


# ---------------------------------------------------------------------------
# 4. DOC FILE WORD (.docx) - CA BANG VA VAN BAN TU DO
# ---------------------------------------------------------------------------

def _text_from_xml_element(element):
    """
    Lay TOAN BO van ban (cac the w:t) nam trong 1 phan tu XML cua Word, KE CA
    van ban nam BEN TRONG hyperlink (w:hyperlink). python-docx mac dinh BO
    QUA phan text nam trong w:hyperlink khi dung .text (ca Paragraph.text lan
    Cell.text), vi Paragraph.runs chi duyet CAC w:r la CON TRUC TIEP cua
    doan van, khong duyet vao ben trong w:hyperlink. Trong thuc te rat nhieu
    file Word chen email/SDT o dang duong dan (hyperlink mailto:) thay vi
    van ban thuong -> neu dung .text nhu binh thuong se bi MAT TRANG toan bo
    cac gia tri nay (thuong la phan lon du lieu trong file).
    """
    from docx.oxml.ns import qn

    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def _full_paragraph_text(paragraph):
    """Nhu paragraph.text nhung KHONG bo sot text nam trong hyperlink."""
    return _text_from_xml_element(paragraph._p)


def _full_cell_text(cell):
    """Nhu cell.text nhung KHONG bo sot text nam trong hyperlink. Noi cac
    doan (neu 1 o co nhieu doan van) bang '\\n' giong hanh vi cua cell.text."""
    from docx.oxml.ns import qn

    return "\n".join(
        _text_from_xml_element(p) for p in cell._tc.iter(qn("w:p"))
    )


def read_docx_file(path, row_stats=None, source_label=None):
    import docx

    label = source_label or os.path.basename(path)
    converted_path = None
    if os.path.splitext(path)[1].lower() == ".doc":
        # File Word 97-2003 cu (.doc, khac .docx) - python-docx KHONG doc
        # duoc dinh dang nay (chi ho tro .docx dang zip/XML moi hon). Tu
        # dong chuyen doi sang .docx tam thoi qua LibreOffice truoc khi doc.
        converted_path = _convert_via_libreoffice(path, "docx")
        if converted_path is None:
            raise ValueError(
                "Không thể đọc tệp .doc này: cần LibreOffice (soffice) trên máy để tự động "
                "chuyển đổi nhưng không tìm thấy. Cách khắc phục: mở tệp bằng Word và 'Save As' "
                "sang định dạng .docx rồi thử lại, hoặc cài LibreOffice (https://www.libreoffice.org/)."
            )
        path = converted_path

    try:
        document = docx.Document(path)
        records = []

        # 4a. Cac bang trong file Word. LUU Y: Word (hoac nguoi soan thao) co the
        # tach 1 danh sach dai thanh NHIEU BANG RIENG BIET trong cung file (vd do
        # ngat trang khi go tay), nhung CHI bang dau tien co dong tieu de - cac
        # bang sau bat dau thang bang du lieu. Dung chung ham gop bang (da dung
        # cho PDF nhieu trang) de tranh mat du lieu / hieu nham dong du lieu dau
        # tien cua cac bang sau la tieu de. Dung _full_cell_text (khong phai
        # cell.text) de khong bo sot email/SDT duoc chen dang hyperlink.
        raw_tables = [
            [[_full_cell_text(cell) for cell in row.cells] for row in table.rows]
            for table in document.tables
        ]
        for logical_table in merge_multi_page_tables(raw_tables):
            records.extend(extract_from_table_rows(logical_table, row_stats=row_stats, source_label=label))

        # 4b. Van ban tu do (paragraph) - gom theo tung "khoi" ban ghi. Cung dung
        # _full_paragraph_text de khong bo sot noi dung trong hyperlink.
        full_text = "\n".join(_full_paragraph_text(p) for p in document.paragraphs)
        records.extend(extract_from_free_text(full_text, row_stats=row_stats, source_label=label))

        return records
    finally:
        if converted_path:
            try:
                os.remove(converted_path)
                os.rmdir(os.path.dirname(converted_path))
            except OSError:
                pass


# ---------------------------------------------------------------------------
# 5. DOC FILE PDF - CA BANG VA VAN BAN TU DO
# ---------------------------------------------------------------------------

def read_pdf_file(path, row_stats=None, source_label=None):
    import pdfplumber

    label = source_label or os.path.basename(path)
    records = []
    raw_tables = []
    free_text_parts = []  # CHI lay text cua trang KHONG co bang, tranh xu ly
                           # trung/nhieu 2 lan tren cung 1 du lieu bang (vd
                           # ten bi tach doi, email bi sai do doan van suy
                           # doan nham khi PDF ngat dong giua o bang).
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables() or []
            if page_tables:
                raw_tables.extend(page_tables)
            else:
                free_text_parts.append(page.extract_text() or "")

    for logical_table in merge_multi_page_tables(raw_tables):
        records.extend(extract_from_table_rows(logical_table, row_stats=row_stats, source_label=label))

    if free_text_parts:
        records.extend(extract_from_free_text("\n".join(free_text_parts), row_stats=row_stats, source_label=label))

    return records


# ---------------------------------------------------------------------------
# 6. TRICH XUAT TU VAN BAN TU DO (khong co cau truc bang)
# ---------------------------------------------------------------------------

def _split_by_label_anchor(text):
    """
    Tach van ban thanh tung ban ghi dua tren vi tri cac nhan lap lai. Uu tien
    dung nhan 'Ho ten' lam diem bat dau ban ghi (vi thuong xuat hien truoc
    email/sdt trong 1 ban ghi); neu khong co thi dung nhan email; cuoi cung
    fallback ve vi tri cac email "tran" (khong nhan) trong van ban.
    Moi ban ghi = tu vi tri anchor hien tai DEN TRUOC vi tri anchor ke tiep,
    dam bao khong bi lech huong (name/email/phone cua cung 1 nguoi nam chung
    1 khoi).
    """
    norm_text = _deaccent_keep_len(text)  # cung do dai voi text goc -> vi tri map 1-1
    name_starts = [m.start() for m in LABEL_PATTERNS["name"].finditer(norm_text)]
    email_label_starts = [m.start() for m in LABEL_PATTERNS["email"].finditer(norm_text)]
    bare_email_starts = [m.start() for m in EMAIL_REGEX.finditer(text)]

    anchors = name_starts if len(name_starts) > 1 else (
        email_label_starts if len(email_label_starts) > 1 else bare_email_starts
    )

    if len(anchors) <= 1:
        return [text]

    anchors = sorted(anchors)
    blocks = []
    for i, start in enumerate(anchors):
        end = anchors[i + 1] if i + 1 < len(anchors) else len(text)
        blocks.append(text[start:end])
    return blocks


def extract_from_free_text(text, row_stats=None, source_label=""):
    """
    Chia van ban thanh cac 'khoi' (block) cach nhau boi dong trong, moi block
    duoc coi la 1 ban ghi. Trong moi block, uu tien tim theo nhan (Email:,
    Ho ten:, SDT:...) - so khop khong phan biet dau tieng Viet, neu khong co
    nhan thi fallback sang regex tong quat. Cac dong "Chuc vu:" / "Don vi:"
    (va tuong tu) bi loai bo, khong duoc dung lam ten.

    source_label: xem extract_from_table_rows - nhan mo ta nguon, ghi kem
    theo dong bi loai de phuc vu tab "Can kiem tra".
    """
    if not text or not text.strip():
        return []

    blocks = re.split(r"\n\s*\n", text)
    # Neu toan bo van ban chi la 1 khoi lon (khong co dong trong phan cach,
    # thuong gap khi PDF loai bo dong trong) -> tach theo vi tri cac nhan
    # lap lai (Ho ten / Email / SDT), moi nhan danh dau DIEM BAT DAU 1 ban ghi.
    if len(blocks) <= 1 and len(EMAIL_REGEX.findall(text)) > 1:
        blocks = _split_by_label_anchor(text)

    records = []
    block_number = 0
    for block in blocks:
        if not block.strip():
            continue
        block_number += 1

        norm_block = _deaccent_keep_len(block)  # cung do dai -> vi tri map ve block goc

        email_match = LABEL_PATTERNS["email"].search(norm_block)
        raw_email = block[email_match.start(1):email_match.end(1)].strip() if email_match else None
        if not raw_email:
            generic = EMAIL_REGEX.search(block)
            raw_email = generic.group(0) if generic else None

        phone_match = LABEL_PATTERNS["phone"].search(norm_block)
        raw_phone = block[phone_match.start(1):phone_match.end(1)].strip() if phone_match else None
        if not raw_phone:
            generic_phone = PHONE_CANDIDATE_REGEX.search(block)
            raw_phone = generic_phone.group(0) if generic_phone else None

        name_match = LABEL_PATTERNS["name"].search(norm_block)
        raw_name = block[name_match.start(1):name_match.end(1)].strip() if name_match else None
        if not raw_name:
            # Fallback: dong dau tien cua block khong chua email/phone/nhan
            # khac va KHONG PHAI la truong "Chuc vu"/"Don vi" (bi loc bo).
            for line in block.split("\n"):
                line = line.strip()
                if not line:
                    continue
                if is_noise_line(line):
                    continue  # bo qua "Chuc vu: ..." / "Don vi: ..."
                if EMAIL_REGEX.search(line):
                    continue
                if PHONE_CANDIDATE_REGEX.search(line) and len(line) < 20:
                    continue
                raw_name = line
                break

        # LUU Y: neu khoi KHONG co bat ky dau hieu nao la 1 ban ghi that su
        # (khong co nhan Email/Ho ten/SDT ro rang, khong co so dien thoai) va
        # cung khong co email -> day nhieu kha nang chi la doan van ban thuong
        # (vd loi mo dau cong van), KHONG phai 1 dong du lieu bi loi -> bo qua
        # HOAN TOAN, khong tinh vao thong ke/"Can kiem tra" de tranh nhieu.
        # Nguoc lai (co it nhat 1 dau hieu ro rang la dang co gang ghi 1
        # nguoi - vd co nhan "SDT:" hoac so dien thoai - nhung lai thieu
        # email) thi VAN tinh la 1 dong bi loi, dung nhu bang du lieu.
        has_record_signal = bool(name_match) or bool(phone_match) or bool(raw_phone)
        if not raw_email and not has_record_signal:
            continue

        row_context = {"source": source_label, "row_number": block_number}
        rec = build_clean_record(raw_email, raw_name, raw_phone, row_stats=row_stats, row_context=row_context)
        if rec:
            records.append(rec)
    return records


# ---------------------------------------------------------------------------
# 7. DIEU PHOI THEO LOAI FILE
# ---------------------------------------------------------------------------

def extract_from_file(path, row_stats=None):
    ext = os.path.splitext(path)[1].lower()
    label = os.path.basename(path)
    if ext in (".xlsx", ".xlsm"):
        return read_excel_file(path, row_stats=row_stats, source_label=label)
    if ext == ".xls":
        return read_xls_file(path, row_stats=row_stats, source_label=label)
    if ext == ".csv":
        return read_csv_file(path, row_stats=row_stats, source_label=label)
    if ext in (".docx", ".doc"):
        return read_docx_file(path, row_stats=row_stats, source_label=label)
    if ext == ".pdf":
        return read_pdf_file(path, row_stats=row_stats, source_label=label)
    raise ValueError(f"Không hỗ trợ định dạng tệp: {ext} ({path})")


def dedupe_by_email(records, row_stats=None):
    """Loai ban ghi trung email (giu ban ghi xuat hien dau tien). Neu duoc
    truyen row_stats, ghi lai CAC BAN GHI BI LOAI vao row_stats['issues']
    (kem nguon/STT/du lieu goc) voi ly do 'duplicate_email', phuc vu tab
    "Can kiem tra" - nguoi dung can biet ban ghi trung nam o dau, khong chi
    biet CO trung."""
    seen = set()
    result = []
    for r in records:
        if r["email"] in seen:
            if row_stats is not None:
                ctx = r.get("_context") or {}
                row_stats.setdefault("issues", []).append({
                    "source": ctx.get("source", ""),
                    "row_number": ctx.get("row_number"),
                    "raw_email": r["email"],
                    "raw_name": r["name"],
                    "raw_phone": r["phone"],
                    "reason": "duplicate_email",
                })
            continue
        seen.add(r["email"])
        result.append(r)
    return result


# ---------------------------------------------------------------------------
# 8. XUAT RA FILE EXCEL THEO TEMPLATE
# ---------------------------------------------------------------------------

def write_output(records, template_path, output_path, password=DEFAULT_PASSWORD):
    """
    Mo file template, xac dinh cot Email / Ten tai khoan / Dien thoai (va Mat
    khau neu co) dua tren header dong 1, sau do ghi du lieu tu dong 2, giu
    nguyen dinh dang/sheet cua template. Cot "Mat khau" (neu template co) se
    duoc dien gia tri mac dinh `password` cho moi ban ghi, vi du du lieu
    nguon khong the cung cap mat khau.
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    col_map = {}
    for cell in header_row:
        field = match_header_field(cell.value)
        if field and field not in col_map:
            col_map[field] = cell.column  # 1-based column index

    missing = [f for f in ("email", "name", "phone") if f not in col_map]
    if missing:
        field_labels_vn = {"email": "Email", "name": "Họ và tên", "phone": "Điện thoại"}
        missing_labels = [field_labels_vn.get(f, f) for f in missing]
        raise ValueError(
            "Không tìm thấy cột phù hợp trong file mẫu cho: "
            + ", ".join(missing_labels)
            + ". Vui lòng kiểm tra lại tiêu đề cột của file mẫu."
        )

    from openpyxl.styles import Font as _Font

    base_font = header_row[0].font
    sample_font = _Font(
        name=base_font.name, size=base_font.sz, bold=False,
        italic=base_font.i, color=base_font.color,
    )

    start_row = 2
    for i, rec in enumerate(records):
        row_idx = start_row + i
        ws.cell(row=row_idx, column=col_map["email"], value=rec["email"]).font = sample_font
        ws.cell(row=row_idx, column=col_map["name"], value=rec["name"]).font = sample_font
        ws.cell(row=row_idx, column=col_map["phone"], value=rec["phone"]).font = sample_font
        if "password" in col_map:
            ws.cell(row=row_idx, column=col_map["password"], value=password).font = sample_font

    wb.save(output_path)


# ---------------------------------------------------------------------------
# 9. HAM CHINH / CLI
# ---------------------------------------------------------------------------

def run(inputs, template_path, output_path, dedupe=True, password=DEFAULT_PASSWORD, verbose=True):
    records, stats = extract_all(inputs, dedupe=dedupe, verbose=verbose)
    write_output(records, template_path, output_path, password=password)
    if verbose:
        print_stats_report(stats)
        print(f"Đã ghi {stats['final_count']} bản ghi vào: {output_path}")
    return records, stats


def print_stats_report(stats):
    """In bao cao thong ke day du ra console (dung cho CLI / debug)."""
    print("--- THỐNG KÊ ---")
    print(f"Tổng số dòng dữ liệu đọc được trong tệp:   {stats['total_rows']}")
    print(f"Số dòng hợp lệ (có email đúng định dạng):  {stats['valid_rows']}")
    print(f"Số dòng bị loại - thiếu email:              {stats['missing_email']}")
    print(f"Số dòng bị loại - email sai định dạng:      {stats['invalid_email_format']}")
    print(f"Số dòng bị loại - trùng email:               {stats['duplicates_removed']}")
    print(f"Còn lại sau cùng (đã lọc trùng):            {stats['final_count']}")


def extract_all(inputs, dedupe=True, verbose=True):
    """
    Doc + trich xuat + chuan hoa du lieu tu tat ca file nguon, KHONG ghi ra
    file Excel (dung cho buoc xem truoc trong GUI truoc khi xuat file).

    Tra ve (records, stats) trong do:
      records: danh sach ban ghi hop le sau cung (da loc trung neu dedupe=True)
      stats:   dict thong ke day du:
        - total_rows:            tong so dong du lieu da xu ly (moi dong co
                                  cau truc/duoc nhan dien la 1 ban ghi tiem
                                  nang, KHONG tinh dong hoan toan trong/dong
                                  tieu de van ban)
        - valid_rows:             so dong co email hop le (TRUOC khi loc trung)
        - missing_email:          so dong bi loai vi KHONG co gia tri email
        - invalid_email_format:   so dong bi loai vi CO gia tri o cot email
                                   nhung khong dung dinh dang email hop le
        - duplicates_removed:     so dong bi loai vi trung email voi 1 dong
                                   da co truoc do
        - final_count:            so dong con lai sau cung (= valid_rows -
                                   duplicates_removed), day la so dong se
                                   duoc ghi vao file ket qua
        - per_file_counts:        so ban ghi hop le (truoc loc trung) theo
                                   tung file nguon
    """
    all_records = []
    per_file_counts = {}
    row_stats = {
        "total_rows": 0,
        "valid_rows": 0,
        "missing_email": 0,
        "invalid_email_format": 0,
        "issues": [],
    }
    for path in inputs:
        if verbose:
            print(f"Đang đọc: {path}")
        try:
            recs = extract_from_file(path, row_stats=row_stats)
        except Exception as e:
            print(f"  !! Lỗi khi đọc {path}: {e}", file=sys.stderr)
            per_file_counts[path] = 0
            continue
        if verbose:
            print(f"  -> Trích xuất được {len(recs)} bản ghi hợp lệ (có email)")
        per_file_counts[path] = len(recs)
        all_records.extend(recs)

    total_extracted = len(all_records)

    if dedupe:
        final_records = dedupe_by_email(all_records, row_stats=row_stats)
    else:
        final_records = all_records

    duplicates_removed = total_extracted - len(final_records)

    # Bo khoa noi bo "_context" (chi dung de theo doi nguon/STT phuc vu tab
    # "Can kiem tra") truoc khi tra ve - khong can thiet cho phan xuat file.
    clean_final_records = [
        {"email": r["email"], "name": r["name"], "phone": r["phone"]} for r in final_records
    ]

    stats = {
        "total_rows": row_stats["total_rows"],
        "valid_rows": row_stats["valid_rows"],
        "missing_email": row_stats["missing_email"],
        "invalid_email_format": row_stats["invalid_email_format"],
        "duplicates_removed": duplicates_removed,
        "final_count": len(clean_final_records),
        "per_file_counts": per_file_counts,
        "issues": row_stats["issues"],
        # Giu lai 2 khoa cu de tuong thich nguoc voi code/GUI dang dung ten cu
        "total_extracted": total_extracted,
    }
    if verbose:
        if duplicates_removed:
            print(f"Đã loại {duplicates_removed} bản ghi trùng email")
        print_stats_report(stats)

    return clean_final_records, stats


def main():
    parser = argparse.ArgumentParser(
        description="Trich xuat Email/Ho ten/SDT tu PDF/Word/Excel va xuat theo template Excel."
    )
    parser.add_argument("--input", nargs="+", required=True, help="Cac file nguon (PDF/DOCX/XLSX/XLS/CSV)")
    parser.add_argument("--template", required=True, help="File Excel template")
    parser.add_argument("--output", required=True, help="Duong dan file Excel ket qua")
    parser.add_argument("--no-dedupe", action="store_true", help="Khong loai bo ban ghi trung email")
    parser.add_argument(
        "--password", default=DEFAULT_PASSWORD,
        help=f"Mat khau mac dinh dien vao cot Mat khau (mac dinh: {DEFAULT_PASSWORD})",
    )
    args = parser.parse_args()

    run(args.input, args.template, args.output, dedupe=not args.no_dedupe, password=args.password)


if __name__ == "__main__":
    main()