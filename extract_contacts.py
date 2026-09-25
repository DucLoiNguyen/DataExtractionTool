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
        "thu cong vu", "dia chi thu",
    ],
    "phone": ["dien thoai", "sdt", "so dien thoai", "phone", "hotline", "tel", "di dong", "mobile"],
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

def _has_empty_email_label(email):
    """True neu email co 1 'nhan' (label - phan giua 2 dau cham, hoac giua
    @ va dau cham/ky tu ke tiep) RONG - vd '@.domain' (nhan dau tien cua
    domain rong), 'domain..com' (2 dau cham lien tiep = 1 nhan rong o
    giua), 'ten.@domain' (nhan cuoi cua phan ten rong)."""
    if ".." in email:
        return True
    local, _, domain = email.partition("@")
    if not local or local.startswith(".") or local.endswith("."):
        return True
    if not domain or domain.startswith(".") or domain.endswith("."):
        return True
    return False


def _repair_email_typos(text):
    """Tu dong sua 1 so LOI GO PHIM PHO BIEN trong email (CHI loai bo/thay
    the ky tu THUA hoac SAI VI TRI ro rang, KHONG doan/them NOI DUNG MOI
    - vd KHONG tu dien them phan ten mien bi thieu hoan toan):
      - Thay dau phay (,) bang dau cham (.) - loi go nham phim ke nhau
        tren ban phim (vd '@gmail,com' -> '@gmail.com').
      - Gop nhieu dau cham lien tiep ('..', '...') thanh 1 dau cham (vd
        'ten..donvi@...' -> 'ten.donvi@...').
      - Bo dau cham NGAY SAU '@' (vd '@.domain' -> '@domain').
      - Bo dau cham NGAY TRUOC '@' (vd 'ten.@domain' -> 'ten@domain').
    Tra ve chuoi da sua (co the van chua hop le neu co loi khac loai nay,
    vd thieu han '@' hoac thieu ca 1 doan ten mien - nhung truong hop do
    KHONG the tu sua an toan nen se van bi loai o buoc kiem tra sau)."""
    text = text.replace(",", ".")
    text = re.sub(r"\.{2,}", ".", text)
    text = re.sub(r"@\.+", "@", text)
    text = re.sub(r"\.+@", "@", text)
    return text


def normalize_email(raw):
    """Tra ve email hop le, viet thuong; None neu khong hop le / rong.
    Bo TOAN BO khoang trang/xuong dong truoc khi tim, vi PDF/Word co the
    ngat dong GIUA 1 dia chi email dai (vd trong 1 o bang bi wrap chu),
    lam email bi tach thanh 2 dong va regex email se bat thieu neu khong
    noi lai truoc.

    Neu ket qua tim duoc CO NHAN RONG (vd '@.domain', 'ten..donvi@...')
    HOAC khong tim thay ket qua nao ca (vd dau phay lam vo cau truc, '@gmail,com'
    khien regex khong nhan ra duoc tu dau) - THU SUA cac loi go phim pho
    bien qua _repair_email_typos() roi tim lai TREN TOAN BO chuoi da sua,
    truoc khi ket luan la khong hop le. Day la SUA LOI GO PHIM RO RANG
    (thua/sai vi tri 1 ky tu), KHONG PHAI tu doan/them noi dung email bi
    thieu hoan toan (vd khong the tu dien ten mien neu no bi mat han)."""
    if not raw:
        return None
    raw = re.sub(r"\s+", "", str(raw))

    match = EMAIL_REGEX.search(raw)
    if match and not _has_empty_email_label(match.group(0)):
        return match.group(0).lower()

    repaired = _repair_email_typos(raw)
    match2 = EMAIL_REGEX.search(repaired)
    if match2 and not _has_empty_email_label(match2.group(0)):
        return match2.group(0).lower()

    return None


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
    # So Excel da bi ep thanh chuoi o dau do (vd "912345678.0" - CSV xuat tu
    # Excel, hoac code goi str() truoc khi truyen vao): bo duoi ".0" truoc,
    # neu khong phan "0" thua se bi gop vao thanh 10 chu so sai va mat SDT.
    if re.fullmatch(r"\+?\d+\.0+", raw):
        raw = raw.split(".")[0]

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


# ---------------------------------------------------------------------------
# 2b. STT (so thu tu) - DUNG CHUNG CHO CA 2 CHE DO
#     Dung de (1) hien STT GOC trong tab "Can kiem tra" (nguoi dung tra nguoc
#     lai tep nguon), (2) phat hien STT bi "nhay so" - dau hieu MAT DONG AM
#     THAM (pdfplumber bo sot ca 1 dong khi tach bang, vd Ta Phin mat STT
#     23/40/55/79) hoac nguon danh so sai (vd Dong Thap: 2315 go nham 2015).
# ---------------------------------------------------------------------------

_STT_HEADERS = {"stt", "tt", "so tt", "so thu tu", "thu tu"}
_ROMAN_GROUP_RE = re.compile(r"^[IVXLC]{1,6}\.?$")

# Dong van ban tho bat dau bang STT (vd "23 Nguyen Van A ...").
LEADING_STT_LINE_RE = re.compile(r"^\s*(\d{1,4})\s+(.+)$")

# Cac tu bat dau pho bien cua ten don vi/phong ban tieng Viet - dung de
# tach "Ten - Don vi" khi ca 2 bi dinh lien trong van ban tho (khong con
# ranh gioi cot) khi khoi phuc dong bi mat.
UNIT_KEYWORD_SPLIT_RE = re.compile(
    r"\b(?:Văn phòng|Ban(?:\s|$)|UBND|HĐND|Ủy ban|Phòng|Trung tâm|Hội(?:\s|$)|"
    r"Đoàn|Sở|Viện|Cơ quan|Đảng ủy)",
    re.IGNORECASE,
)

# Chuc vu hay dung lien sau ten trong van ban tho (vd Dong Thap: "Nguyen
# Thi A Chuyen vien Phong Van hoa..."). CHI dung cum 2 tu tro len de khong
# cat nham ten nguoi (vd ten "Chuyen" don le van giu).
POSITION_SPLIT_RE = re.compile(
    r"\b(?:Chuyên viên|Phó\s+(?:Chủ tịch|Trưởng|Giám đốc|Bí thư|Chánh|phòng)|"
    r"Trưởng\s+(?:phòng|ban|khoa)|Giám đốc|Chủ tịch|Bí thư|Công chức|Viên chức|"
    r"Cán bộ|Nhân viên|Kế toán|Văn thư)",
    re.IGNORECASE,
)


def is_stt_header(cell):
    """True neu o tieu de la cot STT ('STT', 'TT', 'Stt', 'Số TT'...)."""
    norm = re.sub(r"[^a-z0-9 ]", "", _strip_accents_lower(cell)).strip()
    return norm in _STT_HEADERS


def parse_stt(value):
    """Tra ve STT dang so nguyen (int) hoac None. Chap nhan float tu Excel
    (1.0 -> 1) va chuoi so ('12', '12.')."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() and value > 0 else None
    s = str(value).strip().rstrip(".")
    return int(s) if s.isdigit() and 0 < int(s) < 100000 else None


def is_group_stt(value):
    """True neu o STT rong hoac la so La Ma (I, II, III...) - dau hieu cua
    dong TIEU DE NHOM, khong phai dong du lieu 1 nguoi."""
    s = "" if value is None else str(value).strip()
    return s == "" or bool(_ROMAN_GROUP_RE.match(s))


def find_stt_gaps(stts):
    """Tra ve danh sach STT bi thieu trong day 1..max(stts)."""
    stts = set(stts)
    if not stts:
        return []
    return [n for n in range(1, max(stts) + 1) if n not in stts]


def find_stt_lines(full_text, wanted):
    """Tim trong van ban tho (dong theo dong) dong BAT DAU bang tung STT
    trong `wanted`, gop them toi da 2 dong ke tiep (phong khi bi xuong dong
    giua chung). Tra ve {stt: chuoi_da_gop} cho cac STT tim thay."""
    if not full_text or not wanted:
        return {}
    wanted = set(wanted)
    lines = full_text.split("\n")
    found = {}
    for i, line in enumerate(lines):
        m = LEADING_STT_LINE_RE.match(line)
        if not m:
            continue
        num = int(m.group(1))
        if num in wanted and num not in found:
            combined = m.group(2)
            for j in range(1, 3):
                if i + j < len(lines):
                    nxt = lines[i + j].strip()
                    if nxt and not LEADING_STT_LINE_RE.match(nxt):
                        combined += " " + nxt
                    else:
                        break
            found[num] = combined
    return found


def split_recovered_stt_line(combined):
    """Tach 1 dong van ban tho da khoi phuc thanh cac phan. Tra ve dict
    {email, phone, name, unit} (email/phone la chuoi THO tim bang regex,
    None neu khong co; unit la None neu khong tach duoc). Ten = phan truoc
    email, cat tai tu khoa don vi (UNIT_KEYWORD_SPLIT_RE) roi tai chuc vu
    (POSITION_SPLIT_RE) neu co."""
    email_m = EMAIL_REGEX.search(combined)
    phone_m = re.search(r"(?:\+?84|0)\d{8,10}", combined.replace(" ", ""))
    name_part = combined[:email_m.start()] if email_m else combined
    name_part = name_part.strip(" \t-|/,;:")
    unit = None
    m_unit = UNIT_KEYWORD_SPLIT_RE.search(name_part)
    if m_unit and m_unit.start() > 0:
        unit = name_part[m_unit.start():].strip(" \t-|/,;:")
        name_part = name_part[:m_unit.start()].strip(" \t-|/,;:")
    m_pos = POSITION_SPLIT_RE.search(name_part)
    if m_pos and m_pos.start() > 0:
        name_part = name_part[:m_pos.start()].strip(" \t-|/,;:")
    return {
        "email": email_m.group(0) if email_m else None,
        "phone": phone_m.group(0) if phone_m else None,
        "name": name_part,
        "unit": unit,
    }


def format_stt_list(nums, limit=15):
    """'23, 40, 55' - cat bot neu qua dai (dung cho thong bao canh bao)."""
    nums = sorted(nums)
    text = ", ".join(str(n) for n in nums[:limit])
    return text + (f"... (tổng {len(nums)})" if len(nums) > limit else "")


def add_warning(row_stats, message):
    """Ghi 1 canh bao (chuoi co dau, hien cho nguoi dung) vao row_stats.
    Canh bao KHAC 'issue': khong gan voi 1 nguoi cu the ma la dau hieu co
    the MAT DU LIEU AM THAM (STT nhay so, ca cot SDT trong...) - '0 loi'
    khong co nghia la dung."""
    if row_stats is None:
        return
    warnings = row_stats.setdefault("warnings", [])
    if message not in warnings:
        warnings.append(message)


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

def merge_wrapped_continuation_rows(rows, name_col_idx, stt_col_idx=None):
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

    stt_col_idx: KHONG noi o STT (Dong Thap danh STT ca cho dong bi xuong
    hang, vd 909 + 910 'dongthap.gov.vn' -> truoc day o STT thanh '909910').
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
                if idx == stt_col_idx:
                    continue
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


_EMAIL_TAIL_RE = re.compile(r"[A-Za-z0-9.\-]{1,15}")
_PHONE_ONLY_CELL_RE = re.compile(r"[\d\s.+\-()]{9,20}")


def _cell_str(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _find_email_in_row(row, email_idx, skip_idxs=()):
    """Tim email cua 1 dong: uu tien cot Email; neu cot do KHONG co email
    hop le thi tim THEO NOI DUNG o cac o khac. Tra ve (gia_tri_tho, idx);
    idx = None neu phai ghep 2 o. Neu khong tim thay, tra ve gia tri cot
    Email goc (de ghi dung du lieu goc vao "Can kiem tra").

    Ly do (da gap thuc te, Dong Thap STT 742): PDF bi LECH COT - cot Email
    chua SDT '919.099.499', con email nam o cot SDT va bi pdfplumber CAT
    DOI sang o ke ben ('hoquanglon@gmail.c' + 'om'). Doc theo vi tri cot co
    dinh se loai nham nguoi nay.

    Chi GHEP 2 o khi o sau la 1 doan ngan toan ky tu email (vd 'om', 'vn',
    'gov.vn') VA ket qua ghep moi hop le - khong bao gio doan/them noi dung.
    Email thieu '@' hoac thieu ten mien (vd 'nvgiang') van bi loai."""
    primary = row[email_idx] if email_idx is not None and email_idx < len(row) else None
    if normalize_email(primary):
        return primary, email_idx

    for i, cell in enumerate(row):
        if i == email_idx or i in skip_idxs:
            continue
        s = _cell_str(cell)
        if "@" in s and normalize_email(s):
            return cell, i

    for i in range(len(row) - 1):
        a, b = _cell_str(row[i]), _cell_str(row[i + 1])
        if "@" in a and not normalize_email(a) and _EMAIL_TAIL_RE.fullmatch(b):
            if normalize_email(a + b):
                return a + b, None

    return primary, email_idx


def _find_phone_in_row(row, phone_idx, email_used_idx, email_idx, skip_idxs=()):
    """Tim SDT cua 1 dong: uu tien cot SDT. CHI tim sang o khac khi co dau
    hieu LECH COT (email lay tu cot khac cot Email, hoac cot SDT co noi
    dung nhung khong phai SDT) - tranh vo tinh lay 1 con so khong phai SDT
    (vd o Ghi chu) cho nguoi de trong SDT."""
    primary = row[phone_idx] if phone_idx is not None and phone_idx < len(row) else None
    if normalize_phone(primary):
        return primary
    shifted = (email_used_idx != email_idx) or bool(_cell_str(primary))
    if not shifted:
        return primary
    for i, cell in enumerate(row):
        if i in (phone_idx, email_used_idx) or i in skip_idxs:
            continue
        s = _cell_str(cell)
        if s and _PHONE_ONLY_CELL_RE.fullmatch(s) and normalize_phone(cell):
            return cell
    return primary


# Ten nhom/co quan thuong mo dau bang cac tu nay. KHONG co "Đoàn" (la ho
# nguoi, vd "Đoàn Văn Nỉ" o Dong Thap).
_GROUP_UNIT_START_RE = re.compile(
    r"^(?:Ủy ban|Uỷ ban|UBND|HĐND|Hội đồng|Sở|Ban|Trung tâm|Văn phòng|Phòng|Thanh tra|"
    r"Viện|Cơ quan|Đảng ủy|Đảng uỷ|Trường|Bệnh viện|Chi cục|Cục|Tổng cục|Công an)(?:\s|$)",
    re.IGNORECASE,
)


def _is_group_header_row(row, stt_idx, name_idx):
    """Dong TIEU DE NHOM (vd Dong Thap: '| Ủy ban nhân dân Phường Lo | ng
    Thuận |') - khong phai 1 nguoi, khong dua vao thong ke/"Can kiem tra".
    CHI ap dung khi bang CO cot STT (nguoi that luon co STT): o STT rong/so
    La Ma, co ten, khong co email/SDT o bat ky o nao, VA mot trong hai:
      - cac o con lai gan nhu rong (<= 3 ky tu - pdfplumber hay cat vai chu
        cuoi sang o ke ben), hoac
      - o ten bat dau bang ten loai co quan (Ủy ban, Sở, Trung tâm...) - ten
        nhom dai bi cat sang o ke ben ('Trung tâm Khởi nghiệp đổi' | 'mới
        sáng tạo...'). Nguoi that khong STT/email/SDT nhung ten la ten nguoi
        van vao "Can kiem tra" nhu truoc."""
    if stt_idx is None or name_idx is None:
        return False
    stt_val = row[stt_idx] if stt_idx < len(row) else None
    if not is_group_stt(stt_val):
        return False
    name_val = _cell_str(row[name_idx]) if name_idx < len(row) else ""
    if not name_val:
        return False
    others = 0
    for i, cell in enumerate(row):
        s = _cell_str(cell)
        if not s:
            continue
        if "@" in s or PHONE_CANDIDATE_REGEX.search(s) or _PHONE_ONLY_CELL_RE.fullmatch(s):
            return False
        if i not in (stt_idx, name_idx):
            others += len(s)
    return others <= 3 or bool(_GROUP_UNIT_START_RE.match(name_val))


def extract_from_table_rows(rows, row_stats=None, source_label="", seen_stts=None, merge_wrapped=True):
    """
    rows: list cac list o. Tu dong TIM dong tieu de that su (xem
    find_header_row - co the khong phai dong 0 neu co vai dong tieu de van
    ban phia truoc) roi xac dinh cot nao la email / ten / dien thoai dua
    tren HEADER_KEYWORDS. Neu khong tim thay dong tieu de phu hop trong 15
    dong dau, tra ve [] (va ghi canh bao neu bang co chua email).

    source_label: nhan mo ta nguon du lieu (vd "ten_file.xlsx" hoac
    "ten_file.xlsx [Sheet: Thang1]") - duoc ghi kem theo moi dong bi loai
    de nguoi dung biet CHINH XAC no thuoc tep/sheet nao (xem tab "Can kiem
    tra" tren giao dien).

    seen_stts: set (tuy chon) - neu truyen vao, moi STT so nguyen doc duoc
    (ca dong hop le lan bi loai) duoc them vao, de ham goi phat hien STT
    bi nhay so (xem find_stt_gaps).

    Neu bang co cot STT: "row_number" cua dong bi loai la STT GOC trong tep
    (nguoi dung tra nguoc duoc), khong phai so thu tu tu dem.

    merge_wrapped: chi True cho PDF/Word (o bi ngat dong). Excel/CSV truyen
    False - dong thieu ten trong Excel la loi du lieu THAT, ghep nham se
    lam mat 1 nguoi (da gap o Che do 2, file Can Tho).
    """
    if not rows:
        return []

    header_idx, col_map = find_header_row(rows)
    if header_idx is None:
        # Bang co email nhung khong nhan ra tieu de cot -> ca bang bi bo qua.
        # Truoc day im lang; nay canh bao de nguoi dung kiem tra tieu de cot.
        n_email_rows = sum(1 for r in rows if r and any("@" in _cell_str(c) for c in r))
        if n_email_rows:
            add_warning(row_stats, f"{source_label}: có 1 bảng ({n_email_rows} dòng chứa email) "
                                   "không nhận diện được tiêu đề cột Email nên bị bỏ qua. "
                                   "Hãy kiểm tra tiêu đề cột.")
        return []

    header = rows[header_idx]
    stt_idx = next((i for i, c in enumerate(header) if is_stt_header(c)), None)
    name_idx = col_map.get("name")
    skip_idxs = {i for i in (stt_idx, name_idx) if i is not None}

    data_rows = rows[header_idx + 1:]
    # Ghi nhan STT cua MOI dong TRUOC khi gop dong bi ngat - ke ca dong se
    # bi gop vao dong tren (vd Dong Thap STT 910 chi chua duoi email cua
    # 909). Chi STT VANG MAT HOAN TOAN khoi bang moi la dau hieu mat dong;
    # neu khong, buoc khoi phuc STT se tao "nguoi ao" tu cac dong nay.
    if seen_stts is not None and stt_idx is not None:
        for r in data_rows:
            if r and stt_idx < len(r):
                n = parse_stt(r[stt_idx])
                if n is not None:
                    seen_stts.add(n)
    # Gop cac dong bi "gay" do PDF ngat dong giua o (xem
    # merge_wrapped_continuation_rows) TRUOC KHI xu ly tung dong, de khong
    # mat/sai du lieu (vd email bi cat lam doi) va khong tao ra "nguoi ao"
    # tu phan bi ngat.
    if merge_wrapped:
        data_rows = merge_wrapped_continuation_rows(data_rows, name_idx, stt_idx)

    records = []
    row_number = 0
    last_stt = None
    for row in data_rows:
        if row is None or all(c in (None, "") for c in row):
            continue
        if _is_group_header_row(row, stt_idx, name_idx):
            continue
        row_number += 1

        def get(field):
            idx = col_map.get(field)
            if idx is None or idx >= len(row):
                return ""
            return row[idx] if row[idx] is not None else ""

        stt = parse_stt(row[stt_idx]) if stt_idx is not None and stt_idx < len(row) else None

        raw_email, email_used_idx = _find_email_in_row(row, col_map.get("email"), skip_idxs)
        raw_phone = _find_phone_in_row(row, col_map.get("phone"), email_used_idx,
                                       col_map.get("email"), skip_idxs)

        if stt is not None:
            display_row = stt
            last_stt = stt
        elif stt_idx is not None:
            # Bang co cot STT nhung dong nay khong co -> chi vi tri tuong doi,
            # khong dung so tu dem (se trung/nham voi STT that cua dong khac).
            display_row = f"sau STT {last_stt}" if last_stt is not None else "đầu bảng"
        else:
            display_row = row_number
        row_context = {"source": source_label, "row_number": display_row, "stt": stt}
        rec = build_clean_record(raw_email or "", get("name"), raw_phone if raw_phone is not None else "",
                                  row_stats=row_stats, row_context=row_context)
        if rec:
            records.append(rec)
    return records


def warn_stt_gaps(seen_stts, row_stats, source_label, missing=None):
    """Ghi canh bao neu STT bi nhay so (dau hieu mat dong / nguon danh so sai)."""
    missing = find_stt_gaps(seen_stts) if missing is None else missing
    if missing:
        add_warning(row_stats, f"{source_label}: STT bị nhảy số ({format_stt_list(missing)}) - "
                               "không tìm thấy các dòng này. Hãy đối chiếu tệp gốc "
                               "(có thể nguồn đánh số sai, hoặc dòng bị mất khi đọc).")


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
            seen = set()
            all_records.extend(extract_from_table_rows(rows, row_stats=row_stats, source_label=sheet_label,
                                                       seen_stts=seen, merge_wrapped=False))
            warn_stt_gaps(seen, row_stats, sheet_label)
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
        seen = set()
        all_records.extend(extract_from_table_rows(rows, row_stats=row_stats, source_label=sheet_label,
                                                   seen_stts=seen, merge_wrapped=False))
        warn_stt_gaps(seen, row_stats, sheet_label)
    return all_records


def read_csv_file(path, row_stats=None, source_label=None):
    label = source_label or os.path.basename(path)
    with open(path, newline="", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        rows = list(reader)
    seen = set()
    records = extract_from_table_rows(rows, row_stats=row_stats, source_label=label, seen_stts=seen,
                                      merge_wrapped=False)
    warn_stt_gaps(seen, row_stats, label)
    return records


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
        seen = set()
        for logical_table in merge_multi_page_tables(raw_tables):
            records.extend(extract_from_table_rows(logical_table, row_stats=row_stats, source_label=label,
                                                   seen_stts=seen))
        warn_stt_gaps(seen, row_stats, label)

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
    page_texts = []  # van ban tho MOI trang - chi dung de khoi phuc STT bi mat
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables() or []
            page_text = page.extract_text() or ""
            page_texts.append(page_text)
            if page_tables:
                raw_tables.extend(page_tables)
            else:
                free_text_parts.append(page_text)

    seen = set()
    for logical_table in merge_multi_page_tables(raw_tables):
        records.extend(extract_from_table_rows(logical_table, row_stats=row_stats, source_label=label,
                                               seen_stts=seen))
    records = _recover_missing_stt_rows(records, seen, "\n".join(page_texts), row_stats, label)

    if free_text_parts:
        records.extend(extract_from_free_text("\n".join(free_text_parts), row_stats=row_stats, source_label=label))

    return records


def _recover_missing_stt_rows(records, seen_stts, full_text, row_stats, source_label):
    """Che do 1 - LUOI AN TOAN khi pdfplumber BO SOT HOAN TOAN 1 dong khi
    tach bang (Ta Phin mat STT 23/40/55/79 - dong khong xuat hien trong
    bang tach duoc chut nao). Cung co che voi Che do 2
    (process_template_v2._recover_missing_stt_rows): do STT bi nhay so, tim
    lai dong bat dau bang STT do trong van ban tho cua PDF.

      - Tim thay dong -> xu ly nhu 1 dong du lieu binh thuong
        (build_clean_record: co email hop le thi vao ket qua, khong thi vao
        "Can kiem tra"), chen dung vi tri theo STT.
      - Khong tim thay (vd Dong Thap: nguon danh so sai, 2315 go thanh
        2015) -> KHONG bia dong nao, chi ghi CANH BAO de nguoi dung doi chieu.
    Tra ve danh sach records moi."""
    missing = find_stt_gaps(seen_stts)
    if not missing:
        return records
    lines = find_stt_lines(full_text, missing)
    not_found = [n for n in missing if n not in lines]

    for num in missing:
        combined = lines.get(num)
        if not combined:
            continue
        parts = split_recovered_stt_line(combined)
        row_context = {"source": source_label, "row_number": num, "stt": num}
        rec = build_clean_record(parts["email"] or "", parts["name"], parts["phone"] or "",
                                  row_stats=row_stats, row_context=row_context)
        if rec:
            pos = len(records)
            for idx, r in enumerate(records):
                ctx = r.get("_context") or {}
                if ctx.get("source") == source_label and ctx.get("stt") is not None and ctx["stt"] > num:
                    pos = idx
                    break
            records.insert(pos, rec)

    recovered = [n for n in missing if n in lines]
    if recovered and row_stats is not None:
        row_stats.setdefault("recovered_stts", []).extend((source_label, n) for n in recovered)
    warn_stt_gaps(seen_stts, row_stats, source_label, missing=not_found)
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


def _ascii_lower(text):
    text = str(text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn").lower()


def email_name_match_score(name, email):
    """Muc do KHOP giua ho ten va phan truoc @ cua email (0/1/2). Dung khi
    2 nguoi dung CHUNG 1 email (loi nhap lieu o nguon): giu nguoi co ten
    KHOP email hon, thay vi mac dinh giu nguoi xuat hien truoc - da gap
    thuc te nguoi xuat hien truoc la nguoi DAN NHAM email cua nguoi sau
    (vd 'Nguyễn Thị Thu Hà' dung 'hvluyen...' cua 'Hồ Văn Luyến').
      2 = khop mau viet tat pho bien: chu cai dau ho/ten dem + ten
          (vd 'hvluyen' cho 'Hồ Văn Luyến'), ten + chu cai dau
          (vd 'luyenhv'), hoac ho ten viet lien (vd 'hovanluyen').
      1 = chi chua ten (tu cuoi, >= 3 ky tu).
      0 = khong khop."""
    words = [w for w in re.split(r"[^a-z0-9]+", _ascii_lower(name)) if w]
    if not words or not email or "@" not in email:
        return 0
    local = _ascii_lower(email.split("@")[0])
    local_compact = re.sub(r"[^a-z0-9]", "", local)
    first_seg = re.split(r"[._\-]", local)[0]
    given = words[-1]
    initials = "".join(w[0] for w in words[:-1])
    patterns = {initials + given, given + initials, "".join(words)}
    if first_seg in patterns or any(p and local_compact.startswith(p) for p in patterns):
        return 2
    if len(given) >= 3 and given in local_compact:
        return 1
    return 0


def dedupe_by_email(records, row_stats=None):
    """Loai ban ghi trung email. Mac dinh giu ban ghi xuat hien dau tien,
    NHUNG neu ban ghi sau co ten KHOP email ro rang hon (xem
    email_name_match_score) thi GIU ban ghi sau va loai ban ghi truoc.
    Ban ghi bi loai duoc ghi vao row_stats['issues'] (ly do
    'duplicate_email') de hien o tab "Can kiem tra"."""
    kept = {}      # email -> vi tri trong result
    result = []
    dropped = []
    for r in records:
        email = r["email"]
        if email in kept:
            idx = kept[email]
            current = result[idx]
            if email_name_match_score(r["name"], email) > email_name_match_score(current["name"], email):
                dropped.append(current)
                result[idx] = r
            else:
                dropped.append(r)
            continue
        kept[email] = len(result)
        result.append(r)
    if row_stats is not None:
        for r in dropped:
            ctx = r.get("_context") or {}
            row_stats.setdefault("issues", []).append({
                "source": ctx.get("source", ""),
                "row_number": ctx.get("row_number"),
                "raw_email": r["email"],
                "raw_name": r["name"],
                "raw_phone": r["phone"],
                "reason": "duplicate_email",
            })
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
    print_warnings(stats.get("warnings"), stats.get("recovered_stts"))


def print_warnings(warnings, recovered_stts=None):
    """In cac dong khoi phuc duoc + canh bao (dung chung cho 2 che do)."""
    if recovered_stts:
        by_src = {}
        for src, n in recovered_stts:
            by_src.setdefault(src, []).append(n)
        for src, nums in by_src.items():
            print(f"Đã khôi phục STT {format_stt_list(nums)} từ văn bản gốc ({src}) - "
                  "pdfplumber bỏ sót khi tách bảng.")
    for w in warnings or []:
        print(f"CẢNH BÁO: {w}")


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
        - issues:                 cac dong bi loai (tab "Can kiem tra")
        - warnings:               canh bao dau hieu MAT DU LIEU AM THAM (STT
                                   nhay so khong tim lai duoc, ca tep khong co
                                   SDT, bang co email nhung khong nhan ra tieu
                                   de...) - chuoi co dau, hien cho nguoi dung
        - recovered_stts:         [(tep, stt)] cac dong PDF bi pdfplumber bo
                                   sot da khoi phuc tu van ban tho
    """
    all_records = []
    per_file_counts = {}
    row_stats = {
        "total_rows": 0,
        "valid_rows": 0,
        "missing_email": 0,
        "invalid_email_format": 0,
        "issues": [],
        "warnings": [],
        "recovered_stts": [],
    }
    for path in inputs:
        if verbose:
            print(f"Đang đọc: {path}")
        try:
            recs = extract_from_file(path, row_stats=row_stats)
        except Exception as e:
            print(f"  !! Lỗi khi đọc {path}: {e}", file=sys.stderr)
            per_file_counts[path] = 0
            add_warning(row_stats, f"{os.path.basename(path)}: không đọc được tệp ({e}).")
            continue
        if verbose:
            print(f"  -> Trích xuất được {len(recs)} bản ghi hợp lệ (có email)")
        per_file_counts[path] = len(recs)
        all_records.extend(recs)
        # Ca tep khong co SDT nao -> gan nhu chac chan cot SDT khong duoc
        # nhan dien (tieu de la hoac lech cot), khong phai nguon thieu that.
        if len(recs) >= 3 and not any(r["phone"] for r in recs):
            add_warning(row_stats, f"{os.path.basename(path)}: không có số điện thoại nào trong "
                                   f"{len(recs)} bản ghi - kiểm tra tiêu đề cột SĐT.")

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
        "warnings": row_stats["warnings"],
        "recovered_stts": row_stats["recovered_stts"],
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