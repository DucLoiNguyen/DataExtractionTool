#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
process_template_v2.py
=======================
Chuyen doi du lieu dang ky (STT | Ho va ten | Don vi cong tac | Email | SDT)
sang dinh dang template loai 2 (Ten tai khoan | Email | So dien thoai |
Mat khau | Gioi tinh | Ngay sinh | Don vi | To chuc).

File mau loai 2 la PHAN MO RONG cua file mau loai 1 (them cac truong Ngay
sinh, Gioi tinh, Don vi, To chuc), nen script nay AP DUNG DAY DU 3 QUY TAC
GOC cua file mau 1 (tai su dung truc tiep tu extract_contacts.py - can nam
CUNG THU MUC voi file nay), CONG THEM 2 QUY TAC RIENG cho Don vi/To chuc:

0) 3 QUY TAC GOC (tu file mau 1, xem chi tiet trong extract_contacts.py):
   - Email: chuan ve dinh dang hop le, viet thuong. Khong co email hop le
     (thieu hoac sai dinh dang) -> BO QUA ban ghi, KHONG ghi vao ket qua.
   - So dien thoai: chuan ve dang so VN (10 so, bat dau bang 0, khong ky tu
     dac biet, viet lien). Khong chuan hoa duoc -> de trong (van giu ban
     ghi, quy tac bo qua chi ap dung rieng cho email).
   - Ho va ten: viet hoa chu cai dau moi tu, bo ky tu dac biet.
   - Tu dong loai bo ban ghi TRUNG EMAIL (mac dinh bat, dung --no-dedupe de tat).
   - Cac dong bi loai (thieu/sai email, trung email) duoc ghi lai DAY DU vao
     1 tep bao cao rieng (<output>_can_kiem_tra.xlsx) kem STT/du lieu goc/ly
     do - khong chi bao 1 con so tong.

1) DON VI: chuan hoa ve dang "<cap hanh chinh> <ten rieng> - <tinh/thanh>".
   - Voi don vi hanh chinh xa/phuong/thi tran: bo tien to "UBND"/"HDND",
     giu "<Xa/Phuong/Thi tran> <Ten>", roi THEM ten tinh/thanh (dang ngan
     gon, khong tien to) o cuoi, ngan cach boi " - " (vd "UBND phuong Ninh
     Kieu" -> "Phuong Ninh Kieu - Can Tho").
   - Voi cac co quan/don vi khac (So, Ban, Trung tam, Truong...): giu
     nguyen ten rieng (chi mo rong cac tu viet tat pho bien: TT -> Trung
     tam, TP. -> Thanh pho, CDCD -> Cao dang Cong dong, QLDA -> Quan ly
     Du an...), CUNG THEM ten tinh/thanh o cuoi (vd "So Tu phap" -> "So Tu
     phap - Can Tho").
   - Neu ban than "Don vi cong tac" DA LA 1 don vi hanh chinh cap 1 (tinh/
     thanh pho) - vd chi ghi "Tinh Hau Giang" - thi Don vi = CHI ten rieng,
     bo tien to "tinh"/"thanh pho", KHONG them hau to (tranh lap: khong tra
     ve "Hau Giang - Can Tho").

2) TO CHUC: chuan hoa ve DUNG TEN CHINH THUC HIEN NAY cua don vi hanh chinh
   cap 1 (tinh/thanh pho) MA don vi do truc thuoc, VIET HOA TOAN BO. Tu
   dong nhan dien qua ten tinh/thanh CU hoac MOI xuat hien trong van ban
   nguon (vd gap "Soc Trang" hoac "Hau Giang" -> tu dong quy ve "THANH PHO
   CAN THO" vi 2 tinh nay da sap nhap vao TP Can Tho tu 1/7/2025). Neu
   khong tim thay dau hieu nao trong van ban nguon, dung gia tri
   `default_to_chuc` duoc chi dinh khi chay script (vi du toan bo file la
   cua 1 to chuc duy nhat).

6) Cac cot con lai (Mat khau, Gioi tinh, Ngay sinh) de TRONG - khong co du
   lieu tuong ung trong nguon nen khong tu bia.

Du lieu tham chieu PROVINCE_MERGE_MAP duoi day da duoc kiem chung qua tra
cuu Nghi quyet 202/2025/QH15 va bai dang chinh thuc tren Cong TTDT Chinh
phu (xaydungchinhsach.chinhphu.vn), hieu luc tu 12/6/2025 (chinh quyen moi
hoat dong tu 1/7/2025). Vi day la thong tin hanh chinh CO THE TIEP TUC
THAY DOI trong tuong lai, nen kiem tra lai neu dung cho du lieu cua thoi
diem khac.

Cach dung:
    python3 process_template_v2.py \
        --input DangKy_2109.xlsx \
        --template TemplateV2.xlsx \
        --output ket_qua_v2.xlsx \
        --to-chuc "Can Tho"

    --to-chuc: ten tinh/thanh (khong dau tien to) dung lam TO CHUC MAC DINH
    cho cac dong ma khong tu nhan dien duoc tu van ban nguon (thuong la moi
    dong, vi du "So Giao duc va Dao tao" khong tu no cho biet no thuoc tinh
    nao). Bat buoc phai co it nhat 1 trong 2: --to-chuc HOAC du lieu nguon
    tu no da du de nhan dien (hiem khi xay ra).
"""

import argparse
import csv
import os
import re
import sys
import unicodedata

import openpyxl
from openpyxl.styles import Font

# extract_contacts.py (tool file mau loai 1) can nam CUNG THU MUC voi file
# nay - tai su dung nguyen ven 3 quy tac chuan hoa da kiem chung cua no
# (email, so dien thoai, ho ten) vi file mau 2 la PHAN MO RONG cua file 1.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_contacts as core  # noqa: E402

# ---------------------------------------------------------------------------
# 1. BANG THAM CHIEU SAP NHAP TINH/THANH 2025 (hieu luc 12/6/2025)
#    Nguon: Nghi quyet 202/2025/QH15, Cong TTDT Chinh phu.
#    Cau truc: ten CHINH THUC HIEN NAY -> (loai: "tinh"/"thanh pho",
#              danh sach ten CU da sap nhap vao, KHONG ke chinh no).
# ---------------------------------------------------------------------------
PROVINCE_MERGE_MAP = {
    "Tuyên Quang": ("tỉnh", ["Hà Giang"]),
    "Lào Cai": ("tỉnh", ["Yên Bái"]),
    "Thái Nguyên": ("tỉnh", ["Bắc Kạn"]),
    "Phú Thọ": ("tỉnh", ["Vĩnh Phúc", "Hòa Bình"]),
    "Bắc Ninh": ("tỉnh", ["Bắc Giang"]),
    "Hưng Yên": ("tỉnh", ["Thái Bình"]),
    "Hải Phòng": ("thành phố", ["Hải Dương"]),
    "Ninh Bình": ("tỉnh", ["Hà Nam", "Nam Định"]),
    "Quảng Trị": ("tỉnh", ["Quảng Bình"]),
    "Đà Nẵng": ("thành phố", ["Quảng Nam"]),
    "Quảng Ngãi": ("tỉnh", ["Kon Tum"]),
    "Gia Lai": ("tỉnh", ["Bình Định"]),
    "Khánh Hòa": ("tỉnh", ["Ninh Thuận"]),
    "Lâm Đồng": ("tỉnh", ["Đắk Nông", "Bình Thuận"]),
    "Đắk Lắk": ("tỉnh", ["Phú Yên"]),
    "Hồ Chí Minh": ("thành phố", ["Bà Rịa - Vũng Tàu", "Bà Rịa Vũng Tàu", "Bình Dương"]),
    "Đồng Nai": ("tỉnh", ["Bình Phước"]),
    "Tây Ninh": ("tỉnh", ["Long An"]),
    "Cần Thơ": ("thành phố", ["Sóc Trăng", "Hậu Giang"]),
    "Vĩnh Long": ("tỉnh", ["Bến Tre", "Trà Vinh"]),
    "Đồng Tháp": ("tỉnh", ["Tiền Giang"]),
    "Cà Mau": ("tỉnh", ["Bạc Liêu"]),
    "An Giang": ("tỉnh", ["Kiên Giang"]),
    # 11 tinh/thanh KHONG sap nhap (giu nguyen) - van liet ke de tra cuu
    # dong bo, danh sach "cu" rong vi khong hop nhat voi tinh nao khac.
    "Cao Bằng": ("tỉnh", []),
    "Điện Biên": ("tỉnh", []),
    "Hà Tĩnh": ("tỉnh", []),
    "Lai Châu": ("tỉnh", []),
    "Lạng Sơn": ("tỉnh", []),
    "Nghệ An": ("tỉnh", []),
    "Quảng Ninh": ("tỉnh", []),
    "Thanh Hóa": ("tỉnh", []),
    "Sơn La": ("tỉnh", []),
    "Hà Nội": ("thành phố", []),
    "Huế": ("thành phố", []),
}

# Danh sach tat ca "ten cu" (tinh da bi sap nhap, khong con ton tai doc lap)
# tro ve ten CHINH THUC HIEN NAY - dung de tu dong nhan dien Tổ chức tu van
# ban nguon (vd gap "Soc Trang" trong 1 dong -> suy ra "Can Tho").
_OLD_NAME_TO_CURRENT = {}
for _current, (_loai, _old_names) in PROVINCE_MERGE_MAP.items():
    for _old in _old_names:
        _OLD_NAME_TO_CURRENT[_old] = _current
# Ban than ten hien tai cung phai tu nhan dien duoc (vd van ban ghi thang
# "Can Tho" - da la ten hien tai, khong can quy doi).
for _current in PROVINCE_MERGE_MAP:
    _OLD_NAME_TO_CURRENT.setdefault(_current, _current)


def _deaccent(text):
    """Bo dau tieng Viet (dung de so khop khong phan biet dau/khong dau)."""
    if not text:
        return ""
    text = str(text).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


def _norm_key(text):
    return _deaccent(text).lower().strip()


# Tra cuu nhanh (khong dau, thuong) -> ten hien tai, sap xep theo do dai
# GIAM DAN de uu tien khop cum dai truoc (vd "Ba Ria - Vung Tau" truoc "Ba Ria").
_LOOKUP = sorted(
    ((_norm_key(old), current) for old, current in _OLD_NAME_TO_CURRENT.items()),
    key=lambda x: len(x[0]), reverse=True,
)


def get_current_province_name(name):
    """Tra ve (ten_hien_tai, loai) tu 1 ten tinh/thanh (cu hoac moi), hoac
    (None, None) neu khong nhan ra."""
    key = _norm_key(name)
    for old_key, current in _LOOKUP:
        if old_key == key:
            loai = PROVINCE_MERGE_MAP[current][0]
            return current, loai
    return None, None


def detect_to_chuc_from_text(text):
    """Quet 1 doan van ban (vd 'Don vi cong tac') tim ten tinh/thanh (cu
    hoac moi) xuat hien trong do, tra ve TEN CHINH THUC HIEN NAY neu tim
    thay, hoac None neu khong co dau hieu nao."""
    if not text:
        return None
    key = _norm_key(text)
    for old_key, current in _LOOKUP:
        if not old_key:
            continue
        if re.search(r"(?<![a-z0-9])" + re.escape(old_key) + r"(?![a-z0-9])", key):
            return current
    return None


# ---------------------------------------------------------------------------
# 2. CHUAN HOA "DON VI"
# ---------------------------------------------------------------------------
_ADMIN_BODY_PATTERN = re.compile(
    r"^\s*(?:UBND|HĐND|Ủy\s*ban\s*nhân\s*dân|Hội\s*đồng\s*nhân\s*dân)\s+"
    r"(Xã|Phường|Thị\s*trấn)\s+(.+?)\s*$",
    re.IGNORECASE,
)

# Cac tu viet tat pho bien can mo rong trong ten don vi (khong phai admin-body)
_ABBREVIATIONS = [
    (re.compile(r"^\s*TT\b\.?", re.IGNORECASE), "Trung tâm"),
    (re.compile(r"\bTP\.\s*", re.IGNORECASE), "Thành phố "),
    (re.compile(r"\bCĐCĐ\b", re.IGNORECASE), "Cao đẳng Cộng đồng"),
    (re.compile(r"\bQLDA\b", re.IGNORECASE), "Quản lý Dự án"),
]

_ADMIN_LEVEL_DISPLAY = {"xã": "Xã", "phường": "Phường", "thị trấn": "Thị trấn"}

_PROVINCE_PREFIX_RE = re.compile(r"^\s*(?:tỉnh|thành\s*phố)\s+", re.IGNORECASE)
_ALL_PROVINCE_KEYS = {_norm_key(k) for k in _OLD_NAME_TO_CURRENT}


def _title_case_vn(text):
    """Viet hoa chu cai dau moi tu (giu nguyen dau tieng Viet)."""
    words = text.split()
    return " ".join(w[:1].upper() + w[1:] if w else w for w in words)


def _resolve_current_province(raw_text, default_to_chuc=None):
    """Xac dinh ten tinh/thanh HIEN TAI (khong tien to, khong viet hoa) ung
    voi 1 doan van ban Don vi cong tac.

    THU TU UU TIEN quan trong: NEU nguoi dung da chi dinh ro default_to_chuc
    va gia tri do KHONG PHAI 1 tinh/thanh (vd ten 1 vien nghien cuu/bo
    nganh cap trung uong nhu "Viện Năng lượng nguyên tử Việt Nam"), TIN
    TUONG HOAN TOAN default_to_chuc va KHONG tu dong do tim ten tinh/thanh
    trong van ban nua - vi rat de nham voi DIA DANH NAM SAN TRONG TEN
    RIENG cua don vi (vd "Trung tâm Chiếu xạ Hà Nội" khong co nghia nguoi
    do thuoc UBND Ha Noi, ma la ten rieng cua 1 trung tam truc thuoc vien
    trung uong). Chi khi default_to_chuc THUC SU la 1 tinh/thanh (hoac
    khong duoc chi dinh) thi moi dung ca che tu nhan dien qua van ban nhu
    truoc (huu ich cho cac file ma cac tinh/thanh cu con xuat hien trong
    van ban, vd Soc Trang/Hau Giang deu quy ve Can Tho)."""
    if default_to_chuc:
        matched, _loai = get_current_province_name(default_to_chuc)
        if not matched:
            return None  # to chuc la 1 thuc the khac tinh/thanh -> khong suy doan tu van ban nua
        current = detect_to_chuc_from_text(raw_text)
        return current or matched

    return detect_to_chuc_from_text(raw_text)


def normalize_don_vi(raw_don_vi_cong_tac, default_to_chuc=None):
    """Ap dung quy tac 1: tra ve chuoi Don vi da chuan hoa, THEM ten tinh/
    thanh (dang ngan gon, khong tien to, khong viet hoa toan bo) o CUOI,
    ngan cach boi ' - ' (vd 'UBND Xa A' -> 'Xa A - Can Tho', 'So Tu phap'
    -> 'So Tu phap - Can Tho') - CHI KHI thuc su xac dinh duoc 1 tinh/thanh
    THAT SU (tu van ban hoac tu default_to_chuc khop voi bang tra cuu).
    Neu to chuc KHONG PHAI tinh/thanh (vd 1 vien nghien cuu, bo/nganh cap
    trung uong - default_to_chuc khong khop bang tra cuu) thi KHONG them
    hau to nao (vi se thua/vo nghia, vd khong nen tra ve 'Vien Cong nghe
    xa hiem - Vien Nang luong nguyen tu Viet Nam').

    TRU truong hop chinh gia tri Don vi da la 1 don vi hanh chinh cap 1
    (tinh/thanh) - luc do CHI bo tien to "tỉnh"/"thành phố" va GIU NGUYEN
    ten nhu duoc ghi (KHONG doi sang ten hien tai neu la ten cu, KHONG them
    hau to) - vi day chinh la gia tri cap-1 duoc de cap, khong can/khong
    nen tu suy doan/doi ten no."""
    if not raw_don_vi_cong_tac:
        return ""
    text = str(raw_don_vi_cong_tac).strip()

    # Truong hop 1: don vi hanh chinh xa/phuong/thi tran voi tien to UBND/HDND
    m = _ADMIN_BODY_PATTERN.match(text)
    if m:
        level_raw, name = m.group(1), m.group(2)
        level_display = _ADMIN_LEVEL_DISPLAY.get(level_raw.lower().replace(" ", " "), level_raw.capitalize())
        core = f"{level_display} {_title_case_vn(name)}"
    else:
        # Truong hop 2: ban than gia tri la 1 don vi hanh chinh cap 1 (tinh/tp),
        # nhan biet qua tien to "tỉnh"/"thành phố", HOAC ca chuoi khop DUNG
        # BANG 1 ten tinh/thanh da biet (cu hoac moi) ma khong co gi khac.
        prefix_match = _PROVINCE_PREFIX_RE.match(text)
        candidate = text[prefix_match.end():].strip() if prefix_match else text
        if _norm_key(candidate) in _ALL_PROVINCE_KEYS:
            return candidate  # giu nguyen ten nhu ghi, khong doi ten, khong them hau to

        # Truong hop 3: co quan/don vi khac - mo rong viet tat, giu nguyen ten
        for pattern, replacement in _ABBREVIATIONS:
            text = pattern.sub(replacement, text)
        core = re.sub(r"\s+", " ", text).strip()

    province = _resolve_current_province(raw_don_vi_cong_tac, default_to_chuc=default_to_chuc)
    if province:
        return f"{core} - {province}"
    return core


# ---------------------------------------------------------------------------
# 3. XAC DINH "TO CHUC"
# ---------------------------------------------------------------------------
def determine_to_chuc(raw_don_vi_cong_tac, default_to_chuc=None):
    """Ap dung quy tac 2: tra ve chuoi TO CHUC da VIET HOA TOAN BO, hoac ""
    neu khong xac dinh duoc (khong tim thay trong van ban VA khong co
    default_to_chuc).

    - Neu default_to_chuc duoc chi dinh nhung KHONG khop voi tinh/thanh nao
      (vd ten 1 bo/nganh, vien nghien cuu cap trung uong nhu "Viện Năng
      lượng nguyên tử Việt Nam") thi TIN TUONG HOAN TOAN gia tri do, dung
      NGUYEN VAN (chi viet hoa toan bo) - KHONG tu dong do tim ten tinh/
      thanh trong van ban (tranh nham voi dia danh nam san trong ten rieng
      cua don vi, vd "Trung tâm Chiếu xạ Hà Nội" - xem giai thich chi tiet
      trong _resolve_current_province).
    - Nguoc lai (default_to_chuc la 1 tinh/thanh, hoac khong duoc chi dinh):
      nhan dien tinh/thanh tu van ban (uu tien) hoac tu default_to_chuc,
      dinh dang "<tỉnh/thành phố> <TEN>".
    """
    if default_to_chuc:
        matched, loai = get_current_province_name(default_to_chuc)
        if not matched:
            return default_to_chuc.strip().upper()
        current = detect_to_chuc_from_text(raw_don_vi_cong_tac)
        if current:
            loai2 = PROVINCE_MERGE_MAP.get(current, ("tỉnh", []))[0]
            return f"{loai2} {current}".upper()
        return f"{loai} {matched}".upper()

    current = detect_to_chuc_from_text(raw_don_vi_cong_tac)
    if current:
        loai = PROVINCE_MERGE_MAP.get(current, ("tỉnh", []))[0]
        return f"{loai} {current}".upper()
    return ""


# ---------------------------------------------------------------------------
# 4. DOC NGUON + GHI RA TEMPLATE
# ---------------------------------------------------------------------------
# Tu khoa EMAIL/SDT lay THANG tu core.HEADER_KEYWORDS - 2 che do dung 1
# nguon duy nhat, them tieu de moi (vd "Di dong", "Dia chi thu cong vu") o
# extract_contacts.py la tu dong ap dung ca 2 che do (truoc day 2 danh sach
# chep tay, da tung lech nhau).
# Tu khoa TEN giu rieng: Che do 2 KHONG dung tu chung "ten" nhu core (core
# loai cot "Ten don vi" nho NOISE_KEYWORDS truoc; Che do 2 lai CAN cot don vi
# nen "ten" se nhan nham "Tên đơn vị" thanh cot ten nguoi).
SOURCE_HEADER_KEYWORDS = {
    "stt": "stt",
    "tt": "stt",
    "ho va ten": "name",
    "ho ten": "name",
    "hoten": "name",
    "ten tai khoan": "name",
    "ten nguoi dung": "name",
    "full name": "name",
}
for _field in ("email", "phone"):
    for _kw in core.HEADER_KEYWORDS[_field]:
        SOURCE_HEADER_KEYWORDS.setdefault(_kw, _field)
# Cot "Don vi" duoc xu ly RIENG theo 2 tang uu tien (xem _find_source_columns):
# uu tien 1 la cac cot NEU RO TEN DON VI/PHONG BAN (huu ich hon cho viec
# chuan hoa Don vi/To chuc), uu tien 2 (chi dung khi KHONG co cot uu tien 1)
# la cac cot "Chuc vu/Vi tri" - vi cac phieu dang ky thuong co CA 2 cot
# rieng biet (vd "Phong ban/Bo phan" VA "Chuc vu"), va ten phong ban huu ich
# hon nhieu so voi chuc danh cong viec khi dung lam "Don vi cong tac".
UNIT_PRIMARY_KEYWORDS = ["don vi cong tac", "don vi", "phong ban", "bo phan", "co quan", "noi cong tac"]
UNIT_FALLBACK_KEYWORDS = ["chuc vu", "chuc danh", "vi tri cong tac", "vi tri"]


def _find_source_columns(header_row):
    col_map = {}
    email_candidates = []
    for idx, cell in enumerate(header_row):
        key = _norm_key(cell)
        for kw, field in SOURCE_HEADER_KEYWORDS.items():
            if kw in key:
                if field == "email":
                    # Co the co NHIEU cot cung khop tu khoa email (vd "Thư
                    # điện tử công vụ" VA "Thư điện tử Gmail" - da gap thuc
                    # te: nhieu dong de trong cot dau, ghi email o cot sau
                    # thay the) - luu lai TAT CA de dung lam du phong, thay
                    # vi chi lay 1 cot roi bo qua cac cot con lai.
                    if idx not in email_candidates:
                        email_candidates.append(idx)
                elif field not in col_map:
                    col_map[field] = idx
    for idx in range(1, len(email_candidates)):
        # Cot email THU 2 tro di duoc luu la "email_alt" (chi 1 cot du
        # phong la du dung cho hau het truong hop thuc te).
        col_map.setdefault("email_alt", email_candidates[idx])
    if email_candidates:
        col_map["email"] = email_candidates[0]
    # QUAN TRONG: quet THEO THU TU TU KHOA (tu cu the nhat den chung chung
    # nhat) TRUOC, roi moi quet cot - KHONG lam nguoc lai (cot truoc, tu
    # khoa sau) - vi da gap thuc te 2 cot CUNG chua chu "don vi": "Tên đơn
    # vị" (chi co gia tri o DONG DAU moi nhom do gop o trong Excel, phan
    # lon RONG) va "Cơ quan, đơn vị công tác" (day du tren MOI dong). Neu
    # quet theo cot truoc, cot "Tên đơn vị" (thuong nam BEN TRAI hon, khop
    # tu khoa chung "don vi") se duoc chon TRUOC ca khi cot con lai moi la
    # cot dung/day du - quet theo tu khoa cu the truoc (vd "don vi cong
    # tac" truoc "don vi") dam bao chon dung cot ngay ca khi no nam ben
    # phai cot kia.
    for kw in UNIT_PRIMARY_KEYWORDS:
        if "unit" in col_map:
            break
        for idx, cell in enumerate(header_row):
            if kw in _norm_key(cell):
                col_map["unit"] = idx
                break
    for kw in UNIT_FALLBACK_KEYWORDS:
        if "unit" in col_map:
            break
        for idx, cell in enumerate(header_row):
            if kw in _norm_key(cell):
                col_map["unit"] = idx
                break
    # Tieu de GOP O chia thanh 2 cot lien ke CUNG TEN (vd 2 cot "Cơ quan,
    # đơn vị công tác": cot 1 ghi bo phan "HĐND"/"Phòng VHXH", cot 2 ghi
    # "Xã Long Hòa") - danh dau cot thu 2 de GHEP gia tri, tranh mat phan
    # ten xa/phuong (da gap thuc te).
    u = col_map.get("unit")
    if u is not None and u + 1 < len(header_row):
        if _norm_key(header_row[u]) and _norm_key(header_row[u + 1]) == _norm_key(header_row[u]):
            col_map["unit2"] = u + 1
    return col_map


def _find_source_header_row(rows, max_scan=15):
    """Quet toi da max_scan dong dau tien de tim dong TIEU DE THAT SU (co
    nhan dien duoc it nhat cot 'name' hoac 'email'), UU TIEN dong cho ra
    NHIEU cot nhan dien duoc nhat - vi khong phai luc nao dong 0 cung la
    tieu de (nhieu file co dong van ban tieu de/ghi chu truoc do, hoac co
    ca dong 'khoa ky thuat' lan dong 'nhan tieng Viet' - can chon dong day
    du thong tin hon). Tra ve (chi_so_dong, col_map); (None, {}) neu khong
    tim thay dong nao phu hop."""
    best_idx, best_map = None, {}
    for idx, row in enumerate(rows[:max_scan]):
        col_map = _find_source_columns(row)
        if ("name" in col_map or "email" in col_map) and len(col_map) > len(best_map):
            best_idx, best_map = idx, col_map
    return best_idx, best_map


def _cell_display(value):
    """Chuyen 1 gia tri o thanh chuoi HIEN THI gon gang (vd so thuc 1.0 ->
    '1', KHONG phai '1.0') - chi dung de HIEN THI/ghi vao bao cao issues.
    KHONG dung ket qua nay lam dau vao cho core.normalize_email/name/phone
    - cac ham do tu xu ly kieu float/int rieng (vd SDT dang so thuc), neu
    da bi ep thanh chuoi co duoi '.0' truoc thi se bi hong (mat SDT)."""
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()


def _looks_like_group_header(row, col_map):
    """1 dong duoc coi la 'tieu de nhom' (vd danh dau bang so La Ma 'I',
    'II'... ten 1 vien/trung tam lon hon ma cac dong nguoi tiep theo thuoc
    ve) neu dong đó KHONG co email/SDT, va trong 2 cot LIEN QUAN (ten, don
    vi) CHI co dung 1 cot chua noi dung. CO Y BO QUA cot STT khi dem - vi
    dong tieu de nhom thuong VAN co danh dau o cot STT (vd so La Ma 'I'),
    khong phai lac dong day la 1 dong du lieu that."""
    email_idx = col_map.get("email")
    phone_idx = col_map.get("phone")
    has_email = email_idx is not None and email_idx < len(row) and row[email_idx] not in (None, "")
    has_phone = phone_idx is not None and phone_idx < len(row) and row[phone_idx] not in (None, "")
    if has_email or has_phone:
        return False
    relevant_idxs = [i for i in (col_map.get("name"), col_map.get("unit")) if i is not None]
    if not relevant_idxs:
        return False
    non_empty_relevant = [row[i] for i in relevant_idxs if i < len(row) and row[i] not in (None, "")]
    return len(non_empty_relevant) == 1


def _group_header_text(row, col_map):
    """Lay TEN NHOM tu 1 dong tieu de nhom: uu tien cot 'name' hoac 'unit'
    (KHONG lay cot STT, vi cot do chi chua so La Ma/ky hieu danh dau, khong
    phai ten nhom)."""
    for field in ("name", "unit"):
        idx = col_map.get(field)
        if idx is not None and idx < len(row) and row[idx] not in (None, ""):
            return _cell_display(row[idx])
    stt_idx = col_map.get("stt")
    for i, c in enumerate(row):
        if i == stt_idx or c in (None, ""):
            continue
        return _cell_display(c)
    return ""


def _read_tables_and_context(input_path):
    """Doc 1 tep nguon o BAT KY dinh dang nao (.xlsx/.xls/.pdf/.docx/.csv),
    tra ve (list_of_tables, fallback_unit_text, content_based, full_text):
      - list_of_tables: danh sach cac 'bang' (moi bang la list rows) - Excel/
        CSV chi co 1 bang (ca sheet), con PDF/Word co the co NHIEU bang
        (thuong da duoc gop trang qua core.merge_multi_page_tables).
      - fallback_unit_text: ten don vi trich duoc tu PHAN VAN BAN TU DO cua
        tai lieu (vd dong "Tên Cơ quan, đơn vị: Phòng Kinh tế xã Kon Đào"
        thuong thay o dau cac phieu dang ky) - dung lam Don vi MAC DINH cho
        CA FILE khi bang du lieu KHONG co cot Don vi/Phong ban rieng tren
        tung dong. None neu khong tim thay hoac khong ap dung (Excel/CSV).
      - content_based: True NEU dinh dang nay de bi ngat dong
        GIUA 1 o khi doc bang (PDF/Word, do gioi han khong gian trang/wrap
        chu) - CHI khi do moi nen goi core.merge_wrapped_continuation_rows.
        Voi Excel/CSV (du lieu bang GOC, khong bi "ngat dong" theo nghia
        nay), KHONG duoc goi ham do - 1 dong thieu ten (nhung co du lieu
        khac) trong Excel la 1 VAN DE CHAT LUONG DU LIEU THAT (can bao vao
        "Can kiem tra"), KHONG PHAI phan tiep cua dong truoc - goi nham ham
        gop dong se lam MAT 1 nguoi (gop nham email cua ho vao nguoi truoc,
        day la loi da gap va vua duoc sua).
      - full_text: TOAN BO van ban tho cua tai lieu (PDF/Word), dong theo
        dong - dung lam LUOI AN TOAN BO SUNG: pdfplumber doi khi BO SOT
        HOAN TOAN 1 vai dong khi tach bang (khac voi loi lech cot da sua o
        tren - day la truong hop dong khong xuat hien trong bang tach duoc
        chut nao). Sau khi xu ly bang, extract_v2 do STT bi "nhay so" va
        tim lai dong tuong ung trong full_text de khoi phuc (xem
        _recover_missing_stt_rows). None voi Excel/CSV (khong can/ap dung)."""
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".xls":
        core._patch_xlrd_tolerant_datemode()
        import xlrd
        wb = xlrd.open_workbook(input_path)
        sheet = wb.sheet_by_index(0)
        rows = core._rows_from_xlrd_sheet(sheet)
        return [rows], None, False, None

    if ext in (".xlsx", ".xlsm"):
        src_wb = openpyxl.load_workbook(input_path, data_only=True)
        src_ws = src_wb.active
        rows = [list(row) for row in src_ws.iter_rows(values_only=True)]
        return [rows], None, False, None

    if ext == ".csv":
        with open(input_path, newline="", encoding="utf-8-sig", errors="ignore") as f:
            rows = [list(row) for row in csv.reader(f)]
        return [rows], None, False, None

    if ext == ".pdf":
        import pdfplumber
        raw_tables = []
        context_text_parts = []  # TOAN BO van ban moi trang (bat ke co bang hay khong) - CHI
                                  # dung de do "Tên Cơ quan, đơn vị: X" lam Don vi du phong,
                                  # KHONG dung de trich xuat du lieu (tranh trung lap voi bang)
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables() or []
                if page_tables:
                    raw_tables.extend(page_tables)
                context_text_parts.append(page.extract_text() or "")
        flat_table = _flatten_doc_tables(raw_tables)
        full_text = "\n".join(context_text_parts)
        return [flat_table], _extract_doc_header_unit(full_text), True, full_text

    if ext in (".docx", ".doc"):
        import docx
        doc_path = input_path
        converted_path = None
        if ext == ".doc":
            # File Word 97-2003 cu - python-docx khong doc duoc, tu dong
            # chuyen doi qua LibreOffice truoc (xem core._convert_via_libreoffice).
            converted_path = core._convert_via_libreoffice(input_path, "docx")
            if converted_path is None:
                raise ValueError(
                    "Không thể đọc tệp .doc này: cần LibreOffice (soffice) trên máy để tự động "
                    "chuyển đổi nhưng không tìm thấy. Cách khắc phục: mở tệp bằng Word và 'Save As' "
                    "sang định dạng .docx rồi thử lại, hoặc cài LibreOffice (https://www.libreoffice.org/)."
                )
            doc_path = converted_path
        try:
            document = docx.Document(doc_path)
            raw_tables = [
                [[core._full_cell_text(cell) for cell in row.cells] for row in table.rows]
                for table in document.tables
            ]
            flat_table = _flatten_doc_tables(raw_tables)
            free_text = "\n".join(core._full_paragraph_text(p) for p in document.paragraphs)
            return [flat_table], _extract_doc_header_unit(free_text), True, None
        finally:
            if converted_path:
                try:
                    os.remove(converted_path)
                    os.rmdir(os.path.dirname(converted_path))
                except OSError:
                    pass

    raise ValueError(f"Không hỗ trợ định dạng tệp: {ext} ({input_path})")


_DOC_UNIT_CONTEXT_RE = re.compile(
    r"(?:Tên\s*Cơ\s*quan,?\s*đơn\s*vị|Cơ\s*quan,?\s*đơn\s*vị)\s*[:\-]\s*(.+)",
    re.IGNORECASE,
)

# Dong tieu de kieu "Danh sách đăng ký ... của Văn phòng UBND thành phố Đà Nẵng"
# - lay phan SAU tu "của" cuoi cung lam ten don vi.
_TITLE_UNIT_RE = re.compile(r"\bcủa\s+(.+)$", re.IGNORECASE)


def _extract_title_unit_from_rows(rows, header_idx):
    """Voi Excel/CSV: tim ten don vi trong cac DONG PHIA TREN dong tieu de
    cot (thuong la dong tieu de van ban, vd 'Danh sách ... của Văn phòng
    UBND thành phố Đà Nẵng'). Dung lam Don vi MAC DINH khi bang KHONG co
    cot Don vi rieng tren tung dong. Tra ve None neu khong tim thay."""
    if not header_idx:
        return None
    texts = []
    for row in rows[:header_idx]:
        for c in row or []:
            if c not in (None, "") and isinstance(c, str):
                texts.append(c.strip())
    joined = "\n".join(texts)
    unit = _extract_doc_header_unit(joined)
    if unit:
        return unit
    for line in joined.split("\n"):
        m = _TITLE_UNIT_RE.search(line.strip())
        if m:
            value = m.group(1).strip(" .;:)(")
            if value and len(value) <= 150:
                return value
    return None


def _extract_doc_header_unit(free_text):
    """Tim dong kieu 'Tên Cơ quan, đơn vị: X' trong phan van ban tu do o
    dau tai lieu (thuong gap trong cac phieu dang ky dang PDF/Word) - dung
    lam Don vi MAC DINH cho ca file khi bang du lieu khong co cot Don vi
    rieng tren tung dong. Tra ve None neu khong tim thay."""
    if not free_text:
        return None
    m = _DOC_UNIT_CONTEXT_RE.search(free_text)
    if not m:
        return None
    value = m.group(1).strip().split("\n")[0].strip()
    return value or None


_PHONE_LOOSE_REGEX = re.compile(r"(?:\+?84|0)\d{8,10}")


def _flatten_doc_tables(raw_tables):
    """Gop TAT CA bang tho tu 1 tai lieu PDF/Word (nhieu trang) thanh 1
    DANH SACH DONG DUY NHAT - THAY THE cho core.merge_multi_page_tables khi
    dung cho Che do 2 (merge_multi_page_tables coi so cot khac nhau la '1
    bang logic MOI' va cat dut du lieu tu do, lam MAT CA CHUC NGUOI neu
    dieu nay xay ra giua danh sach dai nhieu trang - da gap thuc te).

    KHONG co gang can chinh vi tri cot o day nua (xem ly do trong
    _parse_row_by_content) - chi don gian NOI tat ca dong lai, de buoc xu
    ly phia sau (_process_data_rows_content_based) tu nhan dang tung
    truong THEO NOI DUNG cua o, khong phu thuoc vi tri cot - vi thuc te da
    gap file PDF ma CAC TRANG CO CUNG SO COT nhung Y NGHIA cot KHAC NHAU
    (vd trang nay email o cot 4, trang khac email lai o cot 6).

    Gia dinh: toan bo tai lieu la MOT danh sach lien tuc (dung voi hau het
    phieu/danh sach dang ky thuc te). QUET QUA TAT CA cac bang de tim bang
    DAU TIEN co dong tieu de hop le (KHONG mac dinh la bang dau tien cua
    tai lieu) - vi file Word thuong co 1-2 bang NHO KHONG lien quan DUNG
    TRUOC bang du lieu chinh (vd bang "quoc hieu/tieu ngu" 2 cot ghi ten co
    quan + "CONG HOA XA HOI CHU NGHIA VIET NAM..."), da gap thuc te khien
    TOAN BO du lieu bi bo qua neu chi kiem tra bang dau tien."""
    if not raw_tables:
        return []

    start_idx, header_idx, col_map = None, None, None
    for i, table in enumerate(raw_tables):
        h_idx, c_map = _find_source_header_row(table)
        if h_idx is not None and ("name" in c_map or "email" in c_map):
            start_idx, header_idx, col_map = i, h_idx, c_map
            break

    if start_idx is None:
        return list(raw_tables[0])

    def _is_repeated_header(row):
        # PDF/Word nhieu trang thuong LAP LAI dong tieu de bang o dau moi
        # trang - khong duoc coi la 1 dong du lieu (se thanh 1 "nguoi thieu
        # email" gia trong tab Can kiem tra - da gap thuc te).
        if row is None:
            return False
        if any(c and "@" in str(c) for c in row):
            return False
        cm = _find_source_columns(row)
        return "name" in cm and ("email" in cm or "phone" in cm)

    first_table = raw_tables[start_idx]
    result = [first_table[header_idx]]  # giu dong tieu de o vi tri dau
    result.extend(r for r in first_table[header_idx + 1:] if not _is_repeated_header(r))
    for table in raw_tables[start_idx + 1:]:
        result.extend(r for r in table if r is not None and not _is_repeated_header(r))
    return result


def _looks_like_real_unit_name(text):
    """True neu text TRONG GIONG 1 ten don vi/phong ban THAT (vd 'Văn
    phòng Đảng ủy', 'Sở Xây dựng') - dau hieu don gian nhung hieu qua: ten
    don vi that hau nhu LUON co khoang trang giua cac tu; nguoc lai, ten
    tai khoan/ma nhan vien (vd 'dinhlx_itrre', 'tranchithanh') hau nhu
    KHONG CO khoang trang. Dung de quyet dinh co nen tin cot 'Don vi cong
    tac' tren TUNG DONG hay uu tien dung ten nhom (tieu de La Ma) thay the
    - xem _process_data_rows."""
    if not text:
        return False
    return " " in str(text).strip()


def _parse_row_by_content(row):
    """Trich STT/ten/don vi/email/SDT tu 1 dong THEO NOI DUNG cua tung o,
    KHONG phu thuoc vi tri cot co dinh. Dung cho nguon PDF/Word - da gap
    thuc te NHIEU TRANG CUNG 1 FILE co SO COT GIONG NHAU nhung Y NGHIA cot
    KHAC NHAU (vd trang 2 co email o cot thu 4, trang 4 cua CUNG file do
    lai co email o cot thu 6) - doc theo vi tri co dinh se doc nham truong
    tren nhung trang co bo cuc khac. Nhan dang email/SDT qua regex (dang
    tin cay, it nham lan); STT la o dau tien (so hoac so La Ma); Ten/Don
    vi la 2 o VAN BAN con lai theo THU TU xuat hien (ten truoc, don vi
    sau - dung voi hau het mau bieu thuc te)."""
    cells = list(row)
    used = set()

    email_val, email_pos = None, None
    for i, c in enumerate(cells):
        if c and core.EMAIL_REGEX.search(str(c)):
            email_val, email_pos = c, i
            used.add(i)
            break

    phone_val = None
    for i, c in enumerate(cells):
        if i in used or not c:
            continue
        if _PHONE_LOOSE_REGEX.search(str(c).replace(" ", "")):
            phone_val = c
            used.add(i)
            break

    stt_val = cells[0] if cells else None
    if 0 not in used:
        used.add(0)

    text_cells = []
    for i, c in enumerate(cells):
        if i in used or c is None:
            continue
        text = str(c).strip()
        if text:
            text_cells.append(text)

    name_val = text_cells[0] if len(text_cells) >= 1 else None
    unit_val = text_cells[1] if len(text_cells) >= 2 else None

    return {
        "stt": _cell_display(stt_val), "name": _cell_display(name_val),
        "unit": _cell_display(unit_val), "email": email_val, "phone": phone_val,
    }


def _row_has_identity_content(row):
    """True neu dong CO email hoac SDT (nhan dang qua regex) - dung de
    phan biet 1 dong DU LIEU THAT voi 1 dong tieu de nhom/dong rac."""
    for c in row:
        if not c:
            continue
        s = str(c)
        if core.EMAIL_REGEX.search(s) or _PHONE_LOOSE_REGEX.search(s.replace(" ", "")):
            return True
    return False


def _process_data_rows_content_based(data_rows, default_to_chuc, unit_cache, issues, clean_records,
                                      fallback_unit=None):
    """Ban THEO NOI DUNG (khong dung vi tri cot co dinh) cua
    _process_data_rows - dung cho PDF/Word khi cac trang trong CUNG 1 file
    co the co bo cuc cot khac nhau (xem _parse_row_by_content). Logic con
    lai (3 quy tac goc, Don vi/To chuc, tieu de nhom) giu nguyen tinh than
    nhu ham goc."""
    current_group = None
    group_header_seen = False
    total_read = 0
    pending_records = []  # (index trong clean_records DA THEM, hay chi theo doi ban ghi "dang mo") -
                           # dung con tro toi ban ghi cuoi CHUA bi day vao issues/clean_records de
                           # co the "va" them SDT/email neu gap dong manh vo tiep theo (xem duoi)
    last_added_ref = {"kind": None, "data": None}  # theo doi ban ghi/issue VUA THEM de vah them du lieu neu can

    for row in data_rows:
        if row is None or all(c in (None, "") for c in row):
            continue

        has_identity = _row_has_identity_content(row)
        non_empty = [c for c in row if c not in (None, "")]

        if not has_identity and len(non_empty) <= 2:
            # Khong co email/SDT VA rat it o co noi dung -> nhieu kha nang
            # la dong tieu de nhom (vd 'I  ĐẢNG ỦY') - cap nhat nhom hien
            # tai, khong tinh la 1 dong du lieu.
            text_vals = [str(c).strip() for c in non_empty if str(c).strip()]
            if text_vals:
                current_group = text_vals[-1]  # gia tri van ban (khong phai so La Ma o dau)
                group_header_seen = True
            continue

        parsed = _parse_row_by_content(row)

        # Dong "manh vo" (do pdfplumber tach 1 dong logic thanh nhieu dong
        # vat ly): co dung 1 trong 2 (email HOAC SDT) nhung KHONG co ten -
        # ghep bo sung vao ban ghi/issue VUA THEM o vong lap truoc, thay vi
        # tao 1 "nguoi" gia (chi co SDT, khong ten) trong ket qua/Can kiem tra.
        if not parsed["name"] and last_added_ref["kind"]:
            target = last_added_ref["data"]
            if parsed["email"] and not target.get("email"):
                new_email = core.normalize_email(parsed["email"])
                if new_email and last_added_ref["kind"] == "issue":
                    # Email vua tim thay giup ban ghi truoc TRO NEN HOP LE ->
                    # chuyen tu issues sang clean_records.
                    issues.remove(target)
                    name = core.normalize_name(target["raw_name"])
                    phone = core.normalize_phone(target["raw_phone"])
                    raw_unit = target["raw_unit"]
                    if raw_unit not in unit_cache:
                        unit_cache[raw_unit] = (
                            normalize_don_vi(raw_unit, default_to_chuc=default_to_chuc),
                            determine_to_chuc(raw_unit, default_to_chuc=default_to_chuc),
                        )
                    don_vi, to_chuc = unit_cache[raw_unit]
                    rec = {"stt": target["stt"], "name": name, "email": new_email,
                           "phone": phone, "don_vi": don_vi, "to_chuc": to_chuc}
                    clean_records.append(rec)
                    last_added_ref["kind"], last_added_ref["data"] = "record", rec
                elif last_added_ref["kind"] == "record":
                    target["email"] = new_email or target["email"]
            if parsed["phone"] and not target.get("phone"):
                new_phone = core.normalize_phone(parsed["phone"])
                if last_added_ref["kind"] == "record" and not target.get("phone"):
                    target["phone"] = new_phone
                elif last_added_ref["kind"] == "issue":
                    target["raw_phone"] = _cell_display(parsed["phone"])
            continue

        raw_stt = parsed["stt"]
        raw_name = parsed["name"]
        raw_phone_val = parsed["phone"]
        raw_phone = _cell_display(raw_phone_val)
        raw_unit_from_col = parsed["unit"] or ""
        if group_header_seen and not _looks_like_real_unit_name(raw_unit_from_col):
            raw_unit = current_group or raw_unit_from_col
        else:
            raw_unit = raw_unit_from_col
        if not raw_unit and fallback_unit:
            raw_unit = fallback_unit

        raw_email_val = parsed["email"]
        raw_email = _cell_display(raw_email_val)
        if not raw_name and not raw_email and not raw_phone and not raw_unit:
            continue
        total_read += 1

        email = core.normalize_email(raw_email_val)
        if not email:
            reason = "invalid_email_format" if raw_email else "missing_email"
            issue = {
                "stt": raw_stt, "raw_name": raw_name, "raw_email": raw_email,
                "raw_phone": raw_phone, "raw_unit": raw_unit, "reason": reason,
            }
            issues.append(issue)
            last_added_ref["kind"], last_added_ref["data"] = "issue", issue
            continue

        name = core.normalize_name(raw_name)
        phone = core.normalize_phone(raw_phone_val)

        if raw_unit not in unit_cache:
            unit_cache[raw_unit] = (
                normalize_don_vi(raw_unit, default_to_chuc=default_to_chuc),
                determine_to_chuc(raw_unit, default_to_chuc=default_to_chuc),
            )
        don_vi, to_chuc = unit_cache[raw_unit]

        rec = {"stt": raw_stt, "name": name, "email": email, "phone": phone,
               "don_vi": don_vi, "to_chuc": to_chuc}
        clean_records.append(rec)
        last_added_ref["kind"], last_added_ref["data"] = "record", rec

    return total_read


def _process_data_rows(data_rows, col_map, default_to_chuc, unit_cache, issues,
                        clean_records, fallback_unit=None):
    """Xu ly 1 danh sach data_rows (CUA 1 BANG, sau dong tieu de) theo dung
    3 quy tac goc + 2 quy tac Don vi/To chuc, GOM VAO CHUNG unit_cache/
    issues/clean_records duoc truyen tu ben ngoai (de cong don duoc qua
    NHIEU bang trong cung 1 tep - vd PDF nhieu trang co nhieu bang rieng).
    Tra ve total_read cua RIENG lan goi nay (de cong don o ham goi)."""

    def get_raw(row, field):
        idx = col_map.get(field)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    def get_display(row, field):
        return _cell_display(get_raw(row, field))

    def get_email_raw(row):
        """Uu tien cot email CHINH; neu rong VA co cot email DU PHONG (vd
        'Thư điện tử Gmail' khi 'Thư điện tử công vụ' de trong), dung cot
        du phong thay the - xem _find_source_columns."""
        primary = get_raw(row, "email")
        if primary not in (None, ""):
            return primary
        return get_raw(row, "email_alt")

    current_group = None
    group_header_seen = False
    total_read = 0
    for row in data_rows:
        if row is None or all(c in (None, "") for c in row):
            continue

        if _looks_like_group_header(row, col_map):
            current_group = _group_header_text(row, col_map)
            group_header_seen = True
            continue  # dong tieu de nhom - CHI cap nhat nhom hien tai, khong tinh la 1 dong du lieu

        raw_stt = get_display(row, "stt")
        raw_name = get_display(row, "name")
        raw_email = _cell_display(get_email_raw(row))
        raw_phone = get_display(row, "phone")
        raw_unit_from_col = get_display(row, "unit")
        unit2 = get_display(row, "unit2") if "unit2" in col_map else ""
        if unit2 and _norm_key(unit2) != _norm_key(raw_unit_from_col):
            raw_unit_from_col = f"{raw_unit_from_col} {unit2}".strip() if raw_unit_from_col else unit2
        if group_header_seen and not _looks_like_real_unit_name(raw_unit_from_col):
            # File nay dung tieu de nhom de gom don vi (xem docstring
            # extract_v2), VA gia tri cot "Don vi cong tac" tren dong nay
            # (neu co) KHONG giong 1 ten don vi that (vd la ten tai khoan
            # kieu "dinhlx_itrre", khong co khoang trang) -> uu tien dung
            # ten nhom hien tai thay the. NGUOC LAI, neu cot Don vi tren
            # chinh dong nay DA la 1 ten don vi hop ly (co khoang trang,
            # vd "Văn phòng Đảng ủy") thi GIU NGUYEN gia tri do - CHI TIET
            # HON va DUNG HON la ten nhom chung chung o muc tren (vd chi
            # dung "ĐẢNG ỦY" cho ca nhom se lam mat thong tin phong/ban cu
            # the ma file da co san).
            raw_unit = current_group or raw_unit_from_col
        else:
            raw_unit = raw_unit_from_col
        if not raw_unit and fallback_unit:
            # Khong co cot Don vi rieng tren tung dong (VA khong dung co
            # che nhom o tren) - dung Don vi trich tu phan van ban tu do
            # dau tai lieu (vd "Tên Cơ quan, đơn vị: X") cho MOI nguoi
            # trong bang nay (thuong gap o cac phieu dang ky 1 don vi
            # duy nhat, vd file PDF cua 1 phong/xa cu the).
            raw_unit = fallback_unit
        if not raw_name and not raw_email and not raw_phone and not raw_unit:
            continue  # dong hoan toan trong - bo qua, khong tinh vao thong ke
        total_read += 1

        # --- Ap dung 3 quy tac goc cua file mau 1 (dung GIA TRI THO, chua
        # ep kieu, de khong lam hong SDT/email dang so) ---
        email = core.normalize_email(get_email_raw(row))
        if not email:
            reason = "invalid_email_format" if raw_email else "missing_email"
            issues.append({
                "stt": raw_stt, "raw_name": raw_name, "raw_email": raw_email,
                "raw_phone": raw_phone, "raw_unit": raw_unit, "reason": reason,
            })
            continue  # Quy tac 1: khong co email hop le -> bo qua ban ghi

        name = core.normalize_name(raw_name)
        phone = core.normalize_phone(get_raw(row, "phone"))

        # --- Ap dung 2 quy tac rieng cua file mau 2 ---
        if raw_unit not in unit_cache:
            unit_cache[raw_unit] = (
                normalize_don_vi(raw_unit, default_to_chuc=default_to_chuc),
                determine_to_chuc(raw_unit, default_to_chuc=default_to_chuc),
            )
        don_vi, to_chuc = unit_cache[raw_unit]

        clean_records.append({
            "stt": raw_stt, "name": name, "email": email, "phone": phone,
            "don_vi": don_vi, "to_chuc": to_chuc,
        })

    return total_read


# Da chuyen sang extract_contacts.py de dung chung voi Che do 1 - giu ten cu
# de tuong thich nguoc.
_LEADING_STT_LINE_RE = core.LEADING_STT_LINE_RE
_UNIT_KEYWORD_SPLIT_RE = core.UNIT_KEYWORD_SPLIT_RE


def _known_stts(clean_records, issues):
    """Tap STT so nguyen da xu ly (ca ban ghi hop le lan dong bi loai)."""
    known = set()
    for item in list(clean_records) + list(issues):
        n = core.parse_stt(item.get("stt"))
        if n is not None:
            known.add(n)
    return known


def _recover_missing_stt_rows(clean_records, issues, full_text, default_to_chuc, unit_cache,
                              warnings=None):
    """LUOI AN TOAN BO SUNG cho truong hop pdfplumber BO SOT HOAN TOAN 1
    vai dong khi tach bang (khac voi loi lech cot - o day dong khong xuat
    hien trong bang tach duoc mot chut nao, da gap thuc te tren tep that).

    Cach lam: dua vao cac STT (so nguyen) DA XU LY DUOC (trong ca
    clean_records lan issues) de xac dinh day so lien tuc 1..max VA TIM
    CAC SO BI THIEU ("nhay so"). Voi moi so bi thieu, quet VAN BAN THO
    (dong theo dong) tim dong BAT DAU dung bang so STT do, trich email/SDT
    qua regex tu dong do (gop them 1-2 dong ke tiep phong khi bi xuong
    dong giua chung), roi THEM ban ghi da khoi phuc vao dung vi tri (theo
    STT) trong clean_records - GIU NGUYEN THU TU nhu cac ban ghi khac.

    Don vi cua dong khoi phuc duoc suy tu ban ghi/issue GAN NHAT (STT nho
    hon) da xu ly duoc, vi cac STT lien tiep thuong thuoc cung 1 khoi/don
    vi trong thuc te (dung voi hau het mau bieu). Tim thay dong nhung khong
    co email hop le -> vao issues (nguoi that, can kiem tra). KHONG tim thay
    dong nao trong van ban tho (vd nguon danh so sai - Dong Thap go 2315
    thanh 2015) -> KHONG tao issue rong (truoc day ghi "Thiếu email" voi ten
    trong, gay hieu nham), ma ghi CANH BAO vao `warnings` (list) de nguoi
    dung doi chieu tep goc.

    Tra ve so luong dong THUC SU tim lai duoc trong van ban tho - dung de
    cong don vao total_read."""
    if not full_text:
        return 0

    known_stts = _known_stts(clean_records, issues)
    stt_to_unit = {}
    for rec in clean_records:
        n = core.parse_stt(rec.get("stt"))
        if n is not None:
            stt_to_unit[n] = rec.get("don_vi")

    missing = core.find_stt_gaps(known_stts)
    if not missing:
        return 0

    line_by_stt = core.find_stt_lines(full_text, missing)
    not_found = [n for n in missing if n not in line_by_stt]
    if not_found and warnings is not None:
        warnings.append(f"STT bị nhảy số ({core.format_stt_list(not_found)}) - không tìm thấy các dòng "
                        "này kể cả trong văn bản gốc. Hãy đối chiếu tệp gốc (có thể nguồn đánh số sai).")

    def _insert_position(stt_num):
        pos = len(clean_records)
        for idx, rec in enumerate(clean_records):
            s = rec.get("stt")
            if s and str(s).isdigit() and int(s) > stt_num:
                return idx
        return pos

    recovered_count = 0
    for num in missing:
        combined = line_by_stt.get(num)
        if not combined:
            continue
        recovered_count += 1
        nearest_unit = None
        for candidate in range(num - 1, 0, -1):
            if candidate in stt_to_unit:
                nearest_unit = stt_to_unit[candidate]
                break

        # Tach ten/don vi/email/SDT tu dong van ban tho (khong con ranh gioi
        # cot) - xem core.split_recovered_stt_line: phan TRUOC tu khoa don vi
        # quen thuoc (vd "Văn phòng", "Ban ", "UBND") la ten, phan TU tu khoa
        # la CHINH Don vi cua dong nay (dang tin cay hon don vi cua nguoi gan
        # nhat, vi Don vi co the doi tung nguoi trong 1 khoi - da gap thuc te).
        parts = core.split_recovered_stt_line(combined)
        if not parts["email"]:
            issues.append({
                "stt": str(num), "raw_name": combined.strip(), "raw_email": "", "raw_phone": "",
                "raw_unit": nearest_unit or "", "reason": "missing_email",
            })
            continue

        name_part = parts["name"]
        recovered_unit_text = parts["unit"]
        email = core.normalize_email(parts["email"])
        if not email:
            issues.append({
                "stt": str(num), "raw_name": name_part, "raw_email": parts["email"], "raw_phone": "",
                "raw_unit": nearest_unit or "", "reason": "invalid_email_format",
            })
            continue
        name = core.normalize_name(name_part)
        phone = core.normalize_phone(parts["phone"]) if parts["phone"] else ""

        if recovered_unit_text:
            # Tim thay chinh Don vi cua dong nay trong van ban tho -> chuan
            # hoa nhu 1 raw_unit binh thuong (giong cac dong khac).
            if recovered_unit_text not in unit_cache:
                unit_cache[recovered_unit_text] = (
                    normalize_don_vi(recovered_unit_text, default_to_chuc=default_to_chuc),
                    determine_to_chuc(recovered_unit_text, default_to_chuc=default_to_chuc),
                )
            don_vi, to_chuc = unit_cache[recovered_unit_text]
        elif nearest_unit is not None:
            # Khong tim thay Don vi rieng cho dong nay trong van ban tho
            # (vd trang do bi mat ca phan Don vi) - dung tam Don vi cua
            # nguoi GAN NHAT (STT nho hon) da xu ly duoc, vi cac STT lien
            # tiep thuong (khong phai luon luon) cung khoi/don vi.
            don_vi, to_chuc = nearest_unit, determine_to_chuc("", default_to_chuc=default_to_chuc)
        else:
            don_vi, to_chuc = "", ""

        rec = {"stt": str(num), "name": name, "email": email, "phone": phone,
               "don_vi": don_vi, "to_chuc": to_chuc}
        clean_records.insert(_insert_position(num), rec)

    return recovered_count


def extract_v2(input_path, default_to_chuc=None, dedupe=True, verbose=True):
    """
    Doc tep dang ky nguon VA AP DUNG DAY DU 3 QUY TAC CUA FILE MAU LOAI 1
    (tai su dung truc tiep tu extract_contacts.py) CONG 2 QUY TAC RIENG
    (Don vi, To chuc). KHONG GHI RA FILE (dung cho buoc xem truoc truoc khi
    xuat) - xem write_v2_output() de ghi ket qua, write_v2_issues() de ghi
    bao cao cac dong bi loai.

    HO TRO NHIEU DINH DANG TEP NGUON: .xlsx/.xls/.csv (dang bang co san) VA
    CA .pdf/.docx (tai lieu/cong van tu do, dung chung bo may doc bang cua
    file mau loai 1 - tim va gop bang qua nhieu trang, ghep dong bi ngat...
    - nhung KHAC file mau loai 1 o cho GIU LAI cot "Don vi/Phong ban/Chuc
    vu" tren tung dong thay vi loc bo, vi che do nay CAN thong tin do de
    chuan hoa Don vi/To chuc).

    Ho tro 3 kieu xac dinh "Don vi cong tac" cho tung nguoi, THEO THU TU
    UU TIEN:
      1) Co san 1 cot Don vi/Phong ban/Chuc vu dien tren TUNG DONG (vd
         danh sach cua 1 tinh/thanh, hoac phieu dang ky co cot rieng).
      2) KHONG co cot do (hoac khong dang tin cay), nhung co CAC DONG TIEU
         DE NHOM (vd danh dau so La Ma 'I', 'II'... ten 1 vien/trung tam)
         de gom nhom nguoi phia duoi - dung TEN NHOM lam Don vi.
      3) KHONG co ca 2 dieu tren (thuong gap o PHIEU DANG KY PDF/Word cho
         1 DON VI DUY NHAT, khong co cot Don vi rieng vi ca phieu la cua
         1 co quan) - tu dong tim dong "Tên Cơ quan, đơn vị: X" trong phan
         van ban tu do o dau tai lieu, dung X lam Don vi cho TOAN BO nguoi
         trong phieu.

    Tra ve (clean_records, stats). clean_records: list dict
    {name,email,phone,don_vi,to_chuc}, DA loc trung email neu dedupe=True,
    giu THU TU GOC. stats: dict {total_read, valid_before_dedupe,
    duplicates_removed, final_count, to_chuc_missing, issues, unit_cache}.
    """
    tables, fallback_unit, content_based, full_text = _read_tables_and_context(input_path)
    if not tables or not any(tables):
        raise ValueError("Tệp nguồn không có dữ liệu.")

    unit_cache = {}
    issues = []
    clean_records = []
    total_read = 0
    any_header_found = False
    table_stts = set()  # STT cua MOI dong trong bang (ke ca dong nhom) - Excel/CSV

    for table in tables:
        if not table:
            continue
        header_idx, col_map = _find_source_header_row(table)
        if header_idx is None or ("name" not in col_map and "email" not in col_map):
            continue  # bang nay khong co dong tieu de phu hop - bo qua (vd bang phu/khong lien quan)
        any_header_found = True
        data_rows = table[header_idx + 1:]
        if content_based:
            # PDF/Word: cac trang khac nhau CO THE co bo cuc cot khac nhau
            # (da gap thuc te) - dung bo phan tich THEO NOI DUNG (khong
            # phu thuoc vi tri cot) thay vi col_map co dinh.
            total_read += _process_data_rows_content_based(
                data_rows, default_to_chuc, unit_cache, issues, clean_records,
                fallback_unit=fallback_unit,
            )
        else:
            # Ghi nhan STT ca dong nhom/dong khong phai nguoi (vd Son La STT
            # 89 "Công an tỉnh - chưa có danh sách đăng ký") de khong bao
            # nham "nhay so".
            stt_i = col_map.get("stt")
            if stt_i is not None:
                for r in data_rows:
                    n = core.parse_stt(r[stt_i]) if r and stt_i < len(r) else None
                    if n is not None:
                        table_stts.add(n)
            table_fallback = fallback_unit
            if not table_fallback and "unit" not in col_map:
                # Excel/CSV khong co cot Don vi -> lay tu dong tieu de van ban o tren
                table_fallback = _extract_title_unit_from_rows(table, header_idx)
            total_read += _process_data_rows(
                data_rows, col_map, default_to_chuc, unit_cache, issues, clean_records,
                fallback_unit=table_fallback,
            )

    if not any_header_found:
        raise ValueError(
            "Không tìm thấy bảng dữ liệu phù hợp trong tệp nguồn (cần ít nhất cột "
            "Họ và tên hoặc Email). Vui lòng kiểm tra lại cấu trúc tệp."
        )

    # --- Luoi an toan bo sung: khoi phuc cac STT bi "nhay so" (pdfplumber
    # bo sot hoan toan khi tach bang) tu van ban tho, neu co ---
    warnings = []
    stts_before = _known_stts(clean_records, issues)
    total_read += _recover_missing_stt_rows(clean_records, issues, full_text, default_to_chuc, unit_cache,
                                            warnings=warnings)
    recovered_stts = sorted(_known_stts(clean_records, issues) - stts_before)
    if full_text is None:
        # Excel/CSV/Word: khong co van ban tho de khoi phuc - chi canh bao.
        gaps = core.find_stt_gaps(stts_before | table_stts)
        if gaps:
            warnings.append(f"STT bị nhảy số ({core.format_stt_list(gaps)}) - không tìm thấy các dòng này. "
                            "Hãy đối chiếu tệp gốc (có thể nguồn đánh số sai, hoặc dòng bị mất khi đọc).")

    # --- Loc trung email (mac dinh bat, giong file mau 1) ---
    duplicates_removed = 0
    if dedupe:
        # Giu ban ghi co ten KHOP email hon (xem core.email_name_match_score),
        # mac dinh giu ban ghi xuat hien truoc neu muc do khop ngang nhau.
        kept = {}
        deduped = []
        dropped = []
        for rec in clean_records:
            email = rec["email"]
            if email in kept:
                idx = kept[email]
                cur = deduped[idx]
                if core.email_name_match_score(rec["name"], email) > core.email_name_match_score(cur["name"], email):
                    dropped.append(cur)
                    deduped[idx] = rec
                else:
                    dropped.append(rec)
                continue
            kept[email] = len(deduped)
            deduped.append(rec)
        for rec in dropped:
            issues.append({
                "stt": rec["stt"], "raw_name": rec["name"], "raw_email": rec["email"],
                "raw_phone": rec["phone"], "raw_unit": rec["don_vi"], "reason": "duplicate_email",
            })
        duplicates_removed = len(dropped)
        clean_records = deduped

    to_chuc_missing = sum(1 for r in clean_records if not r["to_chuc"])

    # Ca 1 cot trong tren moi ban ghi -> gan nhu chac chan cot do khong duoc
    # nhan dien (tieu de la / lech cot), khong phai nguon thieu that. Truoc
    # day tool bao "0 loi" trong khi mat ca cot (vd Da Nang: tieu de "Di dong").
    if len(clean_records) >= 3:
        if not any(r["phone"] for r in clean_records):
            warnings.append(f"Không có số điện thoại nào trong {len(clean_records)} bản ghi - "
                            "kiểm tra tiêu đề cột SĐT.")
        if not any(r["don_vi"] for r in clean_records):
            warnings.append(f"Không có Đơn vị nào trong {len(clean_records)} bản ghi - "
                            "kiểm tra cột Đơn vị công tác.")

    stats = {
        "total_read": total_read,
        "valid_before_dedupe": total_read - (len(issues) - duplicates_removed),
        "duplicates_removed": duplicates_removed,
        "final_count": len(clean_records),
        "to_chuc_missing": to_chuc_missing,
        "issues": issues,
        "unit_cache": unit_cache,
        "warnings": warnings,
        "recovered_stts": [(os.path.basename(input_path), n) for n in recovered_stts],
    }

    if verbose:
        print_v2_stats_report(stats)

    return clean_records, stats


def print_v2_stats_report(stats):
    print(f"Tổng số dòng đọc được:                      {stats['total_read']}")
    print(f"Số dòng bị loại - thiếu/sai email:          "
          f"{sum(1 for i in stats['issues'] if i['reason'] in ('missing_email', 'invalid_email_format'))}")
    print(f"Số dòng bị loại - trùng email:               {stats['duplicates_removed']}")
    print(f"Còn lại sau cùng:                            {stats['final_count']}")
    if stats["to_chuc_missing"]:
        print(f"CẢNH BÁO: {stats['to_chuc_missing']} dòng không xác định được 'Tổ chức' "
              f"(để trống) - kiểm tra lại --to-chuc hoặc dữ liệu nguồn.")
    core.print_warnings(stats.get("warnings"), stats.get("recovered_stts"))


def write_v2_output(records, template_path, output_path, password=None):
    """Ghi cac ban ghi (dict name/email/phone/don_vi/to_chuc) ra file Excel
    theo file mau loai 2. Cot Gioi tinh/Ngay sinh luon de trong (khong co
    du lieu nguon tuong ung nen khong tu bia). Cot Mat khau: dien gia tri
    `password` neu duoc truyen vao (vd mat khau mac dinh dung chung), hoac
    de trong neu khong truyen (mac dinh cu, giu tuong thich nguoc)."""
    out_wb = openpyxl.load_workbook(template_path)
    out_ws = out_wb.active
    header_cells = list(out_ws[1])
    sample_font = Font(name=header_cells[0].font.name, size=header_cells[0].font.sz)

    row_idx = 2
    for rec in records:
        values = [rec["name"], rec["email"], rec["phone"], password or "", "", "", rec["don_vi"], rec["to_chuc"]]
        for col_i, value in enumerate(values, start=1):
            out_ws.cell(row=row_idx, column=col_i, value=value or None).font = sample_font
        row_idx += 1
    out_wb.save(output_path)


def write_v2_issues(issues, issues_output_path):
    """Ghi bao cao cac dong bi loai (kem STT/du lieu goc/ly do) ra 1 file
    Excel rieng, phuc vu tab "Can kiem tra" / doi chieu ngoai CLI."""
    issues_wb = openpyxl.Workbook()
    issues_ws = issues_wb.active
    issues_ws.title = "can_kiem_tra"
    issues_ws.append(["STT", "Họ và tên (gốc)", "Email (gốc)", "Điện thoại (gốc)",
                       "Đơn vị công tác (gốc)", "Lý do"])
    for issue in issues:
        reason_label = core.ISSUE_REASON_LABELS.get(issue["reason"], issue["reason"])
        issues_ws.append([issue["stt"], issue["raw_name"], issue["raw_email"],
                           issue["raw_phone"], issue["raw_unit"], reason_label])
    issues_wb.save(issues_output_path)


def process(input_path, template_path, output_path, default_to_chuc=None,
            dedupe=True, issues_output_path=None, password=None, verbose=True):
    """Ham tien ich cho CLI: goi extract_v2() + write_v2_output() +
    write_v2_issues() theo dung thu tu, giu tuong thich nguoc voi cach goi
    cu. Tra ve (so_dong_ghi_duoc, thong_ke)."""
    records, stats = extract_v2(input_path, default_to_chuc=default_to_chuc,
                                 dedupe=dedupe, verbose=False)
    write_v2_output(records, template_path, output_path, password=password)

    issues = stats["issues"]
    if issues_output_path is None and issues:
        base, _ext = os.path.splitext(output_path)
        issues_output_path = f"{base}_can_kiem_tra.xlsx"
    if issues and issues_output_path:
        write_v2_issues(issues, issues_output_path)

    if verbose:
        print_v2_stats_report(stats)
        print(f"Đã ghi: {output_path}")
        if issues and issues_output_path:
            print(f"Đã ghi báo cáo các dòng bị loại ({len(issues)} dòng): {issues_output_path}")
        print()
        print("Các giá trị 'Đơn vị công tác' gốc -> (Đơn vị, Tổ chức) đã chuẩn hoá:")
        for raw, (don_vi, to_chuc) in sorted(stats["unit_cache"].items()):
            print(f"  {raw!r:55s} -> Đơn vị={don_vi!r:35s} Tổ chức={to_chuc!r}")

    return stats["final_count"], stats


def main():
    parser = argparse.ArgumentParser(description="Chuẩn hoá dữ liệu đăng ký sang template loại 2.")
    parser.add_argument("--input", required=True, help="Tệp Excel nguồn (STT/Họ và tên/Đơn vị công tác/Email/SĐT)")
    parser.add_argument("--template", required=True, help="Tệp Excel mẫu (template loại 2)")
    parser.add_argument("--output", required=True, help="Đường dẫn tệp Excel kết quả")
    parser.add_argument("--to-chuc", default=None,
                         help="Tên tỉnh/thành (không tiền tố) dùng mặc định cho 'Tổ chức' khi "
                              "không tự nhận diện được từ dữ liệu nguồn, vd: 'Cần Thơ'")
    parser.add_argument("--no-dedupe", action="store_true", help="Không loại bỏ bản ghi trùng email")
    parser.add_argument("--issues-output", default=None,
                         help="Đường dẫn tệp Excel báo cáo các dòng bị loại (mặc định: "
                              "<output>_can_kiem_tra.xlsx nếu có dòng bị loại)")
    parser.add_argument("--password", default=None,
                         help="Mật khẩu mặc định điền vào cột 'Mật khẩu' (mặc định: để trống). "
                              f"Ví dụ: {core.DEFAULT_PASSWORD}")
    args = parser.parse_args()
    process(args.input, args.template, args.output, default_to_chuc=args.to_chuc,
            dedupe=not args.no_dedupe, issues_output_path=args.issues_output, password=args.password)


if __name__ == "__main__":
    main()