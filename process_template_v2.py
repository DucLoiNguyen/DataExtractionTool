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
def _tone_variants(name):
    """Tra ve ca 2 kieu dat dau tieng Viet (moi/cu): 'hòa'<->'hoà',
    'hóa'<->'hoá', 'thủy'<->'thuỷ'... de so khop co dau khong bo sot."""
    name = unicodedata.normalize("NFC", name).lower()
    pairs = [("òa", "oà"), ("óa", "oá"), ("ỏa", "oả"), ("õa", "oã"), ("ọa", "oạ"),
             ("òe", "oè"), ("óe", "oé"), ("ỏe", "oẻ"), ("õe", "oẽ"), ("ọe", "oẹ"),
             ("ùy", "uỳ"), ("úy", "uý"), ("ủy", "uỷ"), ("ũy", "uỹ"), ("ụy", "uỵ")]
    out = {name}
    for a, b in pairs:
        out |= {v.replace(a, b) for v in out} | {v.replace(b, a) for v in out}
    return out


_LOOKUP_ACCENTED = sorted(
    ((variant, current) for old, current in _OLD_NAME_TO_CURRENT.items() for variant in _tone_variants(old)),
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


def detect_to_chuc_from_text(text, require_prefix=False):
    """Quet 1 doan van ban (vd 'Don vi cong tac') tim ten tinh/thanh (cu
    hoac moi) xuat hien trong do, tra ve TEN CHINH THUC HIEN NAY neu tim
    thay, hoac None neu khong co dau hieu nao."""
    if not text:
        return None
    raw = unicodedata.normalize("NFC", str(text)).lower()
    if _deaccent(raw) != raw:
        # Van ban CO DAU tieng Viet -> so khop GIU NGUYEN DAU. Bo dau se gay
        # trung ten: "Hải Đường" (ten xa/truong o Ninh Bình) -> "hai duong"
        # trung "Hải Dương" (tinh cu, nay thuoc Hải Phòng) - da gap thuc te,
        # 134 nguoi bi gan nham Tổ chức "THÀNH PHỐ HẢI PHÒNG".
        pre = r"(?:tỉnh|thành\s*phố|tp\.?)\s*" if require_prefix else ""
        for old_name, current in _LOOKUP_ACCENTED:
            if re.search(r"(?<!\w)" + pre + re.escape(old_name) + r"(?!\w)", raw):
                return current
        return None
    key = _norm_key(text)
    pre = r"(?:tinh|thanh\s*pho|tp\.?)\s*" if require_prefix else ""
    for old_key, current in _LOOKUP:
        if not old_key:
            continue
        if re.search(r"(?<![a-z0-9])" + pre + re.escape(old_key) + r"(?![a-z0-9])", key):
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
        # Da co to chuc mac dinh la tinh/thanh -> CHI chap nhan tinh khac khi
        # van ban ghi RO "tỉnh X"/"thành phố X"/"TP. X". Chi nhac ten thi giu
        # mac dinh - ten tinh hay trung ten xa/truong/to chuc (da gap thuc
        # te: "Xã Bình Thuận" o Son La bi gan Lam Dong, "Đoàn TNCS Hồ Chí
        # Minh xã Tả Phìn" bi gan TP HCM).
        current = detect_to_chuc_from_text(raw_text, require_prefix=True)
        return current or matched

    return detect_to_chuc_from_text(raw_text)


_COMMUNE_LEVEL_RE = re.compile(r"(?<!\w)(thị\s*trấn|xã|phường)(?!\w)", re.IGNORECASE)
_COMMUNE_LEVEL_DISPLAY = {"xa": "Xã", "phuong": "Phường", "thi tran": "Thị trấn"}
# Tu dung NGAY TRUOC "xã/phường" ma KHONG phai cap xa: "thị xã" (cap huyen
# cu - "Phòng Nội vụ thị xã Bến Cát" khong phai xa), "hợp tác xã" (HTX -
# truoc day ra "Xã Nông"), "các xã"/"cấp xã" (noi chung, khong ten rieng).
_COMMUNE_BAD_PREV_WORDS = {"thị", "tác", "các", "cấp", "những", "mỗi", "toàn"}
# Chi dung khi ca chuoi viet thuong: tu nay KHONG thuoc ten xa -> dung lai.
_COMMUNE_STOP_WORDS = {"huyện", "quận", "tỉnh", "tp", "tp.", "và", "thuộc"}
_COMMUNE_MAX_WORDS = 5


def _find_commune_in_text(text):
    """Tim ten xa/phuong/thi tran trong 1 chuoi don vi, tra ve dang chuan
    "<Xã/Phường/Thị trấn> <Tên>" hoac None. Ten = cac tu VIET HOA chu dau
    (hoac chu so, vd "Phường 5", "Long Phú 1") ngay sau tu chi cap; dung o
    tu viet thuong hoac dau "-", ",", "(". Lay lan xuat hien CUOI co ten
    (vd "Ban kinh tế xã hội HĐND xã Tả Phìn" -> bo "xã hội", lay "xã Tả
    Phìn"). Khong co ten sau tu chi cap (vd "Phòng Kinh tế xã", "Chủ tịch
    UBND phường") -> None, giu nguyen du lieu goc."""
    if not text:
        return None
    found = None
    # Ca chuoi viet thuong (vd nguoi dung go "phòng vh xã nam định") -> khong
    # dung duoc chu hoa de nhan ten rieng; chap nhan tu viet thuong (toi da
    # _COMMUNE_MAX_WORDS tu, dung o dau phan cach).
    lower_mode = text == text.lower()
    name_end = -1  # vi tri ket thuc ten xa vua tim thay
    for m in _COMMUNE_LEVEL_RE.finditer(text):
        if m.start() < name_end:
            continue  # "Phường Xã Đàn": "Xã" nam TRONG ten phuong, khong phai cap xa
        prev = text[:m.start()].split()
        if prev and prev[-1].lower() in _COMMUNE_BAD_PREV_WORDS:
            continue  # "thị xã" (cap huyen cu), "hợp tác xã", "các xã", "cấp xã"
        # Tu chi cap VIET HOA TOAN BO ("XÃ", "PHƯỜNG") -> ten cung phai viet
        # hoa toan bo; tranh lay nham "XÃ NGHĨA ĐÔ Độc lập..." khi PDF gop 2
        # cot tieu de van ban vao 1 dong.
        upper_mode = m.group(1).isupper() and not lower_mode
        words = []
        consumed = m.end()
        for w in text[m.end():].split():
            if w in ("-", "–", "/", "(", ",") or w[0] in "-–(,/;":
                break
            stop_after = False
            if "/" in w:  # "Hưng/ Phòng Văn hóa" -> dung o dau "/"
                w, stop_after = w.split("/", 1)[0], True
            clean = w.strip(",.;:)(")
            if not clean:
                break
            if lower_mode:
                # "thành"/"thị" chi dung khi la "thành phố"/"thị xã" (ten xa co
                # the chua "Thành": "xã yên thành").
                two = " ".join(text[consumed:].split()[:2]).strip(",.;:").lower()
                ok = (len(words) < _COMMUNE_MAX_WORDS and clean.lower() not in _COMMUNE_STOP_WORDS
                      and two not in ("thành phố", "thị xã", "thị trấn"))
            elif upper_mode:
                # Dong gop 2 cot cua PDF/Excel: "PHƯỜNG TAM LONG CỘNG HÒA XÃ HỘI
                # CHỦ NGHĨA VIỆT NAM" - dung o quoc hieu, khong keo dai ten.
                ahead = " ".join(text[consumed:].split()[1:4]).upper().replace("HOÀ", "HÒA")  # cac tu SAU tu hien tai
                ok = (clean.isupper() or clean.isdigit()) and not (
                    (clean.upper() == "CỘNG" and ahead.startswith("HÒA XÃ HỘI"))
                    or (clean.upper() == "ĐỘC" and ahead.startswith("LẬP")))
            else:
                ok = clean[0].isupper() or clean[0].isdigit()
            if not ok:
                break
            words.append(clean)
            consumed = text.index(w, consumed) + len(w)
            if stop_after or (w != clean and w[-1] in ",;:)"):
                break
        if not words or words[0].lower() in ("hội", "viên"):
            continue  # "xã hội", "xã viên" khong phai ten xa
        name_end = consumed
        name = " ".join(words)
        if name.isupper() or lower_mode:
            name = " ".join(x[:1].upper() + x[1:].lower() for x in name.split())
        level = _COMMUNE_LEVEL_DISPLAY[_norm_key(re.sub(r"\s+", " ", m.group(1)))]
        found = f"{level} {name}"
    return found


# "(?<!cán )": "Cán bộ thường trực Trung tâm học tập cộng đồng" khong phai "Bộ" (cap tren xa).
_ABOVE_COMMUNE_RE = re.compile(r"(?<!\w)(?:tỉnh|thành\s*phố|sở|(?<!cán )bộ|cục|tổng\s*cục)(?!\w)", re.IGNORECASE)
# Don vi KHONG thuoc cap xa (cap huyen cu, don vi nganh doc, doanh nghiep)
# - khong tu gan vao xa ban hanh van ban khi o don vi khong ghi ten xa. Vd
# van ban cua xa Nghia Do co the co "Bệnh viện Đa khoa huyện Bảo Yên", "Chi
# nhánh Ngân hàng Chính sách xã hội" -> giu nguyen, khong doan la "Xã Nghĩa Đô".
_NOT_COMMUNE_UNIT_RE = re.compile(
    r"(?<!\w)(?:huyện|quận|thị\s*xã|bệnh\s*viện|ngân\s*hàng|chi\s*nhánh|đại\s*học|cao\s*đẳng|"
    r"kho\s*bạc|bảo\s*hiểm\s*xã\s*hội|chi\s*cục|điện\s*lực|bưu\s*điện|viettel|vnpt|công\s*ty)(?!\w)",
    re.IGNORECASE,
)


def _find_doc_commune(header_text, max_lines=15):
    """Tim xa/phuong/thi tran cua CO QUAN BAN HANH trong phan DAU van ban
    (vd "ỦY BAN NHÂN DÂN / XÃ NGHĨA ĐÔ", "Tên Cơ quan, đơn vị: Phòng Kinh tế
    xã Kon Đào", "UBND PHƯỜNG CHÁNH HƯNG"). Lay ket qua DAU TIEN theo tung
    dong (co quan ban hanh luon dung dau). Tra ve None neu khong co."""
    if not header_text:
        return None
    lines = [l for l in str(header_text).split("\n") if l.strip()][:max_lines]
    for line in lines:
        c = _find_commune_in_text(line)
        if c:
            return c
    return None


# Cac loai co quan thuong gap o dau cong van, nhan dien cu the (khong nhan
# UBND/HDND - la co quan cap tren hoac xa/phuong, xem _find_doc_commune).
_ISSUER_START_RE = re.compile(
    r"^(?:SỞ|BAN|TRUNG\s+TÂM|VĂN\s+PHÒNG|BỆNH\s+VIỆN|VIỆN|CHI\s+CỤC|CỤC|TỔNG\s+CỤC|TRƯỜNG|BẢO\s+TÀNG|THANH\s+TRA)(?!\w)"
)


def _find_doc_issuer(header_text, body_text=None, max_lines=8):
    """Tim CO QUAN BAN HANH khong phai xa/phuong o dau van ban (vd Gia Lai
    PDF: "SỞ NÔNG NGHIỆP VÀ MÔI TRƯỜNG Độc lập - Tự do - Hạnh phúc" -> "Sở
    Nông nghiệp và Môi trường"). Lay cum chu VIET HOA TOAN BO o dau dong (PDF
    gop 2 cot nen phan quoc hieu di sau - dung o tu co chu thuong hoac o
    "CỘNG HÒA XÃ HỘI"), bat dau bang Sở/Ban/Trung tâm/Văn phòng/Bệnh viện/
    Viện/Chi cục/Cục/Trường/Bảo tàng... Bo qua dong "UBND TỈNH..." (co quan
    cap tren).

    Cach viet: neu than van ban co ten do viet thuong dung chuan (vd "Sở
    Nông nghiệp và Môi trường nhận văn bản số..."), dung nguyen van do; neu
    khong, viet hoa chu dau moi tu (_title_case_vn). Tra ve None neu khong
    co."""
    if not header_text:
        return None
    lines = [l.strip() for l in str(header_text).split("\n") if l.strip()][:max_lines]
    for line in lines:
        if not _ISSUER_START_RE.match(line):
            continue
        words = []
        parts = line.split()
        for i, w in enumerate(parts):
            clean = w.strip(",.;:")
            if not (clean.isupper() or clean.lower() in ("và", "&")):
                break
            if clean.upper() == "CỘNG" and " ".join(parts[i:i + 3]).upper().replace("HOÀ", "HÒA").startswith("CỘNG HÒA XÃ HỘI"):
                break
            words.append(clean)
        name = " ".join(words).strip()
        if len(words) < 2 or not _ISSUER_START_RE.match(name):
            continue
        # Thu lay cach viet chuan trong than van ban
        pattern = r"\s+".join(re.escape(w) for w in name.split())
        for source in (body_text, header_text):
            for m in re.finditer(pattern, str(source or ""), re.IGNORECASE):
                found = re.sub(r"\s+", " ", m.group(0))
                if found != found.upper() and found != found.lower():
                    return found
        return _title_case_vn(name)
    return None


def normalize_don_vi(raw_don_vi_cong_tac, default_to_chuc=None, doc_commune=None, doc_issuer=None):
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

    # O don vi CHI GHI CHUC VU thuan tuy ("Giám đốc sở", "Phó Giám đốc Sở") ma
    # van ban co co quan ban hanh ro o dau (khong phai xa/phuong) -> don vi la
    # co quan do (nguoi dung dong y, Gia Lai: "Sở Nông nghiệp và Môi trường").
    # CHI khi chuc vu thuan tuy (_strip_job_title rong): "Chủ tịch UBND phường"
    # (Bến Cát), "Giám đốc, Trung tâm Chiếu xạ Hà Nội" (VINATOM) van giu nguyen.
    if doc_issuer and not doc_commune and _looks_like_job_title(text) and not _strip_job_title(text):
        text = doc_issuer

    # Truong hop 0 (yeu cau nguoi dung): don vi CAP DUOI xa/phuong (phong
    # ban, truong hoc, tram y te, doan the, thon/to dan pho...) -> GOP
    # thanh chinh xa/phuong/thi tran do. Vd "Phòng VH- XH Phường Nam Định"
    # -> "Phường Nam Định", "Văn phòng Đảng ủy xã Tả Phìn" -> "Xã Tả Phìn".
    commune = _find_commune_in_text(text)
    if commune and doc_commune and commune != doc_commune \
            and _norm_key(doc_commune).startswith(_norm_key(commune)) \
            and len(doc_commune.split()) - len(commune.split()) <= 2:
        # "xã Chiềng ken" (chu thuong) chi bat duoc "Xã Chiềng" -> dung ten
        # day du cua xa ban hanh van ban. Gioi han them toi da 2 tu: neu ten
        # xa ban hanh dai bat thuong la do doc nham (vd dinh quoc hieu), giu
        # ten xa ghi trong o.
        commune = doc_commune
    if not commune and doc_commune and not _ABOVE_COMMUNE_RE.search(text) \
            and not _NOT_COMMUNE_UNIT_RE.search(text) and not _ADMIN_BODY_PATTERN.match(text):
        # Van ban do 1 xa/phuong ban hanh (doc_commune) va o don vi chi ghi
        # bo phan, khong kem ten xa (vd "Văn phòng Đảng ủy", "Phòng kinh
        # tế", "Hội nông dân Việt Nam xã") -> bo phan thuoc chinh xa do.
        commune = doc_commune
    m = None if commune else _ADMIN_BODY_PATTERN.match(text)
    if commune:
        core = commune
    # Truong hop 1: don vi hanh chinh xa/phuong/thi tran voi tien to UBND/HDND
    elif m:
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
        current = detect_to_chuc_from_text(raw_don_vi_cong_tac, require_prefix=True)  # xem _resolve_current_province
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
# nguon duy nhat (truoc day 2 danh sach chep tay, da lech nhau nhieu lan).
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
        # Excel hay luu STT bang cong thuc cong don -> sai so dau phay dong
        # (vd 150.00000000000003); gan so nguyen thi hien thi so nguyen.
        if abs(value - round(value)) < 1e-6:
            return str(int(round(value)))
        return str(value)
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


def _looks_like_people_sheet(rows):
    """Sheet co tieu de Ho ten + Email/SDT va >= 3 dong du lieu (giong 1 danh
    sach nguoi) - dung de canh bao khi bo qua 1 sheet khong co email nao."""
    idx, col_map = _find_source_header_row(rows)
    if idx is None:
        return False
    if "name" not in col_map or not ({"email", "phone"} & set(col_map)):
        return False
    return sum(1 for r in rows[idx + 1:] if r and any(c not in (None, "") for c in r)) >= 3


def _select_excel_tables(input_path, sheet=None, warnings=None):
    """Chon cac sheet cua tep Excel de doc -> (list_rows, list_ten_sheet).
      - sheet duoc chi dinh: CHI doc dung sheet do (khong ap quy tac bo sheet
        trung - nguoi dung da chu dong chon);
      - sheet=None: doc cac sheet CO EMAIL, bo sheet trung >= 50% email voi
        cac sheet truoc (ban nhap/phu luc, xem core.analyze_sheets) kem canh
        bao; sheet khong co email (danh muc...) khong doc. Khong sheet nao co
        email -> doc sheet dau nhu cu (de bao 'thieu email')."""
    name = os.path.basename(input_path)
    sheets = core.read_excel_sheets(input_path)
    if sheet is not None:
        for sname, rows in sheets:
            if sname == sheet:
                return [rows], [sname]
        raise ValueError(f"Không tìm thấy sheet '{sheet}' trong tệp {name}.")
    infos = core.analyze_sheets(sheets)
    if not any(i["emails"] for i in infos):
        return ([sheets[0][1]], [sheets[0][0]]) if sheets else ([], [])
    tables, names = [], []
    for i in infos:
        if not i["emails"]:
            if warnings is not None and _looks_like_people_sheet(i["rows"]):
                warnings.append(f"Sheet '{i['sheet']}' không có email nào nên không được đọc.")
            continue
        if i["skip"]:
            if warnings is not None:
                warnings.append(f"Bỏ qua sheet '{i['sheet']}': {core.sheet_skip_reason(i)} (có thể là bản nháp/phụ "
                                f"lục); {i['emails'] - i['overlap']} email mới trong sheet không được đọc. "
                                "Chọn riêng sheet này (Chế độ 2 trên giao diện) nếu cần đọc.")
            continue
        tables.append(i["rows"])
        names.append(i["sheet"])
    return tables, names


def _read_tables_and_context(input_path, sheet=None, warnings=None, sheet_names=None):
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

    if ext in (".xls", ".xlsx", ".xlsm"):
        tables, names = _select_excel_tables(input_path, sheet, warnings)
        if sheet_names is not None:
            sheet_names.extend(names)
        return tables, None, False, None, None

    if ext == ".csv":
        with open(input_path, newline="", encoding="utf-8-sig", errors="ignore") as f:
            rows = [list(row) for row in csv.reader(f)]
        return [rows], None, False, None, None

    if ext == ".pdf":
        import pdfplumber
        raw_tables = []
        context_text_parts = []  # TOAN BO van ban moi trang (bat ke co bang hay khong) - CHI
                                  # dung de do "Tên Cơ quan, đơn vị: X" lam Don vi du phong,
                                  # KHONG dung de trich xuat du lieu (tranh trung lap voi bang)
        with pdfplumber.open(input_path) as pdf:
            per_page_tables = []
            for page in pdf.pages:
                per_page_tables.append(page.extract_tables() or [])
                context_text_parts.append(page.extract_text() or "")
            # Email bi cat theo vien o bang (Phu Tho Nguyet Duc) - xem core.repair_cut_email_cells.
            fixed = core.repair_cut_email_cells(per_page_tables, lambda i: pdf.pages[i].extract_words())
            if fixed and warnings is not None:
                nums = [n for n in (core.parse_stt(s) for s, _e in fixed) if n]
                warnings.append(f"Khôi phục {len(fixed)} email bị cắt theo viền ô bảng PDF"
                                f"{' (STT ' + core.format_stt_list(nums) + ')' if nums else ''} - "
                                "hãy đối chiếu email với tệp gốc.")
            for tables in per_page_tables:
                raw_tables.extend(tables)
        flat_table = _flatten_doc_tables(raw_tables)
        full_text = "\n".join(context_text_parts)
        header_text = context_text_parts[0] if context_text_parts else ""
        return [flat_table], _extract_doc_header_unit(full_text), True, full_text, header_text

    if ext in (".docx", ".doc"):
        import docx
        doc_path = input_path
        converted_path = None
        if ext == ".doc":
            # File Word 97-2003 cu - python-docx khong doc duoc, tu dong
            # chuyen doi truoc: Microsoft Word (pywin32) roi LibreOffice
            # (xem core.convert_office_file).
            converted_path = core.convert_office_file(input_path, "docx")
            doc_path = converted_path
        try:
            document = docx.Document(doc_path)
            raw_tables = [
                [[core._full_cell_text(cell) for cell in row.cells] for row in table.rows]
                for table in document.tables
            ]
            flat_table = _flatten_doc_tables(raw_tables)
            free_text = "\n".join(core._full_paragraph_text(p) for p in document.paragraphs)
            # Phan dau van ban Word: bang NHO (quoc hieu/tieu ngu, <= 3 dong)
            # + cac doan mo dau - noi ghi co quan ban hanh.
            small_tables = ["\n".join("\n".join(c for c in row if c) for row in t) for t in raw_tables if len(t) <= 3]
            header_text = "\n".join(small_tables + [free_text])
            return [flat_table], _extract_doc_header_unit(free_text), True, None, header_text
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


_JOB_TITLE_RE = re.compile(
    r"^(?:chuyên viên|nhân viên|viên chức|công chức|cán bộ|kế toán|văn thư|thủ quỹ|"
    r"giám đốc|phó|trưởng|chủ tịch|bí thư|chánh|chủ nhiệm|ủy viên|uỷ viên|thanh tra viên|"
    r"giáo viên|bác sĩ|y sĩ|điều dưỡng|kỹ thuật viên|lãnh đạo|thư ký|lái xe|bảo vệ|tạp vụ)(?!\w)",
    re.IGNORECASE,
)


def _looks_like_job_title(text):
    """True neu o 'don vi' thuc ra ghi CHUC VU (vd "Chuyên viên", "Phó
    Trưởng phòng", "Nhân viên văn phòng Đảng ủy") - gap o file Dak Lak dot 4,
    Ben Cat. So khop co dau (tranh nham "Trường" (truong hoc) voi "Trưởng")."""
    if not text:
        return False
    return bool(_JOB_TITLE_RE.match(unicodedata.normalize("NFC", str(text)).strip().lower()))


_UNIT_START_RE = re.compile(
    r"^(?:phòng|ban|văn phòng|trung tâm|sở|bệnh viện|chi cục|cục|trạm|trường|viện|hội|đoàn|"
    r"ủy ban|uỷ ban|ubnd|hđnd|đảng ủy|đảng uỷ|mặt trận|công an|ban quản lý|thanh tra|kho bạc)(?!\w)",
    re.IGNORECASE,
)


def _strip_job_title(text):
    """Bo cac tu chi CHUC VU o dau chuoi, tra ve phan TEN DON VI con lai neu
    no bat dau bang tu chi don vi va co >= 2 tu (vd "Giám đốc Trung tâm TGPL
    Nhà nước" -> "Trung tâm TGPL Nhà nước"); nguoc lai tra ve "" (vd
    "Phó Trưởng phòng", "Phó Giám đốc TT" chi la chuc vu)."""
    rest = unicodedata.normalize("NFC", str(text)).strip()
    for _ in range(4):
        m = _JOB_TITLE_RE.match(rest.lower())
        if not m:
            break
        rest = rest[m.end():].strip(" ,-–/")
    if len(rest.split()) >= 2 and _UNIT_START_RE.match(rest.lower()):
        return rest
    return ""


_GROUP_VALUE_RE = re.compile(
    r"(?<![a-z])(xa|phuong|thi tran|so|ubnd|hdnd|uy ban|tinh|thanh pho|trung tam|phong|ban|vien|truong|"
    r"benh vien|chi cuc|cuc|van phong|vp)(?![a-z])"
)


def _detect_sparse_group_col(table, header_idx, col_map):
    """Tim cot 'don vi lon' THUA trong bang CO tieu de (vd cot "GHI CHÚ" chi
    ghi "xã Vụ Bổn", "Sở Y Tế tỉnh Đắk Lắk" o dong DAU moi khoi). Dieu kien:
    cot chua duoc gan, co >= 2 gia tri van ban nhung < 30% so dong du lieu,
    va >= 80% gia tri trong giong ten don vi. Tra ve chi so cot hoac None."""
    data = [r for r in table[header_idx + 1:] if r and any(c not in (None, "") for c in r)]
    if len(data) < 5:
        return None
    used = set(col_map.values())
    ncols = max(len(r) for r in data)
    best = None
    for i in range(ncols):
        if i in used:
            continue
        vals = [str(r[i]).strip() for r in data if i < len(r) and isinstance(r[i], str) and str(r[i]).strip()]
        if len(vals) < 2 or len(vals) >= 0.3 * len(data):
            continue
        unit_like = sum(1 for v in vals if _GROUP_VALUE_RE.search(_norm_key(v)))
        if unit_like >= 0.8 * len(vals) and (best is None or len(vals) > best[1]):
            best = (i, len(vals))
    return best[0] if best else None


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


def _squeeze_at(value):
    """Bo khoang trang/xuong dong SAT hai ben dau "@" (khong dong vao phan
    con lai cua o), vd "abc @laocai.gov.vn" -> "abc@laocai.gov.vn"."""
    return re.sub(r"\s*@\s*", "@", str(value)).strip()


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
        # Bo khoang trang SAT dau "@" (vd "hamanhcuong3 @laocai.gov.vn" - da gap
        # o Yên Thành, 69 nguoi bi loai oan) truoc khi nhan dang.
        if c and core.EMAIL_REGEX.search(_squeeze_at(c)):
            email_val, email_pos = c, i
            used.add(i)
            break
    if email_val is None:
        # O co "@" nhung chua dung chuan (vd "@laocai,gov.vn", "@@laocai.gov.vn")
        # van la o email - de core.normalize_email tu sua loi go.
        for i, c in enumerate(cells):
            if c and "@" in str(c) and " " not in _squeeze_at(c):
                email_val, email_pos = c, i
                used.add(i)
                break

    phone_val = None
    for i, c in enumerate(cells):
        if i in used or not c:
            continue
        if _PHONE_LOOSE_REGEX.search(re.sub(r"[\s.\-]", "", str(c))):
            phone_val = c
            used.add(i)
            break

    stt_val = cells[0] if cells else None
    stt_idx = 0
    if cells and not _cell_display(stt_val):
        # Dong bi LECH 1 COT (Gia Lai: ['', '8', '', '', 'Đoàn Ngọc Có', ...]):
        # STT la o co noi dung DAU TIEN neu do la so ngan (<= 4 chu so, co the
        # co dau cham) hoac so La Ma - khong lay SDT/ten.
        for i, c in enumerate(cells):
            if _cell_display(c):
                if re.fullmatch(r"\d{1,4}\.?|[IVXLC]{1,6}\.?", _cell_display(c)) and i not in used:
                    stt_val, stt_idx = c, i
                break
    if isinstance(stt_val, str) and re.fullmatch(r"\s*\d+\.\s*", stt_val):
        stt_val = stt_val.strip().rstrip(".")  # STT dang "1." -> "1"
    used.add(stt_idx)

    text_cells = []
    for i, c in enumerate(cells):
        if i in used or c is None:
            continue
        text = str(c).strip()
        if text:
            text_cells.append((i, text))

    name_idx, name_val = text_cells[0] if len(text_cells) >= 1 else (None, None)
    unit_val = text_cells[1][1] if len(text_cells) >= 2 else None

    return {
        "stt": _cell_display(stt_val), "name": _cell_display(name_val),
        "unit": _cell_display(unit_val), "email": email_val, "phone": phone_val,
        "name_idx": name_idx, "text_count": len(text_cells),
    }


def _row_has_identity_content(row):
    """True neu dong CO email hoac SDT (nhan dang qua regex) - dung de
    phan biet 1 dong DU LIEU THAT voi 1 dong tieu de nhom/dong rac."""
    for c in row:
        if not c:
            continue
        s = str(c)
        if core.EMAIL_REGEX.search(_squeeze_at(s)) or _PHONE_LOOSE_REGEX.search(re.sub(r"[\s.\-]", "", s)):
            return True
        if "@" in s and " " not in _squeeze_at(s):
            return True
    return False


def _unit_lookup(unit_cache, raw_unit, default_to_chuc, doc_commune, doc_issuer):
    """(Don vi, To chuc) da chuan hoa cua 1 o don vi tho. Khoa cache gom ca
    xa/co quan ban hanh: moi sheet/bang co the co xa/co quan ban hanh RIENG
    - cung 1 chuoi don vi o 2 sheet khong duoc dung chung ket qua."""
    key = (raw_unit, doc_commune, doc_issuer)
    if key not in unit_cache:
        unit_cache[key] = (
            normalize_don_vi(raw_unit, default_to_chuc=default_to_chuc, doc_commune=doc_commune,
                             doc_issuer=doc_issuer),
            determine_to_chuc(raw_unit, default_to_chuc=default_to_chuc),
        )
    return unit_cache[key]


def _is_name_continuation(parsed, last):
    """Dong chi co 1 TU ho ten o dung cot ten cua nguoi vua them (Gia Lai
    PDF: 'Cao Thanh' + dong ke tiep 'Thương', nhieu khi kem SDT lap lai) -
    xem core.merge_name_fragment_rows (cung dieu kien than trong)."""
    if last["kind"] not in ("record", "issue") or last["name_idx"] is None:
        return False
    if parsed["email"] or parsed["stt"] or parsed["text_count"] != 1 or parsed["name_idx"] != last["name_idx"]:
        return False
    word = parsed["name"]
    prev = last["raw_name"]
    if not word or " " in word or not word.isalpha() or core._GROUP_UNIT_START_RE.match(word):
        return False
    if word.isupper() != prev.isupper() or not prev or len(prev.split()) > 3:
        return False
    if parsed["phone"]:
        # O SDT chi duoc phep neu lap lai SDT cua nguoi truoc
        return bool(last["phone"]) and core.normalize_phone(parsed["phone"]) == last["phone"]
    return True


def _process_data_rows_content_based(data_rows, default_to_chuc, unit_cache, issues, clean_records,
                                      fallback_unit=None, doc_commune=None, doc_issuer=None, notes=None):
    """Ban THEO NOI DUNG (khong dung vi tri cot co dinh) cua
    _process_data_rows - dung cho PDF/Word khi cac trang trong CUNG 1 file
    co the co bo cuc cot khac nhau (xem _parse_row_by_content). Logic con
    lai (3 quy tac goc, Don vi/To chuc, tieu de nhom) giu nguyen tinh than
    nhu ham goc.

    notes: list (tuy chon) - nhan STT cac dong ho ten bi ngat 2 dong da duoc
    noi (_is_name_continuation) de ham goi ghi canh bao."""
    current_group = None
    group_header_seen = False
    total_read = 0
    # theo doi ban ghi/issue VUA THEM de va them du lieu neu gap dong manh vo
    # hoac phan tiep cua ho ten o dong ke tiep
    last_added_ref = {"kind": None, "data": None, "name_idx": None, "raw_name": "", "phone": ""}

    for row in data_rows:
        if row is None or all(c in (None, "") for c in row):
            continue

        # Ho ten bi ngat 2 dong vat ly -> noi vao nguoi truoc, KHONG phai tieu
        # de nhom hay nguoi moi (Gia Lai: 7/9 nguoi).
        parsed_pre = _parse_row_by_content(row)
        if _is_name_continuation(parsed_pre, last_added_ref):
            target = last_added_ref["data"]
            new_raw = f"{last_added_ref['raw_name']} {parsed_pre['name']}"
            if last_added_ref["kind"] == "record":
                target["name"] = core.normalize_name(new_raw)
            else:
                target["raw_name"] = new_raw
            last_added_ref["raw_name"] = new_raw
            if notes is not None:
                notes.append(target.get("stt") or new_raw)
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

        parsed = parsed_pre

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
                    don_vi, to_chuc = _unit_lookup(unit_cache, target["raw_unit"], default_to_chuc,
                                                   doc_commune, doc_issuer)
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
            last_added_ref.update(kind="issue", data=issue, name_idx=parsed["name_idx"], raw_name=raw_name,
                                  phone=core.normalize_phone(raw_phone_val) if raw_phone_val else "")
            continue

        name = core.normalize_name(raw_name)
        phone = core.normalize_phone(raw_phone_val)
        don_vi, to_chuc = _unit_lookup(unit_cache, raw_unit, default_to_chuc, doc_commune, doc_issuer)

        rec = {"stt": raw_stt, "name": name, "email": email, "phone": phone,
               "don_vi": don_vi, "to_chuc": to_chuc}
        clean_records.append(rec)
        last_added_ref.update(kind="record", data=rec, name_idx=parsed["name_idx"], raw_name=raw_name,
                              phone=phone or "")

    return total_read


def _process_data_rows(data_rows, col_map, default_to_chuc, unit_cache, issues,
                        clean_records, fallback_unit=None, doc_commune=None, doc_issuer=None,
                        row_offset=None):
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
    group_col_value = None  # gia tri cot "group" (cot thua, gop o) dien tiep xuong duoi
    for row_pos, row in enumerate(data_rows):
        if row is None or all(c in (None, "") for c in row):
            continue
        # So dong THAT trong sheet Excel (1-based) - dung de chi dong bi trung
        # khi nguon khong co cot STT (Son La: nhieu dong de trong STT).
        sheet_row = row_offset + row_pos if row_offset is not None else None

        if "group" in col_map:
            # Bang co cot don vi lon RIENG (chi ghi o dong dau moi nhom do gop
            # o) - khong dung co che "dong tieu de nhom" (se nham 1 nguoi
            # thieu ca email lan SDT thanh tieu de nhom); dien tiep gia tri.
            g = get_display(row, "group")
            if g:
                group_col_value = g
        elif _looks_like_group_header(row, col_map):
            current_group = _group_header_text(row, col_map)
            group_header_seen = True
            continue  # dong tieu de nhom - CHI cap nhat nhom hien tai, khong tinh la 1 dong du lieu

        raw_stt = re.sub(r"^(\d+)\.$", r"\1", get_display(row, "stt"))  # "1." -> "1"
        raw_name = get_display(row, "name")
        raw_email = _cell_display(get_email_raw(row))
        raw_phone = get_display(row, "phone")
        raw_unit_from_col = get_display(row, "unit")
        unit2 = get_display(row, "unit2") if "unit2" in col_map else ""
        if unit2 and _norm_key(unit2) != _norm_key(raw_unit_from_col):
            raw_unit_from_col = f"{raw_unit_from_col} {unit2}".strip() if raw_unit_from_col else unit2
        if group_col_value:
            # Don vi lon cua nhom (cot thua, dien tiep xuong duoi):
            #  - la xa/phuong -> moi bo phan ben duoi GOP thanh xa/phuong do
            #    (quy tac nguoi dung; thong nhat cach viet ten xa theo nhom,
            #    vd "Trạm Y tế xã Ea MDroh" van ra "Xã Ea M’Droh");
            #  - o don vi trong hoac chi ghi CHUC VU ("Chuyên viên", "Phó
            #    Giám đốc TT"...) -> dung nguyen don vi lon;
            #  - con lai ghep bo phan + don vi lon (bo tien to UBND/HĐND),
            #    khong ghep lap neu bo phan da chua ten don vi lon.
            g_commune = _find_commune_in_text(group_col_value)
            if g_commune:
                # Khoi la xa/phuong -> uu tien ten khoi (o tung dong co the
                # sao chep nham: Son La co 32 nguoi email ".taxua@" o khoi
                # "Xã Tà Xùa" nhung o don vi ghi "UBND xã Tạ Khoa").
                raw_unit_from_col = g_commune
            elif not raw_unit_from_col:
                raw_unit_from_col = group_col_value
            elif _looks_like_job_title(raw_unit_from_col):
                # "Giám đốc Trung tâm TGPL Nhà nước" -> "Trung tâm TGPL Nhà nước";
                # chi co chuc vu ("Chuyên viên", "Phó Giám đốc TT") -> dung ten khoi.
                rest = _strip_job_title(raw_unit_from_col)
                raw_unit_from_col = rest if rest else group_col_value
            # con lai: o don vi da la ten don vi that -> giu nguyen, khong ghep ten khoi
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
            issue = {
                "stt": raw_stt, "raw_name": raw_name, "raw_email": raw_email,
                "raw_phone": raw_phone, "raw_unit": raw_unit, "reason": reason,
            }
            if sheet_row is not None:
                issue["row"] = sheet_row
            issues.append(issue)
            continue  # Quy tac 1: khong co email hop le -> bo qua ban ghi

        name = core.normalize_name(raw_name)
        phone = core.normalize_phone(get_raw(row, "phone"))

        # --- Ap dung 2 quy tac rieng cua file mau 2 ---
        don_vi, to_chuc = _unit_lookup(unit_cache, raw_unit, default_to_chuc, doc_commune, doc_issuer)

        rec = {"stt": raw_stt, "name": name, "email": email, "phone": phone,
               "don_vi": don_vi, "to_chuc": to_chuc}
        if sheet_row is not None:
            rec["row"] = sheet_row
        clean_records.append(rec)

    return total_read


_LEADING_STT_LINE_RE = re.compile(r"^\s*(\d{1,4})\s+(.+)$")

# Cac tu bat dau pho bien cua ten don vi/phong ban tieng Viet - dung de
# tach "Ten - Don vi" khi ca 2 bi dinh lien trong van ban tho (khong con
# ranh gioi cot) trong _recover_missing_stt_rows.
_UNIT_KEYWORD_SPLIT_RE = re.compile(
    r"\b(?:Văn phòng|Ban(?:\s|$)|UBND|HĐND|Ủy ban|Phòng|Trung tâm|Hội(?:\s|$)|"
    r"Đoàn|Sở|Viện|Cơ quan|Đảng ủy)",
    re.IGNORECASE,
)


def _known_stts(clean_records, issues):
    """Tap STT so nguyen da xu ly (ca ban ghi hop le lan dong bi loai)."""
    known = set()
    for item in list(clean_records) + list(issues):
        n = core.parse_stt(item.get("stt"))
        if n is not None:
            known.add(n)
    return known


def _recover_missing_stt_rows(clean_records, issues, full_text, default_to_chuc, unit_cache, doc_commune=None,
                              warnings=None, doc_issuer=None):
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
    vi trong thuc te (dung voi hau het mau bieu). Neu khong tim thay email
    hop le trong van ban tho cho 1 STT bi thieu, GHI VAO issues voi ly do
    rieng de nguoi dung biet CAN KIEM TRA THU CONG dong do (khong am tham
    bo qua).

    Tra ve so luong STT da xu ly them duoc (ca thanh cong lan van thieu) -
    dung de cong don vao total_read."""
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

    lines = full_text.split("\n")
    line_by_stt = {}
    for i, line in enumerate(lines):
        m = _LEADING_STT_LINE_RE.match(line)
        if not m:
            continue
        num = int(m.group(1))
        if num in missing and num not in line_by_stt:
            combined = m.group(2)
            for j in range(1, 3):
                if i + j < len(lines):
                    nxt = lines[i + j].strip()
                    if nxt and not _LEADING_STT_LINE_RE.match(nxt):
                        combined += " " + nxt
                    else:
                        break
            line_by_stt[num] = combined

    def _insert_position(stt_num):
        pos = len(clean_records)
        for idx, rec in enumerate(clean_records):
            s = rec.get("stt")
            if s and str(s).isdigit() and int(s) > stt_num:
                return idx
        return pos

    # STT khong tim thay ca trong van ban tho (vd nguon danh so sai - Dong
    # Thap go 2315 thanh 2015) -> KHONG tao issue rong (truoc day ghi "Thiếu
    # email" voi ten trong, gay hieu nham), chi CANH BAO de doi chieu tep goc.
    not_found = [n for n in missing if n not in line_by_stt]
    if not_found and warnings is not None:
        warnings.append(f"STT bị nhảy số ({core.format_stt_list(not_found)}) - không tìm thấy các dòng "
                        "này kể cả trong văn bản gốc. Hãy đối chiếu tệp gốc (có thể nguồn đánh số sai).")

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

        email_m = core.EMAIL_REGEX.search(combined)
        if not email_m:
            issues.append({
                "stt": str(num), "raw_name": combined.strip(), "raw_email": "", "raw_phone": "",
                "raw_unit": nearest_unit or "", "reason": "missing_email",
            })
            continue

        phone_m = _PHONE_LOOSE_REGEX.search(combined.replace(" ", ""))
        name_part = combined[:email_m.start()].strip(" \t-|/,;:")
        # Neu name_part con dinh ca chu Don vi (do van ban tho khong co
        # ranh gioi cot ro rang), tach TU KEYWORD DON VI QUEN THUOC tro di
        # (vd "Văn phòng", "Ban ", "UBND"...) - phan TRUOC keyword la ten,
        # phan TU keyword la CHINH Don vi cua dong nay (dang tin cay hon la
        # doan tu nguoi gan nhat, vi Don vi co the doi tung nguoi trong
        # cung 1 khoi lon - da gap thuc te).
        recovered_unit_text = None
        m_unit_kw = _UNIT_KEYWORD_SPLIT_RE.search(name_part)
        if m_unit_kw and m_unit_kw.start() > 0:
            recovered_unit_text = name_part[m_unit_kw.start():].strip(" \t-|/,;:")
            name_part = name_part[:m_unit_kw.start()].strip(" \t-|/,;:")
        email = core.normalize_email(email_m.group(0))
        if not email:
            issues.append({
                "stt": str(num), "raw_name": name_part, "raw_email": email_m.group(0), "raw_phone": "",
                "raw_unit": nearest_unit or "", "reason": "invalid_email_format",
            })
            continue
        name = core.normalize_name(name_part)
        phone = core.normalize_phone(phone_m.group(0)) if phone_m else ""

        if recovered_unit_text:
            # Tim thay chinh Don vi cua dong nay trong van ban tho -> chuan
            # hoa nhu 1 raw_unit binh thuong (giong cac dong khac).
            don_vi, to_chuc = _unit_lookup(unit_cache, recovered_unit_text, default_to_chuc,
                                           doc_commune, doc_issuer)
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


def _dup_note(kept):
    """Ghi chu 'Trung voi dong nao, tep nao' cho 1 dong bi loai vi trung email."""
    stt, row = kept.get("stt"), kept.get("row")
    if row and stt:
        where = f"dòng {row} (STT {stt})"
    elif row:
        where = f"dòng {row}"
    elif stt:
        where = f"dòng STT {stt}"
    else:
        where = "một dòng khác"
    note = f"Trùng email với {where}"
    if kept.get("name"):
        note += f" - {kept['name']}"
    if kept.get("source"):
        note += f" (tệp {kept['source']})"
    return note


def dedupe_records(records, issues):
    """Loc trung email tren danh sach `records` (co the gom NHIEU tep/sheet -
    moi record co the co khoa "source"). Moi email chi giu 1 nguoi: nguoi co
    ten KHOP email hon (core.email_name_match_score), ngang diem thi giu nguoi
    xuat hien truoc; giu nguoi do o vi tri xuat hien dau tien cua email. Nguoi
    bi loai duoc them vao `issues` (ly do duplicate_email) kem "note": trung
    voi dong STT nao, ten gi, tep/sheet nao. Tra ve (records_moi, so_bi_loai)."""
    groups = {}
    for i, rec in enumerate(records):
        groups.setdefault(rec["email"], []).append(i)
    best = {}
    for email, idxs in groups.items():
        b = idxs[0]
        for i in idxs[1:]:
            if core.email_name_match_score(records[i]["name"], email) > \
                    core.email_name_match_score(records[b]["name"], email):
                b = i
        best[email] = b
    kept_records, dropped = [], 0
    for i, rec in enumerate(records):
        email = rec["email"]
        if i == groups[email][0]:
            kept_records.append(records[best[email]])
        if i != best[email]:
            issue = {
                "stt": rec["stt"], "raw_name": rec["name"], "raw_email": rec["email"],
                "raw_phone": rec["phone"], "raw_unit": rec["don_vi"], "reason": "duplicate_email",
                "note": _dup_note(records[best[email]]),
            }
            if rec.get("source"):
                issue["source"] = rec["source"]
            issues.append(issue)
            dropped += 1
    return kept_records, dropped


def extract_v2(input_path, default_to_chuc=None, dedupe=True, verbose=True, sheet=None):
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

    sheet: ten sheet cua tep Excel. None -> doc cac sheet co email, bo sheet
    trung (xem _select_excel_tables); co gia tri -> CHI doc dung sheet do.

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
    {stt,name,email,phone,don_vi,to_chuc}, DA loc trung email neu
    dedupe=True, giu THU TU GOC. stats: dict {total_read,
    valid_before_dedupe, duplicates_removed, final_count, to_chuc_missing,
    issues, unit_cache, warnings, recovered_stts, sheets}.
    """
    warnings = []
    sheet_names = []
    tables, fallback_unit, content_based, full_text, header_text = _read_tables_and_context(
        input_path, sheet=sheet, warnings=warnings, sheet_names=sheet_names)
    # Xa/phuong va co quan ban hanh (neu co) cua van ban PDF/Word - dung de gop
    # cac bo phan khong ghi ten xa vao chinh xa do / thay o chuc vu thuan tuy.
    # Excel: tinh lai cho TUNG BANG (moi sheet co the thuoc 1 co quan khac).
    head_commune = _find_doc_commune(header_text)
    head_issuer = None if head_commune else _find_doc_issuer(header_text, full_text or header_text)
    if not tables or not any(tables):
        raise ValueError("Tệp nguồn không có dữ liệu.")

    unit_cache = {}
    issues = []
    clean_records = []
    total_read = 0
    any_header_found = False
    table_stt_sets = []  # (ten sheet, tap STT cua MOI dong trong bang) - Excel/CSV
    name_joined = []     # STT cac dong ho ten bi ngat 2 dong da duoc noi (PDF/Word)

    for t_i, table in enumerate(tables):
        if not table:
            continue
        header_idx, col_map = _find_source_header_row(table)
        if (header_idx is None or ("name" not in col_map and "email" not in col_map)) and not content_based:
            # Excel/CSV KHONG co dong tieu de -> suy ra cot tu noi dung
            # (xem core.infer_columns_from_content).
            inferred = core.infer_columns_from_content(table)
            if inferred:
                header_idx, col_map = -1, inferred
        if header_idx is None or ("name" not in col_map and "email" not in col_map):
            continue  # bang nay khong co dong tieu de phu hop - bo qua (vd bang phu/khong lien quan)
        if not content_based and header_idx >= 0 and "group" not in col_map:
            gcol = _detect_sparse_group_col(table, header_idx, col_map)
            if gcol is not None:
                col_map = dict(col_map, group=gcol)
        any_header_found = True
        data_rows = table[header_idx + 1:]
        doc_commune, doc_issuer = head_commune, head_issuer
        if not content_based and header_idx and header_idx > 0 and doc_commune is None:
            # Cac dong phia tren bang. Noi cac o cua cung 1 dong bang XUONG DONG
            # (khong phai dau cach): "PHƯỜNG TAM LONG" (o A1) + "CỘNG HÒA XÃ HỘI
            # CHỦ NGHĨA..." (o D1) neu noi bang dau cach se thanh 1 ten phuong
            # dai sai.
            above = "\n".join("\n".join(str(c) for c in (r or []) if c not in (None, ""))
                              for r in table[:header_idx])
            doc_commune = _find_doc_commune(above, max_lines=40)
            if doc_commune is None:
                doc_issuer = _find_doc_issuer(above, above)
        if content_based:
            # PDF/Word: cac trang khac nhau CO THE co bo cuc cot khac nhau
            # (da gap thuc te) - dung bo phan tich THEO NOI DUNG (khong
            # phu thuoc vi tri cot) thay vi col_map co dinh.
            total_read += _process_data_rows_content_based(
                data_rows, default_to_chuc, unit_cache, issues, clean_records,
                fallback_unit=fallback_unit, doc_commune=doc_commune, doc_issuer=doc_issuer,
                notes=name_joined,
            )
        else:
            # Ghi nhan STT ca dong nhom/dong khong phai nguoi (vd Son La STT
            # 89 "Công an tỉnh - chưa có danh sách đăng ký") de khong bao
            # nham "nhay so". Tinh RIENG tung sheet (moi sheet danh so rieng).
            stt_i = col_map.get("stt")
            if stt_i is not None:
                stts = set()
                for r in data_rows:
                    n = core.parse_stt(r[stt_i]) if r and stt_i < len(r) else None
                    if n is not None:
                        stts.add(n)
                table_stt_sets.append((sheet_names[t_i] if t_i < len(sheet_names) else "", stts))
            table_fallback = fallback_unit
            if not table_fallback and "unit" not in col_map:
                # Excel/CSV khong co cot Don vi -> lay tu dong tieu de van ban o tren
                table_fallback = _extract_title_unit_from_rows(table, header_idx)
            total_read += _process_data_rows(
                data_rows, col_map, default_to_chuc, unit_cache, issues, clean_records,
                fallback_unit=table_fallback, doc_commune=doc_commune, doc_issuer=doc_issuer,
                # Excel: dong dau du lieu o hang header_idx + 2 (1-based) cua sheet;
                # CSV cung vay. Bang suy ra cot khong tieu de: header_idx = -1.
                row_offset=header_idx + 2,
            )

    if not any_header_found:
        raise ValueError(
            "Không tìm thấy bảng dữ liệu phù hợp trong tệp nguồn (cần ít nhất cột "
            "Họ và tên hoặc Email). Vui lòng kiểm tra lại cấu trúc tệp."
        )

    if name_joined:
        nums = [n for n in (core.parse_stt(j) for j in name_joined) if n is not None]
        warnings.append(f"Đã nối {len(name_joined)} họ tên bị ngắt thành 2 dòng"
                        f"{' (STT ' + core.format_stt_list(nums) + ')' if nums else ''} - "
                        "hãy đối chiếu họ tên với tệp gốc.")

    # --- Luoi an toan bo sung: khoi phuc cac STT bi "nhay so" (pdfplumber
    # bo sot hoan toan khi tach bang) tu van ban tho, neu co ---
    stts_before = _known_stts(clean_records, issues)
    total_read += _recover_missing_stt_rows(clean_records, issues, full_text, default_to_chuc, unit_cache,
                                            doc_commune=head_commune, warnings=warnings, doc_issuer=head_issuer)
    recovered_stts = sorted(_known_stts(clean_records, issues) - stts_before)
    if full_text is None:
        # Excel/CSV/Word: khong co van ban tho de khoi phuc - chi canh bao.
        # Excel: tinh theo TUNG SHEET (sheet 1 co STT 1..43, sheet 2 co 1..36 -
        # gop chung se che mat chỗ nhảy số cua sheet ngan hon).
        sets = table_stt_sets or [("", stts_before)]
        for label, stts in sets:
            gaps = core.find_stt_gaps(stts)
            if gaps:
                where = f"Sheet '{label}': " if label and len(sets) > 1 else ""
                warnings.append(f"{where}STT bị nhảy số ({core.format_stt_list(gaps)}) - không tìm thấy các dòng "
                                "này. Hãy đối chiếu tệp gốc (có thể nguồn đánh số sai, hoặc dòng bị mất khi đọc).")

    for rec in clean_records:
        rec["name"] = core.strip_honorific_by_email(rec["name"], rec["email"])

    # --- Email co ten mien bi cat cut (vd "@laocai.go") -> Can kiem tra ---
    truncated = core.find_truncated_domain_emails([r["email"] for r in clean_records])
    if truncated:
        kept = []
        for rec in clean_records:
            if rec["email"] in truncated:
                issues.append({
                    "stt": rec["stt"], "raw_name": rec["name"], "raw_email": rec["email"],
                    "raw_phone": rec["phone"], "raw_unit": rec["don_vi"], "reason": "invalid_email_format",
                })
            else:
                kept.append(rec)
        clean_records = kept

    # --- Loc trung email (mac dinh bat, giong file mau 1) ---
    duplicates_removed = 0
    if dedupe:
        clean_records, duplicates_removed = dedupe_records(clean_records, issues)

    to_chuc_missing = sum(1 for r in clean_records if not r["to_chuc"])

    # Ca 1 cot trong tren moi ban ghi -> gan nhu chac chan cot do khong duoc
    # nhan dien (tieu de la / lech cot), khong phai nguon thieu that.
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
        "sheets": sheet_names,
    }

    if verbose:
        print_v2_stats_report(stats)

    return clean_records, stats


def source_label(path, sheet=None):
    """Nhan nguon hien trong cot 'Tep nguon': 'tep.xlsx' hoac
    'tep.xlsx [Sheet: ten]'."""
    name = os.path.basename(path)
    return f"{name} [Sheet: {sheet}]" if sheet is not None else name


def extract_v2_batch(items, dedupe=True, verbose=True):
    """
    Chay extract_v2 cho NHIEU nguon va GOP thanh 1 danh sach. Moi nguon la
    (duong_dan, to_chuc_mac_dinh) hoac (duong_dan, sheet, to_chuc_mac_dinh)
    - moi nguon co "To chuc mac dinh" rieng (cac tep/sheet trong 1 lo co the
    thuoc cac tinh khac nhau); sheet=None -> doc theo quy tac mac dinh cua
    extract_v2. Nhan nguon ("source") chi kem ten sheet khi cung 1 tep co
    nhieu nguon.

    Loc trung email tren TOAN BO lo (ke ca khac tep/sheet): moi email giu 1
    nguoi (nguoi co ten khop email hon, ngang diem thi nguoi o nguon dung
    truoc); nguoi bi loai vao "Can kiem tra" kem ghi chu trung voi dong STT
    nao, tep/sheet nao (issue["note"]).

    Nguon doc loi -> ghi canh bao va chay tiep cac nguon con lai (khong de 1
    tep hong lam mat ca lo); tat ca deu loi thi bao loi.

    Moi ban ghi/issue co them khoa "source" (xem source_label). Tra ve
    (records, stats) cung dang voi extract_v2, kem stats["files"] = [{source,
    to_chuc, total_read, final_count, error}] cho tung nguon."""
    all_records = []
    all_issues = []
    warnings = []
    recovered = []
    files = []
    unit_cache = {}
    totals = {"total_read": 0, "to_chuc_missing": 0}

    norm_items = []
    for it in items:
        if len(it) == 2:
            norm_items.append((it[0], None, it[1]))
        else:
            norm_items.append((it[0], it[1], it[2]))

    # Chi ghi ten sheet vao nhan nguon khi cung 1 tep xuat hien nhieu lan (tep
    # Excel duoc tach thanh nhieu sheet); tep chi doc 1 sheet thi nhan gon.
    path_counts = {}
    for path, _sheet, _tc in norm_items:
        path_counts[path] = path_counts.get(path, 0) + 1

    for path, sheet, to_chuc in norm_items:
        source = source_label(path, sheet if path_counts[path] > 1 else None)
        if verbose:
            print(f"=== {source} (Tổ chức mặc định: {to_chuc or '(trống)'}) ===")
        try:
            # dedupe=False: loc trung duoc lam 1 lan tren ca lo o duoi
            records, stats = extract_v2(path, default_to_chuc=to_chuc or None, dedupe=False,
                                        verbose=verbose, sheet=sheet)
        except Exception as e:
            warnings.append(f"{source}: không đọc được tệp - {e}")
            files.append({"source": source, "to_chuc": to_chuc, "total_read": 0, "final_count": 0,
                          "error": str(e)})
            if verbose:
                print(f"CẢNH BÁO: {source}: không đọc được tệp - {e}")
            continue
        for rec in records:
            rec["source"] = source
        for issue in stats["issues"]:
            issue["source"] = source
        all_records.extend(records)
        all_issues.extend(stats["issues"])
        warnings.extend(f"{source}: {w}" for w in stats.get("warnings") or [])
        recovered.extend(stats.get("recovered_stts") or [])
        unit_cache.update(stats.get("unit_cache") or {})
        for k in totals:
            totals[k] += stats[k]
        files.append({"source": source, "to_chuc": to_chuc, "total_read": stats["total_read"],
                      "final_count": len(records), "error": None})

    if files and all(f["error"] for f in files):
        raise ValueError("Không đọc được tệp nào:\n" + "\n".join(f"- {f['source']}: {f['error']}" for f in files))

    duplicates_removed = 0
    if dedupe:
        all_records, duplicates_removed = dedupe_records(all_records, all_issues)
    for f in files:
        f["final_count"] = sum(1 for r in all_records if r["source"] == f["source"])

    stats = {
        "total_read": totals["total_read"],
        "valid_before_dedupe": totals["total_read"] - (len(all_issues) - duplicates_removed),
        "duplicates_removed": duplicates_removed,
        "final_count": len(all_records),
        "to_chuc_missing": sum(1 for r in all_records if not r["to_chuc"]),
        "issues": all_issues, "unit_cache": unit_cache, "warnings": warnings,
        "recovered_stts": recovered, "files": files,
    }
    if verbose and len(items) > 1:
        # Canh bao tung nguon da in o tren - o day chi in tong, nguon doc loi
        # (nhac lai de khong bi troi mat) va canh bao cua ca lo.
        print(f"=== TỔNG CỘNG {len(items)} nguồn ===")
        ok_sources = {f["source"] + ":" for f in files if not f["error"]}
        batch_only = [w for w in warnings if not any(w.startswith(s) for s in ok_sources)]
        print_v2_stats_report(dict(stats, warnings=batch_only, recovered_stts=None))
    return all_records, stats


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
    # Xu ly nhieu tep (extract_v2_batch) -> them cot Tep nguon o dau de doi chieu STT.
    with_source = any(issue.get("source") for issue in issues)
    # Cot "Ghi chu": voi dong trung email - trung voi dong STT nao, tep/sheet nao.
    with_note = any(issue.get("note") for issue in issues)
    header = ["STT", "Họ và tên (gốc)", "Email (gốc)", "Điện thoại (gốc)", "Đơn vị công tác (gốc)", "Lý do"]
    issues_ws.append((["Tệp nguồn"] if with_source else []) + header + (["Ghi chú"] if with_note else []))
    for issue in issues:
        reason_label = core.ISSUE_REASON_LABELS.get(issue["reason"], issue["reason"])
        row = [issue["stt"], issue["raw_name"], issue["raw_email"],
               issue["raw_phone"], issue["raw_unit"], reason_label]
        issues_ws.append(([issue.get("source", "")] if with_source else []) + row
                         + ([issue.get("note", "")] if with_note else []))
    issues_wb.save(issues_output_path)


def process(input_path, template_path, output_path, default_to_chuc=None,
            dedupe=True, issues_output_path=None, password=None, verbose=True):
    """Ham tien ich cho CLI: goi extract_v2() + write_v2_output() +
    write_v2_issues() theo dung thu tu, giu tuong thich nguoc voi cach goi
    cu. Tra ve (so_dong_ghi_duoc, thong_ke)."""
    # input_path co the la 1 duong dan hoac danh sach (nhieu tep dung chung --to-chuc).
    if isinstance(input_path, (list, tuple)):
        records, stats = extract_v2_batch([(p, default_to_chuc) for p in input_path],
                                          dedupe=dedupe, verbose=False)
    else:
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
        # Khoa cache = (don vi tho, xa ban hanh, co quan ban hanh)
        for (raw, _commune, _issuer), (don_vi, to_chuc) in sorted(stats["unit_cache"].items(),
                                                                  key=lambda kv: str(kv[0][0])):
            print(f"  {raw!r:55s} -> Đơn vị={don_vi!r:35s} Tổ chức={to_chuc!r}")

    return stats["final_count"], stats


def main():
    parser = argparse.ArgumentParser(description="Chuẩn hoá dữ liệu đăng ký sang template loại 2.")
    parser.add_argument("--input", required=True, nargs="+",
                         help="Một hoặc nhiều tệp nguồn (Excel/PDF/Word/CSV); nhiều tệp được gộp vào 1 tệp kết quả, "
                              "dùng chung --to-chuc, lọc trùng email trong từng tệp")
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
    inputs = args.input[0] if len(args.input) == 1 else args.input
    process(inputs, args.template, args.output, default_to_chuc=args.to_chuc,
            dedupe=not args.no_dedupe, issues_output_path=args.issues_output, password=args.password)


if __name__ == "__main__":
    main()