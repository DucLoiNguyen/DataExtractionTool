#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
regression.py
=============
Chay bang hoi quy (muc 7 CLAUDE.md) tren cac file goc dat trong test_data/.
test_data/ KHONG duoc commit (chua du lieu ca nhan) - xem .gitignore.

Cach dung:
    python regression.py            # chay tat ca, bao OK/SAI tung dong
    python regression.py --detail   # in them kiem tra dinh tinh

File nao khong co trong test_data/ thi bao "BO QUA" (khong tinh la loi).
Ten file trong test_data/ la ten ngan (vd TaPhin.pdf) de tranh loi ma hoa
duong dan tieng Viet tren Windows.
"""

import os
import sys
import warnings

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_contacts as core  # noqa: E402
import process_template_v2 as v2  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
warnings.filterwarnings("ignore")

D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_data")

# (ten file, doc, kq)
MODE1 = [
    ("SonLinh.pdf", 25, 25),
    ("KhuDiTich.docx", 43, 42),
    ("DongThap.pdf", 2381, 2239),  # bo 50 dong tieu de nhom; cuu Ho Quang Lon (email bi cat 2 o)
    ("KonDao.pdf", 10, 10),
    ("TaPhin.pdf", 80, 78),        # khoi phuc STT 23/40/55/79 tu van ban tho
    ("DaNang.xlsx", 82, 82),
    ("LongHoa.xlsx", 96, 95),
    ("CanTho.xlsx", 5000, 4968),   # Excel khong ghep dong (truoc: 4999 - mat 1 nguoi)
    ("NinhBinh.xlsx", 5417, 4327), # Support_excel.xlsx - khong co dong tieu de
    ("ChiengKen.pdf", 71, 67),
    ("YenThanh.pdf", 135, 135),
    ("DakLak.xlsx", 330, 325),
    ("BaoTang.doc", 10, 10),       # can Word + pywin32 hoac LibreOffice
]

# (ten file, to chuc mac dinh, doc, kq, ckt)
MODE2 = [
    ("CanTho.xlsx", "Cần Thơ", 5000, 4968, 32),
    ("VINATOM.xls", "Viện Năng lượng nguyên tử Việt Nam", 539, 532, 7),
    ("ThaiNguyen.xls", "Thái Nguyên", 155, 153, 2),
    ("KonDao.pdf", "Quảng Ngãi", 10, 10, 0),
    ("TaPhin.pdf", "Lào Cai", 80, 78, 2),
    ("NghiaDo.pdf", "Lào Cai", 70, 69, 1),
    ("BenCat.xlsx", "Hồ Chí Minh", 23, 22, 1),
    ("BinhHoa.docx", "Hồ Chí Minh", 38, 0, 38),
    ("ChanhHung.docx", "Hồ Chí Minh", 2, 2, 0),
    ("P_LongHoa.doc", "Hồ Chí Minh", 97, 0, 97),
    ("BaoTang.doc", "Hồ Chí Minh", 10, 10, 0),
    ("SonLa.xlsx", "Sơn La", 5981, 5053, 928),
    ("DaNang.xlsx", "Đà Nẵng", 82, 82, 0),
    ("LongHoa.xlsx", "Hồ Chí Minh", 96, 95, 1),
    ("NinhBinh.xlsx", "Ninh Bình", 5417, 4327, 1090),
    ("ChiengKen.pdf", "Lào Cai", 71, 67, 4),
    ("YenThanh.pdf", "Lào Cai", 135, 135, 0),
    ("DakLak.xlsx", "Đắk Lắk", 330, 325, 5),
]


def _quiet(fn, *a, **kw):
    """Chay ham trich xuat nhung chan output (che do 1 in log ra stdout)."""
    import io
    import contextlib
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        return fn(*a, **kw), buf.getvalue()


def unit_checks():
    """Kiem tra nhanh cac ham chuan hoa/nhan dien (khong can test_data)."""
    fails = 0

    def check(cond, msg):
        nonlocal fails
        fails += not cond
        if not cond:
            print(f"SAI {msg}")

    check(core.normalize_phone("912345678.0") == "0912345678", "SĐT chuỗi '912345678.0'")
    check(core.normalize_phone(912345678.0) == "0912345678", "SĐT float 912345678.0")
    check(core.normalize_phone("0912 345 678") == "0912345678", "SĐT có khoảng trắng")
    check(core.normalize_phone("123456789012") == "", "SĐT 12 số -> trống")
    check(core.normalize_email("ten..x@gmail,com") == "ten.x@gmail.com", "email sửa lỗi gõ")
    check(core.normalize_email("nvgiang") is None, "email thiếu @ -> loại")
    # Dong Thap STT 742: lech cot + email bi cat sang o ke ben
    row = ["742", "Hồ Quang Lớn", "Công chức", "919.099.499", "hoquanglon@gmail.c", "om", ""]
    raw, _ = core._find_email_in_row(row, 3, {0, 1})
    check(core.normalize_email(raw) == "hoquanglon@gmail.com", "ghép email bị cắt 2 ô")
    check(core.normalize_phone(core._find_phone_in_row(row, 4, None, 3, {0, 1})) == "0919099499",
          "lấy SĐT từ cột Email khi lệch cột")
    # Khong ghep khi o sau la chu tieng Viet (tranh bia email)
    raw, _ = core._find_email_in_row(["1", "A", "abc@gmail.c", "Chuyên viên"], 2, {0, 1})
    check(core.normalize_email(raw) is None, "không ghép email với chữ thường")
    check(core._is_group_header_row(["", "Ủy ban nhân dân Phường Lo", "ng Thuận", "", ""], 0, 1),
          "tiêu đề nhóm bị cắt 2 ô")
    check(not core._is_group_header_row(["", "Nguyễn Văn A", "Chuyên viên", "", ""], 0, 1),
          "người thật không STT không phải tiêu đề nhóm")
    check(not core._is_group_header_row(["", "Đoàn Văn Nỉ", "Chuyên viên", "", ""], 0, 1),
          "họ Đoàn không phải tiêu đề nhóm")
    check(core.find_stt_gaps({1, 2, 4, 6}) == [3, 5], "tìm STT nhảy số")
    check(all(core.is_stt_header(h) for h in ("STT", "TT", "Stt", "Số TT")), "nhận tiêu đề STT")
    check(set(core.HEADER_KEYWORDS["email"]) <= set(v2.SOURCE_HEADER_KEYWORDS), "từ khoá email đồng bộ")
    check(set(core.HEADER_KEYWORDS["phone"]) <= set(v2.SOURCE_HEADER_KEYWORDS), "từ khoá SĐT đồng bộ")
    check(core.parse_stt(1.0000000000000848) == 1, "STT số thực lệch (Đắk Lắk)")
    # Gop cap duoi xa/phuong ve chinh xa/phuong (Che do 2); doc_commune = xa ban hanh van ban
    for raw, tc, dc, exp in [
        ("phòng vh xã nam định", "Ninh Bình", None, "Xã Nam Định - Ninh Bình"),
        ("phòng kinh tế xã yên thành huyện x", "Lào Cai", None, "Xã Yên Thành - Lào Cai"),
        ("Phòng VH- XH Phường Nam Định", "Ninh Bình", None, "Phường Nam Định - Ninh Bình"),
        ("Ban kinh tế xã hội HĐND xã Tả Phìn", "Lào Cai", None, "Xã Tả Phìn - Lào Cai"),
        ("Đoàn thanh niên CS Hồ Chí Minh xã\nTả Phìn", "Lào Cai", None, "Xã Tả Phìn - Lào Cai"),
        ("UBND Phường Chánh Hưng/ VP HĐND-UBND", "Hồ Chí Minh", None, "Phường Chánh Hưng - Hồ Chí Minh"),
        ("Phường Xã Đàn", "Hà Nội", None, "Phường Xã Đàn - Hà Nội"),
        ("Phòng Nội vụ thị xã Bến Cát", "Hồ Chí Minh", None, "Phòng Nội vụ thị xã Bến Cát - Hồ Chí Minh"),
        ("Hợp tác xã Nông nghiệp An Bình", "Cần Thơ", None, "Hợp tác xã Nông nghiệp An Bình - Cần Thơ"),
        ("Phòng Văn hóa - Xã hội", "Hồ Chí Minh", None, "Phòng Văn hóa - Xã hội - Hồ Chí Minh"),
        ("Chủ tịch UBND phường", "Hồ Chí Minh", None, "Chủ tịch UBND phường - Hồ Chí Minh"),
        ("Phòng Kinh tế xã", "Quảng Ngãi", "Xã Kon Đào", "Xã Kon Đào - Quảng Ngãi"),
        ("xã Chiềng ken", "Lào Cai", "Xã Chiềng Ken", "Xã Chiềng Ken - Lào Cai"),
        ("Bệnh viện Đa khoa huyện Bảo Yên", "Lào Cai", "Xã Nghĩa Đô", "Bệnh viện Đa khoa huyện Bảo Yên - Lào Cai"),
        ("Chi nhánh Ngân hàng Chính sách xã hội", "Lào Cai", "Xã Tả Phìn",
         "Chi nhánh Ngân hàng Chính sách xã hội - Lào Cai"),
    ]:
        got = v2.normalize_don_vi(raw, tc, doc_commune=dc)
        check(got == exp, f"gộp xã/phường: {raw!r} -> {got!r} (kỳ vọng {exp!r})")
    check(v2.determine_to_chuc("Đoàn thanh niên CS Hồ Chí Minh xã Tả Phìn", "Lào Cai") == "TỈNH LÀO CAI",
          "Tổ chức không bị nhận nhầm từ tên 'Hồ Chí Minh' của Đoàn thanh niên")
    print(f"{'OK ' if not fails else 'SAI'} kiểm tra đơn vị ({fails} lỗi)")
    return fails


def run(detail=False):
    fails = unit_checks()
    print("=== CHE DO 1 ===")
    for f, exp_read, exp_kq in MODE1:
        p = os.path.join(D, f)
        if not os.path.exists(p):
            print(f"BO QUA {f} (khong co trong test_data)")
            continue
        (recs, s), log = _quiet(core.extract_all, [p], dedupe=True, verbose=False)
        got = (s["total_rows"], s["final_count"])
        ok = got == (exp_read, exp_kq)
        fails += not ok
        err = " | LOI: " + log.strip().splitlines()[-1] if "Lỗi" in log else ""
        print(f"{'OK ' if ok else 'SAI'} {f:18s} doc={got[0]} kq={got[1]}  (ky vong {exp_read}/{exp_kq}){err}")

        if detail and f == "DongThap.pdf":
            r = [x for x in recs if x["email"] == "hoquanglon@gmail.com"]
            ok = bool(r) and r[0]["phone"] == "0919099499"
            fails += not ok
            print(f"{'OK ' if ok else 'SAI'} Đồng Tháp: giữ Hồ Quang Lớn (email bị cắt 2 ô, lệch cột)")
            ok = any("724, 743, 915, 920, 2315" in w for w in s["warnings"])
            fails += not ok
            print(f"{'OK ' if ok else 'SAI'} Đồng Tháp: cảnh báo STT nhảy số 724, 743, 915, 920, 2315")
        if detail and f == "TaPhin.pdf":
            got = sorted(n for _src, n in s["recovered_stts"])
            fails += got != [23, 40, 55, 79]
            print(f"{'OK ' if got == [23, 40, 55, 79] else 'SAI'} Tả Phìn chế độ 1: khôi phục STT {got}")

    print("=== CHE DO 2 ===")
    results = {}
    for f, tc, exp_read, exp_kq, exp_ckt in MODE2:
        p = os.path.join(D, f)
        if not os.path.exists(p):
            print(f"BO QUA {f} (khong co trong test_data)")
            continue
        try:
            (recs, s), _log = _quiet(v2.extract_v2, p, default_to_chuc=tc, dedupe=True, verbose=False)
        except Exception as e:  # vd .doc khi khong co LibreOffice
            print(f"LOI {f}: {e}")
            continue
        results[f] = (recs, s)
        got = (s["total_read"], s["final_count"], len(s["issues"]))
        ok = got == (exp_read, exp_kq, exp_ckt)
        fails += not ok
        print(f"{'OK ' if ok else 'SAI'} {f:18s} doc={got[0]} kq={got[1]} ckt={got[2]}  "
              f"(ky vong {exp_read}/{exp_kq}/{exp_ckt})")

    # Xu ly nhieu tep 1 lan (extract_v2_batch): ket qua tung tep phai giong het
    # khi chay rieng, moi tep giu dung To chuc mac dinh cua no.
    batch = [(f, tc) for f, tc, *_ in MODE2 if f in ("TaPhin.pdf", "DaNang.xlsx", "KonDao.pdf") and f in results]
    if len(batch) >= 2:
        (recs, s), _log = _quiet(v2.extract_v2_batch, [(os.path.join(D, f), tc) for f, tc in batch],
                                 dedupe=True, verbose=False)
        key = lambda r: (r["email"], r["name"], r["phone"], r["don_vi"], r["to_chuc"])
        ok = all([key(r) for r in recs if r["source"] == f] == [key(r) for r in results[f][0]] for f, _tc in batch)
        ok = ok and len(s["issues"]) == sum(len(results[f][1]["issues"]) for f, _tc in batch)
        fails += not ok
        print(f"{'OK ' if ok else 'SAI'} Nhiều tệp 1 lần ({', '.join(f for f, _ in batch)}) giống chạy riêng")

    if detail:
        print("=== KIEM TRA DINH TINH ===")
        fails += _qualitative(results)

    print(f"\nTong so muc SAI: {fails}")
    return fails


def _qualitative(results):
    fails = 0

    def check(cond, msg):
        nonlocal fails
        fails += not cond
        print(f"{'OK ' if cond else 'SAI'} {msg}")

    def by_email_prefix(recs, prefix):
        return [r for r in recs if r["email"].startswith(prefix)]

    if "BenCat.xlsx" in results:
        r = by_email_prefix(results["BenCat.xlsx"][0], "lthha")
        check(r and r[0]["name"] == "Lê Thị Hồng Hà", "Bến Cát: lthha giữ Lê Thị Hồng Hà")
    if "NghiaDo.pdf" in results:
        r = by_email_prefix(results["NghiaDo.pdf"][0], "trantheanh2")
        check(r and r[0]["name"] == "Trần Thế Anh", "Nghĩa Đô: trantheanh2 giữ Trần Thế Anh")
    if "TaPhin.pdf" in results:
        recs = results["TaPhin.pdf"][0]
        stts = {r["stt"] for r in recs}
        check({"23", "40", "55", "79"} <= stts, "Tả Phìn: có STT 23, 40, 55, 79")
        r40 = [r for r in recs if r["stt"] == "40"]
        # Goc: "Ban kinh tế xã hội HĐND xã Tả Phìn" (khoi phuc tu van ban tho) -> gop ve xa
        check(r40 and r40[0]["don_vi"] == "Xã Tả Phìn - Lào Cai",
              "Tả Phìn: STT 40 (Ban kinh tế xã hội HĐND xã Tả Phìn) -> Xã Tả Phìn - Lào Cai")
        check(all(r["to_chuc"] == "TỈNH LÀO CAI" for r in recs),
              "Tả Phìn: mọi Tổ chức = TỈNH LÀO CAI (kể cả Đoàn thanh niên CS Hồ Chí Minh xã Tả Phìn)")
    if "VINATOM.xls" in results:
        recs = results["VINATOM.xls"][0]
        check(all(r["to_chuc"] == "VIỆN NĂNG LƯỢNG NGUYÊN TỬ VIỆT NAM" for r in recs),
              "VINATOM: mọi Tổ chức = VIỆN NĂNG LƯỢNG NGUYÊN TỬ VIỆT NAM")
    if "DaNang.xlsx" in results:
        recs = results["DaNang.xlsx"][0]
        check(all(r["phone"] for r in recs), "Đà Nẵng: mọi dòng có SĐT")
        check(all(r["don_vi"] == "Văn phòng UBND thành phố Đà Nẵng - Đà Nẵng" for r in recs),
              "Đà Nẵng: đơn vị 'Văn phòng UBND thành phố Đà Nẵng - Đà Nẵng'")
    if "LongHoa.xlsx" in results:
        r = by_email_prefix(results["LongHoa.xlsx"][0], "hvluyen")
        check(r and r[0]["name"] == "Hồ Văn Luyến", "Long Hòa: hvluyen giữ Hồ Văn Luyến")
    if "P_LongHoa.doc" in results:
        check(not results["P_LongHoa.doc"][0], "P Long Hòa (.doc): nguồn không có email -> 0 bản ghi")
    # Moi file che do 2 chi co DUNG 1 gia tri To chuc (tung loi: "Hải Đường" -> Hải Phòng,
    # "Xã Bình Thuận" -> Lâm Đồng, "Đoàn TNCS Hồ Chí Minh" -> TP HCM)
    for f, (recs, _s) in results.items():
        if recs:
            orgs = {r["to_chuc"] for r in recs}
            check(len(orgs) == 1, f"{f}: chỉ 1 Tổ chức ({', '.join(sorted(orgs))})")
    # Gop xa/phuong: moi file chi con 1 don vi
    for f, exp in [("TaPhin.pdf", "Xã Tả Phìn - Lào Cai"), ("NghiaDo.pdf", "Xã Nghĩa Đô - Lào Cai"),
                   ("KonDao.pdf", "Xã Kon Đào - Quảng Ngãi"), ("ChanhHung.docx", "Phường Chánh Hưng - Hồ Chí Minh"),
                   ("LongHoa.xlsx", "Xã Long Hòa - Hồ Chí Minh"), ("ChiengKen.pdf", "Xã Chiềng Ken - Lào Cai"),
                   ("YenThanh.pdf", "Xã Yên Thành - Lào Cai")]:
        if f in results:
            units = {r["don_vi"] for r in results[f][0]}
            check(units == {exp}, f"{f}: đơn vị duy nhất '{exp}' (thực tế {sorted(units)[:3]})")
    if "SonLa.xlsx" in results:
        tx = [r for r in results["SonLa.xlsx"][0] if ".taxua@" in r["email"]]
        check(len(tx) == 32 and all(r["don_vi"] == "Xã Tà Xùa - Sơn La" for r in tx),
              f"Sơn La: 32 người .taxua@ thuộc Xã Tà Xùa (thực tế {len(tx)})")
    if "BenCat.xlsx" in results:
        check(len({r["don_vi"] for r in results["BenCat.xlsx"][0]}) == 9, "Bến Cát: giữ 9 đơn vị gốc")
    return fails


if __name__ == "__main__":
    sys.exit(1 if run(detail="--detail" in sys.argv) else 0)
