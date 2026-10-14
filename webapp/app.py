"""Web app cho cong cu tao danh sach tai khoan (FastAPI).

Chi la lop vo mong: toan bo logic trich xuat/chuan hoa dung lai nguyen ven
extract_contacts.py (che do 1) va process_template_v2.py (che do 2), nen sua
quy tac o do la tu dong ap dung cho ca GUI Tkinter lan web.

Chay:  uvicorn webapp.app:app --reload     (tu thu muc goc cua repo)
"""
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import uuid

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import extract_contacts as core  # noqa: E402
import process_template_v2 as v2  # noqa: E402

TEMPLATE_V1 = os.path.join(ROOT, "cls_template_users.xlsx")
TEMPLATE_V2 = os.path.join(ROOT, "TemplateV2.xlsx")
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

JOB_TTL_SECONDS = 3600
MAX_UPLOAD_BYTES = 100 * 1024 * 1024
ALLOWED_EXT = {".pdf", ".docx", ".doc", ".xlsx", ".xls", ".csv"}

app = FastAPI(title="Công cụ tạo danh sách tài khoản")

_jobs = {}  # job_id -> {"dir": ..., "created": ..., "files": {...}}
_jobs_lock = threading.Lock()
# Cac ham doc tep dung chung bo nho/COM (Word qua pywin32...) khong an toan khi
# chay song song -> xu ly tung yeu cau mot.
_work_lock = threading.Lock()


def _cleanup_jobs():
    now = time.time()
    with _jobs_lock:
        for jid in [j for j, info in _jobs.items() if now - info["created"] > JOB_TTL_SECONDS]:
            shutil.rmtree(_jobs[jid]["dir"], ignore_errors=True)
            del _jobs[jid]


def _safe_name(name):
    base = os.path.basename(name or "tep")
    base = re.sub(r"[^\w.\- ()À-ỹ]", "_", base)
    return base or "tep"


async def _save_upload(upload, dest_dir, index):
    ext = os.path.splitext(upload.filename or "")[1].lower()
    if ext not in ALLOWED_EXT:
        raise HTTPException(400, f"Định dạng không hỗ trợ: {upload.filename}")
    # Moi tep 1 thu muc con rieng: 2 tep cung ten khong de len nhau va ten tep
    # (hien trong cot "Tep nguon", ghi chu trung email) van dung la ten goc.
    sub_dir = os.path.join(dest_dir, str(index))
    os.makedirs(sub_dir, exist_ok=True)
    path = os.path.join(sub_dir, _safe_name(upload.filename))
    size = 0
    with open(path, "wb") as f:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise HTTPException(413, f"Tệp quá lớn (tối đa 100 MB): {upload.filename}")
            f.write(chunk)
    return path


def _display_source(label):
    return str(label) if label else ""


def _issue_view(issue, mode):
    if mode == "mode1":
        stt = issue.get("row_number")
        unit = ""
    else:
        stt = issue.get("stt")
        unit = issue.get("raw_unit", "")
    return {
        "source": _display_source(issue.get("source", "")),
        "stt": "" if stt is None else str(stt),
        "name": issue.get("raw_name") or "",
        "email": issue.get("raw_email") or "",
        "phone": issue.get("raw_phone") or "",
        "unit": unit or "",
        "reason": core.ISSUE_REASON_LABELS.get(issue.get("reason"), issue.get("reason", "")),
        "note": issue.get("note", "") or "",
    }


def _warnings(stats):
    out = []
    by_src = {}
    for src, n in stats.get("recovered_stts") or []:
        by_src.setdefault(src, []).append(n)
    for src, nums in by_src.items():
        out.append(f"Đã khôi phục STT {core.format_stt_list(nums)} từ văn bản gốc ({_display_source(src)}) "
                   "- pdfplumber bỏ sót khi tách bảng.")
    out.extend(_display_source(w) for w in stats.get("warnings") or [])
    if stats.get("to_chuc_missing"):
        out.append(f"{stats['to_chuc_missing']} dòng không xác định được 'Tổ chức' (để trống) - "
                   "kiểm tra lại ô Tổ chức mặc định hoặc dữ liệu nguồn.")
    return out


@app.get("/api/health")
def health():
    return {"ok": True}


@app.post("/api/sheets")
async def list_sheets(file: UploadFile = File(...)):
    """Liet ke sheet co email cua 1 tep Excel (de nguoi dung gan To chuc rieng tung sheet)."""
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in (".xlsx", ".xls"):
        return {"sheets": []}
    tmp = tempfile.mkdtemp(prefix="sheets_")
    try:
        path = await _save_upload(file, tmp, 0)
        with _work_lock:
            try:
                sheets = core.list_excel_sheets(path)
            except Exception as e:  # tep hong: de buoc trich xuat bao loi ro hon
                return {"sheets": [], "error": str(e)}
        return {"sheets": sheets}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.post("/api/suggest-to-chuc")
async def suggest_to_chuc(file: UploadFile = File(...)):
    """Goi y To chuc mac dinh tu ten mien email (chi Excel/Word, PDF doc cham nen bo qua)."""
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in (".xlsx", ".xls", ".docx", ".doc", ".csv"):
        return {"suggestion": ""}
    tmp = tempfile.mkdtemp(prefix="suggest_")
    try:
        path = await _save_upload(file, tmp, 0)
        with _work_lock:
            try:
                return {"suggestion": v2.suggest_to_chuc(path) or ""}
            except Exception:
                return {"suggestion": ""}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.post("/api/extract")
async def extract(
    mode: str = Form(...),
    files: list[UploadFile] = File(...),
    sources: str = Form("[]"),
    dedupe: bool = Form(True),
    password: str = Form(""),
):
    """mode = mode1 | mode2.
    sources (che do 2): JSON [{"file": <chi so tep>, "sheet": <ten|null>, "to_chuc": "..."}].
    Che do 1: chi nhan 1 tep (nhu GUI)."""
    _cleanup_jobs()
    if mode not in ("mode1", "mode2"):
        raise HTTPException(400, "mode phải là mode1 hoặc mode2")
    if mode == "mode1" and len(files) != 1:
        raise HTTPException(400, "Chế độ 1 chỉ nhận 1 tệp mỗi lần.")
    password = password or core.DEFAULT_PASSWORD

    job_id = uuid.uuid4().hex
    job_dir = tempfile.mkdtemp(prefix=f"job_{job_id[:8]}_")
    try:
        paths = [await _save_upload(f, job_dir, i) for i, f in enumerate(files)]
        try:
            src_list = json.loads(sources) or []
        except json.JSONDecodeError:
            raise HTTPException(400, "Trường sources không phải JSON hợp lệ.")

        with _work_lock:
            if mode == "mode1":
                records, stats = core.extract_all(paths, dedupe=dedupe, verbose=False)
                out_path = os.path.join(job_dir, "ket_qua.xlsx")
                core.write_output(records, TEMPLATE_V1, out_path, password=password)
                issues_path = None
                rows = [{"name": r["name"], "email": r["email"], "phone": r["phone"],
                         "unit": "", "to_chuc": ""} for r in records]
                total_read = stats["total_rows"]
                pipeline = {
                    "total_read": total_read,
                    "missing_email": stats["missing_email"],
                    "invalid_email": stats["invalid_email_format"],
                    "duplicates": stats["duplicates_removed"],
                    "final": stats["final_count"],
                }
            else:
                if not src_list:  # khong khai bao nguon -> moi tep 1 nguon, To chuc trong
                    src_list = [{"file": i, "sheet": None, "to_chuc": ""} for i in range(len(paths))]
                items = []
                for s in src_list:
                    idx = int(s.get("file", 0))
                    if not 0 <= idx < len(paths):
                        raise HTTPException(400, f"Chỉ số tệp không hợp lệ: {idx}")
                    items.append((paths[idx], s.get("sheet") or None, (s.get("to_chuc") or "").strip()))
                records, stats = v2.extract_v2_batch(items, dedupe=dedupe, verbose=False)
                out_path = os.path.join(job_dir, "ket_qua.xlsx")
                v2.write_v2_output(records, TEMPLATE_V2, out_path, password=password)
                issues_path = None
                if stats["issues"]:
                    issues_path = os.path.join(job_dir, "ket_qua_can_kiem_tra.xlsx")
                    v2.write_v2_issues(stats["issues"], issues_path)
                rows = [{"name": r["name"], "email": r["email"], "phone": r["phone"],
                         "unit": r.get("don_vi", ""), "to_chuc": r.get("to_chuc", ""),
                         "source": _display_source(r.get("source", ""))} for r in records]
                pipeline = {
                    "total_read": stats["total_read"],
                    "missing_email": sum(1 for i in stats["issues"] if i["reason"] == "missing_email"),
                    "invalid_email": sum(1 for i in stats["issues"] if i["reason"] == "invalid_email_format"),
                    "duplicates": stats["duplicates_removed"],
                    "final": stats["final_count"],
                }
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(422, str(e))

    with _jobs_lock:
        _jobs[job_id] = {"dir": job_dir, "created": time.time(),
                         "files": {"result": out_path, "issues": issues_path}}
    return {
        "job_id": job_id,
        "mode": mode,
        "stats": pipeline,
        "records": rows,
        "issues": [_issue_view(i, mode) for i in stats["issues"]],
        "warnings": _warnings(stats),
        "files": [{"source": _display_source(f["source"]), "total_read": f["total_read"],
                   "final_count": f["final_count"], "error": f["error"]}
                  for f in stats.get("files") or []],
        "has_issues_file": bool(issues_path),
    }


@app.get("/api/download/{job_id}/{kind}")
def download(job_id: str, kind: str):
    with _jobs_lock:
        info = _jobs.get(job_id)
    if not info or kind not in ("result", "issues"):
        raise HTTPException(404, "Kết quả đã hết hạn hoặc không tồn tại. Hãy trích xuất lại.")
    path = info["files"].get(kind)
    if not path or not os.path.exists(path):
        raise HTTPException(404, "Không có tệp này.")
    name = "danh_sach_tai_khoan.xlsx" if kind == "result" else "danh_sach_can_kiem_tra.xlsx"
    return FileResponse(path, filename=name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/config")
def config():
    return {"default_password": core.DEFAULT_PASSWORD}


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
