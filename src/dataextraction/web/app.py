"""API cho cong cu tao danh sach tai khoan (FastAPI) - chi la lop vo mong.

Toan bo logic trich xuat/chuan hoa dung lai nguyen ven core.py (che do 1) va
mode2.py (che do 2), nen sua quy tac o do la tu dong ap dung cho GUI Tkinter lan web.

Backend KHONG phuc vu giao dien: frontend la du an rieng (thu muc frontend/).

Luong goi:
  1. POST /api/uploads           tai tep len MOT lan, nhan file_id + danh sach sheet + goi y To chuc
  2. POST /api/jobs              gui {mode, sources:[{file_id, sheet, to_chuc}], ...} -> job_id (tra ngay)
  3. GET  /api/jobs/{job_id}     hoi den khi status = done | error
  4. GET  /api/download/{job_id}/{result|issues}

Cau hinh bang bien moi truong (xem docs/web.md):
  DATAEXTRACTION_MAX_UPLOAD_MB   dung luong toi da moi tep (mac dinh 100)
  DATAEXTRACTION_TTL_SECONDS     thoi gian giu tep tai len va ket qua (mac dinh 3600)
  DATAEXTRACTION_WORK_DIR        thu muc tam (mac dinh: thu muc tam cua he dieu hanh)

Chay:  python -m dataextraction.web   (sau pip install -e ".[web]")
"""
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from typing import Literal, Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from .. import TEMPLATE_MODE1, TEMPLATE_MODE2, core
from .. import mode2 as v2


def _env_int(name, default):
    try:
        return int(os.environ.get(name, default))
    except ValueError:
        return default


MAX_UPLOAD_MB = _env_int("DATAEXTRACTION_MAX_UPLOAD_MB", 100)
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
TTL_SECONDS = _env_int("DATAEXTRACTION_TTL_SECONDS", 3600)
WORK_DIR = os.environ.get("DATAEXTRACTION_WORK_DIR") or tempfile.gettempdir()
ALLOWED_EXT = {".pdf", ".docx", ".doc", ".xlsx", ".xls", ".csv"}
SUGGEST_EXT = {".xlsx", ".xls", ".docx", ".doc", ".csv"}  # PDF doc cham nen khong goi y
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="Công cụ tạo danh sách tài khoản")
# JSON ket qua co the vai nghin dong (Son La ~6000) -> nen lai giam manh dung luong tai ve
app.add_middleware(GZipMiddleware, minimum_size=1024)

_uploads = {}  # file_id -> {"path", "name", "size", "dir", "created"}
_jobs = {}  # job_id -> {"status", "created", "dir", "files", "result", "error"}
_state_lock = threading.Lock()
# Cac ham doc tep dung chung bo nho/COM (Word qua pywin32...) khong an toan khi chay
# song song -> xu ly tung viec mot. Viec nang chay trong luong rieng nen /api/health
# va cac yeu cau nhe van tra loi ngay trong luc dang doc tep lon.
_work_lock = threading.Lock()
_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="extract")


def _cleanup():
    now = time.time()
    with _state_lock:
        for fid in [k for k, v in _uploads.items() if now - v["created"] > TTL_SECONDS]:
            shutil.rmtree(_uploads.pop(fid)["dir"], ignore_errors=True)
        for jid in [k for k, v in _jobs.items() if v["status"] != "running" and now - v["created"] > TTL_SECONDS]:
            shutil.rmtree(_jobs.pop(jid)["dir"], ignore_errors=True)


def _safe_name(name):
    base = os.path.basename(name or "tep")
    base = re.sub(r"[^\w.\- ()À-ỹ]", "_", base)
    return base or "tep"


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


# ---------------------------------------------------------------- API ----

@app.get("/api/health")
def health():
    return {"ok": True}


@app.get("/api/config")
def config():
    return {"default_password": core.DEFAULT_PASSWORD, "max_upload_mb": MAX_UPLOAD_MB}


def _inspect_upload(path, ext):
    """Liet ke sheet co email (Excel) + goi y To chuc mac dinh - lam 1 lan ngay luc tai len."""
    sheets, error, suggestion, commune = [], None, "", ""
    with _work_lock:
        if ext in (".xlsx", ".xls"):
            try:
                sheets = core.list_excel_sheets(path)
            except Exception as e:  # tep hong: de buoc trich xuat bao loi ro hon
                error = str(e)
        if ext in SUGGEST_EXT:
            try:
                suggestion = v2.suggest_to_chuc(path) or ""
            except Exception:
                suggestion = ""
        if ext in (".xlsx", ".xls"):
            commune = v2.suggest_commune(path) or ""  # tra None neu khong chac (loi da nuot ben trong)
    return sheets, error, suggestion, commune


@app.post("/api/uploads")
async def upload(file: UploadFile = File(...)):
    _cleanup()
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_EXT:
        raise HTTPException(400, f"Định dạng không hỗ trợ: {file.filename}")
    file_id = uuid.uuid4().hex
    up_dir = tempfile.mkdtemp(prefix=f"up_{file_id[:8]}_", dir=WORK_DIR)
    path = os.path.join(up_dir, _safe_name(file.filename))
    size = 0
    try:
        with open(path, "wb") as f:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise HTTPException(413, f"Tệp quá lớn (tối đa {MAX_UPLOAD_MB} MB): {file.filename}")
                f.write(chunk)
        sheets, error, suggestion, commune = await run_in_threadpool(_inspect_upload, path, ext)
    except BaseException:
        shutil.rmtree(up_dir, ignore_errors=True)
        raise
    with _state_lock:
        _uploads[file_id] = {"path": path, "name": file.filename or os.path.basename(path),
                             "size": size, "dir": up_dir, "created": time.time()}
    return {"file_id": file_id, "name": file.filename, "size": size,
            "sheets": sheets, "suggestion": suggestion,
            "commune_suggestion": commune, "error": error}


@app.delete("/api/uploads/{file_id}", status_code=204)
def delete_upload(file_id: str):
    with _state_lock:
        info = _uploads.pop(file_id, None)
    if info:
        shutil.rmtree(info["dir"], ignore_errors=True)


class Source(BaseModel):
    file_id: str
    sheet: Optional[str] = None
    to_chuc: str = ""
    commune: str = ""  # xa/phuong mac dinh cho don vi khong ghi ten xa (che do 2)


class JobRequest(BaseModel):
    mode: Literal["mode1", "mode2"]
    sources: list[Source] = Field(min_length=1)
    dedupe: bool = True
    password: str = ""


def _run_extraction(req, paths, job_dir, on_progress=None):
    """Chay trong luong nen. paths: list duong dan ung voi req.sources."""
    password = req.password or core.DEFAULT_PASSWORD
    out_path = os.path.join(job_dir, "ket_qua.xlsx")
    issues_path = None
    with _work_lock:
        if req.mode == "mode1":
            records, stats = core.extract_all(paths, dedupe=req.dedupe, verbose=False)
            core.write_output(records, TEMPLATE_MODE1, out_path, password=password)
            rows = [{"name": r["name"], "email": r["email"], "phone": r["phone"],
                     "unit": "", "to_chuc": "", "source": ""} for r in records]
            pipeline = {
                "total_read": stats["total_rows"],
                "missing_email": stats["missing_email"],
                "invalid_email": stats["invalid_email_format"],
                "duplicates": stats["duplicates_removed"],
                "final": stats["final_count"],
            }
        else:
            items = [(p, s.sheet or None, s.to_chuc.strip(), s.commune.strip() or None) for p, s in zip(paths, req.sources)]
            records, stats = v2.extract_v2_batch(items, dedupe=req.dedupe, verbose=False,
                                                progress=on_progress)
            v2.write_v2_output(records, TEMPLATE_MODE2, out_path, password=password)
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
    result = {
        "mode": req.mode,
        "stats": pipeline,
        "records": rows,
        "issues": [_issue_view(i, req.mode) for i in stats["issues"]],
        "warnings": _warnings(stats),
        "files": [{"source": _display_source(f["source"]), "total_read": f["total_read"],
                   "final_count": f["final_count"], "error": f["error"]}
                  for f in stats.get("files") or []],
        "has_issues_file": bool(issues_path),
    }
    return result, {"result": out_path, "issues": issues_path}


def _job_worker(job_id, req, paths):
    job = _jobs[job_id]
    try:
        def on_progress(done, total):
            job["progress"] = {"done": done, "total": total}

        on_progress(0, len(paths))
        result, files = _run_extraction(req, paths, job["dir"], on_progress)
        result["job_id"] = job_id
        job.update(status="done", result=result, files=files)
    except Exception as e:
        job.update(status="error", error=str(e) or e.__class__.__name__)


@app.post("/api/jobs", status_code=202)
def create_job(req: JobRequest):
    _cleanup()
    if req.mode == "mode1" and len(req.sources) != 1:
        raise HTTPException(400, "Chế độ 1 chỉ nhận 1 tệp mỗi lần.")
    paths = []
    with _state_lock:
        for s in req.sources:
            info = _uploads.get(s.file_id)
            if not info:
                raise HTTPException(404, "Tệp tải lên đã hết hạn hoặc không tồn tại. Hãy chọn lại tệp.")
            paths.append(info["path"])
    job_id = uuid.uuid4().hex
    job_dir = tempfile.mkdtemp(prefix=f"job_{job_id[:8]}_", dir=WORK_DIR)
    with _state_lock:
        _jobs[job_id] = {"status": "running", "created": time.time(), "dir": job_dir,
                         "files": {}, "result": None, "error": None}
    _executor.submit(_job_worker, job_id, req, paths)
    return {"job_id": job_id, "status": "running"}


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str):
    with _state_lock:
        job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Kết quả đã hết hạn hoặc không tồn tại. Hãy trích xuất lại.")
    if job["status"] == "done":
        return {"status": "done", "result": job["result"]}
    if job["status"] == "error":
        return {"status": "error", "error": job["error"]}
    return {"status": "running", "progress": job.get("progress")}


@app.get("/api/download/{job_id}/{kind}")
def download(job_id: str, kind: str):
    with _state_lock:
        job = _jobs.get(job_id)
    if not job or kind not in ("result", "issues"):
        raise HTTPException(404, "Kết quả đã hết hạn hoặc không tồn tại. Hãy trích xuất lại.")
    path = job["files"].get(kind)
    if not path or not os.path.exists(path):
        raise HTTPException(404, "Không có tệp này.")
    name = "danh_sach_tai_khoan.xlsx" if kind == "result" else "danh_sach_can_kiem_tra.xlsx"
    return FileResponse(path, filename=name, media_type=XLSX_MEDIA)
