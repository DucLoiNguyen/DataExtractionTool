"""Test tich hop cho API: goi API that, khong mock logic trich xuat."""
import time

import openpyxl
from fastapi.testclient import TestClient

from dataextraction.web.app import app

client = TestClient(app)
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _upload(path):
    r = client.post("/api/uploads", files={"file": (path.name, path.read_bytes(), XLSX)})
    assert r.status_code == 200, r.text
    return r.json()


def _run_job(payload, timeout=30):
    r = client.post("/api/jobs", json=payload)
    assert r.status_code == 202, r.text
    job_id = r.json()["job_id"]
    deadline = time.time() + timeout
    while time.time() < deadline:
        body = client.get(f"/api/jobs/{job_id}").json()
        if body["status"] != "running":
            return job_id, body
        time.sleep(0.05)
    raise AssertionError("job khong ket thuc")


def test_health():
    assert client.get("/api/health").json() == {"ok": True}


def test_backend_does_not_serve_frontend():
    # giao dien la du an rieng (frontend/), backend chi co /api
    assert client.get("/").status_code == 404


def test_config():
    body = client.get("/api/config").json()
    assert body["default_password"] and body["max_upload_mb"] > 0


def test_upload_returns_sheets_and_id(sample_xlsx):
    up = _upload(sample_xlsx)
    assert up["file_id"] and up["name"] == "mau.xlsx"
    assert [s["sheet"] for s in up["sheets"]] == ["DS1"]
    assert up["sheets"][0]["emails"] >= 1


def test_mode1_job(sample_xlsx):
    up = _upload(sample_xlsx)
    _, body = _run_job({"mode": "mode1", "sources": [{"file_id": up["file_id"]}]})
    assert body["status"] == "done"
    result = body["result"]
    assert result["stats"]["final"] == 3
    assert result["stats"]["missing_email"] == 1
    assert result["stats"]["duplicates"] == 1
    emails = {x["email"] for x in result["records"]}
    assert "cuong@gmail.com" in emails  # loi go "gmail,com" duoc tu sua
    assert any(x["phone"] == "0912345678" for x in result["records"])  # thieu so 0 dau


def test_mode2_job_and_download(sample_xlsx):
    up = _upload(sample_xlsx)
    payload = {"mode": "mode2", "sources": [{"file_id": up["file_id"], "sheet": "DS1", "to_chuc": "Cần Thơ"}]}
    job_id, body = _run_job(payload)
    assert body["status"] == "done"
    result = body["result"]
    assert {x["to_chuc"] for x in result["records"]} == {"THÀNH PHỐ CẦN THƠ"}
    assert result["has_issues_file"] and result["job_id"] == job_id

    d = client.get(f"/api/download/{job_id}/result")
    assert d.status_code == 200
    out = sample_xlsx.parent / "out.xlsx"
    out.write_bytes(d.content)
    rows = list(openpyxl.load_workbook(out).active.iter_rows(values_only=True))
    assert rows[0][0] == "Tên tài khoản" and len(rows) == 4
    assert client.get(f"/api/download/{job_id}/issues").status_code == 200


def test_upload_reused_by_several_jobs(sample_xlsx):
    # tai len 1 lan, chay lai nhieu lan (doi mat khau/to chuc ma khong tai lai tep)
    up = _upload(sample_xlsx)
    src = [{"file_id": up["file_id"]}]
    _, a = _run_job({"mode": "mode1", "sources": src})
    _, b = _run_job({"mode": "mode1", "sources": src, "dedupe": False})
    assert a["result"]["stats"]["final"] == 3
    assert b["result"]["stats"]["final"] == 4


def test_mode1_rejects_multiple_files(sample_xlsx):
    up = _upload(sample_xlsx)
    s = {"file_id": up["file_id"]}
    assert client.post("/api/jobs", json={"mode": "mode1", "sources": [s, s]}).status_code == 400


def test_job_unknown_upload():
    r = client.post("/api/jobs", json={"mode": "mode1", "sources": [{"file_id": "khongco"}]})
    assert r.status_code == 404


def test_job_requires_sources():
    assert client.post("/api/jobs", json={"mode": "mode1", "sources": []}).status_code == 422


def test_rejects_unsupported_extension():
    r = client.post("/api/uploads", files={"file": ("x.exe", b"MZ", "application/octet-stream")})
    assert r.status_code == 400


def test_upload_too_large(monkeypatch):
    from dataextraction.web import app as app_module
    monkeypatch.setattr(app_module, "MAX_UPLOAD_BYTES", 10)
    r = client.post("/api/uploads", files={"file": ("a.csv", b"x" * 100, "text/csv")})
    assert r.status_code == 413


def test_delete_upload(sample_xlsx):
    up = _upload(sample_xlsx)
    assert client.delete(f"/api/uploads/{up['file_id']}").status_code == 204
    r = client.post("/api/jobs", json={"mode": "mode1", "sources": [{"file_id": up["file_id"]}]})
    assert r.status_code == 404


def test_unknown_job_and_download():
    assert client.get("/api/jobs/khongco").status_code == 404
    assert client.get("/api/download/khongco/result").status_code == 404


def test_health_stays_responsive_while_job_runs(sample_xlsx, monkeypatch):
    # viec nang chay trong luong rieng: /api/health khong bi chan
    from dataextraction.web import app as app_module
    gate = app_module.threading.Event()
    real = app_module._run_extraction

    def slow(*a, **kw):
        gate.wait(5)
        return real(*a, **kw)

    monkeypatch.setattr(app_module, "_run_extraction", slow)
    up = _upload(sample_xlsx)
    r = client.post("/api/jobs", json={"mode": "mode1", "sources": [{"file_id": up["file_id"]}]})
    job_id = r.json()["job_id"]
    assert client.get(f"/api/jobs/{job_id}").json()["status"] == "running"
    assert client.get("/api/health").json() == {"ok": True}
    gate.set()
    deadline = time.time() + 10
    while time.time() < deadline and client.get(f"/api/jobs/{job_id}").json()["status"] == "running":
        time.sleep(0.05)
    assert client.get(f"/api/jobs/{job_id}").json()["status"] == "done"
