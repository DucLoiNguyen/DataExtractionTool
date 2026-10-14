"""Test tich hop cho web app: goi API that, khong mock logic trich xuat."""
import json

import openpyxl
from fastapi.testclient import TestClient

from webapp.app import app

client = TestClient(app)
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _upload(path):
    return ("files", (path.name, path.read_bytes(), XLSX))


def test_health():
    assert client.get("/api/health").json() == {"ok": True}


def test_index_served():
    r = client.get("/")
    assert r.status_code == 200 and "<html" in r.text


def test_mode1_extract(sample_xlsx):
    r = client.post("/api/extract", data={"mode": "mode1"}, files=[_upload(sample_xlsx)])
    assert r.status_code == 200
    body = r.json()
    assert body["stats"]["final"] == 3
    assert body["stats"]["missing_email"] == 1
    assert body["stats"]["duplicates"] == 1
    emails = {x["email"] for x in body["records"]}
    assert "cuong@gmail.com" in emails  # loi go "gmail,com" duoc tu sua
    assert any(x["phone"] == "0912345678" for x in body["records"])  # thieu so 0 dau


def test_mode2_extract_and_download(sample_xlsx):
    sources = json.dumps([{"file": 0, "sheet": "DS1", "to_chuc": "Cần Thơ"}])
    r = client.post("/api/extract", data={"mode": "mode2", "sources": sources}, files=[_upload(sample_xlsx)])
    assert r.status_code == 200
    body = r.json()
    assert {x["to_chuc"] for x in body["records"]} == {"THÀNH PHỐ CẦN THƠ"}
    assert body["has_issues_file"]

    d = client.get(f"/api/download/{body['job_id']}/result")
    assert d.status_code == 200
    out = sample_xlsx.parent / "out.xlsx"
    out.write_bytes(d.content)
    rows = list(openpyxl.load_workbook(out).active.iter_rows(values_only=True))
    assert rows[0][0] == "Tên tài khoản" and len(rows) == 4
    assert client.get(f"/api/download/{body['job_id']}/issues").status_code == 200


def test_mode1_rejects_multiple_files(sample_xlsx):
    r = client.post("/api/extract", data={"mode": "mode1"}, files=[_upload(sample_xlsx), _upload(sample_xlsx)])
    assert r.status_code == 400


def test_rejects_unsupported_extension(tmp_path):
    bad = tmp_path / "x.exe"
    bad.write_bytes(b"MZ")
    r = client.post("/api/extract", data={"mode": "mode1"}, files=[("files", ("x.exe", bad.read_bytes(), "application/octet-stream"))])
    assert r.status_code == 400


def test_download_unknown_job():
    assert client.get("/api/download/khongco/result").status_code == 404
