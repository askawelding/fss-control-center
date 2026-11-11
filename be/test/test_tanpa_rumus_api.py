from __future__ import annotations

import io
import json
import shutil
import zipfile
from pathlib import Path
from urllib.parse import unquote

import pytest

import tanpa_rumus_testkit as kit

MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.fixture
def client(fake_excel):
    pytest.importorskip("httpx", reason="Pasang httpx untuk menjalankan tes integrasi API.")
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from app.routers.tanpa_rumus import router

    app = FastAPI()
    app.include_router(router)
    return TestClient(app)


def _files(items):
    return [("files", (name, data, MIME)) for name, data in items]


def test_process_returns_zip_with_outputs_and_report(client):
    response = client.post("/api/tanpa-rumus/process", files=_files(kit.many_workbooks(3)))

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/zip"
    assert "FSS_Control_Center_Tanpa_Rumus.zip" in response.headers["content-disposition"]
    summary = json.loads(unquote(response.headers["x-tanpa-rumus-summary"]))
    assert summary == {"total_files": 3, "succeeded": 3, "failed": 0, "formulas_converted": 9, "warning_count": 0}

    archive = zipfile.ZipFile(io.BytesIO(response.content))
    assert archive.testzip() is None
    assert sorted(archive.namelist()) == sorted([
        "schedule_01_TANPA_RUMUS.xlsx", "schedule_02_TANPA_RUMUS.xlsx", "schedule_03_TANPA_RUMUS.xlsx",
        "manifest.json", "laporan_tanpa_rumus.csv",
    ])
    manifest = json.loads(archive.read("manifest.json"))
    assert manifest["status"] == "succeeded" and manifest["total_formulas_converted"] == 9


def test_process_accepts_single_file(client):
    response = client.post("/api/tanpa-rumus/process", files=_files([("satu.xlsx", kit.simple_workbook())]))
    assert response.status_code == 200
    assert "satu_TANPA_RUMUS.xlsx" in zipfile.ZipFile(io.BytesIO(response.content)).namelist()


def test_partial_success_returns_zip_with_only_successful_files(client):
    response = client.post(
        "/api/tanpa-rumus/process",
        files=_files([("bagus.xlsx", kit.simple_workbook()), ("rusak.xlsx", b"bukan excel")]),
    )
    assert response.status_code == 200
    archive = zipfile.ZipFile(io.BytesIO(response.content))
    assert sorted(archive.namelist()) == ["bagus_TANPA_RUMUS.xlsx", "laporan_tanpa_rumus.csv", "manifest.json"]
    manifest = json.loads(archive.read("manifest.json"))
    assert manifest["status"] == "partial"
    assert manifest["files"][1]["status"] == "failed" and "rusak" in manifest["files"][1]["error"]
    assert json.loads(unquote(response.headers["x-tanpa-rumus-summary"]))["failed"] == 1


def test_all_files_failing_returns_422_with_report(client):
    response = client.post("/api/tanpa-rumus/process", files=_files([("rusak.xlsx", b"bukan excel"), ("a.csv", b"x,y")]))
    assert response.status_code == 422
    payload = response.json()
    assert isinstance(payload["detail"], str) and "Tidak ada workbook yang berhasil" in payload["detail"]
    assert payload["report"]["status"] == "failed" and payload["report"]["failed"] == 2
    assert kit.FakeApp.registry == []  # tidak perlu menjalankan Excel


def test_missing_files_is_rejected(client):
    assert client.post("/api/tanpa-rumus/process").status_code == 422


def test_more_than_twenty_files_is_rejected(client):
    response = client.post("/api/tanpa-rumus/process", files=_files(kit.many_workbooks(21)))
    assert response.status_code == 413
    assert "Maksimal 20" in response.json()["detail"]


def test_oversized_file_is_reported_per_file(client, monkeypatch):
    from app.routers import tanpa_rumus as router_module

    monkeypatch.setattr(router_module, "MAX_FILE_SIZE_BYTES", 100)
    response = client.post(
        "/api/tanpa-rumus/process",
        files=_files([("besar.xlsx", kit.simple_workbook()), ("kecil.xlsx", b"x" * 10)]),
    )
    assert response.status_code == 422  # keduanya gagal: satu terlalu besar, satu bukan workbook
    errors = [item["error"] for item in response.json()["report"]["files"]]
    assert "melebihi batas 50 MB" in errors[0]


def test_excel_unavailable_returns_503(client, fake_excel):
    fake_excel.launch_error = True
    response = client.post("/api/tanpa-rumus/process", files=_files([("a.xlsx", kit.simple_workbook())]))
    assert response.status_code == 503
    assert "Excel" in response.json()["detail"]


def test_work_directory_is_removed_after_download(client, monkeypatch):
    from app.routers import tanpa_rumus as router_module

    seen: list[Path] = []
    original = router_module.create_tanpa_rumus_zip

    def recording(inputs):
        result = original(inputs)
        seen.append(result[1])
        return result

    monkeypatch.setattr(router_module, "create_tanpa_rumus_zip", recording)
    response = client.post("/api/tanpa-rumus/process", files=_files([("a.xlsx", kit.simple_workbook())]))
    assert response.status_code == 200
    assert seen and not seen[0].exists()
    shutil.rmtree(seen[0], ignore_errors=True)


def test_route_is_registered_in_main_app_and_header_is_exposed():
    pytest.importorskip("httpx")
    from app.main import app

    assert any(getattr(route, "path", "") == "/api/tanpa-rumus/process" for route in app.routes)
    cors = next(m for m in app.user_middleware if m.cls.__name__ == "CORSMiddleware")
    assert "X-Tanpa-Rumus-Summary" in cors.kwargs["expose_headers"]
