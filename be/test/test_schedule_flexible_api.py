from __future__ import annotations

import pytest

from test_sch_reconciliation import make_schedule


def _client():
    pytest.importorskip("httpx", reason="Pasang httpx untuk menjalankan tes integrasi API.")
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from app.routers.schedule import router

    app = FastAPI()
    app.include_router(router)
    return TestClient(app)


def test_validate_endpoint_accepts_one_schedule_file():
    client = _client()
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    response = client.post(
        "/api/schedule/validate",
        files=[("schedule_files", ("MMKI.xlsx", make_schedule(), mime))],
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["valid"] is True
    assert payload["total_files"] == 1
    assert payload["valid_files"] == 1
    assert payload["files"][0]["filename"] == "MMKI.xlsx"


def test_validate_endpoint_accepts_fifteen_schedule_files():
    client = _client()
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    files = [
        ("schedule_files", (f"schedule_{index:02d}.xlsx", make_schedule(), mime))
        for index in range(1, 16)
    ]

    response = client.post("/api/schedule/validate", files=files)

    assert response.status_code == 200
    payload = response.json()
    assert payload["valid"] is True
    assert payload["total_files"] == 15
    assert payload["valid_files"] == 15
    assert len(payload["files"]) == 15


def test_validate_endpoint_rejects_more_than_twenty_schedule_files():
    client = _client()
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    files = [
        ("schedule_files", (f"schedule_{index:02d}.xlsx", make_schedule(), mime))
        for index in range(1, 22)
    ]

    response = client.post("/api/schedule/validate", files=files)

    assert response.status_code == 413
    assert "Maksimal 20" in response.json()["detail"]
