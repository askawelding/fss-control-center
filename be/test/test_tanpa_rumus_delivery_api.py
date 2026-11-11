from __future__ import annotations

import io

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.routers.tanpa_rumus import router


HEADERS = [
    "Date", "Vendor code", "Vendor name", "Receiving Area", "Delivery date",
    "delivery period", "Firm", "Classification", "PO number", "item",
    "parts number", "part name", "order Quantity",
]


def ckd_bytes(codes: list[str]) -> bytes:
    workbook = Workbook()
    date = workbook.active
    date.title = "Date"
    for column, header in enumerate(HEADERS, start=3):
        date.cell(2, column, header)
    for row, code in enumerate(codes, start=3):
        date.cell(row, 4, "D400153")
        date.cell(row, 6, code)
        date.cell(row, 13, f"PART-{row}")
    for name in ("202610", "202611", "202612", "202701"):
        workbook.create_sheet(name)
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DATA_DIR", str(tmp_path))
    app = FastAPI()
    app.include_router(router)
    return TestClient(app)


def test_groups_api_returns_persisted_defaults(client):
    response = client.get("/api/tanpa-rumus/groups")
    assert response.status_code == 200
    groups = response.json()["groups"]
    assert len(groups) == 5
    assert groups[0]["codes"] == ["7R01"]


def test_groups_api_adds_and_edits_multiple_codes(client):
    created = client.post(
        "/api/tanpa-rumus/groups",
        json={"name": "Singapore", "codes": ["7R12", "7R13"], "outputLabel": "7R12 & 7R13 SINGAPORE"},
    )
    assert created.status_code == 201
    group = created.json()["group"]
    assert group["codes"] == ["7R12", "7R13"]

    updated = client.put(
        f"/api/tanpa-rumus/groups/{group['id']}",
        json={"name": "Singapore & Brunei", "codes": ["7R12", "7R13", "7R14"], "outputLabel": "7R12 7R13 7R14 SEA"},
    )
    assert updated.status_code == 200
    assert updated.json()["group"]["codes"] == ["7R12", "7R13", "7R14"]


def test_groups_api_rejects_duplicate_code(client):
    response = client.post(
        "/api/tanpa-rumus/groups",
        json={"name": "Duplicate", "codes": ["7R01", "7R15"], "outputLabel": "7R15 TEST"},
    )
    assert response.status_code == 422
    assert "sudah digunakan" in response.json()["detail"]


def test_groups_api_soft_deletes(client):
    groups = client.get("/api/tanpa-rumus/groups").json()["groups"]
    response = client.delete(f"/api/tanpa-rumus/groups/{groups[0]['id']}")
    assert response.status_code == 200
    after = client.get("/api/tanpa-rumus/groups").json()["groups"]
    assert all(group["id"] != groups[0]["id"] for group in after)


def test_preview_api_reports_unmapped_7r_code(client):
    response = client.post(
        "/api/tanpa-rumus/preview-delivery",
        files={"ckd_file": ("ckd.xlsx", ckd_bytes(["7R01", "7R12", "1R01"]), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["date_sheet"] == "Date"
    assert payload["month_sheets"] == ["202610", "202611", "202612", "202701"]
    assert payload["unmapped_codes"] == ["7R12"]
    assert payload["blank_receiving_rows"] == 0


def test_source_preview_api_checks_mmki_structure_and_reports_elapsed_time(client):
    response = client.post(
        "/api/tanpa-rumus/preview-source",
        files={"workbook_file": ("mmki.xlsx", ckd_bytes(["7R01"]), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["filename"] == "mmki.xlsx"
    assert payload["date_sheet"] == "Date"
    assert payload["month_sheets"] == ["202610", "202611", "202612", "202701"]
    assert payload["processing_seconds"] >= 0


def test_group_process_is_blocked_for_unmapped_country_code(client):
    response = client.post(
        "/api/tanpa-rumus/process-delivery",
        data={
            "make_mmki": "false",
            "make_ckd_all": "false",
            "make_ckd_groups": "true",
            "selected_group_ids": '["vietnam"]',
        },
        files={"ckd_file": ("ckd.xlsx", ckd_bytes(["7R01", "7R12"]), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 422
    assert "7R12" in response.json()["detail"]
    history = client.get("/api/tanpa-rumus/history").json()["items"]
    assert len(history) == 1
    assert history[0]["status"] == "failed"


def test_preview_api_rejects_workbook_without_date_sheet(client):
    workbook = Workbook()
    workbook.active.title = "202610"
    output = io.BytesIO()
    workbook.save(output)
    response = client.post(
        "/api/tanpa-rumus/preview-delivery",
        files={"ckd_file": ("ckd.xlsx", output.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 422
    assert "Date" in response.json()["detail"]


def test_history_api_returns_saved_group_snapshot(client):
    from app.services.tanpa_rumus_config import list_groups, record_run

    groups = list_groups()
    record_run(
        source_files=["20261018 CKD DELIVERY SCH.xlsx"],
        selected_outputs=[{"kind": "ckd_group", "group_id": "vietnam", "codes": ["7R01"]}],
        group_snapshot=groups,
        status="succeeded",
        report={"total_outputs": 1, "succeeded": 1, "failed": 0, "files": []},
    )
    response = client.get("/api/tanpa-rumus/history")
    assert response.status_code == 200
    items = response.json()["items"]
    assert len(items) == 1
    assert items[0]["sourceFiles"] == ["20261018 CKD DELIVERY SCH.xlsx"]
    assert items[0]["groups"][0]["codes"] == ["7R01"]
