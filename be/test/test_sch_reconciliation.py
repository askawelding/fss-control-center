from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font

from app.services.sch_reconciliation import (
    Category,
    read_master,
    read_schedule,
    reconcile,
)


MASTER_HEADERS = [
    "Date", "Vendor code", "Vendor name", "Recieving Area", "Delivery date",
    "Delivery period", "Firm", "Classification", "PO number", "Item",
    "Parts number", "Parts name", "Order Quantity",
]
SCHEDULE_HEADERS = [
    "Date", "Vendor code", "Vendor name", "Recieving Area", "Delivery date",
    "Delivery time", "Classification", "PO number", "Item", "Parts number",
    "Parts name", "Order Quantity",
]


def _xlsx_bytes(wb: Workbook) -> bytes:
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def make_master() -> bytes:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data FSS"
    ws.append(MASTER_HEADERS)

    # Firm lama sudah ada di schedule.
    ws.append(["20261018", "V1", "Vendor 1", "RA1", "20261019", 1, "", "C1", "PO1", 10, "P1", "Part 1", 100])
    # Firm baru yang akan menggantikan satu Forecast lama (PO/Item di Forecast lama kosong).
    ws.append(["20261018", "V2", "Vendor 2", "RA2", "20261020", 2, "", "C2", "PO2", 20, "P2", "Part 2", 200])
    # Forecast yang tetap tersedia di master terbaru.
    ws.append(["20261018", "V3", "Vendor 3", "RA3", "20261021", 1, "*", "C3", "", "", "P3", "Part 3", 300])
    # Forecast baru.
    ws.append(["20261018", "V4", "Vendor 4", "RA4", "20261022", 1, "*", "C4", "", "", "P4", "Part 4", 400])

    black = Font(color="000000")
    red = Font(color="FF0000")
    for row in ws.iter_rows(min_row=2, max_row=3, min_col=1, max_col=13):
        for cell in row:
            cell.font = black
    for row in ws.iter_rows(min_row=4, max_row=5, min_col=1, max_col=13):
        for cell in row:
            cell.font = red
    return _xlsx_bytes(wb)


def make_schedule() -> bytes:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "DATE"
    # Keep helper columns between payload fields to test header-name mapping.
    headers = ["Date", "Vendor code", "Vendor name", "Recieving Area", "Delivery date", "Delivery time", "Classification", "PO number", "Item", "Parts number", "Parts name", "Order Quantity"]
    ws.append([None, *headers[:6], "HELPER", *headers[6:], None, "Division"])
    # Old Firm that exists in the new master.
    firm_existing = ["20260920", "V1", "Vendor 1", "RA1", "20261019", 1, "C1", "PO1", 10, "P1", "Part 1", 100]
    # Old Firm absent from master: must remain protected.
    firm_legacy = ["20260920", "V9", "Vendor 9", "RA9", "20261025", 1, "C9", "PO9", 90, "P9", "Part 9", 900]
    # Old Forecast transitions to Firm: PO and Item were placeholders.
    fc_transition = ["20260920", "V2", "Vendor 2", "RA2", "20261020", 2, "C2", "", "", "P2", "Part 2", 150]
    # Old Forecast still present in new Forecast.
    fc_stays = ["20260920", "V3", "Vendor 3", "RA3", "20261021", 1, "C3", "", "", "P3", "Part 3", 250]
    # Old Forecast that disappears from the master.
    fc_dropped = ["20260920", "V8", "Vendor 8", "RA8", "20261028", 1, "C8", "", "", "P8", "Part 8", 800]

    def add(values, color: str):
        # Column A is intentionally blank, G is helper, S is extra helper.
        row = [None, *values[:6], "", *values[6:], None, "=IF(TRUE,\"\",\"\")"]
        ws.append(row)
        for col in list(range(2, 8)) + list(range(9, 15)):
            ws.cell(ws.max_row, col).font = Font(color=color)

    add(firm_existing, "000000")
    add(firm_legacy, "000000")
    add(fc_transition, "FF0000")
    add(fc_stays, "FF0000")
    add(fc_dropped, "FF0000")
    return _xlsx_bytes(wb)


def test_master_and_schedule_map_delivery_time_and_classify_colors():
    master_rows, master_map = read_master(make_master(), "20261018.xlsx")
    schedule_rows, schedule_map, summary = read_schedule(make_schedule(), "collect.xlsx")

    assert len(master_rows) == 4
    assert master_map["Delivery period"] == 6
    assert schedule_map["Delivery period"] == 7  # helper column is between payload fields
    assert summary["data_rows"] == 5
    assert summary["firm_rows"] == 2
    assert summary["forecast_rows"] == 3
    assert schedule_rows[0].values["Delivery period"] == "1"


def test_reconciliation_preview_counts_protects_firm_and_detects_transition():
    master_rows, _ = read_master(make_master(), "20261018.xlsx")
    schedule_rows, _, _ = read_schedule(make_schedule(), "collect.xlsx")
    result = reconcile(schedule_rows, master_rows)

    assert result.counts[Category.FIRM_KEPT] == 2
    assert result.counts[Category.FIRM_EXISTS] == 1
    assert result.counts[Category.FIRM_NEW] == 1
    assert result.counts[Category.FORECAST_OLD] == 3
    assert result.counts[Category.FORECAST_NEW] == 2
    assert len(result.transitions) == 1
    assert len(result.dropped_forecast_rows) == 1
    assert not result.blocking


def test_duplicate_firm_master_key_blocks_preview():
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data FSS"
    ws.append(MASTER_HEADERS)
    row = ["20261018", "V1", "Vendor 1", "RA1", "20261019", 1, "", "C1", "PO1", 10, "P1", "Part 1", 100]
    ws.append(row)
    ws.append(row)
    for excel_row in (2, 3):
        for col in range(1, 14):
            ws.cell(excel_row, col).font = Font(color="000000")
    master_rows, _ = read_master(_xlsx_bytes(wb), "duplicate.xlsx")
    schedule_rows, _, _ = read_schedule(make_schedule(), "collect.xlsx")
    result = reconcile(schedule_rows, master_rows)
    assert result.blocking
    assert any("muncul 2x" in message for message in result.blocking)


def test_preview_endpoint_accepts_one_selected_workbook():
    import pytest
    pytest.importorskip("httpx", reason="Pasang httpx untuk menjalankan tes integrasi API.")

    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from app.routers.schedule import router

    app = FastAPI()
    app.include_router(router)
    client = TestClient(app)
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    files = [
        ("master_file", ("20261018.xlsx", make_master(), mime)),
        ("schedule_files", ("MMKI.xlsx", make_schedule(), mime)),
    ]

    response = client.post("/api/schedule/preview", files=files)

    assert response.status_code == 200
    payload = response.json()
    assert payload["valid"] is True
    assert payload["total_files"] == 1
    assert payload["valid_files"] == 1
    assert len(payload["files"]) == 1
    assert payload["files"][0]["filename"] == "MMKI.xlsx"
    assert payload["files"][0]["safe_to_update"] is True
    assert "read-only" in payload["note"]


def test_preview_endpoint_accepts_fifteen_selected_workbooks():
    import pytest
    pytest.importorskip("httpx", reason="Pasang httpx untuk menjalankan tes integrasi API.")

    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from app.routers.schedule import router

    app = FastAPI()
    app.include_router(router)
    client = TestClient(app)
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    files = [("master_file", ("20261018.xlsx", make_master(), mime))]
    files.extend(
        ("schedule_files", (f"schedule_{index:02d}.xlsx", make_schedule(), mime))
        for index in range(1, 16)
    )

    response = client.post("/api/schedule/preview", files=files)

    assert response.status_code == 200
    payload = response.json()
    assert payload["valid"] is True
    assert payload["total_files"] == 15
    assert payload["valid_files"] == 15
    assert len(payload["files"]) == 15
    assert all(item["safe_to_update"] for item in payload["files"])
    assert "read-only" in payload["note"]
