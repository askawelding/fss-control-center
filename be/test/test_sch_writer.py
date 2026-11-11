from __future__ import annotations

import pytest

from app.services.sch_reconciliation import (
    SCHEDULE_COLUMNS,
    read_master,
    read_schedule,
    reconcile,
)
from app.services.sch_writer import (
    SCHUpdateError,
    _coerce_value_for_target,
    _master_to_schedule_values,
    build_write_plan,
    _iter_write_chunks,
)

from test_sch_reconciliation import make_master, make_schedule


def _fixture():
    master_rows, _ = read_master(make_master(), "20261018.xlsx")
    schedule_rows, column_map, summary = read_schedule(make_schedule(), "collect.xlsx")
    result = reconcile(schedule_rows, master_rows)
    return master_rows, schedule_rows, column_map, summary, result


def test_formula_coverage_accepts_full_helper_columns_and_rejects_a_middle_gap():
    from io import BytesIO
    from openpyxl import load_workbook
    from app.services.sch_writer import _check_formula_coverage, _formula_refs_at_rows

    source = make_schedule()
    first_row_refs = _formula_refs_at_rows(source, "DATE", {2})[2]
    assert "P2" in first_row_refs
    assert _check_formula_coverage(source, "DATE", 4, 6, first_row_refs) == ["P"]

    workbook = load_workbook(BytesIO(source))
    worksheet = workbook["DATE"]
    worksheet["P5"] = None
    changed = BytesIO()
    workbook.save(changed)

    with pytest.raises(SCHUpdateError, match="Formula helper belum lengkap"):
        _check_formula_coverage(changed.getvalue(), "DATE", 4, 6, first_row_refs)


def test_write_plan_keeps_helper_gap_out_of_payload_runs():
    _, rows, column_map, summary, result = _fixture()
    plan = build_write_plan(rows, column_map, summary["header_row"], result)

    assert [(run.value_offset, run.first_column, run.last_column) for run in plan.runs] == [
        (0, 2, 7),
        (6, 9, 14),
    ]
    assert plan.start_row == 4
    assert plan.firm_add_count == 1
    assert plan.forecast_add_count == 2
    assert plan.end_row == 6
    assert len(plan.rows) == 3
    assert plan.forecast_flags == (False, True, True)


def test_write_plan_rejects_reconciliation_review_or_blocking():
    _, rows, column_map, summary, result = _fixture()
    result.review.append("Ada perubahan nilai yang perlu diperiksa.")

    with pytest.raises(SCHUpdateError, match="perlu diperiksa"):
        build_write_plan(rows, column_map, summary["header_row"], result)


def test_write_plan_rejects_gap_in_active_schedule_rows():
    _, rows, column_map, summary, result = _fixture()
    rows_with_gap = [row for row in rows if row.row_no != rows[2].row_no]

    with pytest.raises(SCHUpdateError, match="celah"):
        build_write_plan(rows_with_gap, column_map, summary["header_row"], result)


def test_master_item_is_normalized_before_schedule_write():
    master_rows, *_ = _fixture()
    firm = next(row for row in master_rows if not row.is_forecast)
    values = _master_to_schedule_values(firm)
    item_index = SCHEDULE_COLUMNS.index("Item")

    assert values[item_index] == "10"
    assert _coerce_value_for_target(values[item_index], 10, "00000", "Item") == 10


def test_leading_zero_code_stays_text_when_target_does_not_format_it():
    assert _coerce_value_for_target("000123", 123, "General", "PO number") == "000123"
    assert _coerce_value_for_target("000123", 123, "000000", "PO number") == 123


def test_numeric_target_receives_number_and_text_target_receives_text():
    assert _coerce_value_for_target("20261025", 20261018, "General", "Date") == 20261025
    assert _coerce_value_for_target("1", 14, "0", "Delivery period") == 1
    assert _coerce_value_for_target("Part description", "Old part", "General", "Parts name") == "Part description"


def test_write_plan_splits_all_four_real_schedule_layouts():
    from app.services.sch_writer import _runs_for_columns

    layouts = {
        "collect_sch": {
            "Date": 4, "Vendor code": 5, "Vendor name": 6, "Receiving Area": 7,
            "Delivery date": 8, "Delivery period": 9, "Classification": 11,
            "PO number": 12, "Item": 13, "Parts number": 14, "Parts name": 15, "Order Quantity": 16,
        },
        "delivery_sch": {
            "Date": 2, "Vendor code": 3, "Vendor name": 4, "Receiving Area": 5,
            "Delivery date": 6, "Delivery period": 7, "Classification": 9,
            "PO number": 10, "Item": 11, "Parts number": 12, "Parts name": 13, "Order Quantity": 14,
        },
        "collect_ckd": {
            "Date": 4, "Vendor code": 5, "Vendor name": 6, "Receiving Area": 7,
            "Delivery date": 8, "Delivery period": 9, "Classification": 11,
            "PO number": 12, "Item": 13, "Parts number": 14, "Parts name": 15, "Order Quantity": 16,
        },
        "delivery_ckd": {
            "Date": 3, "Vendor code": 4, "Vendor name": 5, "Receiving Area": 6,
            "Delivery date": 7, "Delivery period": 8, "Classification": 10,
            "PO number": 11, "Item": 12, "Parts number": 13, "Parts name": 14, "Order Quantity": 15,
        },
    }
    expected = {
        "collect_sch": [(0, 4, 9), (6, 11, 16)],
        "delivery_sch": [(0, 2, 7), (6, 9, 14)],
        "collect_ckd": [(0, 4, 9), (6, 11, 16)],
        "delivery_ckd": [(0, 3, 8), (6, 10, 15)],
    }

    for role, column_map in layouts.items():
        runs = _runs_for_columns(column_map)
        actual = [(run.value_offset, run.first_column, run.last_column) for run in runs]
        assert actual == expected[role], role


def test_write_chunks_split_at_existing_data_boundary():
    chunks = list(_iter_write_chunks(
        total_rows=45_548,
        start_row=27_248,
        old_last_row=70_199,
        chunk_size=10_000,
    ))

    assert chunks == [
        (0, 10_000),
        (10_000, 20_000),
        (20_000, 30_000),
        (30_000, 40_000),
        (40_000, 42_952),
        (42_952, 45_548),
    ]


def test_create_update_zip_accepts_one_selected_workbook(monkeypatch, tmp_path):
    import zipfile
    import app.services.sch_writer as writer

    master_bytes = make_master()
    schedule_bytes = make_schedule()

    def fake_apply(source_path, output_path, input_bytes, plan, column_map, calculation_timeout_seconds):
        output_path.write_bytes(input_bytes)
        return {"calculated": True, "calculation_seconds": 0.0}

    monkeypatch.setattr(writer, "_apply_plan_with_excel", fake_apply)
    monkeypatch.setattr(writer, "_validate_written_output", lambda *args, **kwargs: {
        "data_rows": 5,
        "firm_rows": 2,
        "forecast_rows": 3,
        "ambiguous_rows": 0,
    })

    zip_path, work_dir, manifest = writer.create_update_zip(
        master_bytes,
        "20261018.xlsx",
        {"workbook_1": ("MMKI.xlsx", schedule_bytes)},
    )

    try:
        assert manifest["total_files"] == 1
        assert len(manifest["files"]) == 1
        assert manifest["files"][0]["source"] == "MMKI.xlsx"
        with zipfile.ZipFile(zip_path) as archive:
            names = archive.namelist()
            assert len([name for name in names if name.endswith(".xlsx")]) == 1
            assert "manifest.json" in names
    finally:
        import shutil
        shutil.rmtree(work_dir, ignore_errors=True)


def test_update_endpoint_accepts_one_file_but_still_requires_confirmation():
    pytest.importorskip("httpx", reason="Pasang httpx untuk menjalankan tes integrasi API.")
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from app.routers.schedule import router

    app = FastAPI()
    app.include_router(router)
    client = TestClient(app)
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    files = [
        ("master_file", ("master.xlsx", make_master(), mime)),
        ("schedule_files", ("MMKI.xlsx", make_schedule(), mime)),
    ]

    response = client.post("/api/schedule/update", files=files, data={"confirm_update": "false"})

    assert response.status_code == 400
    assert "Konfirmasi proses update" in response.json()["detail"]



def test_set_calc_mode_preserves_ignorable_namespace_prefixes(tmp_path):
    import zipfile
    from xml.etree import ElementTree as ET
    from app.services.sch_writer import _set_calc_mode

    workbook_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'mc:Ignorable="x15" '
        b'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
        b'<sheets/><calcPr calcId="191029" calcMode="manual"/></workbook>'
    )
    path = tmp_path / "namespace_test.xlsx"
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("[Content_Types].xml", b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')

    _set_calc_mode(path, "auto")

    with zipfile.ZipFile(path, "r") as archive:
        assert archive.testzip() is None
        updated = archive.read("xl/workbook.xml")
    assert b'mc:Ignorable="x15"' in updated
    assert b'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"' in updated
    assert b'calcMode=' not in updated
    ET.fromstring(updated)


def test_set_calc_mode_adds_calc_pr_without_rewriting_other_namespaces(tmp_path):
    import zipfile
    from app.services.sch_writer import _set_calc_mode

    workbook_xml = (
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'mc:Ignorable="x15" '
        b'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
        b'<sheets/><extLst/></workbook>'
    )
    path = tmp_path / "namespace_test_manual.xlsx"
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)

    _set_calc_mode(path, "manual")

    with zipfile.ZipFile(path, "r") as archive:
        updated = archive.read("xl/workbook.xml")
    assert b'mc:Ignorable="x15"' in updated
    assert b'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"' in updated
    assert b'<calcPr calcMode="manual"/>' in updated
    assert updated.index(b'<calcPr') < updated.index(b'<extLst')



def test_set_calc_mode_rejects_unresolved_ignorable_prefix(tmp_path):
    import zipfile
    from app.services.sch_writer import _set_calc_mode

    workbook_xml = (
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'mc:Ignorable="x15"><sheets/><calcPr/></workbook>'
    )
    path = tmp_path / "invalid_namespace.xlsx"
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)

    with pytest.raises(SCHUpdateError, match="Namespace workbook tidak lengkap"):
        _set_calc_mode(path, "auto")

    # Failed validation must leave the original workbook package untouched.
    with zipfile.ZipFile(path, "r") as archive:
        assert archive.read("xl/workbook.xml") == workbook_xml
