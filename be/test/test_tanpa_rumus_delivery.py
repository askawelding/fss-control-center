from __future__ import annotations

import io
from pathlib import Path

import pytest
from openpyxl import Workbook

from app.services.tanpa_rumus_config import (
    delete_group,
    list_groups,
    save_group,
)
from app.services.tanpa_rumus import ConversionStats
from app.services.tanpa_rumus_delivery import (
    DeliveryWorkbookInput,
    _calculate_workbook_smart,
    _aggregate_delivery_keys,
    _delivery_formula_key,
    _parse_delivery_sumproduct_formula,
    _last_fss_candidate_row,
    _patch_calc_mode,
    build_output_tasks,
    inspect_delivery_source,
    inspect_date_sheet,
)


HEADERS = [
    "Date", "Vendor code", "Vendor name", "Receiving Area", "Delivery date",
    "delivery period", "Firm", "Classification", "PO number", "item",
    "parts number", "part name", "order Quantity",
]


def ckd_workbook_bytes() -> bytes:
    workbook = Workbook()
    date = workbook.active
    date.title = "Date"
    for index, header in enumerate(HEADERS, start=3):
        date.cell(2, index, header)
    date.cell(3, 3, 20261018)
    date.cell(3, 4, "D400153")
    date.cell(3, 6, "7R01")
    date.cell(3, 7, 20261026)
    date.cell(3, 13, "5311D244")
    date.cell(4, 3, 20261018)
    date.cell(4, 4, "D400153")
    date.cell(4, 6, "7R12")
    date.cell(4, 7, 20261026)
    date.cell(4, 13, "5311D245")
    date.cell(5, 3, 20261018)
    date.cell(5, 4, "D400153")
    date.cell(5, 6, "1R01")
    date.cell(5, 7, 20261026)
    date.cell(5, 13, "5311D246")
    date.cell(6, 3, 20261018)
    date.cell(6, 4, "D400153")
    date.cell(6, 7, 20261026)
    date.cell(6, 13, "5311D247")  # data tanpa receiving
    for month in ("202610", "202611", "202612", "202701"):
        workbook.create_sheet(month)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def test_default_receiving_groups_are_persisted(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DATA_DIR", str(tmp_path))
    groups = list_groups()
    assert len(groups) == 5
    assert groups[0]["codes"] == ["7R01"]
    assert groups[1]["codes"] == ["7R02", "7R03"]
    assert groups[3]["codes"] == ["7R07", "7R08", "7R10"]
    assert groups[4]["codes"] == ["7R09", "7R11"]
    assert (tmp_path / "fss_control_center.db").exists()


def test_group_code_cannot_belong_to_two_active_groups(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DATA_DIR", str(tmp_path))
    list_groups()
    with pytest.raises(ValueError, match="sudah digunakan"):
        save_group({"name": "Test", "codes": ["7R01", "7R13"], "outputLabel": "7R13 TEST"})


def test_new_group_can_contain_multiple_codes(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DATA_DIR", str(tmp_path))
    list_groups()
    saved = save_group({"name": "Singapore", "codes": ["7R12", "7R13"], "outputLabel": "7R12 & 7R13 SINGAPORE"})
    assert saved["codes"] == ["7R12", "7R13"]
    assert any(group["id"] == saved["id"] for group in list_groups())


def test_deleted_group_is_hidden_but_defaults_can_be_read_from_history_snapshot(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DATA_DIR", str(tmp_path))
    groups = list_groups()
    assert delete_group(groups[0]["id"]) is True
    assert all(group["id"] != groups[0]["id"] for group in list_groups())
    assert delete_group(groups[0]["id"]) is False


def test_inspect_date_sheet_finds_country_receiving_codes_and_month_tabs():
    content = ckd_workbook_bytes()
    preview = inspect_date_sheet(content, "20261018 CKD DELIVERY SCH.xlsx")
    assert preview.sheet_name == "Date"
    assert preview.header_row == 2
    assert preview.data_start_row == 3
    assert preview.last_data_row == 6
    assert preview.receiving_counts == {"7R01": 1, "7R12": 1, "1R01": 1}
    assert preview.blank_receiving_rows == 1
    assert preview.month_sheets == ["202610", "202611", "202612", "202701"]


def test_inspect_delivery_source_accepts_a_single_mmki_workbook():
    result = inspect_delivery_source(ckd_workbook_bytes(), "20261018 MMKI DELIVERY SCH.xlsx")
    assert result["filename"] == "20261018 MMKI DELIVERY SCH.xlsx"
    assert result["date_sheet"] == "Date"
    assert result["month_sheets"] == ["202610", "202611", "202612", "202701"]


def test_inspect_delivery_source_prefers_date_when_data_sheet_also_exists():
    from openpyxl import load_workbook

    workbook = load_workbook(io.BytesIO(ckd_workbook_bytes()))
    workbook.create_sheet("DATA")
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()

    result = inspect_delivery_source(buffer.getvalue(), "20261018 MMKI DELIVERY SCH.xlsx")

    assert result["date_sheet"] == "Date"


def test_inspect_delivery_source_rejects_workbook_without_month_tabs():
    from openpyxl import Workbook
    from app.services.tanpa_rumus import TanpaRumusInputError

    workbook = Workbook()
    workbook.active.title = "Date"
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()
    with pytest.raises(TanpaRumusInputError, match="worksheet bulanan"):
        inspect_delivery_source(buffer.getvalue(), "mmki.xlsx")


def test_last_fss_candidate_row_ignores_helper_columns_outside_fss_area():
    import zipfile
    import xml.etree.ElementTree as ET

    content = ckd_workbook_bytes()
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        candidate = _last_fss_candidate_row(archive, "xl/worksheets/sheet1.xml")
    assert candidate == 6


def test_inspect_date_sheet_stops_after_long_empty_fss_tail(monkeypatch):
    from openpyxl import load_workbook
    import app.services.tanpa_rumus_delivery as delivery

    monkeypatch.setattr(delivery, "EMPTY_FSS_TAIL_ROWS_TO_STOP", 2)
    workbook = load_workbook(io.BytesIO(ckd_workbook_bytes()))
    date = workbook["Date"]
    for row in range(7, 500):
        date.cell(row, 16, f"=ROW()")  # helper tail outside FSS columns C:O
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()

    preview = inspect_date_sheet(buffer.getvalue(), "20261018 CKD DELIVERY SCH.xlsx")

    assert preview.last_data_row == 6
    assert preview.receiving_counts == {"7R01": 1, "7R12": 1, "1R01": 1}
    assert preview.blank_receiving_rows == 1


def test_lowercase_date_helper_header_is_not_treated_as_duplicate_fss_date():
    from openpyxl import load_workbook

    workbook = load_workbook(io.BytesIO(ckd_workbook_bytes()))
    date = workbook["Date"]
    date.cell(2, 18, "date")
    date.cell(3, 18, "=1")
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()

    preview = inspect_date_sheet(buffer.getvalue(), "20261018 CKD DELIVERY SCH.xlsx")

    assert preview.header_row == 2
    assert preview.fss_columns["Date"] == 3
    assert 18 not in preview.fss_columns.values()


def test_two_capitalized_date_headers_are_still_rejected():
    from openpyxl import load_workbook
    from app.services.tanpa_rumus import TanpaRumusInputError

    workbook = load_workbook(io.BytesIO(ckd_workbook_bytes()))
    workbook["Date"].cell(2, 18, "Date")
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()

    with pytest.raises(TanpaRumusInputError, match="Header 'Date'.*muncul lebih dari sekali"):
        inspect_date_sheet(buffer.getvalue(), "20261018 CKD DELIVERY SCH.xlsx")


@pytest.mark.parametrize(
    "workbook_xml",
    [
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/><calcPr calcId="191029"/></workbook>',
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/><calcPr calcId="191029" calcMode="manual" /></workbook>',
        b'<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets/><x:calcPr calcId="1"></x:calcPr></x:workbook>',
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/><extLst/></workbook>',
    ],
)
def test_patch_calc_mode_keeps_workbook_xml_well_formed(tmp_path, workbook_xml):
    import zipfile
    import xml.etree.ElementTree as ET

    path = tmp_path / "workbook.xlsx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("[Content_Types].xml", b"<Types/>")

    _patch_calc_mode(path, "auto")

    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
        calc = [node for node in root if node.tag.rsplit("}", 1)[-1].split(":")[-1] == "calcPr"]
        assert len(calc) == 1
        assert calc[0].attrib["calcMode"] == "auto"
        assert archive.testzip() is None


def test_output_tasks_allow_mmki_source_only():
    source = DeliveryWorkbookInput("20261018 MMKI DELIVERY SCH.xlsx", b"source")
    tasks = build_output_tasks(source, None, True, False, False, [])
    assert len(tasks) == 1
    assert tasks[0]["kind"] == "mmki_all"
    assert tasks[0]["source_name"] == source.filename


def test_output_tasks_allow_ckd_source_only():
    source = DeliveryWorkbookInput("20261018 CKD DELIVERY SCH.xlsx", b"source")
    tasks = build_output_tasks(None, source, False, True, False, [])
    assert len(tasks) == 1
    assert tasks[0]["kind"] == "ckd_all"
    assert tasks[0]["source_name"] == source.filename


def test_fast_formula_key_matches_delivery_helper_concatenation():
    assert _delivery_formula_key("20261026", "1.0", "5311D244", "7R01", {"7R01", "7R02"}) == "46321" + "1" + "5311D244-KD"
    assert _delivery_formula_key("20261026", "2", "5311D244", "1R02", {"7R01", "7R02"}) == "4632125311D244"
    # New 7R codes do not receive -KD unless the source helper formula explicitly maps them.
    assert _delivery_formula_key("2026-10-26T00:00:00", "1", "5311D244", "7R12", {"7R01", "7R02"}) == "4632115311D244"


def test_fast_formula_parser_recognizes_monthly_sumproduct_template():
    formula = (
        "=SUMPRODUCT((DATE!$B$3:$B$90000='202610(CKD)'!D$8&"
        "'202610(CKD)'!$B10&'202610(CKD)'!$A10)*1,DATE!$O$3:$O$90000)"
    )
    parsed = _parse_delivery_sumproduct_formula(formula, "DATE")
    assert parsed == {
        "data_start_row": 3,
        "data_end_row": 90000,
        "header_column": 4,
        "header_row": 8,
        "cycle_row": 10,
        "parts_row": 10,
    }


def test_inspect_date_sheet_builds_receiving_key_aggregates_for_fast_path():
    from openpyxl import load_workbook

    workbook = load_workbook(io.BytesIO(ckd_workbook_bytes()))
    date = workbook["Date"]
    date["B3"] = '=IFERROR(S3&H3&M3&T3,"")'
    date["T3"] = '=IF(F3="7R01","-KD",IF(F3="7R02","-KD",""))'
    date["O3"] = 20
    date["O4"] = 10
    date["G3"] = 20261026
    date["G4"] = 20261026
    date["H3"] = 1
    date["H4"] = 1
    date["M3"] = "5311D244"
    date["M4"] = "5311D244"
    date["F3"] = "7R01"
    date["F4"] = "7R02"
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()

    preview = inspect_date_sheet(buffer.getvalue(), "20261018 CKD DELIVERY SCH.xlsx")
    assert preview.fast_formula_supported is True
    assert preview.division_suffix_codes == {"7R01", "7R02"}
    key = "4632115311D244-KD"
    assert preview.formula_key_totals_all[key] == 30
    assert _aggregate_delivery_keys(
        preview, {"codes": ["7R01"], "id": "vietnam"}
    )[key] == 20
    assert _aggregate_delivery_keys(
        preview, {"codes": ["7R02"], "id": "malaysia"}
    )[key] == 10


def test_preview_flags_only_unmapped_7r_country_codes():
    preview = inspect_date_sheet(ckd_workbook_bytes(), "ckd.xlsx")
    groups = [{"id": "vietnam", "name": "Vietnam", "codes": ["7R01"], "outputLabel": "7R01 VIETNAM"}]
    result = preview.to_dict(groups)
    assert result["unmapped_codes"] == ["7R12"]
    assert result["blank_receiving_rows"] == 1
    assert any(item["code"] == "1R01" and not item["is_country_code"] for item in result["receiving_areas"])


def test_output_task_creates_one_file_for_each_group_from_same_source():
    source = DeliveryWorkbookInput("20261018 CKD DELIVERY SCH.xlsx", ckd_workbook_bytes())
    groups = [
        {"id": "vietnam", "name": "Vietnam", "codes": ["7R01"], "outputLabel": "7R01 VIETNAM"},
        {"id": "malaysia", "name": "Malaysia & L300", "codes": ["7R02", "7R03"], "outputLabel": "7R02 MALAYSIA & 7R03 L300"},
    ]
    tasks = build_output_tasks(None, source, False, False, True, groups)
    assert len(tasks) == 2
    assert tasks[0]["output_name"] == "20261018 CKD DELIVERY SCH (7R01 VIETNAM).xlsx"
    assert tasks[1]["output_name"] == "20261018 CKD DELIVERY SCH (7R02 MALAYSIA & 7R03 L300).xlsx"
    assert tasks[0]["source_content"] is source.content
    assert tasks[1]["source_content"] is source.content


def test_output_task_can_make_mmki_and_ckd_all_together():
    mmki = DeliveryWorkbookInput("20261018 MMKI DELIVERY SCH.xlsx", ckd_workbook_bytes())
    ckd = DeliveryWorkbookInput("20261018 CKD DELIVERY SCH.xlsx", ckd_workbook_bytes())
    tasks = build_output_tasks(mmki, ckd, True, True, False, [])
    assert [task["output_name"] for task in tasks] == [
        "20261018 MMKI DELIVERY SCH (TANPA RUMUS).xlsx",
        "20261018 CKD DELIVERY SCH (TANPA RUMUS).xlsx",
    ]


def test_delivery_uses_smart_calculation_instead_of_forced_full_recalculation(monkeypatch):
    class FakeExcelApi:
        CalculationState = 0

        def __init__(self):
            self.calculate_calls = 0
            self.calculate_full_calls = 0

        def Calculate(self):
            self.calculate_calls += 1

        def CalculateFull(self):
            self.calculate_full_calls += 1

    monkeypatch.setattr(
        "app.services.tanpa_rumus_delivery._wait_for_calculation",
        lambda api, deadline, seconds: None,
    )
    api = FakeExcelApi()
    stats = ConversionStats()

    _calculate_workbook_smart(api, stats, timeout_seconds=10)

    assert api.calculate_calls == 1
    assert api.calculate_full_calls == 0
    assert stats.calculation_mode == "calculate_smart"
    assert stats.calculation_seconds >= 0


def test_delivery_reuses_large_workbook_inspection_for_multiple_groups(monkeypatch):
    import shutil
    from types import SimpleNamespace

    import app.services.tanpa_rumus_delivery as delivery

    class FakePreview:
        blank_receiving_rows = 0

        def to_dict(self, groups):
            return {"unmapped_codes": []}

    preview = FakePreview()
    inspect_calls = {"date": 0, "workbook": 0}
    output_calls = []

    def fake_inspect_date_sheet(content, filename):
        inspect_calls["date"] += 1
        return preview

    def fake_inspect_workbook(path, extension):
        inspect_calls["workbook"] += 1
        return SimpleNamespace(sheets=[])

    def fake_run_one_output(
        source_path, output_path, source_name, mode, group, timeout_seconds,
        workbook_info=None, date_preview=None,
    ):
        output_calls.append((source_path, workbook_info, date_preview))
        assert date_preview is preview
        output_path.write_bytes(b"test-output")
        return {"status": "succeeded", "output": output_path.name}

    monkeypatch.setattr(delivery, "inspect_date_sheet", fake_inspect_date_sheet)
    monkeypatch.setattr(delivery, "inspect_workbook", fake_inspect_workbook)
    monkeypatch.setattr(delivery, "_run_one_output", fake_run_one_output)

    ckd = DeliveryWorkbookInput("20261018 CKD DELIVERY SCH.xlsx", b"source")
    groups = [
        {"id": "g1", "name": "Vietnam", "codes": ["7R01"], "outputLabel": "7R01 VIETNAM"},
        {"id": "g2", "name": "Malaysia", "codes": ["7R02"], "outputLabel": "7R02 MALAYSIA"},
    ]

    zip_path, work_dir, manifest = delivery.create_delivery_zip(
        None, ckd, False, False, True, groups, groups,
    )
    try:
        assert zip_path.exists()
        assert manifest["succeeded"] == 2
        assert inspect_calls == {"date": 1, "workbook": 1}
        assert len(output_calls) == 2
        assert output_calls[0][0] == output_calls[1][0]
        assert output_calls[0][1] is output_calls[1][1]
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
