"""Tes integrasi Tanpa Rumus dengan Microsoft Excel SUNGGUHAN.

Hanya berjalan di Windows yang memiliki Excel desktop + xlwings; di tempat lain otomatis dilewati.
Workbook uji dibuat oleh Excel sendiri sehingga memiliki formula dan nilai cache yang asli.

Jalankan:  python -m pytest test/test_tanpa_rumus_excel.py -v
"""
from __future__ import annotations

import io
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Membutuhkan Windows + Microsoft Excel")

from app.services import tanpa_rumus as tr  # noqa: E402
from app.services.tanpa_rumus import WorkbookInput  # noqa: E402


@pytest.fixture(scope="module")
def xw():
    xlwings = pytest.importorskip("xlwings")
    try:
        app = xlwings.App(visible=False, add_book=False)
        app.quit()
    except Exception as exc:  # noqa: BLE001
        pytest.skip(f"Microsoft Excel tidak dapat dijalankan: {exc}")
    return xlwings


def _excel_process_count() -> int:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"], capture_output=True, text=True, check=False
    ).stdout
    return sum(1 for line in out.splitlines() if line.strip().upper().startswith("EXCEL.EXE"))


@pytest.fixture
def source_workbook(xw, tmp_path) -> Path:
    """Workbook buatan Excel: format kompleks, lintas sheet, array, sheet hidden/very hidden."""
    path = tmp_path / "sumber.xlsx"
    app = xw.App(visible=False, add_book=False)
    try:
        app.api.DisplayAlerts = False
        book = app.books.add()
        data = book.sheets[0]
        data.name = "Data"
        data.range("A1:A3").value = [[1], [2], [3]]
        data.range("B1").formula = "=SUM(A1:A3)"
        data.range("C1").formula = '=TEXT(A1,"000")'
        data.range("D1").formula = '=IF(A1>5,"x","")'
        data.range("B1").number_format = "#,##0.00"
        data.range("B1").api.Font.Color = 255  # merah (BGR)
        data.range("B1").api.Font.Bold = True
        data.range("F1:G2").api.Merge()
        data.range("B1").column_width = 24
        data.range("H1:H3").api.FormulaArray = "=A1:A3*2"

        ringkasan = book.sheets.add("Ringkasan", after=data)
        ringkasan.range("A1").formula = "=Data!B1*2"

        rahasia = book.sheets.add("Rahasia", after=ringkasan)
        rahasia.range("A1").formula = "=Data!B1+1"
        rahasia.api.Visible = 0  # xlSheetHidden

        sangat = book.sheets.add("SangatRahasia", after=rahasia)
        sangat.range("A1").formula = "=Rahasia!A1+1"
        sangat.api.Visible = 2  # xlSheetVeryHidden

        data.activate()
        book.api.SaveAs(str(path), 51)
        book.close()
    finally:
        app.quit()
    return path


def _run(files):
    inputs = [WorkbookInput(name, content) for name, content in files]
    zip_path, work_dir, manifest = tr.create_tanpa_rumus_zip(inputs, timeout_seconds=300)
    return zip_path, work_dir, manifest


def test_real_excel_removes_all_formulas_and_keeps_structure(source_workbook):
    from openpyxl import load_workbook

    excel_before = _excel_process_count()
    zip_path, work_dir, manifest = _run([("Sumber.xlsx", source_workbook.read_bytes())])
    try:
        report = manifest["files"][0]
        assert report["status"] == "succeeded", report["error"]
        assert report["sheet_count"] == 4
        assert report["formulas_converted"] >= 9  # B1,C1,D1,H1:H3,Ringkasan!A1,Rahasia!A1,SangatRahasia!A1
        assert report["formulas_remaining"] == 0

        with zipfile.ZipFile(zip_path) as archive:
            assert archive.testzip() is None
            output = archive.read("Sumber_TANPA_RUMUS.xlsx")

        formulas = load_workbook(io.BytesIO(output))
        values = load_workbook(io.BytesIO(output), data_only=True)
        for sheet in formulas.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    assert not (isinstance(cell.value, str) and cell.value.startswith("=")), f"{sheet.title}!{cell.coordinate}"
        assert [s.title for s in values.worksheets] == ["Data", "Ringkasan", "Rahasia", "SangatRahasia"]
        assert [s.sheet_state for s in values.worksheets] == ["visible", "visible", "hidden", "veryHidden"]
        assert values["Data"]["B1"].value == 6
        assert values["Data"]["C1"].value == "001"  # tetap teks
        assert values["Ringkasan"]["A1"].value == 12
        assert values["Rahasia"]["A1"].value == 7
        assert values["SangatRahasia"]["A1"].value == 8
        assert [values["Data"][f"H{row}"].value for row in (1, 2, 3)] == [2, 4, 6]

        data = formulas["Data"]
        assert data["B1"].number_format == "#,##0.00"
        assert data["B1"].font.bold
        assert [str(r) for r in data.merged_cells.ranges] == ["F1:G2"]
        assert data.column_dimensions["B"].width >= 23
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    assert _excel_process_count() == excel_before, "proses EXCEL.EXE tertinggal"


def test_real_excel_leaves_source_file_untouched(source_workbook):
    before = source_workbook.read_bytes()
    _, work_dir, manifest = _run([("Sumber.xlsx", before)])
    shutil.rmtree(work_dir, ignore_errors=True)
    assert manifest["status"] == "succeeded"
    assert source_workbook.read_bytes() == before


def test_real_excel_formula_error_is_reported_not_hidden(xw, tmp_path):
    from openpyxl import load_workbook

    path = tmp_path / "err.xlsx"
    app = xw.App(visible=False, add_book=False)
    try:
        book = app.books.add()
        sheet = book.sheets[0]
        sheet.name = "E"
        sheet.range("A1").value = 5
        sheet.range("A2").value = 0
        sheet.range("B1").formula = "=A1/A2"
        book.api.SaveAs(str(path), 51)
        book.close()
    finally:
        app.quit()

    zip_path, work_dir, manifest = _run([("err.xlsx", path.read_bytes())])
    try:
        report = manifest["files"][0]
        assert report["status"] == "succeeded"
        assert report["details"]["formula_errors"]["count"] == 1
        with zipfile.ZipFile(zip_path) as archive:
            output = archive.read("err_TANPA_RUMUS.xlsx")
        assert load_workbook(io.BytesIO(output), data_only=True)["E"]["B1"].value == "#DIV/0!"
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def test_real_excel_multiple_files_with_one_corrupt(source_workbook):
    zip_path, work_dir, manifest = _run([
        ("a.xlsx", source_workbook.read_bytes()),
        ("rusak.xlsx", b"bukan workbook"),
        ("b.xlsx", source_workbook.read_bytes()),
    ])
    try:
        assert [item["status"] for item in manifest["files"]] == ["succeeded", "failed", "succeeded"]
        with zipfile.ZipFile(zip_path) as archive:
            assert set(archive.namelist()) == {
                "a_TANPA_RUMUS.xlsx", "b_TANPA_RUMUS.xlsx", "manifest.json", "laporan_tanpa_rumus.csv",
            }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
