from __future__ import annotations

import io
import json
import shutil
import sys
import zipfile
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

import tanpa_rumus_testkit as kit
from app.services import tanpa_rumus as tr
from app.services.tanpa_rumus import WorkbookInput


@pytest.fixture
def run(fake_excel):
    work_dirs: list[Path] = []

    def _run(files, **kwargs):
        inputs = [WorkbookInput(name, data) for name, data in files]
        zip_path, work_dir, manifest = tr.create_tanpa_rumus_zip(inputs, **kwargs)
        work_dirs.append(work_dir)
        return zip_path, work_dir, manifest

    yield _run
    for work_dir in work_dirs:
        shutil.rmtree(work_dir, ignore_errors=True)


def members(zip_path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(zip_path) as archive:
        assert archive.testzip() is None
        return {name: archive.read(name) for name in archive.namelist()}


def assert_no_formulas(data: bytes) -> None:
    workbook = load_workbook(io.BytesIO(data))  # formula tetap terbaca sebagai teks "=..."
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                assert not isinstance(cell.value, ArrayFormula), f"{sheet.title}!{cell.coordinate}"
                assert not (isinstance(cell.value, str) and cell.value.startswith("=")), f"{sheet.title}!{cell.coordinate}"


def cached(data: bytes, sheet: str, ref: str):
    return load_workbook(io.BytesIO(data), data_only=True)[sheet][ref].value


# ---------------------------------------------------------------- alur utama
def test_simple_workbook_formulas_become_values(run):
    zip_path, _, manifest = run([("Jadwal.xlsx", kit.simple_workbook())])

    assert manifest["status"] == "succeeded"
    report = manifest["files"][0]
    assert report["output"] == "Jadwal_TANPA_RUMUS.xlsx"
    assert report["sheet_count"] == 1
    assert report["formulas_converted"] == 3
    assert report["formulas_remaining"] == 0
    assert report["error"] is None
    assert report["processing_seconds"] >= 0

    files = members(zip_path)
    assert set(files) == {"Jadwal_TANPA_RUMUS.xlsx", "manifest.json", "laporan_tanpa_rumus.csv"}
    output = files["Jadwal_TANPA_RUMUS.xlsx"]
    assert_no_formulas(output)
    assert cached(output, "Data", "B1") == 6
    assert cached(output, "Data", "C1") == "001"  # teks "001" tidak menjadi angka 1
    assert json.loads(files["manifest.json"])["succeeded"] == 1


def test_source_bytes_are_not_modified_and_workdir_is_cleaned(run):
    source = kit.simple_workbook()
    original = bytes(source)
    zip_path, work_dir, _ = run([("a.xlsx", source)])
    assert source == original  # sumber tidak berubah
    leftovers = {path.name for path in work_dir.iterdir()}
    assert not any(name.startswith("input_") for name in leftovers)  # salinan input sudah dihapus
    assert leftovers == {zip_path.name, "output_01.xlsx"}  # hasil tetap ada sampai ZIP terkirim (router menghapus folder)


def test_full_calculation_and_save_format(run):
    run([("a.xlsx", kit.simple_workbook())])
    app = kit.FakeApp.registry[-1]
    assert app.api.calls == ["CalculateFull"]
    assert app.save_formats == [51]
    assert app.api.Calculation == -4105  # mode kalkulasi asli dipulihkan sebelum simpan
    assert app.api.AutomationSecurity == 3  # makro dinonaktifkan
    assert app.quit_called


def test_multiple_sheets_hidden_and_cross_sheet_formulas(run):
    zip_path, _, manifest = run([("Banyak.xlsx", kit.multi_sheet_workbook())])
    report = manifest["files"][0]
    assert report["sheet_count"] == 4
    assert report["formulas_converted"] == 4
    assert {s["name"]: s["state"] for s in report["details"]["sheets"]} == {
        "Ringkasan": "visible", "Data": "visible", "Rahasia": "hidden", "SangatRahasia": "veryHidden",
    }

    output = members(zip_path)["Banyak_TANPA_RUMUS.xlsx"]
    assert_no_formulas(output)
    book = load_workbook(io.BytesIO(output), data_only=True)
    assert [s.title for s in book.worksheets] == ["Ringkasan", "Data", "Rahasia", "SangatRahasia"]
    assert [s.sheet_state for s in book.worksheets] == ["visible", "visible", "hidden", "veryHidden"]
    assert (book["Ringkasan"]["A1"].value, book["Rahasia"]["A1"].value, book["SangatRahasia"]["A1"].value) == (60, 31, 32)


def test_sheet_visibility_is_restored_after_processing(run, fake_excel):
    fake_excel.skip_visibility_restore = True  # mensimulasikan pemulihan gagal -> validasi harus menolak
    _, _, manifest = run([("Banyak.xlsx", kit.multi_sheet_workbook())])
    assert manifest["files"][0]["status"] == "failed"
    assert "visible/hidden" in manifest["files"][0]["error"]


def test_complex_formatting_is_preserved(run):
    zip_path, _, _ = run([("Format.xlsx", kit.complex_format_workbook())])
    output = members(zip_path)["Format_TANPA_RUMUS.xlsx"]
    assert_no_formulas(output)
    sheet = load_workbook(io.BytesIO(output))["Laporan"]
    cell = sheet["B1"]
    assert cell.value == 3001.0
    assert cell.number_format == "#,##0.00"
    assert cell.font.color.rgb == "FFFF0000" and cell.font.bold
    assert cell.fill.fgColor.rgb == "FFFFFF00"
    assert cell.alignment.horizontal == "center"
    assert cell.border.left.style == "thin"
    assert [str(r) for r in sheet.merged_cells.ranges] == ["D1:E2"]
    assert sheet.column_dimensions["B"].width == 22
    assert sheet.row_dimensions[1].height == 30
    assert sheet.freeze_panes == "B2"
    assert len(sheet.conditional_formatting) == 1
    assert len(sheet.data_validations.dataValidation) == 1
    assert sheet.page_setup.orientation == "landscape"
    assert sheet.print_title_rows == "$1:$1"


def test_workbook_without_formulas_succeeds_with_zero_converted(run):
    _, _, manifest = run([("Statis.xlsx", kit.no_formula_workbook())])
    report = manifest["files"][0]
    assert report["status"] == "succeeded"
    assert report["formulas_converted"] == 0
    assert report["formulas_remaining"] == 0


def test_formula_errors_are_kept_as_values_with_warning(run):
    zip_path, _, manifest = run([("Err.xlsx", kit.error_formula_workbook())])
    report = manifest["files"][0]
    assert report["status"] == "succeeded"
    assert report["details"]["formula_errors"] == {"count": 1, "samples": ["Error!B1"]}
    assert any("menghasilkan error" in warning for warning in report["warnings"])
    output = members(zip_path)["Err_TANPA_RUMUS.xlsx"]
    assert_no_formulas(output)
    assert cached(output, "Error", "B1") == "#DIV/0!"
    assert cached(output, "Error", "B2") == 10


def test_array_formula_split_across_areas_is_converted_as_a_unit(run):
    zip_path, _, manifest = run([("Arr.xlsx", kit.array_formula_workbook())])
    report = manifest["files"][0]
    assert report["status"] == "succeeded", report["error"]
    assert report["formulas_converted"] == 7  # 6 sel array + 1 formula biasa
    assert_no_formulas(members(zip_path)["Arr_TANPA_RUMUS.xlsx"])


def test_hidden_or_filtered_rows_are_converted_row_by_row(run):
    zip_path, _, manifest = run([("Filter.xlsx", kit.hidden_rows_workbook())])
    report = manifest["files"][0]
    assert report["status"] == "succeeded", report["error"]
    output = members(zip_path)["Filter_TANPA_RUMUS.xlsx"]
    assert_no_formulas(output)
    assert [cached(output, "Filter", f"B{row}") for row in range(1, 6)] == [11, 21, 31, 41, 51]


def test_unprotected_sheet_is_reprotected(run, sheet_models):
    _, _, manifest = run([("Kunci.xlsx", kit.protected_workbook())])
    assert manifest["files"][0]["status"] == "succeeded"
    # file sumber dimuat sebagai model pertama; sheet dilindungi lagi dengan Contents=True
    sheet = next(m for m in sheet_models if m.name == "Terkunci" and m.protect_calls)
    assert sheet.protected and len(sheet.protect_calls) == 1
    assert sheet.protect_calls[0][2] is True


@pytest.fixture
def sheet_models(monkeypatch):
    """Daftar model sheet palsu yang dimuat, agar tes dapat memeriksa status akhir di dalam Excel palsu."""
    models: list = []
    original = kit.FakeWorkbookModel.__init__

    def tracking(self, app, path):
        original(self, app, path)
        models.extend(self.sheets)

    monkeypatch.setattr(kit.FakeWorkbookModel, "__init__", tracking)
    return models


def test_password_protected_sheet_fails_clearly(run):
    source = zipfile.ZipFile(io.BytesIO(kit.protected_workbook()))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            body = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                patched = body.replace(b"<sheetProtection ", b'<sheetProtection password="CC1A" ')
                assert patched != body, "patch password tidak terpasang"
                body = patched
            target.writestr(item.filename, body)
    _, _, manifest = run([("Kunci.xlsx", out.getvalue())])
    assert manifest["files"][0]["status"] == "failed"
    assert "diproteksi dengan kata sandi" in manifest["files"][0]["error"]


# ---------------------------------------------------------------- banyak file
@pytest.mark.parametrize("count", [1, 2, 5, 15, 20])
def test_any_number_of_files(run, count):
    zip_path, _, manifest = run(kit.many_workbooks(count))
    assert manifest["total_files"] == manifest["succeeded"] == count
    files = members(zip_path)
    outputs = [name for name in files if name.endswith("_TANPA_RUMUS.xlsx")]
    assert len(outputs) == count
    for name in outputs:
        assert_no_formulas(files[name])
    assert len(kit.FakeApp.registry) == count  # satu instance Excel per file
    assert kit.alive_excel_instances() == []


def test_duplicate_and_unsafe_names_are_handled(run):
    zip_path, _, manifest = run([
        ("Sama.xlsx", kit.simple_workbook()),
        ("sama.xlsx", kit.simple_workbook()),
        ("C:\\data\\..\\Sama_TANPA_RUMUS.xlsx", kit.simple_workbook()),
    ])
    outputs = [item["output"] for item in manifest["files"]]
    assert len(set(name.casefold() for name in outputs)) == 3
    assert outputs[0] == "Sama_TANPA_RUMUS.xlsx"
    assert outputs[1] == "sama_TANPA_RUMUS_2.xlsx"
    assert "/" not in "".join(outputs) and "\\" not in "".join(outputs)
    assert set(outputs) <= set(members(zip_path))


def test_one_failing_file_does_not_affect_others(run, fake_excel):
    fake_excel.open_fail_stems = {"input_02"}
    zip_path, _, manifest = run([
        ("ok1.xlsx", kit.simple_workbook()),
        ("rusak_excel.xlsx", kit.simple_workbook()),
        ("ok3.xlsx", kit.multi_sheet_workbook()),
    ])
    assert [item["status"] for item in manifest["files"]] == ["succeeded", "failed", "succeeded"]
    assert manifest["status"] == "partial" and manifest["failed"] == 1
    assert "tidak dapat membuka" in manifest["files"][1]["error"]
    names = set(members(zip_path))
    assert names == {"ok1_TANPA_RUMUS.xlsx", "ok3_TANPA_RUMUS.xlsx", "manifest.json", "laporan_tanpa_rumus.csv"}


def test_report_csv_contents(run):
    zip_path, _, _ = run([("=evil.xlsx", kit.simple_workbook()), ("bad.xlsx", b"bukan excel")])
    text = members(zip_path)["laporan_tanpa_rumus.csv"].decode("utf-8-sig")
    lines = text.strip().splitlines()
    assert lines[0].startswith("No,File Asli,File Hasil,Status,Jumlah Worksheet")
    assert "'=evil.xlsx" in lines[1] and "Berhasil" in lines[1]
    assert "Gagal" in lines[2]


# ---------------------------------------------------------------- validasi input
def _bad(run, name, content):
    _, _, manifest = run([(name, content)])
    report = manifest["files"][0]
    assert manifest["status"] == "failed" and report["status"] == "failed"
    assert report["output"] is None
    return report["error"]


def test_corrupt_and_invalid_inputs_are_rejected_with_clear_messages(run, fake_excel):
    good = kit.simple_workbook()
    assert "rusak atau bukan workbook" in _bad(run, "acak.xlsx", b"ini bukan excel sama sekali")
    assert "rusak" in _bad(run, "terpotong.xlsx", good[: len(good) // 2])
    assert "kata sandi" in _bad(run, "terenkripsi.xlsx", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 100)
    assert "hanya file .xlsx dan .xlsm" in _bad(run, "data.csv", b"a,b")
    assert "hanya file .xlsx dan .xlsm" in _bad(run, "data.xls", good)
    assert "file kosong" in _bad(run, "kosong.xlsx", b"")
    assert "sesuai isi file" in _bad(run, "salah.xlsm", good)  # xlsx berekstensi .xlsm
    assert "sesuai isi file" in _bad(run, "salah.xlsx", kit.make_xlsm(good))
    zip_without_workbook = io.BytesIO()
    with zipfile.ZipFile(zip_without_workbook, "w") as archive:
        archive.writestr("readme.txt", "hai")
    assert "daftar tipe konten" in _bad(run, "zip.xlsx", zip_without_workbook.getvalue())
    assert kit.FakeApp.registry == []  # file tak valid tidak pernah menjalankan Excel


def test_oversized_upload_error_is_reported_per_file(fake_excel):
    zip_path, work_dir, manifest = tr.create_tanpa_rumus_zip([
        WorkbookInput("besar.xlsx", error="besar.xlsx: ukuran file melebihi batas 50 MB."),
        WorkbookInput("ok.xlsx", kit.simple_workbook()),
    ])
    try:
        assert [i["status"] for i in manifest["files"]] == ["failed", "succeeded"]
        assert "50 MB" in manifest["files"][0]["error"]
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def test_file_count_limits(fake_excel):
    with pytest.raises(tr.TanpaRumusError, match="minimal satu"):
        tr.create_tanpa_rumus_zip([])
    with pytest.raises(tr.TanpaRumusError, match="Maksimal 20"):
        tr.create_tanpa_rumus_zip([WorkbookInput(f"{i}.xlsx", b"x") for i in range(21)])


# ---------------------------------------------------------------- makro
def test_xlsm_keeps_macros_and_format(run):
    zip_path, _, manifest = run([("Makro.xlsm", kit.make_xlsm(kit.simple_workbook()))])
    report = manifest["files"][0]
    assert report["status"] == "succeeded", report["error"]
    assert report["output"] == "Makro_TANPA_RUMUS.xlsm"
    assert report["details"]["macros_preserved"] is True
    assert report["details"]["calculation_mode"] == "smart"  # UDF aman: tidak CalculateFull
    assert any("Makro VBA" in warning for warning in report["warnings"])
    output = zipfile.ZipFile(io.BytesIO(members(zip_path)["Makro_TANPA_RUMUS.xlsm"]))
    assert "xl/vbaProject.bin" in output.namelist()
    assert kit.FakeApp.registry[-1].save_formats == [52]


def test_xlsm_that_loses_macros_is_failed_not_silently_converted(run, fake_excel):
    fake_excel.drop_vba_on_save = True
    zip_path, _, manifest = run([("Makro.xlsm", kit.make_xlsm(kit.simple_workbook()))])
    assert manifest["files"][0]["status"] == "failed"
    assert "makro VBA tidak terbawa" in manifest["files"][0]["error"]
    assert not any(name.endswith(".xlsm") for name in members(zip_path))


# ---------------------------------------------------------------- keselamatan
def test_excel_unavailable_aborts_with_specific_error(run, fake_excel):
    fake_excel.launch_error = True
    with pytest.raises(tr.ExcelUnavailable, match="Excel desktop tidak dapat dijalankan"):
        run([("a.xlsx", kit.simple_workbook())])


def test_xlwings_not_installed(monkeypatch):
    monkeypatch.setitem(sys.modules, "xlwings", None)  # import -> ImportError
    with pytest.raises(tr.ExcelUnavailable, match="xlwings belum terpasang"):
        tr.create_tanpa_rumus_zip([WorkbookInput("a.xlsx", kit.simple_workbook())])


def test_calculation_timeout_fails_file_and_closes_excel(run, fake_excel):
    fake_excel.calc_never_done = True
    _, _, manifest = run([("a.xlsx", kit.simple_workbook())], timeout_seconds=0.3)
    assert manifest["files"][0]["status"] == "failed"
    assert "batas waktu" in manifest["files"][0]["error"]


def test_hung_excel_is_killed_by_watchdog(run, fake_excel):
    fake_excel.hang_in_calculate = True
    _, _, manifest = run([("a.xlsx", kit.simple_workbook())], timeout_seconds=0.3)
    assert "batas waktu" in manifest["files"][0]["error"]
    assert kit.FakeApp.registry[-1].killed


def test_excel_that_does_not_exit_after_quit_is_killed(run, fake_excel):
    fake_excel.quit_does_not_exit = True
    _, _, manifest = run([("a.xlsx", kit.simple_workbook())])
    assert manifest["files"][0]["status"] == "succeeded"
    assert kit.FakeApp.registry[-1].killed


def test_wrong_pasted_values_are_detected_and_file_fails(run, fake_excel):
    fake_excel.stolen_clipboard = True
    _, _, manifest = run([("a.xlsx", kit.simple_workbook())])
    report = manifest["files"][0]
    assert report["status"] == "failed"
    assert "berbeda dari hasil hitung" in report["error"]


def test_corrupt_excel_output_is_rejected(run, fake_excel):
    fake_excel.corrupt_save = True
    _, _, manifest = run([("a.xlsx", kit.simple_workbook())])
    assert manifest["files"][0]["status"] == "failed"


def test_external_links_and_leftover_links_use_smart_calc_and_warn(run, fake_excel):
    kit.FakeApp.registry.clear()
    original_init = kit.FakeApp.__init__

    def init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        self.link_sources = ("C:\\data\\sumber.xlsx",)

    kit.FakeApp.__init__ = init
    try:
        _, _, manifest = run([("a.xlsx", kit.simple_workbook())])
    finally:
        kit.FakeApp.__init__ = original_init
    assert any("external link masih tercatat" in w for w in manifest["files"][0]["warnings"])


# ---------------------------------------------------------------- validator hasil independen
def test_inspector_counts_only_cell_formulas_not_x14_extension_formulas(tmp_path):
    source = zipfile.ZipFile(io.BytesIO(kit.simple_workbook()))
    path = tmp_path / "x.xlsx"
    with zipfile.ZipFile(path, "w") as target:
        for item in source.infolist():
            body = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = body.decode("utf-8").replace(
                    "</worksheet>",
                    '<extLst><ext xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
                    "<xm:f>Data!$A$1</xm:f></ext></extLst></worksheet>",
                )
                body = text.encode("utf-8")
            target.writestr(item.filename, body)
    inspection = tr.inspect_output(path, ".xlsx")
    assert inspection.formulas_remaining == 3  # tiga formula sel; <xm:f> tidak dihitung


def test_inspector_detects_leftover_formulas(tmp_path):
    path = tmp_path / "x.xlsx"
    path.write_bytes(kit.simple_workbook())
    info = tr.inspect_workbook(path, ".xlsx")
    output = tr.inspect_output(path, ".xlsx")
    problems = tr._output_problems(info, output)
    assert problems == ["masih ada 3 formula pada worksheet hasil"]


def test_zip_and_every_workbook_inside_reopen(run):
    zip_path, _, _ = run([("a.xlsx", kit.simple_workbook()), ("b.xlsm", kit.make_xlsm(kit.multi_sheet_workbook()))])
    files = members(zip_path)
    for name, data in files.items():
        if name.endswith(".xlsx"):
            load_workbook(io.BytesIO(data))
        if name.endswith(".xlsm"):
            assert zipfile.ZipFile(io.BytesIO(data)).testzip() is None
