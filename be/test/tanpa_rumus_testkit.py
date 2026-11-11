"""Perangkat bantu tes Tanpa Rumus.

1. Pembangun workbook fixture. openpyxl tidak menulis nilai hasil hitung (cached value) untuk formula,
   jadi `with_cached_values` menyuntikkannya seperti yang dilakukan Excel saat menyimpan.
2. Model Excel/xlwings palsu (`FakeXlwings`) yang meniru permukaan COM yang dipakai service:
   Application, Workbook, Worksheet, Range (SpecialCells, Areas, Copy, PasteSpecial, Value2, ...).
   Model dimuat dari file .xlsx sungguhan, dan SaveAs menulis file .xlsx sungguhan (formula yang sudah
   dikonversi dibuang dari XML) sehingga validasi hasil di service berjalan terhadap file nyata.

Catatan: model ini menguji LOGIKA service. Perilaku Excel asli diuji oleh test_tanpa_rumus_excel.py
(hanya berjalan pada Windows yang memiliki Microsoft Excel).
"""
from __future__ import annotations

import io
import re
import threading
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ERROR_CODES = {
    "#DIV/0!": -2146826281,
    "#N/A": -2146826246,
    "#REF!": -2146826265,
    "#VALUE!": -2146826273,
    "#NAME?": -2146826259,
}
STATE_TO_VISIBLE = {"visible": -1, "hidden": 0, "veryHidden": 2}
VISIBLE_TO_STATE = {value: key for key, value in STATE_TO_VISIBLE.items()}


# ----------------------------------------------------------------------------
# Pembangun workbook
# ----------------------------------------------------------------------------
def save_workbook(workbook: Workbook) -> bytes:
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def sheet_part_map(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    result = {}
    for sheet in workbook.iter(f"{{{NS}}}sheet"):
        target = targets[sheet.get(f"{{{RNS}}}id")]
        result[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
    return result


_F_PATTERN = re.compile(r"<f\b[^>]*?(?:/>|>.*?</f>)", re.S)


def _inject_cached(xml: str, values: dict[str, Any]) -> str:
    for ref, value in values.items():
        pattern = re.compile(rf'(<c\b[^>]*?\br="{ref}"[^>]*?)>(.*?)</c>', re.S)
        match = pattern.search(xml)
        assert match, f"sel {ref} tidak ditemukan"
        opening, inner = match.group(1), match.group(2)
        formula = _F_PATTERN.search(inner)
        assert formula, f"sel {ref} bukan formula"
        opening = re.sub(r'\s+t="[^"]*"', "", opening)
        if isinstance(value, tuple):  # (teks, "e") untuk error
            opening += ' t="e"'
            cached = value[0]
        elif isinstance(value, bool):
            opening += ' t="b"'
            cached = "1" if value else "0"
        elif isinstance(value, str):
            opening += ' t="str"'
            cached = value
        else:
            cached = repr(value)
        xml = xml[:match.start()] + f"{opening}>{formula.group(0)}<v>{cached}</v></c>" + xml[match.end():]
    return xml


def with_cached_values(data: bytes, cached: dict[str, dict[str, Any]]) -> bytes:
    """cached = {"NamaSheet": {"B1": 6, "C1": "007", "D1": ("#DIV/0!", "e")}}"""
    source = zipfile.ZipFile(io.BytesIO(data))
    parts = sheet_part_map(source)
    by_part = {parts[name]: values for name, values in cached.items()}
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            body = source.read(item.filename)
            if item.filename in by_part:
                body = _inject_cached(body.decode("utf-8"), by_part[item.filename]).encode("utf-8")
            target.writestr(item.filename, body)
    return out.getvalue()


def make_xlsm(data: bytes, vba: bytes = b"FAKE-VBA-PROJECT") -> bytes:
    """Ubah workbook .xlsx menjadi paket bergaya .xlsm (tipe konten makro + vbaProject.bin)."""
    source = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            body = source.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = body.decode("utf-8")
                text = text.replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                ).replace(
                    "<Default Extension=\"xml\"",
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"',
                )
                body = text.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                text = body.decode("utf-8").replace(
                    "</Relationships>",
                    '<Relationship Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                    'Target="vbaProject.bin" Id="rIdVba"/></Relationships>',
                )
                body = text.encode("utf-8")
            target.writestr(item.filename, body)
        target.writestr("xl/vbaProject.bin", vba)
    return out.getvalue()


def simple_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["A2"], ws["A3"] = 1, 2, 3
    ws["B1"] = "=SUM(A1:A3)"
    ws["C1"] = '=TEXT(A1,"000")'  # hasil teks "001" harus tetap teks
    ws["D1"] = '=IF(A1>5,"x","")'  # hasil string kosong
    return with_cached_values(save_workbook(wb), {"Data": {"B1": 6, "C1": "001", "D1": ""}})


def multi_sheet_workbook() -> bytes:
    wb = Workbook()
    ringkasan = wb.active
    ringkasan.title = "Ringkasan"
    data = wb.create_sheet("Data")
    rahasia = wb.create_sheet("Rahasia")
    sangat = wb.create_sheet("SangatRahasia")
    data["A1"], data["A2"] = 10, 20
    data["B1"] = "=A1+A2"
    ringkasan["A1"] = "=Data!B1*2"  # lintas worksheet
    rahasia["A1"] = "=Data!B1+1"
    rahasia.sheet_state = "hidden"
    sangat["A1"] = "=Rahasia!A1+1"
    sangat.sheet_state = "veryHidden"
    return with_cached_values(
        save_workbook(wb),
        {"Data": {"B1": 30}, "Ringkasan": {"A1": 60}, "Rahasia": {"A1": 31}, "SangatRahasia": {"A1": 32}},
    )


def complex_format_workbook() -> bytes:
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan"
    ws["A1"], ws["A2"] = 1500.5, 2
    ws["B1"] = "=A1*A2"
    ws["B1"].number_format = "#,##0.00"
    ws["B1"].font = Font(color="FFFF0000", bold=True)
    ws["B1"].fill = PatternFill("solid", fgColor="FFFFFF00")
    ws["B1"].border = Border(left=Side(style="thin"), right=Side(style="thin"))
    ws["B1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("D1:E2")
    ws.column_dimensions["B"].width = 22
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "B2"
    ws.conditional_formatting.add("B1", CellIsRule(operator="greaterThan", formula=["100"], fill=PatternFill("solid", bgColor="FFC7CE")))
    validation = DataValidation(type="list", formula1='"Ya,Tidak"', allow_blank=True)
    ws.add_data_validation(validation)
    validation.add("F1")
    ws.page_setup.orientation = "landscape"
    ws.print_title_rows = "1:1"
    return with_cached_values(save_workbook(wb), {"Laporan": {"B1": 3001.0}})


def no_formula_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Statis"
    ws["A1"], ws["A2"] = "teks", 42
    return save_workbook(wb)


def error_formula_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Error"
    ws["A1"], ws["A2"] = 5, 0
    ws["B1"] = "=A1/A2"
    ws["B2"] = "=A1*2"
    return with_cached_values(save_workbook(wb), {"Error": {"B1": ("#DIV/0!", "e"), "B2": 10}})


def array_formula_workbook() -> bytes:
    from openpyxl.worksheet.formula import ArrayFormula

    wb = Workbook()
    ws = wb.active
    ws.title = "Array"
    for row in range(1, 4):
        ws.cell(row, 5, row)
    ws["A1"] = ArrayFormula("A1:B3", "=E1:E3*2")  # array 2 kolom x 3 baris
    ws["A4"] = "=SUM(E1:E3)"  # formula tetangga: area SpecialCells dapat memotong array
    return with_cached_values(save_workbook(wb), {"Array": {"A1": 2, "A4": 6}})


def hidden_rows_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Filter"
    for row in range(1, 6):
        ws.cell(row, 1, row * 10)
        ws.cell(row, 2, f"=A{row}+1")
    ws.row_dimensions[2].hidden = True
    ws.row_dimensions[3].hidden = True
    return with_cached_values(save_workbook(wb), {"Filter": {f"B{row}": row * 10 + 1 for row in range(1, 6)}})


def protected_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Terkunci"
    ws["A1"] = 4
    ws["B1"] = "=A1*2"
    ws.protection.sheet = True
    return with_cached_values(save_workbook(wb), {"Terkunci": {"B1": 8}})


def many_workbooks(count: int) -> list[tuple[str, bytes]]:
    return [(f"schedule_{index:02d}.xlsx", simple_workbook()) for index in range(1, count + 1)]


# ----------------------------------------------------------------------------
# Excel / xlwings palsu
# ----------------------------------------------------------------------------
class ComError(Exception):
    """Pengganti pywintypes.com_error."""


@dataclass
class Behavior:
    launch_error: bool = False
    open_fail_stems: set[str] = field(default_factory=set)  # mis. {"input_02"}
    calc_never_done: bool = False  # CalculationState tidak pernah xlDone
    hang_in_calculate: bool = False  # Calculate() menggantung sampai proses dimatikan
    drop_vba_on_save: bool = False
    quit_does_not_exit: bool = False
    stolen_clipboard: bool = False  # Paste menempelkan isi yang salah
    skip_visibility_restore: bool = False
    corrupt_save: bool = False  # SaveAs menulis berkas rusak


BEHAVIOR = Behavior()


@dataclass
class FakeCell:
    formula: bool = False
    value: Any = None
    is_error: bool = False
    array: tuple[int, int, int, int] | None = None


class FakeSheetModel:
    def __init__(self, workbook: "FakeWorkbookModel", name: str, part: str, state: str) -> None:
        self.workbook = workbook
        self.name = name
        self.part = part
        self.visible = STATE_TO_VISIBLE.get(state, -1)
        self.cells: dict[tuple[int, int], FakeCell] = {}
        self.hidden_rows: set[int] = set()
        self.protected = False
        self.protected_with_password = False
        self.protect_calls: list[tuple] = []

    def cell(self, row: int, col: int) -> FakeCell:
        return self.cells.setdefault((row, col), FakeCell())


class FakeWorkbookModel:
    def __init__(self, app: "FakeApp", path: Path) -> None:
        self.app = app
        self.path = path
        self.source = path.read_bytes()
        archive = zipfile.ZipFile(io.BytesIO(self.source))
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            for si in ET.fromstring(archive.read("xl/sharedStrings.xml")):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{NS}}}t")))
        workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
        parts = sheet_part_map(archive)
        self.sheets: list[FakeSheetModel] = []
        for sheet in workbook_xml.iter(f"{{{NS}}}sheet"):
            model = FakeSheetModel(self, sheet.get("name"), parts[sheet.get("name")], sheet.get("state", "visible"))
            self._load_sheet(archive.read(model.part), model, shared)
            self.sheets.append(model)

    @staticmethod
    def _load_sheet(xml: bytes, model: FakeSheetModel, shared: list[str]) -> None:
        root = ET.fromstring(xml)
        protection = root.find(f"{{{NS}}}sheetProtection")
        if protection is not None and protection.get("sheet") in ("1", "true"):
            model.protected = True
            model.protected_with_password = bool(protection.get("password") or protection.get("hashValue"))
        for row in root.iter(f"{{{NS}}}row"):
            if row.get("hidden") in ("1", "true"):
                model.hidden_rows.add(int(row.get("r")))
            for c in row.findall(f"{{{NS}}}c"):
                letters, number = coordinate_from_string(c.get("r"))
                key = (number, column_index_from_string(letters))
                cell = model.cell(*key)
                f = c.find(f"{{{NS}}}f")
                v = c.find(f"{{{NS}}}v")
                kind = c.get("t")
                text = v.text if v is not None else None
                if kind == "s":
                    cell.value = shared[int(text)]
                elif kind in ("str", "inlineStr"):
                    cell.value = text or ("".join(t.text or "" for t in c.iter(f"{{{NS}}}t")) if kind == "inlineStr" else "")
                elif kind == "b":
                    cell.value = text == "1"
                elif kind == "e":
                    cell.value, cell.is_error = text, True
                elif text not in (None, ""):
                    number_value = float(text)
                    cell.value = int(number_value) if number_value.is_integer() else number_value
                if f is not None:
                    cell.formula = True
                    if f.get("t") == "array" and f.get("ref"):
                        first, _, last = f.get("ref").partition(":")
                        last = last or first
                        (l1, r1), (l2, r2) = coordinate_from_string(first), coordinate_from_string(last)
                        box = (r1, column_index_from_string(l1), r2, column_index_from_string(l2))
                        for rr in range(box[0], box[2] + 1):
                            for cc in range(box[1], box[3] + 1):
                                member = model.cell(rr, cc)
                                member.formula, member.array = True, box

    def save_as(self, path: Path, file_format: int) -> None:
        if BEHAVIOR.corrupt_save:
            path.write_bytes(b"PK\x03\x04 rusak")
            return
        source = zipfile.ZipFile(io.BytesIO(self.source))
        by_part = {sheet.part: sheet for sheet in self.sheets}
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                name, body = item.filename, source.read(item.filename)
                if name in by_part:
                    body = self._strip_converted(body.decode("utf-8"), by_part[name]).encode("utf-8")
                elif name == "xl/workbook.xml":
                    body = self._apply_states(body.decode("utf-8")).encode("utf-8")
                elif name == "xl/vbaProject.bin" and BEHAVIOR.drop_vba_on_save:
                    continue
                elif name == "xl/calcChain.xml":
                    continue
                target.writestr(name, body)
        self.app.save_formats.append(file_format)

    @staticmethod
    def _strip_converted(xml: str, sheet: FakeSheetModel) -> str:
        pattern = re.compile(r'<c\b[^>]*?\br="([A-Z]+)([0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.S)

        def repl(match: re.Match) -> str:
            cell = sheet.cells.get((int(match.group(2)), column_index_from_string(match.group(1))))
            if cell is None or cell.formula:
                return match.group(0)
            return _F_PATTERN.sub("", match.group(0))

        return pattern.sub(repl, xml)

    def _apply_states(self, xml: str) -> str:
        def repl(match: re.Match) -> str:
            tag = match.group(0)
            name = re.search(r'\bname="([^"]*)"', tag).group(1)
            sheet = next(item for item in self.sheets if item.name == name)
            tag = re.sub(r'\s+state="[^"]*"', "", tag)
            state = VISIBLE_TO_STATE[sheet.visible]
            return tag.replace("<sheet ", f'<sheet state="{state}" ', 1) if state != "visible" else tag

        return re.sub(r"<sheet\b[^>]*?/>", repl, xml)


class _Collection:
    def __init__(self, items: list[Any]) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - mengikuti API COM
        return len(self._items)

    def Item(self, index: int) -> Any:  # noqa: N802
        return self._items[index - 1]


class FakeRange:
    def __init__(self, sheet: FakeSheetModel, areas: list[tuple[int, int, int, int]]) -> None:
        self.sheet = sheet
        self.areas = areas

    # ---- geometri
    @property
    def Row(self) -> int:  # noqa: N802
        return min(area[0] for area in self.areas)

    @property
    def Column(self) -> int:  # noqa: N802
        return min(area[1] for area in self.areas)

    @property
    def Rows(self):  # noqa: N802
        r1, c1, r2, c2 = self.areas[0]
        return _Collection([FakeRange(self.sheet, [(row, c1, row, c2)]) for row in range(r1, r2 + 1)])

    @property
    def Columns(self):  # noqa: N802
        _, c1, _, c2 = self.areas[0]
        return _Collection(list(range(c1, c2 + 1)))

    @property
    def Areas(self):  # noqa: N802
        return _Collection([FakeRange(self.sheet, [area]) for area in self.areas])

    @property
    def CountLarge(self) -> int:  # noqa: N802
        return sum((r2 - r1 + 1) * (c2 - c1 + 1) for r1, c1, r2, c2 in self.areas)

    @property
    def Cells(self):  # noqa: N802
        r1, c1, r2, c2 = self.areas[0]
        width = c2 - c1 + 1
        return _Collection([
            FakeRange(self.sheet, [(r1 + i // width, c1 + i % width, r1 + i // width, c1 + i % width)])
            for i in range((r2 - r1 + 1) * width)
        ])

    @property
    def Address(self) -> str:  # noqa: N802
        parts = []
        for r1, c1, r2, c2 in self.areas:
            first = f"${get_column_letter(c1)}${r1}"
            parts.append(first if (r1, c1) == (r2, c2) else f"{first}:${get_column_letter(c2)}${r2}")
        return ",".join(parts)

    @property
    def EntireRow(self):  # noqa: N802
        flags = {
            row in self.sheet.hidden_rows
            for r1, _, r2, _ in self.areas
            for row in range(r1, r2 + 1)
        }
        hidden = None if flags == {True, False} else (True in flags)
        return type("EntireRow", (), {"Hidden": hidden})()

    def _cells(self):
        for r1, c1, r2, c2 in self.areas:
            for row in range(r1, r2 + 1):
                for col in range(c1, c2 + 1):
                    yield (row, col), self.sheet.cell(row, col)

    # ---- nilai
    @property
    def Value2(self):  # noqa: N802
        def convert(cell: FakeCell):
            return ERROR_CODES.get(cell.value, -2146826259) if cell.is_error else cell.value

        r1, c1, r2, c2 = self.areas[0]
        if (r1, c1) == (r2, c2):
            return convert(self.sheet.cell(r1, c1))
        return tuple(
            tuple(convert(self.sheet.cell(row, col)) for col in range(c1, c2 + 1))
            for row in range(r1, r2 + 1)
        )

    @property
    def HasFormula(self) -> bool:  # noqa: N802
        return self.sheet.cell(self.areas[0][0], self.areas[0][1]).formula

    @property
    def HasArray(self) -> bool:  # noqa: N802
        return self.sheet.cell(self.areas[0][0], self.areas[0][1]).array is not None

    @property
    def CurrentArray(self) -> "FakeRange":  # noqa: N802
        return FakeRange(self.sheet, [self.sheet.cell(self.areas[0][0], self.areas[0][1]).array])

    # ---- operasi
    def SpecialCells(self, kind: int, sub: int | None = None) -> "FakeRange":  # noqa: N802
        if kind != -4123:
            raise ComError("hanya xlCellTypeFormulas yang didukung")
        found = sorted(
            (col, row) for (row, col), cell in self._cells()
            if cell.formula and (sub != 16 or cell.is_error)
        )
        if not found:
            raise ComError("No cells were found.")
        areas, start = [], None  # area = run vertikal per kolom (sengaja dapat memotong array 2 kolom)
        for col, row in found:
            if start and start[0] == col and start[2] == row - 1:
                start = (col, start[1], row)
            else:
                if start:
                    areas.append((start[1], start[0], start[2], start[0]))
                start = (col, row, row)
        areas.append((start[1], start[0], start[2], start[0]))
        return FakeRange(self.sheet, areas)

    def Copy(self) -> None:  # noqa: N802
        if len(self.areas) != 1:
            raise ComError("That command cannot be used on multiple selections.")
        rows = {row for r1, _, r2, _ in self.areas for row in range(r1, r2 + 1)}
        hidden = rows & self.sheet.hidden_rows
        mixed = bool(hidden) and hidden != rows
        snapshot = [
            (cell.value, cell.is_error)
            for (row, _), cell in self._cells()
            if not (mixed and row in hidden)
        ]
        self.sheet.workbook.app.clipboard = (self.areas, snapshot)

    def PasteSpecial(self, kind: int) -> None:  # noqa: N802
        app = self.sheet.workbook.app
        if kind != -4163 or app.clipboard is None or app.clipboard[0] != self.areas:
            raise ComError("PasteSpecial gagal")
        r1, c1, r2, c2 = self.areas[0]
        for _, cell in self._cells():
            if cell.array and not (r1 <= cell.array[0] and cell.array[2] <= r2 and c1 <= cell.array[1] and cell.array[3] <= c2):
                raise ComError("You can't change part of an array.")
        snapshot = app.clipboard[1]
        if BEHAVIOR.stolen_clipboard:
            snapshot = [("SALAH", False)] * len(snapshot)
        for (_, cell), (value, is_error) in zip(self._cells(), snapshot):
            cell.value, cell.is_error, cell.formula, cell.array = value, is_error, False, None


class FakeProtection:
    def __init__(self) -> None:
        for flag in (
            "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns",
            "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows",
            "AllowSorting", "AllowFiltering", "AllowUsingPivotTables",
        ):
            setattr(self, flag, flag == "AllowSorting")


class FakeWorksheet:
    def __init__(self, model: FakeSheetModel) -> None:
        self.model = model
        self.Protection = FakeProtection()
        self.ProtectDrawingObjects = True
        self.ProtectScenarios = True

    Name = property(lambda self: self.model.name)  # noqa: N815

    @property
    def Visible(self) -> int:  # noqa: N802
        return self.model.visible

    @Visible.setter
    def Visible(self, value: int) -> None:  # noqa: N802
        if BEHAVIOR.skip_visibility_restore and value != -1:
            return
        self.model.visible = value

    @property
    def ProtectContents(self) -> bool:  # noqa: N802
        return self.model.protected

    def Unprotect(self) -> None:  # noqa: N802
        if self.model.protected_with_password:
            raise ComError("The password you supplied is not correct.")
        self.model.protected = False

    def Protect(self, *args: Any) -> None:  # noqa: N802
        self.model.protect_calls.append(args)
        self.model.protected = True

    @property
    def UsedRange(self) -> FakeRange:  # noqa: N802
        keys = [key for key, cell in self.model.cells.items() if cell.formula or cell.value is not None]
        if not keys:
            return FakeRange(self.model, [(1, 1, 1, 1)])
        rows, cols = [k[0] for k in keys], [k[1] for k in keys]
        return FakeRange(self.model, [(min(rows), min(cols), max(rows), max(cols))])

    @property
    def Cells(self):  # noqa: N802
        model = self.model
        return type("Cells", (), {"Item": lambda _self, row, col: FakeRange(model, [(row, col, row, col)])})()

    def Range(self, first: FakeRange, last: FakeRange) -> FakeRange:  # noqa: N802
        return FakeRange(self.model, [(first.areas[0][0], first.areas[0][1], last.areas[0][2], last.areas[0][3])])


class FakeWorkbookApi:
    def __init__(self, model: FakeWorkbookModel | None, app: "FakeApp") -> None:
        self.model = model
        self.app = app
        self._sheets = [FakeWorksheet(sheet) for sheet in model.sheets] if model else []

    @property
    def Sheets(self):  # noqa: N802
        return _Collection(self._sheets)

    @property
    def Worksheets(self):  # noqa: N802
        return _Collection(self._sheets)

    def LinkSources(self, kind: int):  # noqa: N802
        return self.app.link_sources

    def SaveAs(self, path: str, file_format: int) -> None:  # noqa: N802
        self.model.save_as(Path(path), file_format)


class FakeBook:
    def __init__(self, app: "FakeApp", model: FakeWorkbookModel | None) -> None:
        self.app = app
        self.api = FakeWorkbookApi(model, app)
        self.closed = False

    def close(self) -> None:
        self.closed = True
        if self in self.app.open_books:
            self.app.open_books.remove(self)


class FakeBooks:
    def __init__(self, app: "FakeApp") -> None:
        self.app = app

    def add(self) -> FakeBook:
        book = FakeBook(self.app, None)
        self.app.open_books.append(book)
        return book

    def open(self, path: str, update_links: Any = None, read_only: Any = None) -> FakeBook:
        stem = Path(path).stem
        if stem in BEHAVIOR.open_fail_stems:
            raise ComError("Excel tidak dapat membuka file")
        self.app.opened_paths.append((path, read_only))
        book = FakeBook(self.app, FakeWorkbookModel(self.app, Path(path)))
        self.app.open_books.append(book)
        return book


class FakeApplication:
    """Objek yang dikembalikan `app.api` (Excel.Application)."""

    def __init__(self, app: "FakeApp") -> None:
        self._app = app
        self.Calculation = -4105
        self.CutCopyMode = False
        self.calls: list[str] = []

    @property
    def CalculationState(self) -> int:  # noqa: N802
        return 1 if BEHAVIOR.calc_never_done else 0

    def _calculate(self, name: str) -> None:
        self.calls.append(name)
        if BEHAVIOR.hang_in_calculate:
            while self._app.alive:
                time.sleep(0.01)
            raise ComError("The RPC server is unavailable")

    def Calculate(self) -> None:  # noqa: N802
        self._calculate("Calculate")

    def CalculateFull(self) -> None:  # noqa: N802
        self._calculate("CalculateFull")


class FakeApp:
    registry: list["FakeApp"] = []
    _pids = iter(range(5000, 10**6))

    def __init__(self, visible: bool = False, add_book: bool = True) -> None:
        if BEHAVIOR.launch_error:
            raise RuntimeError("Excel tidak terpasang")
        self.pid = next(self._pids)
        self.alive = True
        self.killed = False
        self.quit_called = False
        self.clipboard: Any = None
        self.link_sources: Any = None
        self.open_books: list[FakeBook] = []
        self.opened_paths: list[tuple[str, Any]] = []
        self.save_formats: list[int] = []
        self.books = FakeBooks(self)
        self.api = FakeApplication(self)
        self.registry.append(self)

    def quit(self) -> None:
        self.quit_called = True
        if not BEHAVIOR.quit_does_not_exit:
            self.alive = False

    def kill(self) -> None:
        self.killed = True
        self.alive = False


class FakeXlwings:
    """Pengganti modul `xlwings` untuk monkeypatch sys.modules."""

    App = FakeApp


def reset_fake_excel() -> None:
    global BEHAVIOR
    BEHAVIOR.__dict__.update(Behavior().__dict__)
    FakeApp.registry.clear()


def alive_excel_instances() -> list[FakeApp]:
    return [app for app in FakeApp.registry if app.alive]
