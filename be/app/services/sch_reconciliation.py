"""Read-only reconciliation preview for FSS master vs selected SCH workbooks.

The schedule files use 12 physical payload columns. Their `Delivery time` maps to the
master's `Delivery period`; status is classified from payload font colors. `Division`
and other helper/formula columns are intentionally excluded.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Sequence
from zipfile import ZipFile
import posixpath
import re
import xml.etree.ElementTree as ET


MASTER_COLUMNS = [
    "Date", "Vendor code", "Vendor name", "Receiving Area", "Delivery date",
    "Delivery period", "Firm", "Classification", "PO number", "Item",
    "Parts number", "Parts name", "Order Quantity",
]

SCHEDULE_COLUMNS = [
    "Date", "Vendor code", "Vendor name", "Receiving Area", "Delivery date",
    "Delivery period", "Classification", "PO number", "Item", "Parts number",
    "Parts name", "Order Quantity",
]

MASTER_ALIASES = {
    "date": "Date", "vendorcode": "Vendor code", "vendorname": "Vendor name",
    "receivingarea": "Receiving Area", "recievingarea": "Receiving Area",
    "deliverydate": "Delivery date", "deliveryperiod": "Delivery period",
    "firm": "Firm", "classification": "Classification", "ponumber": "PO number",
    "ponum": "PO number", "item": "Item", "partsnumber": "Parts number",
    "partnumber": "Parts number", "partsname": "Parts name", "partname": "Parts name",
    "orderquantity": "Order Quantity", "orderqty": "Order Quantity",
}
SCHEDULE_ALIASES = {
    **{k: v for k, v in MASTER_ALIASES.items() if k != "firm" and k != "deliveryperiod"},
    "deliveryperiod": "Delivery period",
}

MATCH_KEY = [
    "Vendor code", "PO number", "Item", "Parts number", "Receiving Area",
    "Delivery date", "Delivery period",
]
LOOSE_KEY = ["Vendor code", "PO number", "Item", "Parts number", "Receiving Area"]
FORECAST_OVERLAP_KEY = [
    "Vendor code", "Receiving Area", "Parts number", "Delivery date", "Delivery period",
]
TRANSITION_KEY = [
    "Vendor code", "Receiving Area", "Delivery date", "Delivery period", "Parts number",
]
TRANSITION_SOFT_COLUMNS = ["PO number", "Item"]
COMPARE_COLUMNS = ["Order Quantity", "Classification", "Vendor name", "Parts name"]
PLACEHOLDERS = {"", "0", "00", "000", "0000", "00000", "000000", "0000000000", "-", "--", "*", "N/A", "NA", "NULL", "TBD"}
EMPTY_RUN_STOP = 5_000
HEADER_SCAN_LIMIT = 15


class SCHReconciliationError(ValueError):
    """Input workbook is unreadable or its schema is not safe to reconcile."""


class RowType(str, Enum):
    FIRM = "Firm"
    FORECAST = "Forecast"
    AMBIGUOUS = "Ambigu"
    EMPTY = "Kosong"


class Category(str, Enum):
    FIRM_KEPT = "Firm lama dipertahankan"
    FIRM_EXISTS = "Firm terbaru sudah ada (tidak ditambahkan)"
    FIRM_NEW = "Firm baru akan ditambahkan"
    FORECAST_OLD = "Forecast lama akan diganti"
    FORECAST_NEW = "Forecast terbaru akan ditambahkan"
    DUPLICATE = "Kandidat duplikat"
    VALUE_CHANGED = "Perubahan nilai order (perlu diperiksa)"
    STATUS_CONFLICT = "Konflik klasifikasi Firm/Forecast"
    AMBIGUOUS = "Identitas ambigu"
    SCHEDULE_SHIFT = "Kemungkinan perubahan jadwal"
    FORECAST_TO_FIRM = "Forecast lama menjadi Firm (transisi)"


@dataclass
class SheetRow:
    row_no: int
    values: dict[str, str]
    row_type: RowType
    note: str = ""


@dataclass
class MasterRow:
    index: int
    values: dict[str, str]
    raw: list[Any]
    is_forecast: bool
    ambiguous: bool = False


@dataclass
class ReconEntry:
    category: Category
    message: str = ""
    sheet_row: int | None = None
    master_row: int | None = None
    old_value: str = ""
    new_value: str = ""
    blocking: bool = False


@dataclass
class Transition:
    sheet_row: int
    master_row: int
    quantity_changed: bool
    old_quantity: str = ""
    new_quantity: str = ""


@dataclass
class ReconciliationResult:
    entries: list[ReconEntry] = field(default_factory=list)
    counts: dict[Category, int] = field(default_factory=dict)
    old_firm: list[SheetRow] = field(default_factory=list)
    old_forecast: list[SheetRow] = field(default_factory=list)
    firm_to_add: list[MasterRow] = field(default_factory=list)
    forecast_to_add: list[MasterRow] = field(default_factory=list)
    transitions: list[Transition] = field(default_factory=list)
    dropped_forecast_rows: list[int] = field(default_factory=list)
    blocking: list[str] = field(default_factory=list)
    review: list[str] = field(default_factory=list)
    info: list[str] = field(default_factory=list)


@dataclass
class ScheduleRules:
    match_key: list[str] = field(default_factory=lambda: MATCH_KEY.copy())
    loose_key: list[str] = field(default_factory=lambda: LOOSE_KEY.copy())
    forecast_overlap_key: list[str] = field(default_factory=lambda: FORECAST_OVERLAP_KEY.copy())
    transition_key: list[str] = field(default_factory=lambda: TRANSITION_KEY.copy())
    transition_soft_columns: list[str] = field(default_factory=lambda: TRANSITION_SOFT_COLUMNS.copy())
    compare_columns: list[str] = field(default_factory=lambda: COMPARE_COLUMNS.copy())
    placeholder_values: list[str] = field(default_factory=lambda: sorted(PLACEHOLDERS))
    forecast_overlap_action: str = "exclude"


def _normalize_header(value: Any) -> str:
    return "".join(ch for ch in str(value or "").strip().casefold() if ch.isalnum())


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y%m%d")
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if value != value:
            return ""
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value).strip()


def _is_excel_date_serial(text: str) -> bool:
    if not text.replace(".", "", 1).isdigit():
        return False
    whole = text.split(".", 1)[0]
    return len(whole) == 5 and 30_000 <= int(whole) <= 80_000


def _serial_to_ymd(text: str) -> str:
    from datetime import timedelta
    return (datetime(1899, 12, 30) + timedelta(days=int(float(text)))).strftime("%Y%m%d")


def _normalize_value(column: str, value: Any, number_format: str = "General") -> str:
    text = _cell_text(value)
    if not text:
        return ""
    if column in {"Date", "Delivery date"}:
        if _is_excel_date_serial(text):
            return _serial_to_ymd(text)
        if not text.isdigit():
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(text, fmt).strftime("%Y%m%d")
                except ValueError:
                    pass
        return text
    if column in {"Delivery period", "Order Quantity"}:
        try:
            number = float(text.replace(",", ""))
            return str(int(number)) if number.is_integer() else str(number)
        except ValueError:
            return text
    if column == "Item":
        return (text.lstrip("0") or "0") if text.isdigit() else text
    if isinstance(value, (int, float)) and text.lstrip("-").isdigit() and number_format and set(number_format) == {"0"}:
        return text.zfill(len(number_format))
    return text


def _color_kind(color_type: str, color_value: str | None) -> str:
    if color_type == "rgb":
        rgb = (color_value or "")[-6:].upper()
        if rgb == "FF0000":
            return "forecast"
        if rgb in {"000000", "202124"}:
            return "firm"
        return "other"
    if color_type == "indexed":
        if color_value == "10":
            return "forecast"
        if color_value in {"8", "64"}:
            return "firm"
        return "other"
    if color_type == "theme":
        return "firm" if color_value in {"0", "1"} else "other"
    if color_type in {"auto", "none"}:
        return "firm"
    return "other"


def _classify(values: Sequence[Any], colors: Sequence[str]) -> tuple[RowType, str]:
    kinds = {color for value, color in zip(values, colors) if _cell_text(value) != ""}
    if not kinds:
        return RowType.EMPTY, ""
    if kinds == {"firm"}:
        return RowType.FIRM, ""
    if kinds == {"forecast"}:
        return RowType.FORECAST, ""
    if "other" in kinds:
        return RowType.AMBIGUOUS, "warna font tidak dikenal"
    return RowType.AMBIGUOUS, "warna font Firm/Forecast bercampur"


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

def _tag(name: str) -> str:
    return f"{{{_MAIN_NS}}}{name}"


def _column_number(cell_ref: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_ref.upper())
    if not letters:
        return 0
    number = 0
    for char in letters.group(1):
        number = number * 26 + ord(char) - 64
    return number


def _shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall(_tag("si")):
        # Handles both simple and rich-text shared strings.
        values.append("".join(node.text or "" for node in item.iter(_tag("t"))))
    return values


def _style_tables(archive: ZipFile) -> tuple[list[str], list[str]]:
    """Return style-id -> category and style-id -> number format."""
    if "xl/styles.xml" not in archive.namelist():
        return ["firm"], ["General"]
    root = ET.fromstring(archive.read("xl/styles.xml"))
    fonts_node = root.find(_tag("fonts"))
    font_kinds: list[str] = []
    for font in (list(fonts_node) if fonts_node is not None else []):
        color = font.find(_tag("color"))
        if color is None:
            font_kinds.append("firm")
        elif "rgb" in color.attrib:
            font_kinds.append(_color_kind("rgb", color.attrib.get("rgb")))
        elif "indexed" in color.attrib:
            font_kinds.append(_color_kind("indexed", color.attrib.get("indexed")))
        elif "theme" in color.attrib:
            font_kinds.append(_color_kind("theme", color.attrib.get("theme")))
        elif color.attrib.get("auto") == "1":
            font_kinds.append("firm")
        else:
            font_kinds.append("other")

    custom_formats: dict[int, str] = {}
    num_fmts = root.find(_tag("numFmts"))
    if num_fmts is not None:
        for fmt in num_fmts.findall(_tag("numFmt")):
            try:
                custom_formats[int(fmt.attrib["numFmtId"])] = fmt.attrib.get("formatCode", "General")
            except (KeyError, ValueError):
                continue

    xfs = root.find(_tag("cellXfs"))
    style_kinds: list[str] = []
    number_formats: list[str] = []
    for xf in (list(xfs) if xfs is not None else []):
        try:
            font_id = int(xf.attrib.get("fontId", "0"))
        except ValueError:
            font_id = 0
        style_kinds.append(font_kinds[font_id] if font_id < len(font_kinds) else "other")
        try:
            num_fmt_id = int(xf.attrib.get("numFmtId", "0"))
        except ValueError:
            num_fmt_id = 0
        number_formats.append(custom_formats.get(num_fmt_id, "General"))
    return style_kinds or ["firm"], number_formats or ["General"]


def _worksheet_path(archive: ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {item.attrib["Id"]: item.attrib["Target"] for item in relationships.findall(f"{{{_PKG_REL_NS}}}Relationship")}
    for sheet in (workbook.find(_tag("sheets")) if workbook.find(_tag("sheets")) is not None else []):
        if sheet.attrib.get("name", "").casefold() == sheet_name.casefold():
            relationship_id = sheet.attrib.get(f"{{{_REL_NS}}}id", "")
            target = targets.get(relationship_id)
            if not target:
                break
            if target.startswith("/"):
                return target.lstrip("/")
            return posixpath.normpath(posixpath.join("xl", target))
    raise SCHReconciliationError(f"Sheet '{sheet_name}' tidak ditemukan.")


def _iter_xml_rows(archive: ZipFile, sheet_path: str):
    with archive.open(sheet_path) as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            if element.tag == _tag("row"):
                yield element
                element.clear()


def _decode_cell(cell: ET.Element, shared: Sequence[str]) -> Any:
    cell_type = cell.attrib.get("t", "n")
    value_node = cell.find(_tag("v"))
    if cell_type == "inlineStr":
        inline = cell.find(_tag("is"))
        return "".join(node.text or "" for node in inline.iter(_tag("t"))) if inline is not None else ""
    if value_node is None or value_node.text is None:
        return None
    raw = value_node.text
    if cell_type == "s":
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "b":
        return raw == "1"
    if cell_type in {"str", "e"}:
        return raw
    # Numeric Excel cells: keep integers as Python int so identifiers are not decorated with .0.
    try:
        number = float(raw)
        return int(number) if number.is_integer() else number
    except ValueError:
        return raw


def _row_cell_maps(
    row_element: ET.Element,
    shared: Sequence[str],
    style_kinds: Sequence[str],
    number_formats: Sequence[str],
    wanted_columns: set[int],
) -> tuple[dict[int, Any], dict[int, str], dict[int, str]]:
    values: dict[int, Any] = {}
    colors: dict[int, str] = {}
    formats: dict[int, str] = {}
    for cell in row_element.findall(_tag("c")):
        col = _column_number(cell.attrib.get("r", ""))
        if col not in wanted_columns:
            continue
        value = _decode_cell(cell, shared)
        style_id = int(cell.attrib.get("s", "0"))
        values[col] = value
        colors[col] = style_kinds[style_id] if style_id < len(style_kinds) else "other"
        formats[col] = number_formats[style_id] if style_id < len(number_formats) else "General"
    return values, colors, formats


def _discover_header(
    archive: ZipFile,
    sheet_path: str,
    shared: Sequence[str],
    style_kinds: Sequence[str],
    number_formats: Sequence[str],
    aliases: dict[str, str],
    required: Sequence[str],
    is_schedule: bool,
) -> tuple[int, dict[str, int]]:
    best_row = 0
    best_map: dict[str, int] = {}
    best_duplicates: list[str] = []
    for row_element in _iter_xml_rows(archive, sheet_path):
        row_number = int(row_element.attrib.get("r", "0"))
        if row_number > HEADER_SCAN_LIMIT:
            break
        values, _, _ = _row_cell_maps(row_element, shared, style_kinds, number_formats, set(range(1, 61)))
        mapping: dict[str, int] = {}
        duplicates: list[str] = []
        for column, value in values.items():
            if is_schedule and isinstance(value, str) and value.strip() == "date":
                continue
            canonical = aliases.get(_normalize_header(value))
            if canonical not in required:
                continue
            if canonical in mapping:
                duplicates.append(canonical)
            else:
                mapping[canonical] = column
        if len(mapping) > len(best_map):
            best_row, best_map, best_duplicates = row_number, mapping, duplicates
    missing = [name for name in required if name not in best_map]
    if not best_row or missing:
        raise SCHReconciliationError("Header wajib tidak lengkap: " + ", ".join(missing or required))
    if best_duplicates:
        raise SCHReconciliationError("Header ganda: " + ", ".join(sorted(set(best_duplicates))))
    positions = [best_map[name] for name in required]
    if positions != sorted(positions):
        details = ", ".join(f"{name}={_excel_column(best_map[name])}" for name in required)
        raise SCHReconciliationError("Urutan header tidak sesuai: " + details)
    return best_row, best_map


def _read_rows_from_xml(
    archive: ZipFile,
    sheet_path: str,
    header_row: int,
    column_map: dict[str, int],
    shared: Sequence[str],
    style_kinds: Sequence[str],
    number_formats: Sequence[str],
    columns: Sequence[str],
    schedule_mode: bool,
) -> tuple[list[Any], int]:
    wanted = set(column_map.values())
    first_col = min(wanted)
    rows: list[Any] = []
    empty_run = 0
    last_seen_row = header_row
    scanned_rows = 0
    for row_element in _iter_xml_rows(archive, sheet_path):
        row_number = int(row_element.attrib.get("r", "0"))
        if row_number <= header_row:
            continue
        if row_number > last_seen_row + 1:
            gap = row_number - last_seen_row - 1
            empty_run += gap
            scanned_rows += gap
            if empty_run >= EMPTY_RUN_STOP:
                break
        last_seen_row = row_number
        scanned_rows += 1
        raw_by_col, color_by_col, fmt_by_col = _row_cell_maps(
            row_element, shared, style_kinds, number_formats, wanted
        )
        raw = [raw_by_col.get(column_map[name]) for name in columns]
        if not any(value not in (None, "") for value in raw):
            empty_run += 1
            if empty_run >= EMPTY_RUN_STOP:
                break
            continue
        empty_run = 0
        colors = [color_by_col.get(column_map[name], "firm") for name in columns]
        formats = [fmt_by_col.get(column_map[name], "General") for name in columns]
        row_type, note = _classify(raw, colors)
        values = {
            name: _normalize_value(name, value, fmt)
            for name, value, fmt in zip(columns, raw, formats)
        }
        if schedule_mode:
            values["Firm"] = ""
            rows.append(SheetRow(row_no=row_number, values=values, row_type=row_type, note=note))
        else:
            rows.append(MasterRow(
                index=row_number,
                values=values,
                raw=raw,
                is_forecast=row_type == RowType.FORECAST,
                ambiguous=row_type in {RowType.AMBIGUOUS, RowType.EMPTY},
            ))
    return rows, scanned_rows


def read_master(content: bytes, filename: str) -> tuple[list[MasterRow], dict[str, int]]:
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise SCHReconciliationError("Data FSS master harus berformat .xlsx atau .xlsm.")
    try:
        archive = ZipFile(BytesIO(content))
    except Exception as exc:
        raise SCHReconciliationError(f"Data FSS master tidak dapat dibuka ({type(exc).__name__}).") from exc
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_names = [item.attrib.get("name", "") for item in (workbook.find(_tag("sheets")) if workbook.find(_tag("sheets")) is not None else [])]
        master_sheet = next((name for name in sheet_names if name.casefold() == "data fss"), sheet_names[0] if sheet_names else "")
        if not master_sheet:
            raise SCHReconciliationError("Workbook master tidak memiliki worksheet.")
        path = _worksheet_path(archive, master_sheet)
        shared = _shared_strings(archive)
        style_kinds, number_formats = _style_tables(archive)
        header_row, cmap = _discover_header(archive, path, shared, style_kinds, number_formats,
                                            MASTER_ALIASES, MASTER_COLUMNS, False)
        rows, _ = _read_rows_from_xml(archive, path, header_row, cmap, shared, style_kinds,
                                      number_formats, MASTER_COLUMNS, False)
    except SCHReconciliationError:
        raise
    except Exception as exc:
        raise SCHReconciliationError(f"Data FSS master gagal dibaca ({type(exc).__name__}).") from exc
    finally:
        archive.close()
    if not rows:
        raise SCHReconciliationError("Data FSS master tidak memiliki baris order.")
    return rows, cmap


def read_schedule(content: bytes, filename: str) -> tuple[list[SheetRow], dict[str, int], dict[str, Any]]:
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise SCHReconciliationError(f"{filename}: format harus .xlsx atau .xlsm.")
    try:
        archive = ZipFile(BytesIO(content))
    except Exception as exc:
        raise SCHReconciliationError(f"{filename}: workbook tidak dapat dibuka ({type(exc).__name__}).") from exc
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_names = [item.attrib.get("name", "") for item in (workbook.find(_tag("sheets")) if workbook.find(_tag("sheets")) is not None else [])]
        sheet_name = next((name for name in sheet_names if name.casefold() == "date"), None)
        if sheet_name is None:
            sheet_name = next((name for name in sheet_names if name.casefold() == "data"), None)
        if sheet_name is None:
            raise SCHReconciliationError(f"{filename}: sheet DATE/DATA tidak ditemukan.")
        path = _worksheet_path(archive, sheet_name)
        shared = _shared_strings(archive)
        style_kinds, number_formats = _style_tables(archive)
        header_row, cmap = _discover_header(archive, path, shared, style_kinds, number_formats,
                                            SCHEDULE_ALIASES, SCHEDULE_COLUMNS, True)
        rows, scanned_rows = _read_rows_from_xml(archive, path, header_row, cmap, shared, style_kinds,
                                                number_formats, SCHEDULE_COLUMNS, True)
    except SCHReconciliationError:
        raise
    except Exception as exc:
        raise SCHReconciliationError(f"{filename}: gagal membaca worksheet ({type(exc).__name__}).") from exc
    finally:
        archive.close()
    if not rows:
        raise SCHReconciliationError(f"{filename}: tidak ada baris data pada sheet {sheet_name}.")
    summary = {
        "sheet_name": sheet_name,
        "header_row": header_row,
        "data_rows": len(rows),
        "firm_rows": sum(row.row_type == RowType.FIRM for row in rows),
        "forecast_rows": sum(row.row_type == RowType.FORECAST for row in rows),
        "ambiguous_rows": sum(row.row_type == RowType.AMBIGUOUS for row in rows),
        "scanned_rows": scanned_rows,
        "column_positions": {name: _excel_column(cmap[name]) for name in SCHEDULE_COLUMNS},
    }
    return rows, cmap, summary

def _excel_column(number: int) -> str:
    out = ""
    while number:
        number, rem = divmod(number - 1, 26)
        out = chr(65 + rem) + out
    return out


def _key(values: dict[str, str], cols: Sequence[str]) -> tuple[str, ...]:
    return tuple(values.get(col, "") for col in cols)


def _group(rows: Iterable[Any], cols: Sequence[str]) -> dict[tuple[str, ...], list[Any]]:
    grouped: dict[tuple[str, ...], list[Any]] = defaultdict(list)
    for row in rows:
        grouped[_key(row.values, cols)].append(row)
    return grouped


def _fmt(cols: Sequence[str], key: tuple[str, ...]) -> str:
    return ", ".join(f"{col}={value or '∅'}" for col, value in zip(cols, key))


def _numeric_key(values: dict[str, str], cols: Sequence[str]) -> tuple[str, ...]:
    return tuple((value.lstrip("0") or "0") if value.isdigit() else value for value in (values.get(col, "") for col in cols))


def _contradicts(old: SheetRow, new: MasterRow) -> bool:
    for col in TRANSITION_SOFT_COLUMNS:
        old_value = old.values.get(col, "")
        new_value = new.values.get(col, "")
        if old_value.upper() not in PLACEHOLDERS and new_value.upper() not in PLACEHOLDERS and old_value != new_value:
            return True
    return False


def _transitions(
    result: ReconciliationResult,
    old_fc: Sequence[SheetRow],
    new_firm: Sequence[MasterRow],
    new_fc: Sequence[MasterRow],
) -> None:
    candidate_index = _group(result.firm_to_add, TRANSITION_KEY)
    new_fc_keys = {_key(row.values, TRANSITION_KEY) for row in new_fc}
    old_groups = _group(old_fc, TRANSITION_KEY)
    claims: dict[int, list[SheetRow]] = defaultdict(list)
    incomplete = 0
    for key, old_rows in old_groups.items():
        if any(value == "" for value in key):
            incomplete += len(old_rows)
            continue
        for candidate in candidate_index.get(key, []):
            for old in old_rows:
                if not _contradicts(old, candidate):
                    claims[candidate.index].append(old)

    handled_old: set[int] = set()
    for key, old_rows in old_groups.items():
        if any(value == "" for value in key):
            continue
        candidates = [candidate for candidate in candidate_index.get(key, [])
                      if any(not _contradicts(old, candidate) for old in old_rows)]
        if not candidates:
            continue
        row_text = ", ".join(str(row.row_no) for row in old_rows)
        master_text = ", ".join(str(row.index) for row in candidates)
        if len(old_rows) > 1 or len(candidates) > 1 or any(len(claims[row.index]) > 1 for row in candidates):
            result.entries.append(ReconEntry(
                Category.AMBIGUOUS,
                f"Transisi Forecast→Firm ambigu: {len(old_rows)} Forecast lama (schedule baris {row_text}) dan "
                f"{len(candidates)} Firm terbaru (master baris {master_text}) memiliki kunci sama ({_fmt(TRANSITION_KEY, key)}).",
                sheet_row=old_rows[0].row_no, master_row=candidates[0].index, blocking=True,
            ))
            handled_old.update(row.row_no for row in old_rows)
            continue
        old, new = old_rows[0], candidates[0]
        handled_old.add(old.row_no)
        old_qty = old.values.get("Order Quantity", "")
        new_qty = new.values.get("Order Quantity", "")
        result.transitions.append(Transition(old.row_no, new.index, old_qty != new_qty, old_qty, new_qty))
        result.entries.append(ReconEntry(
            Category.FORECAST_TO_FIRM,
            f"Forecast lama menjadi Firm: {_fmt(TRANSITION_KEY, key)}; quantity {old_qty or '∅'} → {new_qty or '∅'}.",
            sheet_row=old.row_no, master_row=new.index, old_value=old_qty, new_value=new_qty,
        ))

    dropped = [old.row_no for old in old_fc if old.row_no not in handled_old
               and (any(value == "" for value in _key(old.values, TRANSITION_KEY))
                    or _key(old.values, TRANSITION_KEY) not in new_fc_keys)]
    result.dropped_forecast_rows = dropped
    if dropped:
        result.info.append(f"{len(dropped)} Forecast lama tidak ditemukan lagi di Forecast/Firm master terbaru.")
    if incomplete:
        result.info.append(f"{incomplete} Forecast lama memiliki kolom kunci transisi kosong.")


def reconcile(old_rows: Sequence[SheetRow], master_rows: Sequence[MasterRow], rules: ScheduleRules | None = None) -> ReconciliationResult:
    rules = rules or ScheduleRules()
    result = ReconciliationResult()
    old_firm = [row for row in old_rows if row.row_type == RowType.FIRM]
    old_fc = [row for row in old_rows if row.row_type == RowType.FORECAST]
    result.old_firm, result.old_forecast = old_firm, old_fc

    for row in old_rows:
        if row.row_type == RowType.AMBIGUOUS:
            result.entries.append(ReconEntry(Category.AMBIGUOUS,
                f"Baris schedule {row.row_no}: klasifikasi tidak jelas ({row.note}).", sheet_row=row.row_no, blocking=True))
    for master in master_rows:
        if master.ambiguous:
            result.entries.append(ReconEntry(Category.AMBIGUOUS,
                f"Baris master {master.index}: klasifikasi Firm/Forecast tidak jelas.", master_row=master.index, blocking=True))
    if old_firm and old_fc and max(row.row_no for row in old_firm) > min(row.row_no for row in old_fc):
        result.entries.append(ReconEntry(Category.AMBIGUOUS,
            "Firm dan Forecast lama berselang-seling di worksheet target (Firm harus berada di atas Forecast).", blocking=True))

    new_firm = [row for row in master_rows if not row.is_forecast and not row.ambiguous]
    new_fc = [row for row in master_rows if row.is_forecast and not row.ambiguous]

    for label, rows, getter in (("schedule", old_firm, lambda row: row.row_no), ("master", new_firm, lambda row: row.index)):
        for row in rows:
            blank = [col for col in rules.match_key if row.values.get(col, "") == ""]
            if blank:
                result.entries.append(ReconEntry(Category.AMBIGUOUS,
                    f"Baris {label} {getter(row)}: kolom kunci kosong ({', '.join(blank)}).",
                    sheet_row=row.row_no if label == "schedule" else None,
                    master_row=row.index if label == "master" else None, blocking=True))
        for key, group in _group(rows, rules.match_key).items():
            if len(group) > 1:
                ids = ", ".join(str(getter(row)) for row in group)
                identical = len({tuple(sorted(row.values.items())) for row in group}) == 1
                result.entries.append(ReconEntry(
                    Category.DUPLICATE if identical else Category.AMBIGUOUS,
                    f"Kunci order muncul {len(group)}x di {label} (baris {ids}): {_fmt(rules.match_key, key)}", blocking=True))

    old_index: dict[tuple[str, ...], SheetRow] = {}
    for row in old_firm:
        old_index.setdefault(_key(row.values, rules.match_key), row)
    loose_old = _group(old_firm, rules.loose_key)
    numeric_old: dict[tuple[str, ...], SheetRow] = {}
    for row in old_firm:
        numeric_old.setdefault(_numeric_key(row.values, rules.match_key), row)

    matched_old: set[int] = set()
    firm_exists = 0
    for master in new_firm:
        key = _key(master.values, rules.match_key)
        old = old_index.get(key)
        if old is not None:
            matched_old.add(old.row_no)
            diffs = [(col, old.values.get(col, ""), master.values.get(col, ""))
                     for col in rules.compare_columns if old.values.get(col, "") != master.values.get(col, "")]
            if diffs:
                for col, old_value, new_value in diffs:
                    result.entries.append(ReconEntry(Category.VALUE_CHANGED,
                        f"{col} berbeda (Firm lama dipertahankan): {_fmt(rules.match_key, key)}",
                        sheet_row=old.row_no, master_row=master.index, old_value=old_value, new_value=new_value))
            else:
                firm_exists += 1
            continue

        near = numeric_old.get(_numeric_key(master.values, rules.match_key))
        if near is not None:
            differences = [f"{col}: schedule '{near.values.get(col, '')}' vs master '{master.values.get(col, '')}'"
                           for col in rules.match_key if near.values.get(col, "") != master.values.get(col, "")]
            result.entries.append(ReconEntry(Category.AMBIGUOUS,
                f"Master baris {master.index} mirip Firm lama (schedule baris {near.row_no}) setelah nol di depan diabaikan "
                f"({'; '.join(differences)}). Periksa format; tidak ditambahkan agar tidak terjadi duplikasi.",
                sheet_row=near.row_no, master_row=master.index, blocking=True))
            continue

        result.firm_to_add.append(master)
        shifted = loose_old.get(_key(master.values, rules.loose_key))
        if shifted:
            old_shift = shifted[0]
            result.entries.append(ReconEntry(Category.SCHEDULE_SHIFT,
                "Order tampak sama dengan Firm lama tetapi tanggal/periode berbeda; tetap dihitung sebagai Firm baru.",
                sheet_row=old_shift.row_no, master_row=master.index,
                old_value=f"{old_shift.values.get('Delivery date', '')} / P{old_shift.values.get('Delivery period', '')}",
                new_value=f"{master.values.get('Delivery date', '')} / P{master.values.get('Delivery period', '')}"))
        else:
            result.entries.append(ReconEntry(Category.FIRM_NEW, _fmt(rules.match_key, key),
                master_row=master.index, new_value=master.values.get("Order Quantity", "")))

    missing_old = len(old_firm) - len(matched_old)
    if missing_old:
        result.info.append(f"{missing_old} Firm lama tidak ditemukan di master terbaru; tetap dipertahankan.")

    _transitions(result, old_fc, new_firm, new_fc)

    firm_soft = {_key(row.values, rules.forecast_overlap_key): row for row in [*old_firm, *result.firm_to_add]}
    for master in new_fc:
        key = _key(master.values, rules.forecast_overlap_key)
        hit = firm_soft.get(key)
        if hit is not None:
            result.entries.append(ReconEntry(Category.STATUS_CONFLICT,
                f"Forecast bentrok dengan Firm (schedule baris {hit.row_no}): {_fmt(rules.forecast_overlap_key, key)} — tidak dimasukkan.",
                master_row=master.index, new_value=master.values.get("Order Quantity", "")))
            continue
        result.forecast_to_add.append(master)

    counts: dict[Category, int] = defaultdict(int)
    for entry in result.entries:
        counts[entry.category] += 1
        if entry.blocking:
            result.blocking.append(entry.message)
    counts[Category.FIRM_KEPT] = len(old_firm)
    counts[Category.FIRM_EXISTS] = firm_exists
    counts[Category.FIRM_NEW] = len(result.firm_to_add)
    counts[Category.FORECAST_OLD] = len(old_fc)
    counts[Category.FORECAST_NEW] = len(result.forecast_to_add)
    result.counts = dict(counts)
    for category, description in (
        (Category.VALUE_CHANGED, "order Firm lama memiliki nilai berbeda di master terbaru"),
        (Category.STATUS_CONFLICT, "Forecast bentrok dengan Firm"),
        (Category.SCHEDULE_SHIFT, "order Firm tampak berubah jadwal"),
    ):
        count = counts.get(category, 0)
        if count:
            result.review.append(f"{count} {description}.")
    return result


def _entry_json(entry: ReconEntry) -> dict[str, Any]:
    return {
        "category": entry.category.value,
        "message": entry.message,
        "schedule_row": entry.sheet_row,
        "master_row": entry.master_row,
        "old_value": entry.old_value,
        "new_value": entry.new_value,
        "blocking": entry.blocking,
    }


def preview_schedule_reconciliation(
    master_content: bytes,
    master_filename: str,
    schedule_inputs: dict[str, tuple[str, bytes]],
    max_entries: int = 100,
) -> dict[str, Any]:
    """Create a non-mutating preview for any selected set of schedule workbooks."""
    if not schedule_inputs:
        raise SCHReconciliationError("Pilih minimal satu workbook schedule untuk preview.")
    if len(schedule_inputs) > 20:
        raise SCHReconciliationError("Maksimal 20 workbook schedule dapat dipreview sekaligus.")

    master_rows, _ = read_master(master_content, master_filename)
    master_summary = {
        "filename": master_filename,
        "total_rows": len(master_rows),
        "firm_rows": sum(not row.is_forecast and not row.ambiguous for row in master_rows),
        "forecast_rows": sum(row.is_forecast and not row.ambiguous for row in master_rows),
        "ambiguous_rows": sum(row.ambiguous for row in master_rows),
    }

    files_result: list[dict[str, Any]] = []
    for index, (role, (filename, content)) in enumerate(schedule_inputs.items(), start=1):
        try:
            old_rows, _, file_summary = read_schedule(content, filename)
            result = reconcile(old_rows, master_rows)
            count_json = {category.value: count for category, count in result.counts.items()}
            preview_entries = [_entry_json(entry) for entry in result.entries[:max_entries]]
            files_result.append({
                "role": role,
                "label": f"Workbook {index}",
                "filename": filename,
                "valid": True,
                "status": "blocked" if result.blocking else ("review" if result.review else "ready"),
                "safe_to_update": not result.blocking and not result.review,
                **file_summary,
                "reconciliation": {
                    "counts": count_json,
                    "transition_count": len(result.transitions),
                    "firm_new_non_transition_count": max(0, len(result.firm_to_add) - len(result.transitions)),
                    "dropped_forecast_count": len(result.dropped_forecast_rows),
                    "blocking_count": len(result.blocking),
                    "review_count": len(result.review),
                    "info": result.info,
                    "review_messages": result.review,
                    "blocking_messages": result.blocking[:25],
                    "entries_total": len(result.entries),
                    "entries_preview": preview_entries,
                    "entries_truncated": max(0, len(result.entries) - len(preview_entries)),
                },
            })
        except SCHReconciliationError as exc:
            files_result.append({
                "role": role,
                "label": f"Workbook {index}",
                "filename": filename,
                "valid": False,
                "message": str(exc),
            })

    total_files = len(files_result)
    valid_files = sum(item.get("valid") is True for item in files_result)
    return {
        "status": "ok" if valid_files == total_files else "needs_review",
        "valid": valid_files == total_files,
        "master": master_summary,
        "total_files": total_files,
        "valid_files": valid_files,
        "invalid_files": total_files - valid_files,
        "files": files_result,
        "note": "Preview rekonsiliasi bersifat read-only. Tidak ada workbook yang ditulis atau diubah.",
    }
