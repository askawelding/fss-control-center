"""Production writer for the selected SCH workbooks.

Only the 12 mapped payload columns are written. Helper/formula columns between
those fields (and Division at the right edge) are never included in write ranges.
Production execution requires Windows + Microsoft Excel desktop + xlwings.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Sequence
from xml.etree import ElementTree as ET

from app.services.sch_reconciliation import (
    MASTER_COLUMNS,
    SCHEDULE_COLUMNS,
    MasterRow,
    ReconciliationResult,
    RowType,
    SCHReconciliationError,
    read_master,
    read_schedule,
    reconcile,
)

MAX_EXCEL_ROW = 1_048_576
WRITE_CHUNK_ROWS = 10_000
XL_CALC_MANUAL = -4135
XL_CALC_DONE = 0
XL_PASTE_FORMATS = -4122
ZERO_FORMAT = re.compile(r"^0+$")
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKBOOK_XML = "xl/workbook.xml"
MAX_SCHEDULE_FILES = 20


class SCHUpdateError(RuntimeError):
    """A safe Update SCH operation could not be completed."""


class ExcelUnavailable(SCHUpdateError):
    """Microsoft Excel desktop or xlwings is unavailable."""


@dataclass(frozen=True)
class ColumnRun:
    value_offset: int
    first_column: int
    last_column: int

    @property
    def width(self) -> int:
        return self.last_column - self.first_column + 1


@dataclass(frozen=True)
class SCHWritePlan:
    sheet_name: str
    header_row: int
    data_start_row: int
    start_row: int
    end_row: int
    old_last_row: int
    clear_to_row: int
    firm_add_count: int
    forecast_add_count: int
    runs: tuple[ColumnRun, ...]
    rows: tuple[tuple[Any, ...], ...]
    forecast_flags: tuple[bool, ...]


def _iter_write_chunks(
    total_rows: int,
    start_row: int,
    old_last_row: int,
    chunk_size: int = WRITE_CHUNK_ROWS,
):
    """Yield (start_offset, end_offset) chunks without crossing the old-data edge.

    Excel COM writes are deliberately split at old_last_row. This makes the
    transition from the pre-existing used payload range to newly appended rows
    a separate operation, which is especially important for delivery templates.
    End offsets are exclusive.
    """
    if chunk_size <= 0:
        raise ValueError("chunk_size harus lebih besar dari nol.")
    if total_rows <= 0:
        return

    offset = 0
    old_data_boundary = old_last_row - start_row + 1
    while offset < total_rows:
        next_offset = min(offset + chunk_size, total_rows)
        if offset < old_data_boundary < next_offset:
            next_offset = old_data_boundary
        if next_offset <= offset:
            raise SCHUpdateError("Batas chunk penulisan Excel tidak valid.")
        yield offset, next_offset
        offset = next_offset


def _runs_for_columns(column_map: dict[str, int]) -> tuple[ColumnRun, ...]:
    columns = [column_map[name] for name in SCHEDULE_COLUMNS]
    if len(set(columns)) != len(SCHEDULE_COLUMNS):
        raise SCHUpdateError("Pemetaan kolom schedule memiliki posisi ganda.")
    if columns != sorted(columns):
        raise SCHUpdateError("Urutan kolom schedule tidak sesuai urutan data master.")

    runs: list[ColumnRun] = []
    run_start_offset = 0
    for index in range(1, len(columns)):
        if columns[index] != columns[index - 1] + 1:
            runs.append(ColumnRun(run_start_offset, columns[run_start_offset], columns[index - 1]))
            run_start_offset = index
    runs.append(ColumnRun(run_start_offset, columns[run_start_offset], columns[-1]))
    if len(runs) < 2:
        raise SCHUpdateError(
            "Layout workbook ini tidak memiliki pemisah helper seperti yang diharapkan. "
            "Penulisan otomatis dihentikan demi melindungi kolom formula."
        )
    return tuple(runs)


def _master_to_schedule_values(master: MasterRow) -> tuple[Any, ...]:
    raw_map = dict(zip(MASTER_COLUMNS, master.raw))
    # Gunakan nilai Item yang sudah dinormalisasi untuk menghindari nol di depan.
    raw_map["Item"] = master.values.get("Item", raw_map.get("Item"))
    return tuple(raw_map.get(column) for column in SCHEDULE_COLUMNS)


def build_write_plan(
    old_rows: Sequence[Any],
    column_map: dict[str, int],
    header_row: int,
    reconciliation: ReconciliationResult,
) -> SCHWritePlan:
    """Prepare deterministic in-memory writes; this function never opens Excel."""
    if reconciliation.blocking:
        raise SCHUpdateError(
            f"Rekonsiliasi memiliki {len(reconciliation.blocking)} konflik pemblokir. "
            "Workbook tidak boleh ditulis."
        )
    if reconciliation.review:
        raise SCHUpdateError(
            f"Rekonsiliasi memiliki {len(reconciliation.review)} temuan yang perlu diperiksa. "
            "Workbook tidak boleh ditulis otomatis."
        )

    ordered = sorted(old_rows, key=lambda row: row.row_no)
    data_start = header_row + 1
    old_last = max((row.row_no for row in ordered), default=header_row)
    expected_rows = list(range(data_start, old_last + 1))
    actual_rows = [row.row_no for row in ordered]
    if actual_rows != expected_rows:
        raise SCHUpdateError(
            "Baris data schedule memiliki celah atau baris kosong di tengah area aktif. "
            "Periksa worksheet DATE sebelum update."
        )

    old_firm = sorted(reconciliation.old_firm, key=lambda row: row.row_no)
    old_forecast = sorted(reconciliation.old_forecast, key=lambda row: row.row_no)
    if old_firm and old_forecast:
        last_firm = max(row.row_no for row in old_firm)
        first_forecast = min(row.row_no for row in old_forecast)
        if first_forecast != last_firm + 1:
            raise SCHUpdateError(
                "Forecast lama tidak dimulai tepat setelah Firm lama. "
                "Penulisan otomatis dihentikan agar tidak menimpa baris yang tidak dikenal."
            )
        start_row = last_firm + 1
    elif old_firm:
        start_row = max(row.row_no for row in old_firm) + 1
    else:
        start_row = data_start

    additions = [*reconciliation.firm_to_add, *reconciliation.forecast_to_add]
    firm_count = len(reconciliation.firm_to_add)
    forecast_count = len(reconciliation.forecast_to_add)
    end_row = start_row + len(additions) - 1 if additions else start_row - 1
    if end_row > MAX_EXCEL_ROW:
        raise SCHUpdateError(
            f"Hasil memerlukan baris {end_row:,}, melewati batas Excel {MAX_EXCEL_ROW:,}."
        )

    rows = tuple(_master_to_schedule_values(row) for row in additions)
    flags = tuple([False] * firm_count + [True] * forecast_count)
    return SCHWritePlan(
        sheet_name="DATE",
        header_row=header_row,
        data_start_row=data_start,
        start_row=start_row,
        end_row=end_row,
        old_last_row=old_last,
        clear_to_row=max(old_last, end_row),
        firm_add_count=firm_count,
        forecast_add_count=forecast_count,
        runs=_runs_for_columns(column_map),
        rows=rows,
        forecast_flags=flags,
    )


def _excel_column(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _worksheet_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read(WORKBOOK_XML))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    sheets_node = workbook.find(f"{{{MAIN_NS}}}sheets")
    for sheet in (list(sheets_node) if sheets_node is not None else []):
        if sheet.attrib.get("name", "").casefold() == sheet_name.casefold():
            rid = sheet.attrib.get(f"{{{REL_NS}}}id", "")
            target = targets.get(rid)
            if not target:
                break
            if target.startswith("/"):
                return target.lstrip("/")
            import posixpath
            return posixpath.normpath(posixpath.join("xl", target))
    raise SCHUpdateError(f"Sheet '{sheet_name}' tidak ditemukan saat validasi output.")


def _formula_refs_at_rows(content: bytes, sheet_name: str, row_numbers: set[int]) -> dict[int, set[str]]:
    """Return formula cell refs outside Excel values for selected row numbers."""
    try:
        with zipfile.ZipFile(__import__("io").BytesIO(content)) as archive:
            sheet_path = _worksheet_path(archive, sheet_name)
            found: dict[int, set[str]] = {row: set() for row in row_numbers}
            with archive.open(sheet_path) as stream:
                for _, element in ET.iterparse(stream, events=("end",)):
                    if element.tag != f"{{{MAIN_NS}}}row":
                        continue
                    row_no = int(element.attrib.get("r", "0"))
                    if row_no in row_numbers:
                        for cell in element.findall(f"{{{MAIN_NS}}}c"):
                            if cell.find(f"{{{MAIN_NS}}}f") is not None:
                                found[row_no].add(cell.attrib.get("r", ""))
                    element.clear()
                    if row_no >= max(row_numbers, default=0):
                        break
            return found
    except SCHUpdateError:
        raise
    except Exception as exc:
        raise SCHUpdateError(
            f"Tidak dapat memeriksa formula helper ({type(exc).__name__})."
        ) from exc


def _formula_column_from_ref(cell_ref: str) -> str:
    match = re.match(r"([A-Z]+)", cell_ref.upper())
    return match.group(1) if match else ""


def _check_formula_coverage(
    content: bytes,
    sheet_name: str,
    row_start: int,
    row_end: int,
    first_data_row_formula_refs: set[str],
) -> list[str]:
    """Require every helper formula column to have a formula on each target row.

    Supported SCH templates pre-fill helper formulas below the active data.
    We verify the entire region that will be written/cleared before Excel is opened,
    so a missing formula in the middle cannot silently pass an edge-only check.
    """
    if row_end < row_start:
        return sorted({_formula_column_from_ref(ref) for ref in first_data_row_formula_refs if _formula_column_from_ref(ref)})

    helper_columns = sorted({
        _formula_column_from_ref(ref)
        for ref in first_data_row_formula_refs
        if _formula_column_from_ref(ref)
    })
    if not helper_columns:
        raise SCHUpdateError(
            "Tidak ditemukan formula helper pada baris data pertama. "
            "Penulisan otomatis dihentikan untuk mencegah hilangnya rumus."
        )

    missing_counts = {column: 0 for column in helper_columns}
    missing_samples: dict[str, list[int]] = {column: [] for column in helper_columns}
    next_expected_row = row_start
    try:
        with zipfile.ZipFile(__import__("io").BytesIO(content)) as archive:
            sheet_path = _worksheet_path(archive, sheet_name)
            with archive.open(sheet_path) as stream:
                for _, element in ET.iterparse(stream, events=("end",)):
                    if element.tag != f"{{{MAIN_NS}}}row":
                        continue
                    row_no = int(element.attrib.get("r", "0"))
                    if row_no < row_start:
                        element.clear()
                        continue
                    if row_no > row_end:
                        break

                    # An absent row element is also a missing formula row.
                    while next_expected_row < row_no:
                        for column in helper_columns:
                            missing_counts[column] += 1
                            if len(missing_samples[column]) < 5:
                                missing_samples[column].append(next_expected_row)
                        next_expected_row += 1

                    present = set()
                    for cell in element.findall(f"{{{MAIN_NS}}}c"):
                        if cell.find(f"{{{MAIN_NS}}}f") is None:
                            continue
                        column = _formula_column_from_ref(cell.attrib.get("r", ""))
                        if column in missing_counts:
                            present.add(column)

                    for column in helper_columns:
                        if column not in present:
                            missing_counts[column] += 1
                            if len(missing_samples[column]) < 5:
                                missing_samples[column].append(row_no)
                    next_expected_row = row_no + 1
                    element.clear()

        while next_expected_row <= row_end:
            for column in helper_columns:
                missing_counts[column] += 1
                if len(missing_samples[column]) < 5:
                    missing_samples[column].append(next_expected_row)
            next_expected_row += 1
    except SCHUpdateError:
        raise
    except Exception as exc:
        raise SCHUpdateError(
            f"Tidak dapat memeriksa kelengkapan formula helper ({type(exc).__name__})."
        ) from exc

    incomplete = {column: count for column, count in missing_counts.items() if count}
    if incomplete:
        details = "; ".join(
            f"{column}: {count:,} baris kosong (contoh baris {', '.join(map(str, missing_samples[column]))})"
            for column, count in incomplete.items()
        )
        raise SCHUpdateError(
            f"Formula helper belum lengkap pada rentang baris {row_start:,}–{row_end:,}: {details}. "
            "Perbaiki/extend formula di workbook sumber terlebih dahulu. Tidak ada output yang dibuat."
        )
    return helper_columns


def _read_calc_mode(content: bytes) -> str:
    try:
        with zipfile.ZipFile(__import__("io").BytesIO(content)) as archive:
            workbook_xml = archive.read(WORKBOOK_XML)
        root = ET.fromstring(workbook_xml)
        calc = root.find(f"{{{MAIN_NS}}}calcPr")
        return calc.attrib.get("calcMode", "auto") if calc is not None else "auto"
    except Exception as exc:
        raise SCHUpdateError("Tidak dapat membaca pengaturan kalkulasi workbook sumber.") from exc


def _set_calc_mode(path: Path, mode: str) -> None:
    """Restore the source workbook's calculation mode without rewriting XML namespaces.

    Do not use ElementTree to serialize xl/workbook.xml here. Excel workbooks may
    declare prefixes only through mc:Ignorable (for example x15). ElementTree
    drops namespace declarations that it does not see in element/attribute names,
    leaving mc:Ignorable with an undeclared prefix and making the XLSX unreadable.
    """
    if mode not in {"auto", "manual", "autoNoTable"}:
        raise SCHUpdateError(f"Mode kalkulasi workbook tidak dikenal: {mode!r}.")

    temp_path = path.with_name(path.name + ".calc.tmp")
    calc_tag_pattern = re.compile(
        rb"<(?P<qname>(?:[A-Za-z_][A-Za-z0-9_.-]*:)?calcPr)\b[^<>]*?/?>",
        re.DOTALL,
    )
    calc_mode_attribute = re.compile(
        rb"\s+calcMode\s*=\s*(?:\"[^\"]*\"|'[^']*')",
        re.DOTALL,
    )

    def validate_ignorable_prefixes(xml_bytes: bytes) -> None:
        root = ET.fromstring(xml_bytes)
        ignorable = root.attrib.get(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable",
            "",
        )
        if not ignorable:
            return
        declared = set(re.findall(rb"xmlns:([A-Za-z_][A-Za-z0-9_.-]*)\s*=", xml_bytes))
        missing = sorted(set(ignorable.split()) - {item.decode("ascii") for item in declared})
        if missing:
            raise SCHUpdateError(
                "Namespace workbook tidak lengkap untuk mc:Ignorable: " + ", ".join(missing)
            )

    def update_workbook_xml(xml_bytes: bytes) -> bytes:
        # Validate the existing part, but preserve its original byte representation.
        validate_ignorable_prefixes(xml_bytes)
        match = calc_tag_pattern.search(xml_bytes)
        if match:
            tag = match.group(0)
            tag = calc_mode_attribute.sub(b"", tag)
            if mode != "auto":
                close_at = tag.rfind(b"/>")
                if close_at >= 0:
                    tag = tag[:close_at] + b' calcMode="' + mode.encode("ascii") + b'"' + tag[close_at:]
                else:
                    close_at = tag.rfind(b">")
                    if close_at < 0:
                        raise SCHUpdateError("Elemen calcPr pada workbook.xml tidak valid.")
                    tag = tag[:close_at] + b' calcMode="' + mode.encode("ascii") + b'"' + tag[close_at:]
            updated = xml_bytes[:match.start()] + tag + xml_bytes[match.end():]
            validate_ignorable_prefixes(updated)
            return updated

        if mode == "auto":
            return xml_bytes

        root_match = re.search(
            rb"<(?P<qname>(?:[A-Za-z_][A-Za-z0-9_.-]*:)?workbook)\b[^<>]*>",
            xml_bytes,
        )
        if not root_match:
            raise SCHUpdateError("Elemen workbook tidak ditemukan pada workbook.xml.")
        qname = root_match.group("qname")
        prefix = qname.rsplit(b":", 1)[0] + b":" if b":" in qname else b""
        new_calc = b"<" + prefix + b'calcPr calcMode="' + mode.encode("ascii") + b'"/>'

        # OOXML places calcPr immediately before extLst when extLst exists.
        ext_match = re.search(rb"<" + re.escape(prefix) + rb"extLst\b", xml_bytes)
        if ext_match:
            insert_at = ext_match.start()
        else:
            closing = re.search(rb"</" + re.escape(qname) + rb"\s*>", xml_bytes)
            if not closing:
                raise SCHUpdateError("Elemen penutup workbook tidak ditemukan pada workbook.xml.")
            insert_at = closing.start()
        updated = xml_bytes[:insert_at] + new_calc + xml_bytes[insert_at:]
        validate_ignorable_prefixes(updated)
        return updated

    try:
        found_workbook_xml = False
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temp_path, "w") as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == WORKBOOK_XML:
                    found_workbook_xml = True
                    data = update_workbook_xml(data)
                target.writestr(item, data)
        if not found_workbook_xml:
            raise SCHUpdateError("Bagian xl/workbook.xml tidak ditemukan pada file hasil Excel.")
        # Check all member CRCs before replacing the valid Excel output.
        with zipfile.ZipFile(temp_path, "r") as check:
            bad_member = check.testzip()
            if bad_member:
                raise SCHUpdateError(f"Validasi ZIP workbook gagal pada {bad_member}.")
        os.replace(temp_path, path)
    except SCHUpdateError:
        temp_path.unlink(missing_ok=True)
        raise
    except Exception as exc:
        temp_path.unlink(missing_ok=True)
        raise SCHUpdateError(f"Tidak dapat memulihkan mode kalkulasi workbook ({type(exc).__name__}).") from exc


def _coerce_value_for_target(value: Any, target_sample: Any, number_format: str, column: str) -> Any:
    if value is None or value == "":
        return None
    if isinstance(value, (datetime, date, bool, int, float)):
        return value
    text = str(value).strip()
    if not text:
        return None
    if column == "Item":
        try:
            number = float(text.replace(",", ""))
            return int(number) if number.is_integer() else number
        except ValueError:
            return text
    if isinstance(target_sample, (int, float)) and not isinstance(target_sample, bool):
        if (text.isdigit() and len(text) > 1 and text.startswith("0")
                and not (ZERO_FORMAT.fullmatch(number_format or "") and len(number_format) >= len(text))):
            return text
        try:
            number = float(text.replace(",", ""))
            return int(number) if number.is_integer() else number
        except ValueError:
            return text
    if isinstance(target_sample, (datetime, date)):
        for fmt in ("%Y%m%d", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
    return text


def _make_values_for_plan(ws: Any, plan: SCHWritePlan) -> tuple[list[list[Any]], dict[int, str]]:
    if not plan.rows:
        return [], {}
    # Prefer the last existing Firm row for type inference when the first Forecast has blank PO/Item.
    sample_row = max(plan.data_start_row, plan.start_row - 1)
    sample_values: dict[int, Any] = {}
    formats: dict[int, str] = {}
    for run in plan.runs:
        for col in range(run.first_column, run.last_column + 1):
            sample = ws.range((sample_row, col))
            sample_values[col] = sample.value
            formats[col] = str(sample.number_format or "General")
    # Build rows in the same logical order as SCHEDULE_COLUMNS.
    prepared: list[list[Any]] = [list(row) for row in plan.rows]
    for row in prepared:
        for index, column in enumerate(SCHEDULE_COLUMNS):
            physical_col = _physical_column_for_offset(plan.runs, index)
            raw = row[index]
            row[index] = _coerce_value_for_target(raw, sample_values.get(physical_col), formats.get(physical_col, "General"), column)
    return prepared, formats


def _physical_column_for_offset(runs: Sequence[ColumnRun], offset: int) -> int:
    for run in runs:
        if run.value_offset <= offset < run.value_offset + run.width:
            return run.first_column + offset - run.value_offset
    raise SCHUpdateError("Pemetaan offset kolom tidak konsisten.")


def _verify_helper_formulas(ws: Any, plan: SCHWritePlan, max_column: int) -> set[int]:
    payload = {col for run in plan.runs for col in range(run.first_column, run.last_column + 1)}
    first_row_formulas = ws.range((plan.data_start_row, 1), (plan.data_start_row, max_column)).formula
    if not isinstance(first_row_formulas, (list, tuple)):
        first_row_formulas = [first_row_formulas]
    if first_row_formulas and isinstance(first_row_formulas[0], (list, tuple)):
        first_row_formulas = list(first_row_formulas[0])
    formula_columns = {
        col for col, formula in enumerate(first_row_formulas, start=1)
        if col not in payload and isinstance(formula, str) and formula.startswith("=")
    }
    if plan.end_row < plan.start_row:
        return formula_columns
    missing: list[str] = []
    for col in sorted(formula_columns):
        for row in {plan.start_row, plan.end_row}:
            formula = ws.range((row, col)).formula
            if not (isinstance(formula, str) and formula.startswith("=")):
                missing.append(f"{_excel_column(col)}{row}")
    if missing:
        raise SCHUpdateError(
            "Formula helper belum tersedia pada seluruh baris yang dibutuhkan: "
            + ", ".join(missing[:20])
            + ". Perluas formula di Excel terlebih dahulu; tidak ada output yang dibuat."
        )
    return formula_columns


def _open_excel_copy(path: Path):
    try:
        import xlwings as xw
    except ImportError as exc:
        raise ExcelUnavailable(
            "Paket xlwings belum terpasang. Jalankan 'python -m pip install xlwings' di environment backend."
        ) from exc
    try:
        app = xw.App(visible=False, add_book=False)
    except Exception as exc:
        raise ExcelUnavailable(
            "Microsoft Excel desktop tidak dapat dijalankan. Pastikan Excel terpasang dan dapat dibuka."
        ) from exc
    saved: dict[str, Any] = {}
    blank = None
    workbook = None
    try:
        api = app.api
        blank = app.books.add()
        for key in ("ScreenUpdating", "DisplayAlerts", "EnableEvents", "Calculation", "CalculateBeforeSave"):
            try:
                saved[key] = getattr(api, key)
            except Exception:
                pass
        api.ScreenUpdating = False
        api.DisplayAlerts = False
        api.EnableEvents = False
        api.AutomationSecurity = 3  # disable macros while opening uploaded file
        api.Calculation = XL_CALC_MANUAL
        api.CalculateBeforeSave = False
        try:
            mt = api.MultiThreadedCalculation
            saved["MultiThreaded"] = mt.Enabled
            mt.Enabled = True
        except Exception:
            pass
        workbook = app.books.open(str(path), update_links=False, read_only=False)
        return app, blank, workbook, saved
    except Exception:
        try:
            if workbook is not None:
                workbook.close()
            if blank is not None:
                blank.close()
            app.quit()
        except Exception:
            try:
                app.kill()
            except Exception:
                pass
        raise


def _close_excel(app: Any, blank: Any, workbook: Any, saved: dict[str, Any]) -> None:
    if workbook is not None:
        try:
            workbook.close()
        except Exception:
            pass
    if app is not None:
        try:
            api = app.api
            if "MultiThreaded" in saved:
                try:
                    api.MultiThreadedCalculation.Enabled = saved["MultiThreaded"]
                except Exception:
                    pass
            for key in ("Calculation", "CalculateBeforeSave", "EnableEvents", "ScreenUpdating", "DisplayAlerts"):
                if key in saved:
                    try:
                        setattr(api, key, saved[key])
                    except Exception:
                        pass
            if blank is not None:
                try:
                    blank.close()
                except Exception:
                    pass
            app.quit()
        except Exception:
            try:
                app.kill()
            except Exception:
                pass


def _apply_plan_with_excel(
    source_path: Path,
    output_path: Path,
    input_bytes: bytes,
    plan: SCHWritePlan,
    column_map: dict[str, int],
    calculation_timeout_seconds: float = 3600.0,
) -> dict[str, Any]:
    app = blank = workbook = None
    saved: dict[str, Any] = {}
    original_calc_mode = _read_calc_mode(input_bytes)
    try:
        app, blank, workbook, saved = _open_excel_copy(source_path)
        ws = workbook.sheets[plan.sheet_name]
        max_col = max(max(column_map.values()), ws.used_range.last_cell.column)
        formula_columns = _verify_helper_formulas(ws, plan, max_col)

        if plan.end_row >= plan.start_row:
            # Reuse the first old Forecast row's formatting as a template. Only the payload runs are touched.
            template_row = plan.start_row if plan.start_row <= plan.old_last_row else max(plan.data_start_row, plan.start_row - 1)
            target_last = plan.clear_to_row
            if target_last >= plan.start_row:
                for run in plan.runs:
                    source = ws.range((template_row, run.first_column), (template_row, run.last_column))
                    target = ws.range((plan.start_row, run.first_column), (target_last, run.last_column))
                    source.api.Copy()
                    target.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
                    app.api.CutCopyMode = False
                    # Assign the template's number format explicitly as a safeguard for newly used rows.
                    for col in range(run.first_column, run.last_column + 1):
                        logical_offset = run.value_offset + col - run.first_column
                        logical_name = SCHEDULE_COLUMNS[logical_offset]
                        fmt = "General" if logical_name == "Item" else ws.range((template_row, col)).number_format
                        ws.range((plan.start_row, col), (target_last, col)).number_format = fmt

        prepared_rows, _ = _make_values_for_plan(ws, plan)
        if plan.rows:
            # Set Item to General so a legacy number format like 00000 does not reintroduce leading zeroes.
            item_col = column_map["Item"]
            ws.range((plan.start_row, item_col), (plan.end_row, item_col)).number_format = "General"
            for offset, next_offset in _iter_write_chunks(
                len(prepared_rows), plan.start_row, plan.old_last_row
            ):
                chunk = prepared_rows[offset:next_offset]
                chunk_start = plan.start_row + offset
                chunk_end = plan.start_row + next_offset - 1
                for run in plan.runs:
                    matrix = [row[run.value_offset:run.value_offset + run.width] for row in chunk]
                    target = ws.range(
                        (chunk_start, run.first_column),
                        (chunk_end, run.last_column),
                    )
                    target.options(ndim=2).value = matrix

                    # Read back the first, middle, and last row of each write chunk.
                    # This catches silent Excel COM truncation before the workbook is saved.
                    sample_indexes = sorted({0, len(chunk) // 2, len(chunk) - 1})
                    for sample_index in sample_indexes:
                        row_number = chunk_start + sample_index
                        actual = ws.range(
                            (row_number, run.first_column),
                            (row_number, run.last_column),
                        ).value
                        if not isinstance(actual, (list, tuple)):
                            actual = [actual]
                        expected = matrix[sample_index]
                        if len(actual) != len(expected) or any(
                            expected_value not in (None, "") and actual_value in (None, "")
                            for actual_value, expected_value in zip(actual, expected)
                        ):
                            raise SCHUpdateError(
                                f"Verifikasi penulisan Excel gagal pada baris {row_number}, "
                                f"kolom {_excel_column(run.first_column)}–{_excel_column(run.last_column)}. "
                                "Sebagian nilai tidak tersimpan; output dibatalkan."
                            )

            firm_end = plan.start_row + plan.firm_add_count - 1
            if plan.firm_add_count:
                for run in plan.runs:
                    ws.range((plan.start_row, run.first_column), (firm_end, run.last_column)).font.color = (0, 0, 0)
            forecast_start = firm_end + 1
            if plan.forecast_add_count:
                for run in plan.runs:
                    ws.range((forecast_start, run.first_column), (plan.end_row, run.last_column)).font.color = (255, 0, 0)

        # Clear any previous Forecast rows left below the new output. Only payload columns are cleared.
        clear_from = max(plan.start_row, plan.end_row + 1)
        if plan.clear_to_row >= clear_from:
            for run in plan.runs:
                target = ws.range((clear_from, run.first_column), (plan.clear_to_row, run.last_column))
                target.clear_contents()
                target.font.color = (0, 0, 0)

        # Recalculate once in the isolated Excel instance. Do not mark output final if calculation fails.
        started = time.perf_counter()
        app.api.Calculate()
        deadline = time.monotonic() + calculation_timeout_seconds
        while app.api.CalculationState != XL_CALC_DONE:
            if time.monotonic() > deadline:
                raise SCHUpdateError("Kalkulasi Excel melewati batas waktu. Tidak ada output final yang dibuat.")
            time.sleep(0.25)
        calculation_seconds = time.perf_counter() - started

        fmt = 52 if output_path.suffix.lower() == ".xlsm" else 51
        workbook.api.SaveAs(str(output_path.resolve()), FileFormat=fmt)
        workbook.close()
        workbook = None
        # Excel locks SaveAs output until workbook is closed; restore mode after close.
        _set_calc_mode(output_path, original_calc_mode)

        return {
            "calculated": True,
            "calculation_seconds": round(calculation_seconds, 2),
            "helper_formula_columns": [_excel_column(col) for col in sorted(formula_columns)],
            "start_row": plan.start_row,
            "end_row": plan.end_row,
            "firm_added": plan.firm_add_count,
            "forecast_added": plan.forecast_add_count,
        }
    except SCHUpdateError:
        raise
    except Exception as exc:
        raise SCHUpdateError(
            f"Gagal menulis workbook melalui Microsoft Excel ({type(exc).__name__}). "
            "File sumber tidak diubah."
        ) from exc
    finally:
        _close_excel(app, blank, workbook, saved)


def _validate_written_output(
    output_bytes: bytes,
    filename: str,
    original_rows: Sequence[Any],
    master_rows: Sequence[MasterRow],
    result: ReconciliationResult,
    plan: SCHWritePlan,
    formula_refs_before: dict[int, set[str]],
) -> dict[str, int]:
    output_rows, _, summary = read_schedule(output_bytes, filename)
    if summary["firm_rows"] != len(result.old_firm) + len(result.firm_to_add):
        raise SCHUpdateError(
            f"Validasi hasil gagal: Firm terbaca {summary['firm_rows']}, "
            f"seharusnya {len(result.old_firm) + len(result.firm_to_add)}."
        )
    if summary["forecast_rows"] != len(result.forecast_to_add):
        raise SCHUpdateError(
            f"Validasi hasil gagal: Forecast terbaca {summary['forecast_rows']}, "
            f"seharusnya {len(result.forecast_to_add)}."
        )
    if summary["ambiguous_rows"]:
        raise SCHUpdateError("Validasi hasil gagal: ditemukan baris dengan warna font ambigu.")

    before_firm = {row.row_no: row for row in result.old_firm}
    after_by_row = {row.row_no: row for row in output_rows}
    for row_no, old in before_firm.items():
        current = after_by_row.get(row_no)
        if current is None or current.row_type != RowType.FIRM or current.values != old.values:
            raise SCHUpdateError(f"Validasi hasil gagal: Firm lama pada baris {row_no} berubah atau hilang.")

    additions = [*result.firm_to_add, *result.forecast_to_add]
    output_additions = [row for row in output_rows if row.row_no >= plan.start_row]
    if len(output_additions) != len(additions):
        raise SCHUpdateError("Validasi hasil gagal: jumlah baris tambahan tidak sesuai rencana.")
    for index, (actual, expected) in enumerate(zip(output_additions, additions), start=0):
        expected_type = RowType.FORECAST if index >= plan.firm_add_count else RowType.FIRM
        if actual.row_type != expected_type:
            raise SCHUpdateError(f"Validasi hasil gagal: warna/status baris tambahan ke-{index + 1} tidak sesuai.")
        for column in SCHEDULE_COLUMNS:
            if actual.values.get(column, "") != expected.values.get(column, ""):
                raise SCHUpdateError(
                    f"Validasi hasil gagal: nilai {column} pada baris output {actual.row_no} tidak sesuai master."
                )

    formula_rows = set(formula_refs_before)
    if formula_rows:
        formula_refs_after = _formula_refs_at_rows(output_bytes, plan.sheet_name, formula_rows)
        for row_no in formula_rows:
            missing = formula_refs_before[row_no] - formula_refs_after.get(row_no, set())
            if missing:
                raise SCHUpdateError(
                    f"Validasi hasil gagal: formula helper hilang pada baris {row_no}: "
                    + ", ".join(sorted(missing)[:15])
                )
    return {
        "data_rows": summary["data_rows"],
        "firm_rows": summary["firm_rows"],
        "forecast_rows": summary["forecast_rows"],
        "ambiguous_rows": summary["ambiguous_rows"],
    }


def create_update_zip(
    master_content: bytes,
    master_filename: str,
    schedule_inputs: dict[str, tuple[str, bytes]],
    calculation_timeout_seconds: float = 3600.0,
) -> tuple[Path, Path, dict[str, Any]]:
    """Create updated copies for the selected one-to-twenty workbooks in one ZIP.

    Only supplied workbooks are processed; unselected workbooks are not required
    and are never read or changed.
    """
    if not schedule_inputs:
        raise SCHUpdateError("Pilih minimal satu workbook schedule untuk di-update.")
    if len(schedule_inputs) > MAX_SCHEDULE_FILES:
        raise SCHUpdateError(f"Maksimal {MAX_SCHEDULE_FILES} workbook schedule dapat di-update sekaligus.")
    work_dir = Path(tempfile.mkdtemp(prefix="fss_sch_update_"))
    zip_path = work_dir / "Update_SCH_Data.zip"
    try:
        master_rows, _ = read_master(master_content, master_filename)
        if any(row.ambiguous for row in master_rows):
            raise SCHUpdateError("Data FSS master mengandung baris dengan klasifikasi Firm/Forecast ambigu.")

        manifest: dict[str, Any] = {
            "status": "succeeded",
            "master_file": Path(master_filename).name,
            "total_files": len(schedule_inputs),
            "files": [],
            "note": "File sumber tidak ditimpa. Hanya workbook yang dipilih yang diproses sebagai salinan baru.",
        }
        for index, (role, (original_name, content)) in enumerate(schedule_inputs.items(), start=1):
            safe_name = Path(original_name).name
            suffix = Path(safe_name).suffix.lower()
            if suffix not in {".xlsx", ".xlsm"}:
                raise SCHUpdateError(f"{safe_name}: hanya format XLSX/XLSM yang didukung.")
            old_rows, column_map, summary = read_schedule(content, safe_name)
            result = reconcile(old_rows, master_rows)
            if result.blocking or result.review:
                raise SCHUpdateError(
                    f"{safe_name}: rekonsiliasi belum aman untuk ditulis. "
                    f"Pemblokir={len(result.blocking)}, perlu diperiksa={len(result.review)}."
                )
            plan = build_write_plan(old_rows, column_map, summary["header_row"], result)
            # Capture formulas at the first appended row and final row before opening Excel.
            formula_rows = {plan.data_start_row, plan.end_row} if plan.end_row >= plan.start_row else {plan.data_start_row}
            formulas_before = _formula_refs_at_rows(content, summary["sheet_name"], formula_rows)
            formula_check_end = max(plan.end_row, plan.clear_to_row)
            helper_formula_columns = _check_formula_coverage(
                content,
                summary["sheet_name"],
                plan.start_row,
                formula_check_end,
                formulas_before.get(plan.data_start_row, set()),
            )

            input_path = work_dir / f"input_{index:02d}_{role}{suffix}"
            input_path.write_bytes(content)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"{index:02d}_{Path(safe_name).stem}_UPDATED_{timestamp}{suffix}"
            output_path = work_dir / output_name
            excel_summary = _apply_plan_with_excel(
                input_path, output_path, content, plan, column_map, calculation_timeout_seconds
            )
            if not output_path.exists() or output_path.stat().st_size == 0:
                raise SCHUpdateError(f"{safe_name}: Excel tidak menghasilkan file output.")
            output_bytes = output_path.read_bytes()
            validation_summary = _validate_written_output(
                output_bytes, safe_name, old_rows, master_rows, result, plan, formulas_before
            )
            manifest["files"].append({
                "index": index,
                "role": role,
                "source": safe_name,
                "output": output_name,
                "sheet": summary["sheet_name"],
                "start_row": plan.start_row,
                "end_row": plan.end_row,
                "firm_kept": len(result.old_firm),
                "firm_added_including_transitions": len(result.firm_to_add),
                "firm_transitions": len(result.transitions),
                "firm_added_non_transition": max(0, len(result.firm_to_add) - len(result.transitions)),
                "forecast_added": len(result.forecast_to_add),
                "forecast_removed": len(result.old_forecast),
                "helper_formula_columns_checked": helper_formula_columns,
                "helper_formula_range_checked": [plan.start_row, formula_check_end],
                "calculation": excel_summary,
                "validation": validation_summary,
            })

        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=3) as archive:
            for item in manifest["files"]:
                archive.write(work_dir / item["output"], arcname=item["output"])
            archive.writestr("manifest.json", json.dumps(manifest, indent=2, ensure_ascii=False))
        return zip_path, work_dir, manifest
    except Exception:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise
