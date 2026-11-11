"""Build Data FSS service for the FSS Control Center web backend.

Adapted from the legacy desktop application's processing rules; no source file is modified.
"""
from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CANONICAL_COLUMNS = [
    "Date", "Vendor code", "Vendor name", "Receiving Area", "Delivery date",
    "Delivery period", "Firm", "Classification", "PO number", "Item",
    "Parts number", "Parts name", "Order Quantity",
]
REQUIRED_FIRM_COLUMNS = [c for c in CANONICAL_COLUMNS if c != "Firm"]
HEADER_ALIASES = {
    "date": "Date",
    "vendorcode": "Vendor code",
    "vendorname": "Vendor name",
    "receivingarea": "Receiving Area",
    "recievingarea": "Receiving Area",
    "recvarea": "Receiving Area",
    "deliverydate": "Delivery date",
    "deliveryperiod": "Delivery period",
    "firm": "Firm",
    "classification": "Classification",
    "ponumber": "PO number",
    "ponum": "PO number",
    "item": "Item",
    "partsnumber": "Parts number",
    "partnumber": "Parts number",
    "partsno": "Parts number",
    "partno": "Parts number",
    "partsname": "Parts name",
    "partname": "Parts name",
    "orderquantity": "Order Quantity",
    "orderqty": "Order Quantity",
}
OUTPUT_HEADERS = CANONICAL_COLUMNS.copy()
OUTPUT_HEADERS[3] = "Recieving Area"  # Legacy app's default/manual output profile.
HEADER_SCAN_LIMIT = 30
MIN_HEADER_MATCH = 8
MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024


class FSSBuildError(ValueError):
    """A user-safe validation or processing error."""
ParsedInput = tuple[
    list[dict[str, Any]],
    bool,
    int,
    str,
    list[str],
]

def _normal_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y%m%d")
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value).strip()


def _map_header(row: list[Any]) -> tuple[dict[int, str], list[str], list[str]]:
    mapping: dict[int, str] = {}
    unknown: list[str] = []
    duplicates: list[str] = []
    seen: set[str] = set()
    for index, raw in enumerate(row):
        label = _cell_text(raw)
        canonical = HEADER_ALIASES.get(_normal_header(label))
        if not label:
            continue
        if canonical is None:
            unknown.append(label)
            continue
        if canonical in seen:
            duplicates.append(canonical)
            continue
        seen.add(canonical)
        mapping[index] = canonical
    return mapping, unknown, duplicates


def _find_header(rows: list[list[Any]]) -> int | None:
    best_index: int | None = None
    best_score = 0
    for index, row in enumerate(rows[:HEADER_SCAN_LIMIT]):
        mapping, _, _ = _map_header(row)
        score = len(set(mapping.values()))
        if score > best_score:
            best_index, best_score = index, score
    return best_index if best_score >= MIN_HEADER_MATCH else None


def _read_tsv(name: str, raw: bytes) -> ParsedInput:
    if not raw.strip():
        raise FSSBuildError(f"{name}: file kosong.")
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        decoded = raw.decode("utf-16")
    else:
        decoded = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                decoded = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            raise FSSBuildError(f"{name}: encoding file tidak dikenali.")
    lines = [line for line in decoded.splitlines() if line.strip()]
    if not lines:
        raise FSSBuildError(f"{name}: file kosong.")
    probe = lines[0]
    delimiter = max(("\t", ";", ","), key=lambda d: probe.count(d))
    if probe.count(delimiter) == 0:
        raise FSSBuildError(f"{name}: pemisah kolom tidak dikenali; untuk saat ini gunakan TSV.")
    rows = [[cell.strip() for cell in row] for row in csv.reader(io.StringIO(decoded), delimiter=delimiter)]
    header_idx = _find_header(rows)
    if header_idx is None:
        raise FSSBuildError(f"{name}: header kolom FSS tidak ditemukan pada 30 baris pertama.")
    return _rows_to_frame_data(name, rows, header_idx, "(teks)")


def _rows_to_frame_data(
    name: str,
    rows: list[list[Any]],
    header_idx: int,
    sheet: str,
) -> ParsedInput:
    header = [_cell_text(v) for v in rows[header_idx]]
    while header and not header[-1]:
        header.pop()
    mapping, unknown, duplicates = _map_header(header)
    if duplicates:
        raise FSSBuildError(f"{name}: header ganda/ambigu: {', '.join(sorted(set(duplicates)))}.")
    missing = [col for col in REQUIRED_FIRM_COLUMNS if col not in mapping.values()]
    if missing:
        raise FSSBuildError(f"{name}: kolom wajib tidak ditemukan: {', '.join(missing)}.")
    is_forecast = "Firm" in mapping.values()
    width = len(header)
    records: list[dict[str, Any]] = []
    for source_row, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        if not any(_cell_text(v) for v in row):
            continue
        if len(row) > width and any(_cell_text(v) for v in row[width:]):
            raise FSSBuildError(
                f"{name}: baris {source_row} memiliki sel berisi di luar jumlah header; periksa kemungkinan data bergeser."
            )
        padded = list(row) + [""] * max(0, width - len(row))
        record: dict[str, Any] = {
            col: "" for col in CANONICAL_COLUMNS
            }
        for index, canonical in mapping.items():
            record[canonical] = _cell_text(padded[index])
        if not is_forecast:
            record["Firm"] = ""
        record["_source_file"] = name
        record["_source_row"] = source_row
        records.append(record)
    warnings = []
    if unknown:
        warnings.append(f"{name}: kolom tidak dikenal tidak dimasukkan ke output: {', '.join(unknown)}.")
    return records, is_forecast, header_idx + 1, sheet, warnings


def _read_excel(name: str, raw: bytes) -> ParsedInput:
    try:
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True, keep_vba=False)
    except Exception as exc:
        raise FSSBuildError(f"{name}: workbook tidak dapat dibuka; pastikan file Excel tidak rusak.") from exc
    candidates = []
    try:
        for ws in wb.worksheets:
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            rows = [[_cell_text(v) for v in row] for row in ws.iter_rows(values_only=True)]
            if not any(any(cell for cell in row) for row in rows):
                continue
            idx = _find_header(rows)
            if idx is None:
                continue
            try:
                parsed = _rows_to_frame_data(name, rows, idx, ws.title)
            except FSSBuildError:
                # If a header-like row is found but invalid, retain the diagnostic rather than silently selecting another sheet.
                continue
            candidates.append((len(parsed[0]), parsed))
    finally:
        wb.close()
    if not candidates:
        raise FSSBuildError(f"{name}: tidak ada worksheet dengan header FSS yang valid.")
    candidates.sort(key=lambda item: item[0], reverse=True)
    result = candidates[0][1]
    if len(candidates) > 1:
        result[4].append(f"{name}: lebih dari satu worksheet berisi tabel FSS; dipakai '{result[3]}' karena jumlah barisnya terbanyak.")
    return result


def _read_input(name: str, raw: bytes) -> ParsedInput:
    if len(raw) > MAX_FILE_SIZE_BYTES:
        raise FSSBuildError(f"{name}: ukuran file melebihi 50 MB.")
    ext = Path(name).suffix.lower()
    if ext == ".tsv":
        return _read_tsv(name, raw)
    if ext in (".xlsx", ".xlsm"):
        return _read_excel(name, raw)
    raise FSSBuildError(f"{name}: format belum didukung. Gunakan .tsv, .xlsx, atau .xlsm.")


def _number_or_text(raw: Any, label: str, warnings: list[str], file_name: str, source_row: int):
    text = _cell_text(raw)
    if text == "":
        return None
    try:
        number = float(text.replace(",", ""))
    except ValueError:
        warnings.append(f"{file_name} baris {source_row}: {label} bukan angka ('{text}'); nilai dipertahankan.")
        return text
    return int(number) if number.is_integer() else number


def _normalize_item(raw: Any, warnings: list[str], file_name: str, source_row: int):
    text = _cell_text(raw)
    if text == "":
        return None
    if text.isdigit():
        return int(text)
    warnings.append(f"{file_name} baris {source_row}: Item bukan angka ('{text}'); nilai dipertahankan.")
    return text


def _parse_sort_date(text: str):
    text = str(text or "").strip()
    if not text:
        return pd.NaT
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return pd.NaT


def build_fss(files: list[tuple[str, bytes]], release_date: str) -> tuple[bytes, dict[str, Any]]:
    """Build an .xlsx workbook from uploaded (filename, bytes) pairs.

    Returns (xlsx_bytes, summary). All processing happens in memory; input files are untouched.
    """
    if not files:
        raise FSSBuildError("Pilih minimal satu file Firm dan satu file Forecast.")
    if len(files) > 20:
        raise FSSBuildError("Maksimal 20 file dapat diproses sekaligus.")
    if not re.fullmatch(r"\d{8}", release_date or ""):
        raise FSSBuildError("Tanggal rilis harus berformat YYYYMMDD, misalnya 20261018.")
    try:
        datetime.strptime(release_date, "%Y%m%d")
    except ValueError as exc:
        raise FSSBuildError("Tanggal rilis tidak valid.") from exc

    firm_rows: list[dict[str, Any]] = []
    forecast_rows: list[dict[str, Any]] = []
    firm_files: list[str] = []
    forecast_files: list[str] = []
    warnings: list[str] = []
    for name, raw in files:
        records, is_forecast, header_row, sheet, file_warnings = _read_input(name, raw)
        warnings.extend(file_warnings)
        if is_forecast:
            forecast_files.append(name)
            forecast_rows.extend(records)
        else:
            firm_files.append(name)
            firm_rows.extend(records)
    if not firm_files:
        raise FSSBuildError("Tidak ditemukan file Firm (file tanpa kolom Firm).")
    if len(forecast_files) != 1:
        raise FSSBuildError(f"Harus ada tepat satu file Forecast; ditemukan {len(forecast_files)}.")

    # The legacy desktop app defaults: fill blank periods with 1 for both Firm and Forecast.
    filled_firm = filled_forecast = 0
    for row in firm_rows:
        if not str(row["Delivery period"]).strip():
            row["Delivery period"] = "1"
            filled_firm += 1
    for row in forecast_rows:
        if not str(row["Delivery period"]).strip():
            row["Delivery period"] = "1"
            filled_forecast += 1

    # Stable sort Firm by Delivery date then numeric Delivery period, ascending (legacy default).
    for row in firm_rows:
        row["_sort_date"] = _parse_sort_date(row["Delivery date"])
        try:
            row["_sort_period"] = float(str(row["Delivery period"]).strip())
        except (TypeError, ValueError):
            row["_sort_period"] = float("nan")
        if pd.isna(row["_sort_date"]):
            warnings.append(f"{row['_source_file']} baris {row['_source_row']}: Delivery date tidak valid; baris dipertahankan dan diurutkan paling akhir.")
        if pd.isna(row["_sort_period"]):
            warnings.append(f"{row['_source_file']} baris {row['_source_row']}: Delivery period bukan angka; nilai dipertahankan.")
    firm_rows.sort(key=lambda r: (pd.isna(r["_sort_date"]), r["_sort_date"] if not pd.isna(r["_sort_date"]) else pd.Timestamp.max,
                                  pd.isna(r["_sort_period"]), r["_sort_period"] if not pd.isna(r["_sort_period"]) else float("inf")))

    # Legacy default removes Forecast rows only when the Firm column is exactly '*'.
    before_forecast = len(forecast_rows)
    removed_rows = [r for r in forecast_rows if str(r["Firm"]).strip() == "*"]
    forecast_rows = [r for r in forecast_rows if str(r["Firm"]).strip() != "*"]

    combined = firm_rows + forecast_rows
    for row in combined:
        row["Delivery period"] = _number_or_text(row["Delivery period"], "Delivery period", warnings,
                                                 row["_source_file"], row["_source_row"])
        row["Order Quantity"] = _number_or_text(row["Order Quantity"], "Order Quantity", warnings,
                                                row["_source_file"], row["_source_row"])
        row["Item"] = _normalize_item(row["Item"], warnings, row["_source_file"], row["_source_row"])
    # Exact duplicate rows are reported only; never removed automatically.
    seen: set[tuple[str, ...]] = set()
    duplicate_count = 0
    for row in combined:
        key = tuple(_cell_text(row.get(c, "")) for c in CANONICAL_COLUMNS)
        if key in seen:
            duplicate_count += 1
        seen.add(key)
    if duplicate_count:
        warnings.append(f"Ditemukan {duplicate_count} baris duplikat identik; tidak ada yang dihapus otomatis.")

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data FSS")
    ws.freeze_panes = "A2"
    widths = [11, 12, 28, 14, 14, 15, 7, 14, 13, 8, 15, 36, 15]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    header_font = Font(name="Calibri", bold=True, color="FFFFFF")
    firm_font = Font(name="Calibri", color="202124")
    forecast_font = Font(name="Calibri", color="FF0000")
    header_fill = PatternFill("solid", fgColor="E87524")
    header_cells = []
    for label in OUTPUT_HEADERS:
        cell = WriteOnlyCell(ws, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")
        header_cells.append(cell)
    ws.append(header_cells)
    for row_index, row in enumerate(combined):
        is_forecast = row_index >= len(firm_rows)
        font = forecast_font if is_forecast else firm_font
        cells = []
        for col_index, col_name in enumerate(CANONICAL_COLUMNS):
            value = row.get(col_name)
            if value is None:
                value = ""
            cell = WriteOnlyCell(ws, value=value)
            cell.font = font
            if col_name in ("Delivery period", "Order Quantity") and isinstance(value, (int, float)):
                cell.number_format = "0"
            cells.append(cell)
        ws.append(cells)
    ws.auto_filter.ref = f"A1:M{len(combined) + 1}"
    output = io.BytesIO()
    wb.save(output)
    workbook_bytes = output.getvalue()

    # Re-open the result and verify the most important output invariants before returning it.
    check_wb = load_workbook(io.BytesIO(workbook_bytes), read_only=True, data_only=True)
    try:
        check_ws = check_wb["Data FSS"]
        actual_header = [cell.value for cell in next(check_ws.iter_rows(min_row=1, max_row=1))]
        if actual_header != OUTPUT_HEADERS:
            raise FSSBuildError("Validasi output gagal: header hasil tidak sesuai skema 13 kolom.")
        if len(actual_header) != 13:
            raise FSSBuildError("Validasi output gagal: jumlah kolom bukan 13.")
    finally:
        check_wb.close()

    summary = {
        "filename": f"{release_date}.xlsx",
        "firm_files": firm_files,
        "forecast_file": forecast_files[0],
        "firm_rows": len(firm_rows),
        "forecast_rows_before_filter": before_forecast,
        "forecast_rows_removed": len(removed_rows),
        "forecast_rows_after_filter": len(forecast_rows),
        "total_rows": len(combined),
        "firm_periods_filled": filled_firm,
        "forecast_periods_filled": filled_forecast,
        "duplicate_rows_warned": duplicate_count,
        "warnings": warnings[:100],
        "warnings_truncated": max(0, len(warnings) - 100),
        "removed_forecast_rows": [
            {"file": r["_source_file"], "row": r["_source_row"],
             "receiving_area": r["Receiving Area"], "delivery_date": r["Delivery date"],
             "parts_number": r["Parts number"], "order_quantity": r["Order Quantity"]}
            for r in removed_rows[:5000]
        ],
        "removed_rows_truncated": max(0, len(removed_rows) - 5000),
    }
    return workbook_bytes, summary
