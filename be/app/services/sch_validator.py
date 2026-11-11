"""Read-only validator for the four SCH workbooks used by Update SCH Data.

The schedule sheets have 12 business-data fields. `Division` is a formula/helper
field, not the `Firm` classifier from the Data FSS master. Firm/Forecast status
is represented by font color in the schedule workbook and is handled separately
by the reconciliation reader.
"""
from __future__ import annotations

from datetime import date, datetime, time
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SHEET_NAME = "DATE"
HEADER_SCAN_LIMIT = 10
EMPTY_RUN_STOP = 5_000
PREVIEW_LIMIT = 5
MAX_SCAN_COLUMN = 50

# Data fields physically present in the DATE sheets. The master field
# `Delivery period` is expected to map to target `Delivery time` in the next
# reconciliation stage; this validator only validates the target schema.
SCHEDULE_DATA_HEADERS = [
    "Date",
    "Vendor code",
    "Vendor name",
    "Receiving Area",
    "Delivery date",
    "Delivery period",
    "Classification",
    "PO number",
    "Item",
    "Parts number",
    "Parts name",
    "Order Quantity",
]

# The following header is not payload data. It is a formula/helper output
# derived from Receiving Area, so it must be preserved and never overwritten.
OPTIONAL_FORMULA_HEADERS = ["Division"]

PREVIEW_HEADERS = [
    "Date",
    "Vendor code",
    "Receiving Area",
    "Delivery date",
    "Delivery period",
    "PO number",
    "Item",
    "Parts number",
    "Order Quantity",
]

HEADER_ALIASES = {
    "date": "Date",
    "vendorcode": "Vendor code",
    "vendorname": "Vendor name",
    "receivingarea": "Receiving Area",
    "recievingarea": "Receiving Area",  # typo present in source workbooks
    "deliverydate": "Delivery date",
    "deliveryperiod": "Delivery period",
    "classification": "Classification",
    "ponumber": "PO number",
    "ponum": "PO number",
    "item": "Item",
    "partsnumber": "Parts number",
    "partnumber": "Parts number",
    "partsname": "Parts name",
    "partname": "Parts name",
    "orderquantity": "Order Quantity",
    "orderqty": "Order Quantity",
    "division": "Division",
}

ROLE_LABELS = {
    "collect_sch": "Collect SCH",
    "delivery_sch": "MMKI Delivery SCH",
    "collect_ckd": "Collect CKD",
    "delivery_ckd": "CKD Delivery SCH",
}


class SCHValidationError(ValueError):
    """Workbook structure does not match the expected SCH layout."""


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return "".join(ch for ch in str(value).strip().casefold() if ch.isalnum())


def _canonical_header(value: Any) -> str | None:
    # Delivery workbooks also contain a lowercase helper header `date`.
    # Keep this distinct from the actual `Date` data header.
    if isinstance(value, str) and value.strip() == "date":
        return None
    return HEADER_ALIASES.get(_normalize_header(value))


def _json_value(value: Any) -> str | int | float | bool | None:
    if value is None:
        return None
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _find_header(ws) -> tuple[int, dict[str, int]]:
    """Find the header row and map only the 12 actual data fields."""
    best_row = 0
    best_map: dict[str, int] = {}
    best_duplicates: list[str] = []
    max_col = min(max(int(ws.max_column or 1), 1), MAX_SCAN_COLUMN)

    upper_row = min(int(ws.max_row or HEADER_SCAN_LIMIT), HEADER_SCAN_LIMIT)
    for row_number in range(1, upper_row + 1):
        values = next(
            ws.iter_rows(
                min_row=row_number,
                max_row=row_number,
                min_col=1,
                max_col=max_col,
                values_only=True,
            ),
            (),
        )
        mapping: dict[str, int] = {}
        duplicates: list[str] = []
        for col_number, value in enumerate(values, start=1):
            canonical = _canonical_header(value)
            if canonical is None or canonical not in SCHEDULE_DATA_HEADERS:
                continue
            if canonical in mapping:
                duplicates.append(canonical)
            else:
                mapping[canonical] = col_number

        if len(mapping) > len(best_map):
            best_row = row_number
            best_map = mapping
            best_duplicates = duplicates

    missing = [header for header in SCHEDULE_DATA_HEADERS if header not in best_map]
    if best_row == 0 or missing:
        raise SCHValidationError(
            "Header data wajib tidak lengkap pada sheet DATE. "
            + ("Belum ditemukan: " + ", ".join(missing) if missing else "Periksa 10 baris pertama.")
        )

    if best_duplicates:
        raise SCHValidationError(
            "Ada header data yang terdeteksi lebih dari sekali: "
            + ", ".join(sorted(set(best_duplicates)))
            + "."
        )

    actual_positions = [best_map[header] for header in SCHEDULE_DATA_HEADERS]
    if actual_positions != sorted(actual_positions):
        details = ", ".join(
            f"{header}={get_column_letter(best_map[header])}" for header in SCHEDULE_DATA_HEADERS
        )
        raise SCHValidationError(
            "Urutan header data tidak sesuai. Posisi ditemukan: " + details
        )

    return best_row, best_map


def inspect_sch_workbook(content: bytes, filename: str, role: str) -> dict[str, Any]:
    """Validate one workbook and preview payload without changing the source."""
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise SCHValidationError("Format file harus .xlsx atau .xlsm.")
    if not content:
        raise SCHValidationError("File kosong.")

    workbook = None
    try:
        workbook = load_workbook(
            BytesIO(content),
            read_only=True,
            data_only=False,  # preserve formula text for helper-column inspection
            keep_vba=(extension == ".xlsm"),
        )
        sheet_lookup = {name.casefold(): name for name in workbook.sheetnames}
        actual_sheet_name = sheet_lookup.get(SHEET_NAME.casefold())
        if actual_sheet_name is None:
            raise SCHValidationError(
                f"Sheet '{SHEET_NAME}' tidak ditemukan. Sheet tersedia: "
                + ", ".join(workbook.sheetnames)
            )

        ws = workbook[actual_sheet_name]
        worksheet_max_row = int(ws.max_row or 0)
        header_row, column_map = _find_header(ws)

        min_col = min(column_map.values())
        max_data_col = max(column_map.values())
        optional_helper_map: dict[str, int] = {}
        header_values = next(
            ws.iter_rows(
                min_row=header_row,
                max_row=header_row,
                min_col=1,
                max_col=min(max(int(ws.max_column or 1), 1), MAX_SCAN_COLUMN),
                values_only=True,
            ),
            (),
        )
        for col_number, value in enumerate(header_values, start=1):
            canonical = _canonical_header(value)
            if canonical in OPTIONAL_FORMULA_HEADERS:
                if canonical in optional_helper_map:
                    raise SCHValidationError(f"Header helper '{canonical}' ditemukan lebih dari sekali.")
                optional_helper_map[canonical] = col_number

        # Bound scan to relevant data/helper columns. No need to scan Excel's
        # full 1,048,576-row dimension for formula tails.
        scan_max_col = max_data_col
        index_map = {header: column - min_col for header, column in column_map.items()}

        data_rows = 0
        last_data_row = header_row
        blank_run = 0
        scanned_rows = 0
        preview: list[dict[str, Any]] = []
        for row_number, values in enumerate(
            ws.iter_rows(
                min_row=header_row + 1,
                min_col=min_col,
                max_col=scan_max_col,
                values_only=True,
            ),
            start=header_row + 1,
        ):
            scanned_rows += 1
            has_payload = any(values[index_map[header]] not in (None, "") for header in SCHEDULE_DATA_HEADERS)

            if not has_payload:
                blank_run += 1
                if blank_run >= EMPTY_RUN_STOP:
                    break
                continue

            blank_run = 0
            data_rows += 1
            last_data_row = row_number

            if len(preview) < PREVIEW_LIMIT:
                preview.append({
                    header: _json_value(values[index_map[header]])
                    for header in PREVIEW_HEADERS
                })

        if data_rows == 0:
            raise SCHValidationError("Sheet DATE ditemukan, tetapi tidak ada baris data yang terbaca.")

        # Inspect only a tiny sample of the helper formula columns. Do not scan
        # these columns down to the worksheet's max row: CKD templates may
        # pre-fill formulas through Excel's 1,048,576-row limit.
        helper_formula_samples: dict[str, dict[str, int]] = {}
        sample_end_row = min(last_data_row, header_row + PREVIEW_LIMIT)
        for helper_header, helper_column in optional_helper_map.items():
            found = 0
            sampled = 0
            for (helper_value,) in ws.iter_rows(
                min_row=header_row + 1,
                max_row=sample_end_row,
                min_col=helper_column,
                max_col=helper_column,
                values_only=True,
            ):
                sampled += 1
                if isinstance(helper_value, str) and helper_value.startswith("="):
                    found += 1
            helper_formula_samples[helper_header] = {
                "sampled_cells": sampled,
                "formula_cells_found": found,
            }

        column_positions = {
            header: get_column_letter(column_map[header])
            for header in SCHEDULE_DATA_HEADERS
        }
        helper_span_max = max([max_data_col, *optional_helper_map.values()])
        helper_columns = [
            get_column_letter(col)
            for col in range(min_col, helper_span_max + 1)
            if col not in set(column_map.values()) and col not in set(optional_helper_map.values())
        ]
        helper_header_positions = {
            header: get_column_letter(column)
            for header, column in optional_helper_map.items()
        }

        warnings: list[str] = []
        if header_row != 2:
            warnings.append(
                f"Header ditemukan pada baris {header_row}; template contoh menggunakan baris 2."
            )
        if worksheet_max_row > last_data_row + EMPTY_RUN_STOP:
            warnings.append(
                "Worksheet memiliki rentang tambahan setelah data aktif (kemungkinan formula/helper). "
                "Rentang ini tidak dihitung sebagai baris order."
            )
        if helper_columns:
            warnings.append(
                "Terdapat kolom sela/helper di antara field data: " + ", ".join(helper_columns) + "."
            )
        if helper_header_positions:
            warnings.append(
                "Kolom Division dikenali sebagai formula/helper dan tidak termasuk data yang akan ditimpa."
            )
        else:
            warnings.append(
                "Header Division tidak ditemukan; periksa apakah template ini memang tidak memakai kolom tersebut."
            )

        return {
            "role": role,
            "label": ROLE_LABELS.get(role, role),
            "filename": filename,
            "valid": True,
            "sheet_name": actual_sheet_name,
            "header_row": header_row,
            "data_start_row": header_row + 1,
            "data_rows": data_rows,
            "last_data_row": last_data_row,
            "worksheet_max_row": worksheet_max_row,
            "scanned_rows": scanned_rows,
            "column_positions": column_positions,
            "helper_columns_between_data_fields": helper_columns,
            "helper_header_positions": helper_header_positions,
            "helper_formula_samples": helper_formula_samples,
            "row_type_strategy": "font_color (black=Firm, red=Forecast); master column Firm is not copied into a worksheet field",
            "preview": preview,
            "warnings": warnings,
            "message": "Struktur 12 field data pada sheet DATE valid dan terpetakan berurutan.",
        }

    except SCHValidationError:
        raise
    except Exception as exc:
        raise SCHValidationError(
            f"Workbook tidak dapat diperiksa ({type(exc).__name__}). Pastikan file tidak rusak."
        ) from exc
    finally:
        if workbook is not None:
            workbook.close()
