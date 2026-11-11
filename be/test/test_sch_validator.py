from io import BytesIO
import pytest
from openpyxl import Workbook

from app.services.sch_validator import (
    SCHEDULE_DATA_HEADERS,
    SCHValidationError,
    inspect_sch_workbook,
)


def make_workbook(
    headers_by_column: dict[int, str],
    row_values: dict[int, str | int | float | bool | None],
) -> bytes:
    workbook = Workbook()
    ws = workbook.active
    assert ws is not None
    ws.title = "DATE"

    for column, header in headers_by_column.items():
        ws.cell(row=2, column=column, value=header)
    for column, value in row_values.items():
        ws.cell(row=3, column=column, value=value)

    # Division is a formula/helper, not a field in the FSS master data block.
    division_column = next((column for column, header in headers_by_column.items() if header == "Division"), None)
    if division_column is not None:
        ws.cell(row=3, column=division_column, value='=IF(F3="7R01","-KD","")')

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def collect_headers() -> dict[int, str]:
    # D:P with J intentionally skipped, mirroring Collect SCH/CKD.
    positions = [4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16]
    return {column: header for column, header in zip(positions, SCHEDULE_DATA_HEADERS)}


def test_valid_collect_layout_with_formula_helper() -> None:
    headers = collect_headers() | {22: "Division"}
    values = {4: 20260920, 5: "D400153", 7: "7R01", 8: 20260928, 9: 1, 12: "2004541206", 13: 10, 14: "5301H188", 16: 225}
    result = inspect_sch_workbook(make_workbook(headers, values), "collect.xlsx", "collect_sch")

    assert result["valid"] is True
    assert len(result["column_positions"]) == 12
    assert result["helper_header_positions"] == {"Division": "V"}
    assert result["helper_formula_samples"]["Division"]["formula_cells_found"] == 1
    assert result["preview"][0]["Order Quantity"] == 225
    assert "Division" not in result["preview"][0]


def test_valid_delivery_layout_with_lowercase_date_helper() -> None:
    headers = {2: "Date"}
    delivery_headers = [3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14]
    for column, header in zip(delivery_headers, SCHEDULE_DATA_HEADERS[1:]):
        headers[column] = header
    headers[18] = "date"  # lowercase helper header must not duplicate Date
    headers[19] = "Division"
    values = {2: 20260823, 3: "D400153", 5: "7R01", 6: 20260901, 7: 1, 10: "2004486106", 11: 10, 12: "5301H188", 14: 275}
    result = inspect_sch_workbook(make_workbook(headers, values), "delivery.xlsx", "delivery_sch")

    assert result["valid"] is True
    assert result["column_positions"]["Date"] == "B"
    assert result["column_positions"]["Delivery time"] == "G"
    assert result["helper_header_positions"] == {"Division": "S"}


def test_missing_required_data_header_is_rejected() -> None:
    headers = collect_headers() | {22: "Division"}
    headers.pop(16)  # Order Quantity
    with pytest.raises(SCHValidationError, match="Order Quantity"):
        inspect_sch_workbook(make_workbook(headers, {4: 20260920}), "bad.xlsx", "collect_sch")


