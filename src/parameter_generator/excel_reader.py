"""Read Excel workbooks through Microsoft Excel via pywin32."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def _to_text(value: Any) -> str:
    """Convert an Excel cell value to the stable text used by the generator."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _rows_from_values(values: Any) -> list[tuple[Any, ...]]:
    """Normalize Excel's scalar, row, and matrix return shapes."""
    if values is None:
        return []
    if not isinstance(values, tuple):
        return [(values,)]
    if values and not isinstance(values[0], tuple):
        return [values]
    return list(values)


def read_workbook(workbook_path: Path, sheet_start_rows: dict[str, int]) -> dict[str, list[dict[str, str]]]:
    """Return requested sheets as column-number-keyed records.

    Microsoft Excel must be installed because the project deliberately uses
    pywin32 rather than a file-format-only reader.
    """
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as error:
        raise RuntimeError("pywin32 is not installed. Run: pip install -r requirements.txt") from error

    workbook_path = workbook_path.resolve()
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(workbook_path), ReadOnly=True)
        available = {sheet.Name for sheet in workbook.Worksheets}
        missing = set(sheet_start_rows) - available
        if missing:
            raise ValueError(f"Workbook is missing required sheet(s): {', '.join(sorted(missing))}")

        result: dict[str, list[dict[str, str]]] = {}
        for sheet_name, start_row in sheet_start_rows.items():
            worksheet = workbook.Worksheets(sheet_name)
            used_range = worksheet.UsedRange
            rows = _rows_from_values(used_range.Value)
            first_used_row = int(used_range.Row)
            first_data_index = start_row - first_used_row
            if first_data_index < 0:
                raise ValueError(f"Sheet '{sheet_name}' has no cells before configured start row {start_row}.")
            records: list[dict[str, str]] = []
            for row_number, row in enumerate(rows[first_data_index:], start=start_row):
                values = [_to_text(cell) for cell in row]
                if not any(values):
                    continue
                record = {f"__column_{index + 1}__": value for index, value in enumerate(values)}
                record["__row_number__"] = str(row_number)
                records.append(record)
            result[sheet_name] = records
        return result
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
