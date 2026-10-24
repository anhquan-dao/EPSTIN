#!/usr/bin/env python
"""Benchmark parameter-file builds across Variant, Signal, value, and function counts."""
from __future__ import annotations
import csv
import subprocess
import sys
import time
from pathlib import Path
from xml.sax.saxutils import escape

PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = PROJECT_ROOT / "examples" / "benchmark_parameter_template.par"
MAPPING = PROJECT_ROOT / "config" / "keyword_mapping.yaml"
BENCHMARK_DIR = PROJECT_ROOT / "data" / "benchmark"
REPORT_DIR = PROJECT_ROOT / "reports"
SCENARIOS = (("Small", 2, 10, 1, 1), ("Medium", 10, 50, 2, 2), ("Large", 20, 100, 3, 2))

def create_workbook(path: Path, variants: int, signals: int, invalids: int, disabled_count: int) -> None:
    import win32com.client  # type: ignore[import-not-found]
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        variant = book.Worksheets(1); variant.Name = "Variant"
        signal = book.Worksheets.Add(After=variant); signal.Name = "Signals"
        firmware = book.Worksheets.Add(After=signal); firmware.Name = "Firmware_DTC"
        variant_rows = [("variant ID", "variant code", "available functions", "unavailable functions")]
        variant_rows.extend((f"VAR-{i:03d}", f"VARIANT_{i:03d}", "", "") for i in range(1, variants + 1))
        for row_number, row in enumerate(variant_rows, start=3):
            for column_number, value in enumerate(row, start=1):
                variant.Cells(row_number, column_number).Value = value
        signal_rows = [("CAN Frame ID", "CAN Bus", "Signal name", "Preconditions", "invalid values", "expected FW", "debounce", "functions will be disabled", "Non-affected variant", "REQ ID")]
        functions = ",".join(f"Function_{i}" for i in range(1, disabled_count + 1)); values = ",".join(str(i) for i in range(1, invalids + 1))
        signal_rows.extend((f"0x18FF{i:04X}", "1", f"Signal_{i:03d}", "PreconditionSignal,1", values, "FW_BENCH", "100;100", functions, f"VARIANT_{((i - 1) % variants) + 1:03d}", f"REQ-{i:03d}") for i in range(1, signals + 1))
        for row_number, row in enumerate(signal_rows, start=3):
            for column_number, value in enumerate(row, start=1):
                signal.Cells(row_number, column_number).Value = value
        for row_number, row in enumerate((("FW name", "FW DTC"), ("FW_BENCH", "B0001")), start=3):
            for column_number, value in enumerate(row, start=1):
                firmware.Cells(row_number, column_number).Value = value
        book.SaveAs(str(path.resolve()), FileFormat=51)
    finally:
        book.Close(SaveChanges=False); excel.Quit()

def write_svg(rows: list[dict[str, object]], path: Path) -> None:
    maximum = max(float(row["seconds"]) for row in rows) or 1; height = 105 + len(rows) * 64; parts = []
    for i, row in enumerate(rows):
        y = 55 + i * 64; width = 560 * float(row["seconds"]) / maximum
        parts.append(f'<text x="18" y="{y + 22}" font-family="Arial" font-size="15">{escape(str(row["scenario"]))} ({row["test_cases"]} tests)</text><rect x="280" y="{y}" width="{width:.1f}" height="30" fill="#1F4E78"/><text x="{290 + width:.1f}" y="{y + 22}" font-family="Arial" font-size="14">{float(row["seconds"]):.3f} s</text>')
    path.write_text(f'<svg xmlns="http://www.w3.org/2000/svg" width="920" height="{height}"><rect width="100%" height="100%" fill="white"/><text x="18" y="30" font-family="Arial" font-size="20" font-weight="bold">Parameter build time by scenario</text>{"".join(parts)}</svg>', encoding="utf-8")

def main() -> int:
    BENCHMARK_DIR.mkdir(parents=True, exist_ok=True); REPORT_DIR.mkdir(parents=True, exist_ok=True); rows = []
    for name, variants, signals, invalids, disabled_count in SCENARIOS:
        workbook = BENCHMARK_DIR / f"{name.lower()}_benchmark.xlsx"; output = BENCHMARK_DIR / f"{name.lower()}_output"
        create_workbook(workbook, variants, signals, invalids, disabled_count)
        started = time.perf_counter()
        subprocess.run([sys.executable, str(PROJECT_ROOT / "scripts" / "create_parameters.py"), str(workbook), str(TEMPLATE), "--keyword-mapping", str(MAPPING), "--output-dir", str(output)], check=True, cwd=PROJECT_ROOT)
        rows.append({"scenario": name, "variant_count": variants, "signal_count": signals, "invalid_value_count": invalids, "disabled_function_count": disabled_count, "test_cases": variants * signals, "seconds": round(time.perf_counter() - started, 3)})
    with (REPORT_DIR / "build_time_results.csv").open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=rows[0].keys()); writer.writeheader(); writer.writerows(rows)
    write_svg(rows, REPORT_DIR / "build_time_bar_chart.svg")
    lines = ["# Parameter build-time benchmark", "", "| Scenario | Variants | Signals | Invalid values | Disabled functions | Test cases | Build time |", "| --- | ---: | ---: | ---: | ---: | ---: | ---: |"]
    lines.extend(f"| {r['scenario']} | {r['variant_count']} | {r['signal_count']} | {r['invalid_value_count']} | {r['disabled_function_count']} | {r['test_cases']} | {float(r['seconds']):.3f} s |" for r in rows)
    lines.extend(("", "The large scenario creates 2,000 CAN-signal test cases in one parameter file.", "", "![Build time bar chart](build_time_bar_chart.svg)"))
    (REPORT_DIR / "build_time_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Created {REPORT_DIR / 'build_time_report.md'}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
