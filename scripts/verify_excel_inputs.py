#!/usr/bin/env python
"""Verify Excel CAN signals and disabled functions against selected source files."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Iterable

import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from parameter_generator.excel_reader import read_workbook

HASH_KEY_PATTERN = re.compile(r"(?:['\"]([^'\"]+)['\"]|\b([A-Za-z_][A-Za-z0-9_:.-]*)\b)\s*=>")
DBC_SIGNAL_PATTERN = re.compile(r"^\s*SG_\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.MULTILINE)
A2L_FUNCTION_PATTERN = re.compile(r"/begin\s+FUNCTION\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
A2L_SYMBOL_PATTERN = re.compile(r"/begin\s+(?:MEASUREMENT|CHARACTERISTIC)\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", type=Path, help="Input workbook; opens a picker when omitted")
    parser.add_argument("--signal-source", type=Path, help="DBC or Perl .pm signal source; opens a picker when omitted")
    parser.add_argument("--function-source", type=Path, help="A2L or Perl .pm function source; opens a picker when omitted")
    parser.add_argument("--mapping", type=Path, default=Path("config/keyword_mapping.yaml"))
    parser.add_argument("--report", type=Path, help="Optional JSON verification report")
    return parser.parse_args()


def pick_file(title: str, patterns: Iterable[tuple[str, str]]) -> Path:
    """Open a native file picker only when a command-line path was not given."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as error:
        raise RuntimeError("Provide the file path on the command line because Tk file selection is unavailable.") from error
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(title=title, filetypes=list(patterns))
    root.destroy()
    if not selected:
        raise RuntimeError(f"No file selected for {title}.")
    return Path(selected)


def parse_hash_keys(text: str) -> set[str]:
    return {quoted or bare for quoted, bare in HASH_KEY_PATTERN.findall(text)}


def parse_signals(path: Path) -> set[str]:
    text = path.read_text(encoding="utf-8", errors="replace")
    suffix = path.suffix.lower()
    if suffix == ".dbc":
        return set(DBC_SIGNAL_PATTERN.findall(text))
    if suffix == ".pm":
        return parse_hash_keys(text)
    raise ValueError("Signal source must be a .dbc or .pm file.")


def parse_functions(path: Path) -> set[str]:
    text = path.read_text(encoding="utf-8", errors="replace")
    suffix = path.suffix.lower()
    if suffix == ".a2l":
        return set(A2L_FUNCTION_PATTERN.findall(text)) | set(A2L_SYMBOL_PATTERN.findall(text))
    if suffix == ".pm":
        return parse_hash_keys(text)
    raise ValueError("Function source must be an .a2l or .pm file.")


def column(record: dict[str, str], number: int) -> str:
    return record.get(f"__column_{number}__", "").strip()


def split_list(value: str, delimiter: str) -> list[str]:
    return [item.strip().strip('"') for item in value.split(delimiter) if item.strip().strip('"')]


def mapping_errors(config: dict[str, object]) -> list[str]:
    errors: list[str] = []
    for name, settings in config.get("source_navigators", {}).items():
        if not settings.get("sheet_name") or int(settings.get("start_row", 0)) < 1:
            errors.append(f"source_navigators.{name} requires sheet_name and a positive start_row.")
    for keyword, specification in config.get("template_keywords", {}).items():
        for navigator_name in ("navigator", "source_navigator", "lookup_navigator"):
            navigator = specification.get(navigator_name)
            if navigator and (not navigator.get("sheet") or int(navigator.get("column_number", 0)) < 1 or int(navigator.get("start_row", 0)) < 1):
                errors.append(f"{keyword}.{navigator_name} requires sheet, positive column_number, and positive start_row.")
    return errors


def verify_records(workbook: dict[str, list[dict[str, str]]], config: dict[str, object], known_signals: set[str], known_functions: set[str]) -> list[str]:
    errors = mapping_errors(config)
    source_navigators = config["source_navigators"]
    variants = workbook.get(source_navigators["Variant"]["sheet_name"], [])
    signals = workbook.get(source_navigators["Signals"]["sheet_name"], [])
    if not variants:
        return errors + ["Variant has no data rows."]
    if not signals:
        return errors + ["Signals has no data rows."]

    rule = config["net_logic"]["function_test_outcome"]
    variant_code_column = int(rule["variant_code"]["column_number"])
    known_variant_codes = {column(row, variant_code_column) for row in variants if column(row, variant_code_column)}
    for row in variants:
        if not column(row, 1):
            errors.append(f"Variant row {row['__row_number__']}: missing variant ID.")
        if not column(row, variant_code_column):
            errors.append(f"Variant row {row['__row_number__']}: missing variant code.")

    invalid_settings = config["test_set"]["invalid_values"]
    precondition_settings = config["test_set"]["preconditions"]
    function_column = int(rule["signal_function"]["column_number"])
    non_affected_column = int(rule["non_affected_variant"]["column_number"])
    invalid_column = int(invalid_settings["navigator"]["column_number"])
    precondition_column = int(precondition_settings["navigator"]["column_number"])
    test_type_settings = config["test_type"]
    test_type_column = int(test_type_settings["navigator"]["column_number"])
    signal_column = int(config["template_keywords"]["XXXCAN_SIGNALXXX"]["navigator"]["column_number"])
    for row in signals:
        row_number = row["__row_number__"]
        signal_name = column(row, signal_column)
        if not signal_name:
            errors.append(f"Signals row {row_number}: missing signal name.")
        elif signal_name not in known_signals:
            errors.append(f"Signals row {row_number}: signal '{signal_name}' is absent from the selected signal source.")
        if not split_list(column(row, invalid_column), invalid_settings["delimiter"]):
            errors.append(f"Signals row {row_number}: no invalid values.")
        for function in split_list(column(row, function_column), rule["disabled_function_assignments"]["list_delimiter"]):
            if function not in known_functions:
                errors.append(f"Signals row {row_number}: function '{function}' is absent from the selected function source.")
        non_affected = column(row, non_affected_column)
        if non_affected and non_affected not in known_variant_codes:
            errors.append(f"Signals row {row_number}: non-affected Variant '{non_affected}' does not exist in Variant.")
        preconditions = column(row, precondition_column)
        if preconditions:
            for assignment in preconditions.split(precondition_settings["assignment_delimiter"]):
                if len(assignment.split(precondition_settings["signal_value_delimiter"], maxsplit=1)) != 2:
                    errors.append(f"Signals row {row_number}: invalid precondition '{assignment}'. Use signal,value;signal,value.")
        test_type = column(row, test_type_column)
        if test_type not in test_type_settings["input_mappings"]:
            errors.append(f"Signals row {row_number}: unsupported Test type '{test_type}'.")
    return errors


def main() -> int:
    args = parse_args()
    excel = args.excel or pick_file("Choose Excel input workbook", (("Excel workbook", "*.xlsx *.xlsm"),))
    signal_source = args.signal_source or pick_file("Choose DBC or Perl signal source", (("DBC or Perl", "*.dbc *.pm"),))
    function_source = args.function_source or pick_file("Choose A2L or Perl function source", (("A2L or Perl", "*.a2l *.pm"),))
    config = yaml.safe_load(args.mapping.read_text(encoding="utf-8"))
    start_rows = {settings["sheet_name"]: int(settings["start_row"]) for settings in config["source_navigators"].values()}
    errors = verify_records(read_workbook(excel, start_rows), config, parse_signals(signal_source), parse_functions(function_source))
    result = {"excel": str(excel), "signal_source": str(signal_source), "function_source": str(function_source), "valid": not errors, "errors": errors}
    if args.report:
        args.report.write_text(json.dumps(result, indent=2), encoding="utf-8")
    if errors:
        print("Verification failed:")
        print("\n".join(f"- {error}" for error in errors))
        return 2
    print("Verification passed: Excel values and YAML navigation match the selected signal and function sources.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
