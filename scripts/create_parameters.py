#!/usr/bin/env python
"""Create reusable CAN test parameters and a Variant-ordered test list."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from parameter_generator.excel_reader import read_workbook

KEYWORD_PATTERN = re.compile(r"XXX[A-Z][A-Z0-9_]*XXX")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="Excel input workbook")
    parser.add_argument("template", type=Path, help="PAR template containing XXX...XXX placeholders")
    parser.add_argument("--keyword-mapping", type=Path, default=Path("config/keyword_mapping.yaml"))
    parser.add_argument("--output-dir", type=Path, default=Path("data/output"))
    return parser.parse_args()


def column_value(record: dict[str, str], column_number: int) -> str:
    return record.get(f"__column_{column_number}__", "")


def resolve_navigator(navigator: dict[str, object], variant: dict[str, str], signal: dict[str, str]) -> str:
    records = {"Variant": variant, "Signals": signal}
    return column_value(records[navigator["sheet"]], int(navigator["column_number"]))


def split_function_list(value: str, delimiter: str) -> list[str]:
    return [item.strip().strip('"') for item in value.split(delimiter) if item.strip().strip('"')]


def resolve_net_logic(
    variant: dict[str, str], signal: dict[str, str], config: dict[str, object], variants: list[dict[str, str]]
) -> dict[str, str]:
    """Resolve test type and expected disabled-function assignments."""
    rule = config["net_logic"]["function_test_outcome"]
    assignment_settings = rule["disabled_function_assignments"]
    disabled_functions = split_function_list(
        resolve_navigator(rule["signal_function"], variant, signal), assignment_settings["list_delimiter"]
    )
    if not disabled_functions:
        raise ValueError(f"Signals row {signal['__row_number__']} has no functions to disable.")

    non_affected_variant = resolve_navigator(rule["non_affected_variant"], variant, signal)
    if not non_affected_variant:
        inference = rule["non_affected_variant_inference"]
        candidate_column = int(inference["candidate_unavailable_functions"]["column_number"])
        candidates = [
            resolve_navigator(rule["variant_code"], candidate, signal)
            for candidate in variants
            if set(disabled_functions).issubset(
                set(split_function_list(column_value(candidate, candidate_column), inference["list_delimiter"]))
            )
        ]
        if not candidates and inference["no_candidate"] == "error":
            raise ValueError(f"Cannot infer a non-affected Variant for Signals row {signal['__row_number__']}.")
        if len(candidates) > 1 and inference["multiple_candidates"] == "error":
            raise ValueError(f"Ambiguous non-affected Variant for Signals row {signal['__row_number__']}: {', '.join(candidates)}.")
        non_affected_variant = candidates[0] if candidates else ""

    variant_code = resolve_navigator(rule["variant_code"], variant, signal)
    outcome_key = "non_affected_variant" if variant_code == non_affected_variant else "affected_variant"
    outcome = {
        key: str(value).format(signal_function=disabled_functions[0])
        for key, value in rule["outcomes"][outcome_key].items()
        if key != "condition"
    }
    configured_values = assignment_settings["values_in_order"]
    if len(disabled_functions) > len(configured_values):
        raise ValueError(f"Signals row {signal['__row_number__']} has more disabled functions than configured NET values.")
    outcome["disabled_function_assignments"] = assignment_settings["output_joiner"].join(
        assignment_settings["output_format"].format(function=function, value=configured_values[index])
        for index, function in enumerate(disabled_functions)
    )
    return outcome


def render_preconditions(signal: dict[str, str], config: dict[str, object]) -> str:
    settings = config["test_set"]["preconditions"]
    assignments = column_value(signal, int(settings["navigator"]["column_number"])).split(settings["assignment_delimiter"])
    if settings.get("trim_whitespace", False):
        assignments = [assignment.strip() for assignment in assignments]
    if settings["empty_value_policy"] == "error" and any(not assignment for assignment in assignments):
        raise ValueError(f"Signals row {signal['__row_number__']} contains an empty precondition assignment.")
    rendered = []
    for assignment in assignments:
        parts = assignment.split(settings["signal_value_delimiter"], maxsplit=1)
        if len(parts) != 2 or not all(part.strip() for part in parts):
            raise ValueError(f"Signals row {signal['__row_number__']} has invalid precondition '{assignment}'.")
        rendered.append(settings["output_format"].format(signal=parts[0].strip(), value=parts[1].strip()))
    return settings["output_joiner"].join(rendered)


def render_invalid_values(signal: dict[str, str], config: dict[str, object]) -> str:
    settings = config["test_set"]["invalid_values"]
    values = column_value(signal, int(settings["navigator"]["column_number"])).split(settings["delimiter"])
    if settings.get("trim_whitespace", False):
        values = [value.strip() for value in values]
    if settings["empty_value_policy"] == "error" and any(not value for value in values):
        raise ValueError(f"Signals row {signal['__row_number__']} contains an empty invalid value.")
    return settings["output_joiner"].join(settings["output_format"].format(value=value) for value in values)


def variant_test_id(variant: dict[str, str], config: dict[str, object]) -> str:
    specification = config["template_keywords"]["XXXVARIANT_TEST_IDXXX"]
    value = column_value(variant, int(specification["navigator"]["column_number"]))
    for old, new in specification.get("replace_characters", {}).items():
        value = value.replace(old, new)
    return value


def firmware_index(rows: list[dict[str, str]], lookup_column: int, return_column: int) -> dict[str, str]:
    index = {column_value(row, lookup_column): column_value(row, return_column) for row in rows}
    if len(index) != len(rows):
        raise ValueError("Firmware_DTC contains duplicate FW names.")
    return index


def keyword_values(
    variant: dict[str, str], signal: dict[str, str], preconditions: str, invalid_values: str,
    firmware_dtc: list[dict[str, str]], config: dict[str, object], variants: list[dict[str, str]],
) -> dict[str, str]:
    values: dict[str, str] = {}
    fw_index: dict[str, str] | None = None
    outcome: dict[str, str] | None = None
    for keyword, specification in config["template_keywords"].items():
        source_type = specification["source_type"]
        if source_type == "navigator":
            values[keyword] = resolve_navigator(specification["navigator"], variant, signal)
        elif source_type == "variant_test_id":
            values[keyword] = variant_test_id(variant, config)
        elif source_type in {"invalid_value", "invalid_values"}:
            values[keyword] = invalid_values
        elif source_type == "precondition":
            values[keyword] = preconditions
        elif source_type == "lookup":
            if fw_index is None:
                lookup = specification["lookup_navigator"]
                fw_index = firmware_index(firmware_dtc, int(lookup["column_number"]), int(specification["return_column_number"]))
            lookup_key = resolve_navigator(specification["source_navigator"], variant, signal)
            if lookup_key not in fw_index:
                raise ValueError(f"No Firmware_DTC entry for expected FW '{lookup_key}'.")
            values[keyword] = fw_index[lookup_key]
        elif source_type == "net_logic":
            outcome = outcome or resolve_net_logic(variant, signal, config, variants)
            values[keyword] = outcome[specification["output_field"]]
        else:
            raise ValueError(f"Unsupported source_type '{source_type}' for {keyword}")
    return values


def split_template_variations(template: str, config: dict[str, object]) -> tuple[str, dict[str, list[str]]]:
    marker_prefix = re.escape(config["template_variations"]["marker_prefix"])
    parts = re.split(rf"(?m)^{marker_prefix}([A-Za-z_]+)\s*$", template)
    change_template = parts[0].rstrip() + "\n"
    variations: dict[str, list[str]] = {}
    for index in range(1, len(parts), 2):
        variations.setdefault(parts[index], []).append(parts[index + 1].strip() + "\n")
    return change_template, variations


def render(template: str, replacements: dict[str, str]) -> str:
    content = template
    for keyword, value in replacements.items():
        content = content.replace(keyword, value)
    remaining = KEYWORD_PATTERN.findall(content)
    if remaining:
        raise ValueError(f"Unreplaced keyword(s): {', '.join(sorted(set(remaining)))}")
    return content


def test_id(test_case_type: str, signal: dict[str, str], variant: dict[str, str], config: dict[str, object]) -> str:
    signal_name = resolve_navigator(config["template_keywords"]["XXXCAN_SIGNALXXX"]["navigator"], variant, signal)
    return f"TC.{test_case_type}_{signal_name}.{variant_test_id(variant, config)}"


def main() -> int:
    args = parse_args()
    template = args.template.read_text(encoding="utf-8")
    config = yaml.safe_load(args.keyword_mapping.read_text(encoding="utf-8"))
    start_rows = {settings["sheet_name"]: int(settings["start_row"]) for settings in config["source_navigators"].values()}
    workbook = read_workbook(args.workbook, start_rows)
    variants, signals, firmware_dtc = workbook["Variant"], workbook["Signals"], workbook["Firmware_DTC"]
    if not variants or not signals or not firmware_dtc:
        raise ValueError("Variant, Signals, and Firmware_DTC must each contain at least one data row.")
    unmapped = set(KEYWORD_PATTERN.findall(template)) - set(config["template_keywords"])
    if unmapped:
        raise ValueError(f"Template contains unmapped keyword(s): {', '.join(sorted(unmapped))}")

    change_template, variation_blocks = split_template_variations(template, config)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    parameter_sections: list[str] = []
    test_list: list[str] = ["# Generated test list"]
    for variant in variants:
        change_replacements = keyword_values(variant, signals[0], "", "", firmware_dtc, config, variants)
        parameter_sections.append(render(change_template, change_replacements))
        test_list.append(f"TC.CHANGE_VARCODE_{variant_test_id(variant, config)}")
        for signal in signals:
            preconditions = render_preconditions(signal, config)
            invalid_values = render_invalid_values(signal, config)
            replacements = keyword_values(variant, signal, preconditions, invalid_values, firmware_dtc, config, variants)
            outcome = resolve_net_logic(variant, signal, config, variants)
            variation_name = config["template_variations"]["with_precondition"] if preconditions else config["template_variations"]["without_precondition"]
            blocks = variation_blocks.get(variation_name, [])
            matching = [block for block in blocks if f"[TC.{outcome['test_case_type']}_" in block]
            if len(matching) != 1:
                raise ValueError(f"Template needs exactly one {variation_name} block for {outcome['test_case_type']}.")
            parameter_sections.append(render(matching[0], replacements))
            test_list.append(test_id(outcome["test_case_type"], signal, variant, config))

    outputs = config["test_set"]["outputs"]
    parameter_path = args.output_dir / outputs["parameter_file"]
    test_list_path = args.output_dir / outputs["test_list_file"]
    parameter_path.write_text("\n".join(parameter_sections), encoding="utf-8", newline="\n")
    test_list_path.write_text("\n".join(test_list) + "\n", encoding="utf-8", newline="\n")
    print(f"Created {parameter_path}")
    print(f"Created {test_list_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
