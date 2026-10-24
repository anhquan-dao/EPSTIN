# CAN Signal Parameter File Generator

Starter project for a Windows tool that will read CAN Rx/Tx signal information from an Excel workbook, validate it, and generate test-parameter `.par` files from a supplied template parameter file.

## Planned scope

- Extract variants and CAN Rx/Tx test parameters from Excel.
- Generate one or more `.par` files while preserving the structure of a template parameter file.
- Perform low-level consistency checks, including missing values, invalid formats, unknown or duplicate signals, and firmware-name-to-DTC mapping issues.
- Perform high-level checks of each variant's Value-Added Function (VAF) disabling strategy.

## Current status

This repository is intentionally only a project skeleton. It contains no Python application scripts, parser, validation rules, or executable configuration yet.

## Planned layout

```text
.
├── data/
│   ├── input/                 # User-provided workbooks (ignored by Git)
│   └── output/                # Generated .par files (ignored by Git)
├── examples/
│   └── can_signal_input_example.xlsx
├── docs/                      # Future data-contract and rule documentation
├── requirements.txt
└── README.md
```

## Example workbook

`examples/can_signal_input_example.xlsx` is a small, illustrative input contract. It has:

- `Signals`: sample Rx and Tx signal records grouped by variant.
- `Firmware_DTC`: an example firmware-to-DTC mapping table.
- `VAF_Strategy`: example expectations for VAF disabling by variant.
- `Instructions`: brief field definitions and conventions.

The workbook is a discussion aid, not a final schema. Its headers and rules can be revised before implementation.

## Intended environment

- Windows and a locally installed Microsoft Excel instance.
- Python virtual environment.
- `pywin32` for Excel automation.
- `PyInstaller` for packaging the future application as an executable.

## Setup (after implementation starts)

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Open design decisions

- Final Excel schema, mandatory fields, and allowed value formats.
- Authoritative CAN-signal catalog and duplicate-signal policy.
- Firmware name and DTC mapping source and ownership.
- `.par` template placeholders, section rules, and output naming convention.
- Variant-level VAF disabling rules and exception policy.
