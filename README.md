# Excel Parameter Sequencing for TestIng Net

Starter project for a Windows tool that will read CAN Rx/Tx signal information from an Excel workbook, validate it, and generate test-parameter `.par` files from a supplied template parameter file.

## Planned scope

- Extract variants and CAN Rx/Tx data from Excel to create test parameters while preserving the structure of a template parameter file.
- Generate corresponding test specs.
- Perform low-level consistency checks, including missing values, invalid formats, unknown or duplicate signals, and firmware-name-to-DTC mapping issues.
- Perform high-level checks of each variant's Value-Added Function (VAF) disabling strategy.

## Current status

The repository now includes a starter, Windows-only Excel-to-PAR generation pipeline. It reads Excel through `pywin32`, validates template placeholders against a YAML mapping, and creates one `.par` file for every Variant × Signal × invalid-value test case.

## Planned layout

```text
.
├── data/
│   ├── input/                 # User-provided workbooks (ignored by Git)
│   └── output/                # Generated .par files (ignored by Git)
├── examples/
│   └── can_signal_input_example.xlsx
├── config/
│   ├── keyword_mapping.yaml
├── scripts/
│   ├── detect_template_keywords.py
│   └── create_parameters.py
├── src/parameter_generator/
│   └── excel_reader.py
├── docs/                      # Future data-contract and rule documentation
├── requirements.txt
└── README.md
```

## Example workbook

`examples/can_signal_input_example.xlsx` is a small, illustrative input contract. It has:

- `Variant`: variant identifiers, codes, and available/unavailable functions.
- `Signals`: CAN frame, signal, precondition, invalid-value, firmware, debounce, disablement, and non-affected-variant fields.
- `Firmware_DTC`: an FW-to-DTC mapping table.
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

## Generate parameter files

Run the complete example with one command:

```powershell
.\.venv\Scripts\python.exe scripts\create_parameters.py
```

It generates one reusable parameter file and its corresponding test list in `data/output`.

To validate placeholders separately:

```powershell
.\.venv\Scripts\python.exe scripts\detect_template_keywords.py examples/parameter_template.par
```

To override the example paths:

```powershell
.\.venv\Scripts\python.exe scripts\create_parameters.py <workbook.xlsx> <template.par>
```

## Benchmark build time

Run the benchmark scenarios, including the 2,000-test large scenario:

```powershell
.\.venv\Scripts\python.exe scripts\benchmark_build_time.py
```

It creates the Markdown and CSV reports plus an SVG bar chart in `reports/`.

### Latest example results

| Scenario | Variants | Signals | Invalid values | Disabled functions | Test cases | Build time |
| -------- | -------: | ------: | -------------: | -----------------: | ---------: | ---------: |
| Small    |        2 |      10 |              1 |                  1 |         20 |    1.765 s |
| Medium   |       10 |      50 |              2 |                  2 |        500 |    1.715 s |
| Large    |       20 |     100 |              3 |                  2 |      2,000 |    1.741 s |

These measurements include Excel workbook reading and generation of the final parameter file and test list on the benchmark machine. Re-run the benchmark to measure the current environment.

The mapping defines a Variant × Signals test set; separate invalid values in the `invalid values` cell with `,`. The `preconditions` cell is an ordered list of `signal,value` pairs separated by `;`, which resolves to `set_CAN_signal` commands through `XXXPRECONDITIONSXXX`. `net_logic.function_test_outcome` selects `NoMonitoring` when the Variant code equals the signal's non-affected Variant, and `Monitoring` otherwise.

## Open design decisions

- Final Excel schema, mandatory fields, and allowed value formats.
- Authoritative CAN-signal catalog and duplicate-signal policy.
- Firmware name and DTC mapping source and ownership.
- `.par` template placeholders, section rules, and output naming convention.
- Variant-level VAF disabling rules and exception policy.
