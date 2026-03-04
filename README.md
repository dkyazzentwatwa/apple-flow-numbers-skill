# Apple Flow Numbers

> **Apple Numbers automation for AI agents.** Create, modify, and manage `.numbers` spreadsheets programmatically through an elegant CLI interface.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)](https://www.apple.com/macos/)

---

## Overview

**Apple Flow Numbers** is a standalone skill that enables AI agents to interact with Apple Numbers spreadsheets directly. Whether you need to create new workbooks, append data to existing sheets, or automate complex spreadsheet workflows, this tool provides a simple, reliable interface.

> Inspired by [apple-flow](https://github.com/dkyazzentwatwa/apple-flow) — a broader Apple ecosystem automation framework.

---

## Features

| Feature | Description |
|---------|-------------|
| **Create Workbooks** | Generate new `.numbers` files with custom sheets and headers |
| **Append Data** | Add structured rows with automatic type detection |
| **Flexible Insertion** | Choose where data goes: after headers, after existing data, or at the end |
| **Read & Verify** | Read back sheet contents to validate operations |
| **Standalone** | No dependencies — works independently as a skill or CLI tool |

---

## Installation

### As a Skill

```bash
./scripts/build_skill.sh
./scripts/install_skill.sh
```

### Direct CLI Usage

```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py --help
```

---

## Quick Start

### Create a New Workbook

```python
from skill.apple_flow_numbers.scripts.numbers_tools import create_workbook

create_workbook(
    "/path/to/workbook.numbers",
    sheet_name="Expenses",
    headers=["Date", "Category", "Amount", "Notes"]
)
```

### Append Rows

```python
from skill.apple_flow_numbers.scripts.numbers_tools import append_rows

append_rows(
    "/path/to/workbook.numbers",
    rows=[
        ["2024-03-01", "Office", 150.00, "Printer paper"],
        ["2024-03-02", "Travel", 45.50, "Taxi to airport"]
    ],
    insert_behavior="after-data"
)
```

### Convert Markdown Logs to Workbooks

Transform structured markdown files into multi-tab spreadsheets:

```bash
./scripts/md_log_to_numbers_workbook.py \
  --input /path/to/automation-log.md \
  --output /path/to/workbook.numbers \
  --sheets 3 \
  --rows-per-sheet 20 \
  --overwrite
```

---

## API Reference

### Core Functions

#### `create_workbook(path, sheet_name=None, headers=None)`

Creates a new Numbers workbook.

| Parameter | Type | Description |
|-----------|------|-------------|
| `path` | `str` | Output path for the `.numbers` file |
| `sheet_name` | `str` | Optional sheet name (defaults to "Sheet 1") |
| `headers` | `list` | Optional column headers |

#### `append_rows(path, rows, insert_behavior="after-data", target_sheet_name=None)`

Appends rows to an existing workbook.

| Parameter | Type | Description |
|-----------|------|-------------|
| `path` | `str` | Path to the `.numbers` file |
| `rows` | `list[list]` | 2D array of row data |
| `insert_behavior` | `str` | `"after-data"`, `"after-headers"`, or `"at-end"` |
| `target_sheet_name` | `str` | Optional target sheet name |

#### `read_sheet(path, rows=None, sheet_name=None)`

Reads contents from a sheet for verification.

| Parameter | Type | Description |
|-----------|------|-------------|
| `path` | `str` | Path to the `.numbers` file |
| `rows` | `int` | Number of rows to read |
| `sheet_name` | `str` | Optional sheet name |

---

## Project Structure

```
apple-flow-numbers-skill/
├── skill/
│   └── apple-flow-numbers/
│       ├── SKILL.md                    # Skill instructions & workflows
│       └── scripts/
│           └── numbers_tools.py        # Core Numbers automation library
├── scripts/
│   ├── build_skill.sh                  # Build distributable .skill file
│   ├── install_skill.sh                # Install to Codex skills directory
│   ├── verify_skill.sh                 # Validate skill structure
│   └── md_log_to_numbers_workbook.py   # Markdown → Numbers converter
└── dist/
    └── apple-flow-numbers.skill        # Built artifact (zip)
```

---

## Requirements

- **macOS** with Apple Numbers installed
- **Python** 3.8 or later

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Acknowledgments

This project was inspired by [apple-flow](https://github.com/dkyazzentwatwa/apple-flow), a comprehensive Apple ecosystem automation framework by [@dkyazzentwatwa](https://github.com/dkyazzentwatwa).
