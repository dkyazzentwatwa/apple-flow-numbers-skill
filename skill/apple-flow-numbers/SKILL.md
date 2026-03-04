---
name: apple-flow-numbers
description: General Apple Numbers automation with a local CLI for `.numbers` files. Use when creating workbooks, adding sheets, appending rows with insertion control (`after-data`, `after-headers`, `at-end`), styling ranges, and validating data placement.
---

# Apple Numbers Automation (Local CLI)

Use this skill to reliably create and update Apple Numbers documents with:
- `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight`
- `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create`
- `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create_workbook`
- `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_add_sheet`
- `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_append_rows`
- `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_style_apply`

Favor deterministic CLI workflows over ad hoc AppleScript. Use direct AppleScript only for read-back verification and debugging.

## Current Capability Snapshot

- Supports wide tables:
  - `numbers_create` auto-expands columns to fit all headers.
  - `numbers_create_workbook` builds multi-sheet files from one JSON spec.
  - `numbers_add_sheet` adds initialized sheets to existing workbooks.
  - `numbers_append_rows` auto-expands columns to fit the widest incoming row.
- Supports insertion modes:
  - `after-data`, `after-headers`, `at-end`
- Supports styling operations:
  - colors (`background_color`, `text_color`)
  - font (`font_name`, `font_size`)
  - alignment (`left|center|right|justified|natural`)
  - number format (`automatic|currency|percentage|scientific|fraction|text`)
  - wrapping (`text_wrap`)
  - dimensions (`row_height`, `column_width`)

## Quick Start

1. Run preflight and confirm `"ok": true`:
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight
```

2. Create a Numbers file with headers:
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create \
  "/abs/path/tracker.numbers" \
  '["Date","Item","Category","Amount","Notes"]' \
  --sheet "Sheet 1" \
  --table "Table 1" \
  --overwrite true
```

3. Append rows (recommended default `after-data`):
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_append_rows \
  "/abs/path/tracker.numbers" \
  '[["2026-03-04","Coffee","Food",15,"Morning"],["2026-03-04","Burger","Food",30,"Lunch"]]' \
  --sheet "Sheet 1" \
  --table "Table 1" \
  --position after-data
```

4. Verify insertion response:
- Expect JSON with `"ok": true`
- Check `start_row` and `insert_after_row`

5. Apply formatting/style:
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_style_apply \
  "/abs/path/tracker.numbers" \
  '{"scope":"range","start_row":2,"end_row":20,"start_column":1,"end_column":5}' \
  '{"background_color":[255,245,230],"font_size":12,"alignment":"center","row_height":28,"column_width":160}'
```

6. Build a full workbook (multiple sheets):
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create_workbook \
  "/abs/path/workbook.numbers" \
  '{"sheets":[{"sheet_name":"Transactions","table_name":"Tx","headers":["Date","Item","Amount"],"rows":[["2026-03-04","Coffee",15]]},{"sheet_name":"Summary","table_name":"Summary","headers":["Metric","Value"],"rows":[["Total",15]]}]}' \
  --overwrite true
```

7. Add one more sheet to an existing workbook:
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_add_sheet \
  "/abs/path/workbook.numbers" \
  '{"sheet_name":"Dashboard","table_name":"DashboardTable","headers":["Metric","Value"],"rows":[["Count",1]]}'
```

## Input Rules

- Always use an absolute path.
- File extension must be `.numbers`.
- `numbers_create` headers must be a JSON array of strings.
- `numbers_append_rows` payload must be a JSON array.
- `numbers_create_workbook` requires `{"sheets":[...]}` with unique `sheet_name` values.
- `numbers_add_sheet` requires a sheet JSON object with `sheet_name` and non-empty `headers`.
- Safest append shape: array-of-arrays (for example: `[[...],[...]]`).
- `numbers_style_apply` target/style args must be JSON objects.
- Style target indices are 1-based.

## Position Strategy

- `after-data`:
  - Best for logs and trackers.
  - Inserts right after the last non-empty data row.
  - Fills the top data region instead of jumping to visual bottom rows.
- `after-headers`:
  - Inserts at first data row.
  - Shifts existing data down.
- `at-end`:
  - Always appends to the physical end of the table.
  - Use when you explicitly want bottom append behavior.

## Wide-Column Imports

If a CSV has more than the default table width, import directly with full headers and rows. The tool will auto-add required columns before writing data.

## Standard Workflow

1. Define columns first.
2. Create or reuse target file.
3. Build rows as JSON.
4. Append with `--position after-data` unless user asks otherwise.
5. Verify first and last inserted rows.

## Read-Back Verification

Use this AppleScript probe after appending:
```bash
APP_TARGET_EXPR="${NUMBERS_APP_TARGET:-application \"Numbers Creator Studio\"}"
osascript <<APPLESCRIPT
set p to POSIX file "/abs/path/tracker.numbers"
tell ${APP_TARGET_EXPR}
  set d to open p
  set t to first table of first sheet of d
  tell t
    set firstRow to (value of cell 1 of row 2 as text) & "|" & (value of cell 2 of row 2 as text)
    set lastRow to (value of cell 1 of last row as text) & "|" & (value of cell 2 of last row as text)
  end tell
  close d saving no
  return firstRow & "\n" & lastRow
end tell
APPLESCRIPT
```

## Troubleshooting

- `absolute path required`:
  - Convert to absolute path before local `numbers_tools.py` call.
- `target document does not exist`:
  - Create with `numbers_create` first or confirm path typo.
- `Can't get sheet` or `Can't get table`:
  - Provide exact `--sheet` and `--table` names.
- `Connection invalid` / AppleScript runtime failures:
  - Run preflight first:
    - `python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight`
  - `numbers_tools.py` target resolution order is:
    - `NUMBERS_APP_TARGET` (direct AppleScript target override)
    - `NUMBERS_APP_BUNDLE_ID` (converted to `application id "..."`)
    - Numbers Creator Studio app path/name/bundle-id candidates
    - Legacy Apple Numbers fallbacks: `com.apple.Numbers`, `com.apple.iWork.Numbers`, `"Numbers"`
  - If needed, set one explicit target:
    - `export NUMBERS_APP_TARGET='application "Numbers Creator Studio"'`
    - `export NUMBERS_APP_BUNDLE_ID='your.bundle.id'`
  - Retry command outside restrictive sandbox context when needed.
- Rows appear too far down:
  - Use `--position after-data` and verify table has expected headers/data.

## Done Criteria

- Tool command returns `"ok": true`.
- Inserted row range is sensible (`start_row`, `insert_after_row`).
- Read-back confirms expected top and tail data placement.
