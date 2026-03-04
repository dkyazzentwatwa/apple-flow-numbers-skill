---
name: apple-flow-numbers
description: General Apple Numbers automation for Apple Flow. Use when working with `.numbers` files through the local `scripts/numbers_tools.py` CLI, including creating new sheets, appending structured rows, choosing insertion behavior (`after-data`, `after-headers`, `at-end`), and validating/debugging row placement with read-back checks.
---

# Apple Flow Numbers

Use this skill to reliably create and update Apple Numbers documents with:
- `python3 scripts/numbers_tools.py numbers_create`
- `python3 scripts/numbers_tools.py numbers_create_workbook`
- `python3 scripts/numbers_tools.py numbers_add_sheet`
- `python3 scripts/numbers_tools.py numbers_append_rows`
- `python3 scripts/numbers_tools.py numbers_style_apply`

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

1. Create a Numbers file with headers:
```bash
python3 scripts/numbers_tools.py numbers_create \
  "/abs/path/tracker.numbers" \
  '["Date","Item","Category","Amount","Notes"]' \
  --sheet "Sheet 1" \
  --table "Table 1" \
  --overwrite true
```

2. Append rows (recommended default `after-data`):
```bash
python3 scripts/numbers_tools.py numbers_append_rows \
  "/abs/path/tracker.numbers" \
  '[["2026-03-04","Coffee","Food",15,"Morning"],["2026-03-04","Burger","Food",30,"Lunch"]]' \
  --sheet "Sheet 1" \
  --table "Table 1" \
  --position after-data
```

3. Verify insertion response:
- Expect JSON with `"ok": true`
- Check `start_row` and `insert_after_row`

4. Apply formatting/style:
```bash
python3 scripts/numbers_tools.py numbers_style_apply \
  "/abs/path/tracker.numbers" \
  '{"scope":"range","start_row":2,"end_row":20,"start_column":1,"end_column":5}' \
  '{"background_color":[255,245,230],"font_size":12,"alignment":"center","row_height":28,"column_width":160}'
```

5. Build a full workbook (multiple sheets):
```bash
python3 scripts/numbers_tools.py numbers_create_workbook \
  "/abs/path/workbook.numbers" \
  '{"sheets":[{"sheet_name":"Transactions","table_name":"Tx","headers":["Date","Item","Amount"],"rows":[["2026-03-04","Coffee",15]]},{"sheet_name":"Summary","table_name":"Summary","headers":["Metric","Value"],"rows":[["Total",15]]}]}' \
  --overwrite true
```

6. Add one more sheet to an existing workbook:
```bash
python3 scripts/numbers_tools.py numbers_add_sheet \
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
osascript <<'APPLESCRIPT'
set p to POSIX file "/abs/path/tracker.numbers"
tell application id "com.apple.iWork.Numbers"
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
  - Ensure Apple Numbers is installed (`com.apple.Numbers`; older installs may use `com.apple.iWork.Numbers`) and automation permissions are granted.
  - `numbers_tools.py` probes multiple targets (`com.apple.Numbers`, `com.apple.iWork.Numbers`, and Creator Studio path/name) and picks the first scriptable target.
  - Retry command outside restrictive sandbox context when needed.
- Rows appear too far down:
  - Use `--position after-data` and verify table has expected headers/data.

## Done Criteria

- Tool command returns `"ok": true`.
- Inserted row range is sensible (`start_row`, `insert_after_row`).
- Read-back confirms expected top and tail data placement.
