# Apple Numbers Skill

Local AppleScript automation tools for `.numbers` files, designed for AI-agent workflows.

## What This Repo Provides

- A local CLI at `skill/apple-flow-numbers/scripts/numbers_tools.py`
- A skill definition at `skill/apple-flow-numbers/SKILL.md`
- Packaging scripts to build and distribute `dist/apple-flow-numbers.skill`

This repo does not require a global `apple-flow` binary.

## Requirements

- macOS
- Python 3.8+
- A scriptable Numbers app:
- `Numbers Creator Studio` (preferred)
- Legacy Apple Numbers fallback (`com.apple.Numbers`, `com.apple.iWork.Numbers`, `Numbers`)
- Automation permissions granted for your terminal app

## Quick Start

1. Run preflight:

```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight
```

Expected outcome: JSON containing `"ok": true`.

2. If preflight is not OK, set an explicit target and retry:

```bash
export NUMBERS_APP_TARGET='application "Numbers Creator Studio"'
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight
```

Alternative override by bundle id:

```bash
export NUMBERS_APP_BUNDLE_ID='your.bundle.id'
```

Optional custom app path discovery input:

```bash
export NUMBERS_CREATOR_STUDIO_APP="$HOME/Applications/Numbers Creator Studio.app"
```

3. Create a workbook:

```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create_workbook \
  "/abs/path/workbook.numbers" \
  '{"sheets":[{"sheet_name":"Tasks","table_name":"Tasks","headers":["Date","Task","Status"],"rows":[["2026-03-04","Draft plan","Active"]]}]}' \
  --overwrite true
```

## Target Resolution Order

`numbers_tools.py` resolves app targets in this order:

1. `NUMBERS_APP_TARGET`
2. `NUMBERS_APP_BUNDLE_ID`
3. Creator Studio app path/name/bundle-id candidates
4. Legacy fallbacks: `com.apple.Numbers`, `com.apple.iWork.Numbers`, `Numbers`

## Common Commands

```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create ...
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_create_workbook ...
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_add_sheet ...
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_append_rows ...
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_style_apply ...
```

## Build And Verify The Skill Artifact

Build:

```bash
./scripts/build_skill.sh
```

Verify:

```bash
./scripts/verify_skill.sh
```

Install locally to Codex skills directory:

```bash
./scripts/install_skill.sh
```

Artifact path:

- `dist/apple-flow-numbers.skill`

## Project Layout

```text
apple-flow-numbers-skill/
├── skill/
│   └── apple-flow-numbers/
│       ├── SKILL.md
│       └── scripts/
│           └── numbers_tools.py
├── scripts/
│   ├── build_skill.sh
│   ├── install_skill.sh
│   ├── verify_skill.sh
│   └── md_log_to_numbers_workbook.py
└── dist/
    └── apple-flow-numbers.skill
```
