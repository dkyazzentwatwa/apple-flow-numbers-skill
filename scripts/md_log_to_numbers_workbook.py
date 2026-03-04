#!/usr/bin/env python3
"""Convert automation-log markdown into a Numbers workbook via local numbers_tools."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


def parse_runs(md_path: Path) -> list[list[str]]:
    lines = md_path.read_text(encoding="utf-8").splitlines()
    in_runs = False
    rows: list[list[str]] = []
    for line in lines:
        if line.strip() == "## Runs":
            in_runs = True
            continue
        if not in_runs:
            continue
        text = line.strip()
        if not text:
            continue
        if text.startswith("- "):
            text = text[2:].strip()
        if " | " not in text:
            continue
        parts = text.split(" | ", 4)
        if len(parts) < 5:
            parts += [""] * (5 - len(parts))
        rows.append(parts[:5])
    return rows


def chunk_rows(rows: list[list[str]], sheets: int, rows_per_sheet: int) -> list[list[list[str]]]:
    needed = sheets * rows_per_sheet
    selected = rows[:needed]
    while len(selected) < needed:
        selected.append(["", "", "", "", ""])
    return [selected[i : i + rows_per_sheet] for i in range(0, needed, rows_per_sheet)]


def build_spec(chunks: list[list[list[str]]], sheet_prefix: str) -> dict:
    headers = ["Timestamp", "Source", "Action", "Result", "Notes"]
    sheets: list[dict] = []
    for i, chunk in enumerate(chunks, start=1):
        sheets.append(
            {
                "sheet_name": f"{sheet_prefix} {i}",
                "table_name": f"AutomationLog{i}",
                "headers": headers,
                "rows": chunk,
            }
        )
    return {"sheets": sheets}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Path to automation-log markdown file")
    parser.add_argument("--output", required=True, help="Output .numbers file path")
    parser.add_argument("--sheets", type=int, default=3, help="Number of tabs/sheets")
    parser.add_argument("--rows-per-sheet", type=int, default=20, help="Rows per sheet")
    parser.add_argument("--sheet-prefix", default="Log", help="Sheet naming prefix")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite output if exists")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2
    if output_path.suffix.lower() != ".numbers":
        print("Output path must end with .numbers", file=sys.stderr)
        return 2

    parsed = parse_runs(input_path)
    chunks = chunk_rows(parsed, args.sheets, args.rows_per_sheet)
    spec = build_spec(chunks, args.sheet_prefix)

    tool_script = Path(__file__).resolve().parents[1] / "skill" / "apple-flow-numbers" / "scripts" / "numbers_tools.py"
    if not tool_script.exists():
        print(f"numbers_tools.py not found: {tool_script}", file=sys.stderr)
        return 2

    cmd = [
        sys.executable,
        str(tool_script),
        "numbers_create_workbook",
        str(output_path),
        json.dumps(spec, ensure_ascii=False),
        "--overwrite",
        "true" if args.overwrite else "false",
    ]

    print(f"Parsed rows: {len(parsed)}")
    print(f"Creating workbook: {output_path}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.returncode != 0 and result.stderr.strip():
        print(result.stderr.strip(), file=sys.stderr)
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
