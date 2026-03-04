#!/usr/bin/env python3
"""Standalone Numbers tools for local AppleScript-driven workbook automation."""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

logger = logging.getLogger("numbers_tools")
DEFAULT_NUMBERS_APP_TARGETS = (
    'application "Numbers Creator Studio"',
    'application id "com.apple.NumbersCreatorStudio"',
    'application id "com.apple.Numbers"',
    'application id "com.apple.iWork.Numbers"',
    'application "Numbers"',
)


def _normalize_user_target(raw_target: str) -> str:
    """Normalize user input into a valid AppleScript application target."""
    target = raw_target.strip()
    if not target:
        return ""
    if target.startswith("application "):
        return target
    if target.endswith(".app"):
        return f'application "{target}"'
    if "." in target and " " not in target:
        return f'application id "{target}"'
    return f'application "{target}"'


def _read_bundle_id(app_path: Path) -> str:
    """Read CFBundleIdentifier for an app bundle path."""
    try:
        result = subprocess.run(
            ["mdls", "-name", "kMDItemCFBundleIdentifier", "-raw", str(app_path)],
            capture_output=True,
            text=True,
            timeout=5.0,
        )
        if result.returncode != 0:
            return ""
        bundle_id = (result.stdout or "").strip()
        if not bundle_id or bundle_id == "(null)":
            return ""
        return bundle_id
    except Exception:
        return ""


def _discover_creator_studio_targets() -> list[str]:
    """Build Creator Studio candidates from env + common install locations."""
    targets: list[str] = []
    explicit_app_path = os.getenv("NUMBERS_CREATOR_STUDIO_APP", "").strip()
    possible_paths = [
        Path("/Applications/Numbers Creator Studio.app"),
        Path.home() / "Applications" / "Numbers Creator Studio.app",
    ]
    if explicit_app_path:
        possible_paths.insert(0, Path(explicit_app_path).expanduser())

    for app_path in possible_paths:
        if not app_path.exists():
            continue
        targets.append(f'application "{str(app_path)}"')
        targets.append(f'application "{app_path.stem}"')
        bundle_id = _read_bundle_id(app_path)
        if bundle_id:
            targets.append(f'application id "{bundle_id}"')
    return targets


def _build_numbers_app_targets() -> tuple[str, ...]:
    """Assemble candidate targets with Creator Studio first."""
    candidates: list[str] = []

    explicit_target = _normalize_user_target(os.getenv("NUMBERS_APP_TARGET", ""))
    if explicit_target:
        candidates.append(explicit_target)

    explicit_bundle = os.getenv("NUMBERS_APP_BUNDLE_ID", "").strip()
    if explicit_bundle:
        candidates.append(f'application id "{explicit_bundle}"')

    candidates.extend(_discover_creator_studio_targets())
    candidates.extend(DEFAULT_NUMBERS_APP_TARGETS)

    deduped: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        normalized = candidate.strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        deduped.append(normalized)
    return tuple(deduped)

def _probe_applescript_target(target: str, timeout: float = 5.0) -> bool:
    """Best-effort probe for whether a document-based app is scriptable."""
    ok, _ = _probe_applescript_target_with_error(target, timeout=timeout)
    return ok


def _probe_applescript_target_with_error(target: str, timeout: float = 5.0) -> tuple[bool, str]:
    """Probe target and return (ok, error_text)."""
    script = f"tell {target} to count of documents"
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if result.returncode == 0:
            return True, ""
        stderr = (result.stderr or "").strip()
        stdout = (result.stdout or "").strip()
        return False, stderr or stdout or f"osascript exited with {result.returncode}"
    except subprocess.TimeoutExpired:
        return False, f"osascript probe timed out after {timeout:.1f}s"
    except FileNotFoundError:
        return False, "osascript not found"
    except Exception as exc:
        return False, str(exc)

def _resolve_applescript_target_candidates(
    candidates: tuple[str, ...],
    *,
    fallback: str,
    app_label: str,
) -> str:
    for candidate in candidates:
        if _probe_applescript_target(candidate):
            return candidate
    logger.warning(
        "Unable to verify AppleScript target for %s from candidates; falling back to %s",
        app_label,
        fallback,
    )
    return fallback

def _numbers_app_target() -> str:
    # Prefer explicit/env-driven and Creator Studio targets before Apple Numbers.
    candidates = _build_numbers_app_targets()
    if not candidates:
        candidates = DEFAULT_NUMBERS_APP_TARGETS
    return _resolve_applescript_target_candidates(
        candidates,
        fallback=candidates[0],
        app_label="Numbers",
    )


def numbers_preflight() -> dict[str, Any]:
    """Run environment and AppleScript-target checks before write operations."""
    osascript_path = shutil.which("osascript")
    candidates = _build_numbers_app_targets()
    selected_target = candidates[0] if candidates else ""

    probe_results: list[dict[str, Any]] = []
    selected_probe_ok = False
    selected_probe_error = ""
    for candidate in candidates:
        ok, error = _probe_applescript_target_with_error(candidate)
        probe_results.append({"target": candidate, "ok": ok, "error": error})
        if ok and not selected_probe_ok:
            selected_target = candidate
            selected_probe_ok = True
            selected_probe_error = ""
        if candidate == selected_target and not ok and not selected_probe_error:
            selected_probe_error = error

    hints: list[str] = []
    if not osascript_path:
        hints.append("Install/run on macOS with osascript available.")
    if osascript_path and not selected_probe_ok:
        hints.append("Grant Automation permission to Terminal/iTerm for the Numbers app target.")
        hints.append('Set NUMBERS_APP_TARGET, e.g. application "Numbers Creator Studio".')
        hints.append('Or set NUMBERS_APP_BUNDLE_ID to the app bundle identifier.')

    return {
        "ok": bool(osascript_path) and selected_probe_ok,
        "platform": sys.platform,
        "osascript_available": bool(osascript_path),
        "osascript_path": osascript_path or "",
        "env": {
            "NUMBERS_APP_TARGET": os.getenv("NUMBERS_APP_TARGET", ""),
            "NUMBERS_APP_BUNDLE_ID": os.getenv("NUMBERS_APP_BUNDLE_ID", ""),
            "NUMBERS_CREATOR_STUDIO_APP": os.getenv("NUMBERS_CREATOR_STUDIO_APP", ""),
        },
        "candidate_targets": list(candidates),
        "selected_target": selected_target,
        "selected_target_probe_ok": selected_probe_ok,
        "selected_target_probe_error": selected_probe_error,
        "probe_results": probe_results,
        "hints": hints,
    }

def _run_script(script: str, timeout: float = 30.0) -> str | None:
    """Run an osascript -e command. Returns stdout string or None on any failure."""
    transient_markers = (
        "Connection Invalid error for service com.apple.hiservices-xpcservice",
        "Error received in message reply handler: Connection invalid",
        "Expected class name but found identifier. (-2741)",
    )
    max_attempts = 8
    for attempt in range(1, max_attempts + 1):
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=timeout,
            )
            if result.returncode == 0:
                return result.stdout.strip("\r\n")

            stderr = (result.stderr or "").strip()
            is_transient = any(marker in stderr for marker in transient_markers)
            if is_transient and attempt < max_attempts:
                time.sleep(0.5 * attempt)
                continue

            logger.warning("AppleScript failed (rc=%s): %s", result.returncode, stderr)
            return None
        except subprocess.TimeoutExpired:
            logger.warning("AppleScript timed out after %.1fs", timeout)
            return None
        except FileNotFoundError:
            logger.warning("osascript not found — numbers_tools requires macOS")
            return None
        except Exception as exc:
            logger.warning("Unexpected error running AppleScript: %s", exc)
            return None
    return None

def numbers_create(
    file_path: str,
    headers: list[str],
    sheet_name: str = "",
    table_name: str = "",
    overwrite: bool = False,
) -> str | None:
    """Create a Numbers document and initialize header row."""
    path = Path(file_path).expanduser()
    if not path.is_absolute():
        logger.warning("numbers_create requires an absolute path: %s", file_path)
        return None
    if path.suffix.lower() != ".numbers":
        logger.warning("numbers_create requires a .numbers path: %s", file_path)
        return None
    if path.exists() and not overwrite:
        logger.warning("numbers_create target exists and overwrite=false: %s", file_path)
        return None
    if not headers:
        logger.warning("numbers_create requires at least one header")
        return None

    def _esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")

    esc_path = _esc(str(path))
    esc_sheet = _esc(sheet_name)
    esc_table = _esc(table_name)

    header_lines: list[str] = []
    for idx, header in enumerate(headers, start=1):
        header_lines.append(f'set value of cell {idx} of row 1 to "{_esc(str(header))}"')
    headers_block = "\n            ".join(header_lines)

    sheet_name_setter = (
        f'set name of first sheet of newDoc to "{esc_sheet}"'
        if sheet_name
        else ""
    )
    table_name_setter = (
        f'set name of first table of first sheet of newDoc to "{esc_table}"'
        if table_name
        else ""
    )
    numbers_app = _numbers_app_target()

    script = f'''
    tell {numbers_app}
        try
            activate
            set newDoc to make new document
            save newDoc in POSIX file "{esc_path}"
            {sheet_name_setter}
            {table_name_setter}
            set targetTable to first table of first sheet of newDoc
            tell targetTable
                set totalCols to count of columns
                set requiredCols to {len(headers)}
                repeat while totalCols < requiredCols
                    make new column at end of columns
                    set totalCols to totalCols + 1
                end repeat
                {headers_block}
            end tell
            save newDoc
            close newDoc saving yes
            return "{esc_path}"
        on error errMsg
            return "error: " & errMsg
        end try
    end tell
    '''
    result = _run_script(script, timeout=60.0)
    if not result or result.startswith("error:"):
        logger.warning("numbers_create failed: %s", result)
        return None
    return result

def _normalize_numbers_rows_payload(rows: Any) -> list[list[Any]] | None:
    if rows is None:
        return []
    if not isinstance(rows, list):
        return None
    normalized_rows: list[list[Any]] = []
    for row in rows:
        if isinstance(row, (list, tuple)):
            normalized_rows.append(list(row))
        else:
            normalized_rows.append([row])
    return normalized_rows

def _validate_numbers_sheet_spec(sheet_spec: Any) -> tuple[dict[str, Any], str | None]:
    if not isinstance(sheet_spec, dict):
        return {}, "sheet spec must be a JSON object"

    sheet_name = str(sheet_spec.get("sheet_name", "")).strip()
    if not sheet_name:
        return {}, "sheet_name is required"

    headers_raw = sheet_spec.get("headers")
    if not isinstance(headers_raw, list) or not headers_raw:
        return {}, "headers must be a non-empty JSON array"
    headers = [str(header) for header in headers_raw if str(header).strip()]
    if len(headers) != len(headers_raw):
        return {}, "headers must not contain empty values"

    table_name = str(sheet_spec.get("table_name", "")).strip()
    rows = _normalize_numbers_rows_payload(sheet_spec.get("rows"))
    if rows is None:
        return {}, "rows must be a JSON array when provided"

    return {
        "sheet_name": sheet_name,
        "table_name": table_name,
        "headers": headers,
        "rows": rows,
    }, None

def numbers_add_sheet(file_path: str, sheet_spec: dict[str, Any]) -> dict[str, Any]:
    """Add one initialized sheet to an existing Numbers workbook."""
    path = Path(file_path).expanduser()
    if not path.is_absolute():
        return {"ok": False, "error": "absolute path required"}
    if path.suffix.lower() != ".numbers":
        return {"ok": False, "error": ".numbers path required"}
    if not path.exists():
        return {"ok": False, "error": "target document does not exist"}

    normalized_spec, spec_error = _validate_numbers_sheet_spec(sheet_spec)
    if spec_error:
        return {"ok": False, "error": spec_error}

    sheet_name = normalized_spec["sheet_name"]
    table_name = normalized_spec["table_name"]
    headers = normalized_spec["headers"]
    rows = normalized_spec["rows"]
    required_cols = max(1, max(len(headers), max((len(row) for row in rows), default=0)))

    def _esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")

    esc_path = _esc(str(path))
    esc_sheet = _esc(sheet_name)
    esc_table = _esc(table_name)

    header_lines: list[str] = []
    for idx, header in enumerate(headers, start=1):
        header_lines.append(f'set value of cell {idx} of row 1 to "{_esc(str(header))}"')
    headers_block = "\n                    ".join(header_lines)

    row_lines: list[str] = []
    for row in rows:
        row_lines.extend(
            [
                "if insertionRow <= totalRows then",
                "set targetRow to row insertionRow",
                "else",
                "set targetRow to make new row at end of rows",
                "set totalRows to totalRows + 1",
                "end if",
            ]
        )
        for idx, value in enumerate(row, start=1):
            if value is None:
                row_lines.append(f"set value of cell {idx} of targetRow to \"\"")
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                row_lines.append(f"set value of cell {idx} of targetRow to {value}")
            else:
                row_lines.append(f'set value of cell {idx} of targetRow to "{_esc(str(value))}"')
        row_lines.append("set insertionRow to insertionRow + 1")
    rows_block = "\n                    ".join(row_lines) if row_lines else ""

    table_name_setter = f'set name of targetTable to "{esc_table}"' if table_name else ""
    rows_section = rows_block if rows_block else "-- no initial rows"
    numbers_app = _numbers_app_target()

    script = f'''
    tell {numbers_app}
        try
            activate
            set targetDoc to open POSIX file "{esc_path}"
            tell targetDoc
                set existingSheet to missing value
                try
                    set existingSheet to first sheet whose name is "{esc_sheet}"
                end try
                if existingSheet is not missing value then
                    close targetDoc saving no
                    return "error: sheet already exists"
                end if

                set newSheet to make new sheet at end of sheets
                set name of newSheet to "{esc_sheet}"
                tell newSheet
                    set targetTable to first table
                    {table_name_setter}
                    tell targetTable
                        set totalCols to count of columns
                        set requiredCols to {required_cols}
                        repeat while totalCols < requiredCols
                            make new column at end of columns
                            set totalCols to totalCols + 1
                        end repeat
                        set totalRows to count of rows
                        set insertionRow to 2
                        {headers_block}
                        {rows_section}
                    end tell
                end tell
            end tell
            save targetDoc
            close targetDoc saving yes
            return "ok|{len(rows)}"
        on error errMsg
            return "error: " & errMsg
        end try
    end tell
    '''
    result = _run_script(script, timeout=90.0)
    if not result:
        return {"ok": False, "error": "no response from Numbers"}
    if result.startswith("error:"):
        return {"ok": False, "error": result}

    rows_inserted = len(rows)
    parts = result.split("|")
    if len(parts) >= 2:
        try:
            rows_inserted = int(parts[1])
        except ValueError:
            pass
    return {
        "ok": True,
        "sheet_name": sheet_name,
        "table_name": table_name or "Table 1",
        "rows_inserted": rows_inserted,
    }

def numbers_create_workbook(
    file_path: str,
    workbook_spec: dict[str, Any],
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a multi-sheet workbook from a workbook spec."""
    if not isinstance(workbook_spec, dict):
        return {"ok": False, "error": "workbook_json must be a JSON object"}
    sheets_raw = workbook_spec.get("sheets")
    if not isinstance(sheets_raw, list) or not sheets_raw:
        return {"ok": False, "error": "workbook_json.sheets must be a non-empty JSON array"}

    normalized_sheets: list[dict[str, Any]] = []
    seen_sheet_names: set[str] = set()
    for sheet_spec in sheets_raw:
        normalized_spec, spec_error = _validate_numbers_sheet_spec(sheet_spec)
        if spec_error:
            return {"ok": False, "error": spec_error}
        sheet_key = normalized_spec["sheet_name"].strip().lower()
        if sheet_key in seen_sheet_names:
            return {"ok": False, "error": f'duplicate sheet_name: "{normalized_spec["sheet_name"]}"'}
        seen_sheet_names.add(sheet_key)
        normalized_sheets.append(normalized_spec)

    first_sheet = normalized_sheets[0]
    created = numbers_create(
        file_path,
        headers=first_sheet["headers"],
        sheet_name=first_sheet["sheet_name"],
        table_name=first_sheet["table_name"],
        overwrite=overwrite,
    )
    if not created:
        return {"ok": False, "error": "failed to create workbook"}

    rows_inserted_total = 0
    first_rows = first_sheet["rows"]
    if first_rows:
        first_insert = numbers_append_rows(
            file_path,
            rows=first_rows,
            sheet_name=first_sheet["sheet_name"],
            table_name=first_sheet["table_name"],
            insert_position="after-data",
        )
        if not first_insert.get("ok"):
            return {"ok": False, "error": str(first_insert.get("error", "failed to insert initial rows"))}
        rows_inserted_total += int(first_insert.get("inserted_rows", len(first_rows)))

    for sheet_spec in normalized_sheets[1:]:
        add_result = numbers_add_sheet(file_path, sheet_spec)
        if not add_result.get("ok"):
            return {
                "ok": False,
                "error": str(add_result.get("error", "failed to add sheet")),
                "sheets_created": len(normalized_sheets[: normalized_sheets.index(sheet_spec)]),
            }
        rows_inserted_total += int(add_result.get("rows_inserted", len(sheet_spec["rows"])))

    return {
        "ok": True,
        "path": str(Path(file_path).expanduser()),
        "sheets_created": len(normalized_sheets),
        "rows_inserted_total": rows_inserted_total,
    }

def numbers_append_rows(
    file_path: str,
    rows: list[list[Any]],
    sheet_name: str = "",
    table_name: str = "",
    insert_position: str = "after-data",
) -> dict[str, Any]:
    """Append one or more rows to a Numbers table."""
    path = Path(file_path).expanduser()
    if not path.is_absolute():
        return {"ok": False, "error": "absolute path required"}
    if path.suffix.lower() != ".numbers":
        return {"ok": False, "error": ".numbers path required"}
    if not path.exists():
        return {"ok": False, "error": "target document does not exist"}
    if insert_position not in {"after-headers", "after-data", "at-end"}:
        return {"ok": False, "error": "invalid insert position"}
    if not rows:
        return {"ok": False, "error": "rows must not be empty"}

    normalized_rows: list[list[Any]] = []
    for row in rows:
        if isinstance(row, (list, tuple)):
            normalized_rows.append(list(row))
        else:
            normalized_rows.append([row])
    required_cols = max(1, max((len(row) for row in normalized_rows), default=1))

    def _esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")

    esc_path = _esc(str(path))
    esc_sheet = _esc(sheet_name)
    esc_table = _esc(table_name)
    esc_position = _esc(insert_position)

    if insert_position == "after-headers":
        target_row_block = '''if insertionRow <= totalRows then
                    set anchorRow to row insertionRow
                    set targetRow to make new row at before anchorRow
                else
                    set targetRow to make new row at end of rows
                end if
                set totalRows to totalRows + 1'''
    elif insert_position == "at-end":
        target_row_block = '''set targetRow to make new row at end of rows
                set totalRows to totalRows + 1'''
    else:
        target_row_block = '''if insertionRow > totalRows then
                    set targetRow to make new row at end of rows
                    set totalRows to totalRows + 1
                else
                    set targetRow to row insertionRow
                end if'''

    row_lines: list[str] = []
    for row in normalized_rows:
        row_lines.append(target_row_block)
        for idx, value in enumerate(row, start=1):
            if value is None:
                row_lines.append(f"set value of cell {idx} of targetRow to \"\"")
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                row_lines.append(f"set value of cell {idx} of targetRow to {value}")
            else:
                row_lines.append(f'set value of cell {idx} of targetRow to "{_esc(str(value))}"')
        row_lines.append("set insertionRow to insertionRow + 1")
    rows_block = "\n                ".join(row_lines)

    sheet_lookup = (
        f'set targetSheet to (first sheet of targetDoc whose name is "{esc_sheet}")'
        if sheet_name
        else "set targetSheet to first sheet of targetDoc"
    )
    table_lookup = (
        f'set targetTable to (first table of targetSheet whose name is "{esc_table}")'
        if table_name
        else "set targetTable to first table of targetSheet"
    )
    numbers_app = _numbers_app_target()

    script = f'''
    tell {numbers_app}
        try
            activate
            set targetDoc to open POSIX file "{esc_path}"
            {sheet_lookup}
            {table_lookup}

            tell targetTable
                set totalRows to count of rows
                set totalCols to count of columns
                set requiredCols to {required_cols}
                repeat while totalCols < requiredCols
                    make new column at end of columns
                    set totalCols to totalCols + 1
                end repeat
                set scanCols to totalCols
                if scanCols < 1 then set scanCols to 1
                set headerRows to 1
                try
                    set headerRows to header row count
                end try
                if headerRows < 1 then set headerRows to 1
                set dataStartRow to headerRows + 1

                if "{esc_position}" is "after-headers" then
                    set insertionRow to dataStartRow
                else if "{esc_position}" is "at-end" then
                    set insertionRow to totalRows + 1
                else
                    set lastDataRow to headerRows
                    if totalRows >= dataStartRow then
                        repeat with r from dataStartRow to totalRows
                            set rowHasData to false
                            repeat with c from 1 to scanCols
                                set cellVal to missing value
                                try
                                    set cellVal to value of cell c of row r
                                on error
                                    set cellVal to missing value
                                end try
                                if cellVal is not missing value then
                                    try
                                        if (cellVal as text) is not "" then
                                            set rowHasData to true
                                            exit repeat
                                        end if
                                    on error
                                        set rowHasData to true
                                        exit repeat
                                    end try
                                end if
                            end repeat
                            if rowHasData then set lastDataRow to r
                        end repeat
                    end if
                    set insertionRow to lastDataRow + 1
                end if

                set startRow to insertionRow
                {rows_block}
            end tell

            save targetDoc
            close targetDoc saving yes
            return "ok|" & startRow & "|" & (insertionRow - 1)
        on error errMsg
            return "error: " & errMsg
        end try
    end tell
    '''
    result = _run_script(script, timeout=60.0)
    if not result:
        return {"ok": False, "error": "no response from Numbers"}
    if result.startswith("error:"):
        return {"ok": False, "error": result}

    parts = result.split("|")
    start_row = -1
    insert_after_row = -1
    if len(parts) >= 3:
        try:
            start_row = int(parts[1])
            insert_after_row = int(parts[2])
        except ValueError:
            pass
    return {
        "ok": True,
        "insert_position": insert_position,
        "attempted_rows": len(normalized_rows),
        "inserted_rows": len(normalized_rows),
        "start_row": start_row,
        "insert_after_row": insert_after_row,
    }

def _normalize_numbers_color_triplet(value: Any) -> tuple[int, int, int] | None:
    if not isinstance(value, (list, tuple)) or len(value) != 3:
        return None
    channels: list[float] = []
    for channel in value:
        if isinstance(channel, bool):
            return None
        try:
            numeric = float(channel)
        except (TypeError, ValueError):
            return None
        if numeric < 0:
            return None
        channels.append(numeric)
    if all(channel <= 255 for channel in channels):
        return tuple(int(round(channel * 257)) for channel in channels)
    if all(channel <= 65535 for channel in channels):
        return tuple(int(round(channel)) for channel in channels)
    return None

def _validate_numbers_style_target(target: Any) -> tuple[dict[str, int | str], str | None]:
    if not isinstance(target, dict):
        return {}, "target_json must be a JSON object"
    scope = str(target.get("scope", "")).strip().lower()
    if scope not in {"table", "row", "column", "cell", "range"}:
        return {}, "target_json.scope must be one of: table|row|column|cell|range"

    def _positive_int(key: str) -> tuple[int | None, str | None]:
        value = target.get(key)
        if isinstance(value, bool):
            return None, f"target_json.{key} must be a positive integer"
        try:
            int_value = int(value)
        except (TypeError, ValueError):
            return None, f"target_json.{key} must be a positive integer"
        if int_value < 1:
            return None, f"target_json.{key} must be >= 1"
        return int_value, None

    normalized: dict[str, int | str] = {"scope": scope}
    if scope == "table":
        return normalized, None
    if scope == "row":
        index, err = _positive_int("index")
        if err:
            return {}, err
        normalized["index"] = int(index)
        return normalized, None
    if scope == "column":
        index, err = _positive_int("index")
        if err:
            return {}, err
        normalized["index"] = int(index)
        return normalized, None
    if scope == "cell":
        row, err = _positive_int("row")
        if err:
            return {}, err
        col, err = _positive_int("column")
        if err:
            return {}, err
        normalized["row"] = int(row)
        normalized["column"] = int(col)
        return normalized, None

    start_row, err = _positive_int("start_row")
    if err:
        return {}, err
    end_row, err = _positive_int("end_row")
    if err:
        return {}, err
    start_col, err = _positive_int("start_column")
    if err:
        return {}, err
    end_col, err = _positive_int("end_column")
    if err:
        return {}, err
    if int(start_row) > int(end_row):
        return {}, "target_json start_row must be <= end_row"
    if int(start_col) > int(end_col):
        return {}, "target_json start_column must be <= end_column"
    normalized["start_row"] = int(start_row)
    normalized["end_row"] = int(end_row)
    normalized["start_column"] = int(start_col)
    normalized["end_column"] = int(end_col)
    return normalized, None

def _validate_numbers_style(style: Any, target_scope: str) -> tuple[dict[str, Any], str | None]:
    if not isinstance(style, dict):
        return {}, "style_json must be a JSON object"
    if not style:
        return {}, "style_json must not be empty"

    allowed_keys = {
        "background_color",
        "text_color",
        "font_name",
        "font_size",
        "alignment",
        "number_format",
        "text_wrap",
        "row_height",
        "column_width",
    }
    unknown = sorted(set(style.keys()) - allowed_keys)
    if unknown:
        return {}, f"unsupported style key(s): {', '.join(unknown)}"

    normalized: dict[str, Any] = {}
    for color_key in ("background_color", "text_color"):
        if color_key in style:
            color = _normalize_numbers_color_triplet(style[color_key])
            if color is None:
                return {}, f"style_json.{color_key} must be [r,g,b] with values in 0-255 or 0-65535"
            normalized[color_key] = color

    if "font_name" in style:
        font_name = str(style["font_name"]).strip()
        if not font_name:
            return {}, "style_json.font_name must be a non-empty string"
        normalized["font_name"] = font_name

    for numeric_key in ("font_size", "row_height", "column_width"):
        if numeric_key not in style:
            continue
        value = style[numeric_key]
        if isinstance(value, bool):
            return {}, f"style_json.{numeric_key} must be a positive number"
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return {}, f"style_json.{numeric_key} must be a positive number"
        if numeric <= 0:
            return {}, f"style_json.{numeric_key} must be > 0"
        normalized[numeric_key] = numeric

    if "alignment" in style:
        alignment = str(style["alignment"]).strip().lower()
        if alignment not in {"left", "center", "right", "justified", "natural"}:
            return {}, "style_json.alignment must be one of: left|center|right|justified|natural"
        normalized["alignment"] = alignment

    if "number_format" in style:
        number_format = str(style["number_format"]).strip().lower()
        if number_format not in {"automatic", "currency", "percentage", "scientific", "fraction", "text"}:
            return {}, "style_json.number_format must be one of: automatic|currency|percentage|scientific|fraction|text"
        normalized["number_format"] = number_format

    if "text_wrap" in style:
        text_wrap = style["text_wrap"]
        if not isinstance(text_wrap, bool):
            return {}, "style_json.text_wrap must be true or false"
        normalized["text_wrap"] = text_wrap

    if target_scope == "row" and "column_width" in normalized:
        return {}, "column_width is not supported for row target scope"
    if target_scope == "column" and "row_height" in normalized:
        return {}, "row_height is not supported for column target scope"

    return normalized, None

def numbers_style_apply(
    file_path: str,
    target: dict[str, Any],
    style: dict[str, Any],
    sheet_name: str = "",
    table_name: str = "",
) -> dict[str, Any]:
    """Apply formatting/style to a Numbers target scope."""
    path = Path(file_path).expanduser()
    if not path.is_absolute():
        return {"ok": False, "error": "absolute path required"}
    if path.suffix.lower() != ".numbers":
        return {"ok": False, "error": ".numbers path required"}
    if not path.exists():
        return {"ok": False, "error": "target document does not exist"}

    normalized_target, target_error = _validate_numbers_style_target(target)
    if target_error:
        return {"ok": False, "error": target_error}

    target_scope = str(normalized_target["scope"])
    normalized_style, style_error = _validate_numbers_style(style, target_scope=target_scope)
    if style_error:
        return {"ok": False, "error": style_error}

    cell_style_keys = {
        "background_color",
        "text_color",
        "font_name",
        "font_size",
        "alignment",
        "number_format",
        "text_wrap",
    }
    has_cell_styles = any(key in normalized_style for key in cell_style_keys)

    def _esc(text: str) -> str:
        return text.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")

    def _num_literal(value: float) -> str:
        return str(int(value)) if float(value).is_integer() else str(value)

    cell_style_lines: list[str] = []
    if "background_color" in normalized_style:
        r, g, b = normalized_style["background_color"]
        cell_style_lines.append(f"set background color of cellRef to {{{r}, {g}, {b}}}")
    if "text_color" in normalized_style:
        r, g, b = normalized_style["text_color"]
        cell_style_lines.append(f"set text color of cellRef to {{{r}, {g}, {b}}}")
    if "font_name" in normalized_style:
        cell_style_lines.append(f'set font name of cellRef to "{_esc(normalized_style["font_name"])}"')
    if "font_size" in normalized_style:
        cell_style_lines.append(f"set font size of cellRef to {_num_literal(normalized_style['font_size'])}")
    if "alignment" in normalized_style:
        cell_style_lines.append(f"set alignment of cellRef to {normalized_style['alignment']}")
    if "number_format" in normalized_style:
        cell_style_lines.append(f"set format of cellRef to {normalized_style['number_format']}")
    if "text_wrap" in normalized_style:
        cell_style_lines.append(
            f"set text wrap of cellRef to {'true' if normalized_style['text_wrap'] else 'false'}"
        )
    cell_styles_block = "\n                        ".join(cell_style_lines)

    has_row_height = "row_height" in normalized_style
    has_column_width = "column_width" in normalized_style
    row_height_line = _num_literal(float(normalized_style["row_height"])) if has_row_height else ""
    column_width_line = _num_literal(float(normalized_style["column_width"])) if has_column_width else ""

    if target_scope == "table":
        table_row_height_block = ""
        if has_row_height:
            table_row_height_block = f'''
                repeat with r from 1 to totalRows
                    set height of row r to {row_height_line}
                end repeat
                set rowsResized to totalRows
            '''
        table_column_width_block = ""
        if has_column_width:
            table_column_width_block = f'''
                repeat with c from 1 to totalCols
                    set width of column c to {column_width_line}
                end repeat
                set columnsResized to totalCols
            '''
        scope_block = f'''
                if {str(has_cell_styles).lower()} then
                    repeat with r from 1 to totalRows
                        repeat with c from 1 to totalCols
                            set cellRef to cell c of row r
                            {cell_styles_block}
                        end repeat
                    end repeat
                    set cellsTouched to totalRows * totalCols
                end if
                {table_row_height_block}
                {table_column_width_block}
        '''
    elif target_scope == "row":
        row_index = int(normalized_target["index"])
        row_height_block = ""
        if has_row_height:
            row_height_block = f'''
                set height of row {row_index} to {row_height_line}
                set rowsResized to 1
            '''
        scope_block = f'''
                if {row_index} > totalRows then return "error: target row out of bounds"
                if {str(has_cell_styles).lower()} then
                    repeat with c from 1 to totalCols
                        set cellRef to cell c of row {row_index}
                        {cell_styles_block}
                    end repeat
                    set cellsTouched to totalCols
                end if
                {row_height_block}
        '''
    elif target_scope == "column":
        column_index = int(normalized_target["index"])
        column_width_block = ""
        if has_column_width:
            column_width_block = f'''
                set width of column {column_index} to {column_width_line}
                set columnsResized to 1
            '''
        scope_block = f'''
                if {column_index} > totalCols then return "error: target column out of bounds"
                if {str(has_cell_styles).lower()} then
                    repeat with r from 1 to totalRows
                        set cellRef to cell {column_index} of row r
                        {cell_styles_block}
                    end repeat
                    set cellsTouched to totalRows
                end if
                {column_width_block}
        '''
    elif target_scope == "cell":
        row_index = int(normalized_target["row"])
        column_index = int(normalized_target["column"])
        cell_row_height_block = ""
        if has_row_height:
            cell_row_height_block = f'''
                set height of row {row_index} to {row_height_line}
                set rowsResized to 1
            '''
        cell_column_width_block = ""
        if has_column_width:
            cell_column_width_block = f'''
                set width of column {column_index} to {column_width_line}
                set columnsResized to 1
            '''
        scope_block = f'''
                if {row_index} > totalRows then return "error: target row out of bounds"
                if {column_index} > totalCols then return "error: target column out of bounds"
                if {str(has_cell_styles).lower()} then
                    set cellRef to cell {column_index} of row {row_index}
                    {cell_styles_block}
                    set cellsTouched to 1
                end if
                {cell_row_height_block}
                {cell_column_width_block}
        '''
    else:
        start_row = int(normalized_target["start_row"])
        end_row = int(normalized_target["end_row"])
        start_column = int(normalized_target["start_column"])
        end_column = int(normalized_target["end_column"])
        range_row_height_block = ""
        if has_row_height:
            range_row_height_block = f'''
                repeat with r from {start_row} to {end_row}
                    set height of row r to {row_height_line}
                end repeat
                set rowsResized to rangeRowCount
            '''
        range_column_width_block = ""
        if has_column_width:
            range_column_width_block = f'''
                repeat with c from {start_column} to {end_column}
                    set width of column c to {column_width_line}
                end repeat
                set columnsResized to rangeColCount
            '''
        scope_block = f'''
                if {end_row} > totalRows then return "error: range row out of bounds"
                if {end_column} > totalCols then return "error: range column out of bounds"
                set rangeRowCount to ({end_row} - {start_row}) + 1
                set rangeColCount to ({end_column} - {start_column}) + 1
                if {str(has_cell_styles).lower()} then
                    repeat with r from {start_row} to {end_row}
                        repeat with c from {start_column} to {end_column}
                            set cellRef to cell c of row r
                            {cell_styles_block}
                        end repeat
                    end repeat
                    set cellsTouched to rangeRowCount * rangeColCount
                end if
                {range_row_height_block}
                {range_column_width_block}
        '''

    esc_path = _esc(str(path))
    esc_sheet = _esc(sheet_name)
    esc_table = _esc(table_name)
    sheet_lookup = (
        f'set targetSheet to (first sheet of targetDoc whose name is "{esc_sheet}")'
        if sheet_name
        else "set targetSheet to first sheet of targetDoc"
    )
    table_lookup = (
        f'set targetTable to (first table of targetSheet whose name is "{esc_table}")'
        if table_name
        else "set targetTable to first table of targetSheet"
    )
    numbers_app = _numbers_app_target()

    script = f'''
    tell {numbers_app}
        try
            activate
            set targetDoc to open POSIX file "{esc_path}"
            {sheet_lookup}
            {table_lookup}

            tell targetTable
                set totalRows to count of rows
                set totalCols to count of columns
                set cellsTouched to 0
                set rowsResized to 0
                set columnsResized to 0
                {scope_block}
            end tell

            save targetDoc
            close targetDoc saving yes
            return "ok|" & cellsTouched & "|" & rowsResized & "|" & columnsResized
        on error errMsg
            return "error: " & errMsg
        end try
    end tell
    '''
    result = _run_script(script, timeout=90.0)
    if not result:
        return {"ok": False, "error": "no response from Numbers"}
    if result.startswith("error:"):
        return {"ok": False, "error": result}

    cells_touched = 0
    rows_resized = 0
    columns_resized = 0
    parts = result.split("|")
    if len(parts) >= 4:
        try:
            cells_touched = int(parts[1])
            rows_resized = int(parts[2])
            columns_resized = int(parts[3])
        except ValueError:
            pass

    return {
        "ok": True,
        "target_scope": target_scope,
        "applied_keys": list(normalized_style.keys()),
        "cells_touched": cells_touched,
        "rows_resized": rows_resized,
        "columns_resized": columns_resized,
    }


# ---------------------------------------------------------------------------
# Apple Mail
# ---------------------------------------------------------------------------

def _parse_bool(value: str) -> bool:
    lowered = value.strip().lower()
    if lowered in {"1", "true", "yes", "y", "on"}:
        return True
    if lowered in {"0", "false", "no", "n", "off"}:
        return False
    raise ValueError(f"invalid boolean value: {value}")


def _emit(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False))


def main() -> int:
    parser = argparse.ArgumentParser(description="Standalone local Numbers tools")
    sub = parser.add_subparsers(dest="command", required=True)

    p_create = sub.add_parser("numbers_create")
    p_create.add_argument("file_path")
    p_create.add_argument("headers_json")
    p_create.add_argument("--sheet", dest="sheet_name", default="")
    p_create.add_argument("--table", dest="table_name", default="")
    p_create.add_argument("--overwrite", default="false")

    p_workbook = sub.add_parser("numbers_create_workbook")
    p_workbook.add_argument("file_path")
    p_workbook.add_argument("workbook_json")
    p_workbook.add_argument("--overwrite", default="false")

    p_add = sub.add_parser("numbers_add_sheet")
    p_add.add_argument("file_path")
    p_add.add_argument("sheet_json")

    p_append = sub.add_parser("numbers_append_rows")
    p_append.add_argument("file_path")
    p_append.add_argument("rows_json")
    p_append.add_argument("--sheet", dest="sheet_name", default="")
    p_append.add_argument("--table", dest="table_name", default="")
    p_append.add_argument("--position", dest="insert_position", default="after-data")

    p_style = sub.add_parser("numbers_style_apply")
    p_style.add_argument("file_path")
    p_style.add_argument("target_json")
    p_style.add_argument("style_json")
    p_style.add_argument("--sheet", dest="sheet_name", default="")
    p_style.add_argument("--table", dest="table_name", default="")

    sub.add_parser("numbers_preflight")

    args = parser.parse_args()

    try:
        if args.command == "numbers_create":
            headers = json.loads(args.headers_json)
            overwrite = _parse_bool(args.overwrite)
            result = numbers_create(
                args.file_path,
                headers=headers,
                sheet_name=args.sheet_name,
                table_name=args.table_name,
                overwrite=overwrite,
            )
            _emit({"ok": bool(result), "path": result})
            return 0 if result else 1

        if args.command == "numbers_create_workbook":
            workbook = json.loads(args.workbook_json)
            overwrite = _parse_bool(args.overwrite)
            result = numbers_create_workbook(args.file_path, workbook, overwrite=overwrite)
            _emit(result)
            return 0 if result.get("ok") else 1

        if args.command == "numbers_add_sheet":
            sheet = json.loads(args.sheet_json)
            result = numbers_add_sheet(args.file_path, sheet)
            _emit(result)
            return 0 if result.get("ok") else 1

        if args.command == "numbers_append_rows":
            rows = json.loads(args.rows_json)
            result = numbers_append_rows(
                args.file_path,
                rows=rows,
                sheet_name=args.sheet_name,
                table_name=args.table_name,
                insert_position=args.insert_position,
            )
            _emit(result)
            return 0 if result.get("ok") else 1

        if args.command == "numbers_style_apply":
            target = json.loads(args.target_json)
            style = json.loads(args.style_json)
            result = numbers_style_apply(
                args.file_path,
                target=target,
                style=style,
                sheet_name=args.sheet_name,
                table_name=args.table_name,
            )
            _emit(result)
            return 0 if result.get("ok") else 1

        if args.command == "numbers_preflight":
            result = numbers_preflight()
            _emit(result)
            return 0 if result.get("ok") else 1

        _emit({"ok": False, "error": "unknown command"})
        return 2
    except json.JSONDecodeError as exc:
        _emit({"ok": False, "error": f"invalid json: {exc}"})
        return 2
    except Exception as exc:
        _emit({"ok": False, "error": str(exc)})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
