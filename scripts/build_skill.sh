#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SKILL_DIR="$ROOT_DIR/skill/apple-flow-numbers"
DIST_DIR="$ROOT_DIR/dist"
OUT_FILE="$DIST_DIR/apple-flow-numbers.skill"

if [[ ! -d "$SKILL_DIR" ]]; then
  echo "Missing skill directory: $SKILL_DIR" >&2
  exit 1
fi

mkdir -p "$DIST_DIR"
tmpdir="$(mktemp -d)"
trap 'rm -rf "$tmpdir"' EXIT

mkdir -p "$tmpdir/apple-flow-numbers"
cp -R "$SKILL_DIR/." "$tmpdir/apple-flow-numbers/"

(cd "$tmpdir" && zip -q -r "$OUT_FILE" apple-flow-numbers)

echo "Built: $OUT_FILE"
unzip -l "$OUT_FILE"
