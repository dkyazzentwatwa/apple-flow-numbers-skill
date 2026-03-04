#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SKILL_FILE="$ROOT_DIR/dist/apple-flow-numbers.skill"
CODEX_HOME_DIR="${CODEX_HOME:-$HOME/.codex}"
DEST_DIR="$CODEX_HOME_DIR/skills"

if [[ ! -f "$SKILL_FILE" ]]; then
  echo "Missing artifact: $SKILL_FILE" >&2
  echo "Run ./scripts/build_skill.sh first." >&2
  exit 1
fi

mkdir -p "$DEST_DIR"
unzip -oq "$SKILL_FILE" -d "$DEST_DIR"

echo "Installed skill into: $DEST_DIR/apple-flow-numbers"
