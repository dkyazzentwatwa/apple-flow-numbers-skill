#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SKILL_MD="$ROOT_DIR/skill/apple-flow-numbers/SKILL.md"
TOOL_SCRIPT="$ROOT_DIR/skill/apple-flow-numbers/scripts/numbers_tools.py"
SKILL_FILE="$ROOT_DIR/dist/apple-flow-numbers.skill"

if [[ ! -f "$SKILL_MD" ]]; then
  echo "Missing: $SKILL_MD" >&2
  exit 1
fi

if [[ ! -f "$TOOL_SCRIPT" ]]; then
  echo "Missing: $TOOL_SCRIPT" >&2
  exit 1
fi

if ! rg -q '^name:\s*apple-flow-numbers' "$SKILL_MD"; then
  echo "Frontmatter name missing or incorrect in $SKILL_MD" >&2
  exit 1
fi

if ! rg -q '^description:\s*' "$SKILL_MD"; then
  echo "Frontmatter description missing in $SKILL_MD" >&2
  exit 1
fi

echo "SKILL.md frontmatter check: OK"
echo "numbers_tools.py check: OK"

if [[ -f "$SKILL_FILE" ]]; then
  echo "Artifact check: $SKILL_FILE"
  unzip -l "$SKILL_FILE"
  if ! unzip -l "$SKILL_FILE" | rg -q 'apple-flow-numbers/scripts/numbers_tools.py'; then
    echo "Artifact is missing numbers_tools.py" >&2
    exit 1
  fi
else
  echo "Artifact not found yet: $SKILL_FILE"
  echo "Run ./scripts/build_skill.sh to generate it."
fi
