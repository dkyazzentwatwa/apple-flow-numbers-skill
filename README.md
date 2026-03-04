# Apple Flow Numbers Skill (Standalone)

Standalone package for the `apple-flow-numbers` skill so you can keep it in its own repo.

## What is included

- `skill/apple-flow-numbers/SKILL.md` - skill instructions and workflows
- `skill/apple-flow-numbers/scripts/numbers_tools.py` - standalone Numbers CLI extracted from `apple_tools.py` (Numbers-only)
- `scripts/build_skill.sh` - builds `dist/apple-flow-numbers.skill`
- `scripts/install_skill.sh` - installs the built skill into your Codex skills directory
- `scripts/verify_skill.sh` - validates structure/frontmatter and artifact contents
- `scripts/md_log_to_numbers_workbook.py` - converts `automation-log.md` style entries into a multi-tab `.numbers` workbook through local `numbers_tools.py`

## No apple-flow dependency required

You can run Numbers tools directly with:

```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py --help
```

## Quick start

```bash
cd apple-flow-numbers-skill
./scripts/build_skill.sh
./scripts/install_skill.sh
```

## Convert automation-log.md into 3 tabs x 20 rows

```bash
cd apple-flow-numbers-skill
./scripts/md_log_to_numbers_workbook.py \
  --input /path/to/automation-log.md \
  --output /path/to/automation-log-3tabs.numbers \
  --sheets 3 \
  --rows-per-sheet 20 \
  --overwrite
```

## Repo notes

- The `.skill` artifact is a zip containing `apple-flow-numbers/SKILL.md` and `apple-flow-numbers/scripts/numbers_tools.py`.
- Keep the source of truth in `skill/apple-flow-numbers/`.
