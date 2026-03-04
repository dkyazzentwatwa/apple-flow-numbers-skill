# Apple Numbers Skill (Beginner Guide)

This skill lets your AI assistant (Claude, Codex, Cline, etc.) create and update Apple Numbers spreadsheets for you.

You do not need to run terminal commands.

## What This Skill Helps You Do

- Create new `.numbers` files
- Add multiple sheets (tabs)
- Add rows to the top, after data, or bottom
- Apply simple table styling
- Convert structured notes/logs into spreadsheets

## What You Need

- A Mac
- A Numbers app installed
- `Numbers Creator Studio` (recommended)
- Older Apple Numbers versions are also supported as fallback
- An AI coding assistant you already use (Claude/Codex/Cline/etc.)

## How To Install (No Technical Steps)

Tell your AI assistant to install this repo’s skill for you.

Use a prompt like:

- "Install the Apple Numbers skill from this repository into my skills folder."
- "Set up this Apple Numbers skill so you can create and edit .numbers files for me."

If the assistant asks which app target to use, say:

- "Use Numbers Creator Studio first, then fallback to legacy Numbers if needed."

## First-Time Setup Prompt

After install, ask your assistant:

- "Run a preflight check for the Apple Numbers skill and fix any app-target issues automatically."

If there is a macOS permission popup, approve it.

## Example Prompts You Can Copy

- "Create a new expense tracker Numbers file with columns Date, Item, Category, Amount, Payment Method, Notes and add 20 test rows."
- "Add 2 rows at the top and 2 rows at the bottom of each sheet in my second brain workbook."
- "Create a workbook with 8 tabs for projects, people, inbox, decisions, habits, learning, resources, and archive."
- "Style the header row so it is easy to read and set practical column widths."
- "Verify that rows were inserted correctly and tell me which row numbers were used."

## If Something Fails

Tell your assistant exactly this:

- "Use Numbers Creator Studio as the app target and retry outside sandbox if needed."
- "Run preflight again and show me what app target you selected."
- "If Creator Studio fails, try legacy Numbers fallback automatically."

## Notes For Beginners

- You can speak in plain English; no special syntax is required.
- Ask the assistant to explain what it changed after each action.
- If you want safe testing, ask it to create a new `*_test.numbers` file first.

## Developer Instructions (Optional)

If you prefer to run the skill directly during development:

1. Run preflight:
```bash
python3 skill/apple-flow-numbers/scripts/numbers_tools.py numbers_preflight
```

2. Force Creator Studio target if needed:
```bash
export NUMBERS_APP_TARGET='application "Numbers Creator Studio"'
```

3. Build the distributable artifact:
```bash
./scripts/build_skill.sh
```

4. Verify skill integrity:
```bash
./scripts/verify_skill.sh
```

5. Install into local Codex skills:
```bash
./scripts/install_skill.sh
```

If automation fails, grant macOS Automation permission for your terminal app and retry preflight.

## License

MIT
