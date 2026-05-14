# build-ppt

`build-ppt` is a Codex skill for generating editable Swiss-style PowerPoint decks from structured JSON. It is designed for AIGC course decks and presentation workflows that need native `.pptx` output instead of HTML screenshots.

## What It Does

- Builds editable PowerPoint files with `python-pptx`.
- Uses a strict canvas-style coordinate system instead of HTML flow layout.
- Enforces Swiss-style visual controls: large titles, tiny chrome, strong whitespace, hairline rules, and a single accent color.
- Supports registered layouts `S01` through `S22`.
- Validates `deck_spec.json` before generation.

## Install

Copy the skill folder into your Codex skills directory:

```powershell
Copy-Item -Recurse .\build-ppt C:\Users\Administrator\.codex\skills\build-ppt
```

Or on Unix-like systems:

```bash
cp -R build-ppt ~/.codex/skills/build-ppt
```

## Usage

Create a `deck_spec.json` that follows:

```text
build-ppt/references/deck-spec.schema.json
```

Validate:

```bash
python build-ppt/scripts/build_pptx.py deck_spec.json --validate-only
```

Generate:

```bash
python build-ppt/scripts/build_pptx.py deck_spec.json --out output.pptx
```

## Agent Prompt

Use this prompt with another agent:

```text
Use the build-ppt skill to create an editable Swiss-style PPTX deck.

Read build-ppt/SKILL.md first. Then create deck_spec.json using build-ppt/references/deck-spec.schema.json. Use only registered layouts S01-S22 and one theme: lime, ikb, lemon, lemon-green, or safety-orange. Prefer lime for AIGC course decks. Validate with:

python build-ppt/scripts/build_pptx.py deck_spec.json --validate-only

Then generate:

python build-ppt/scripts/build_pptx.py deck_spec.json --out output.pptx

The output must be a native editable .pptx, not screenshots or HTML.
```

## Requirements

- Python 3.10+
- `python-pptx`
- `Pillow`
- `lxml`

## License

MIT
