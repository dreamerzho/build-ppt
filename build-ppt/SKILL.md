---
name: build-ppt
description: Generate editable Swiss International Style PowerPoint decks (.pptx) from a structured deck_spec.json. Use when Codex needs to create a real PPTX file, convert an outline into a Swiss-style presentation, render editable slides with text boxes/shapes/images, or produce a PowerPoint alternative to HTML slide decks.
---

# Build PPT Skill

## Overview

Create editable PowerPoint decks in the Swiss International Style and output native `.pptx` files instead of HTML. The primary workflow is: clarify the deck, write `deck_spec.json`, run `scripts/build_pptx.py`, inspect or iterate.

## Workflow

1. Choose only the Swiss style for v1. Do not use the electronic magazine style, WebGL backgrounds, HTML transitions, or custom colors.
2. Use one theme for the whole deck: `lime`, `ikb`, `lemon`, `lemon-green`, or `safety-orange`. Default to `lime` for AIGC, AI image generation, courses, and decks that should resemble the original guizang Swiss visual language.
3. Draft a `deck_spec.json` that follows `references/deck-spec.schema.json`. Use only registered layout IDs `S01` through `S22`.
4. Prefer the core layouts first: `S01`, `S02`, `S03`, `S04`, `S05`, `S08`, `S11`, `S15`, `S16`, `S19`, `S20`, `S21`, `S22`.
5. Run validation before writing the final PPTX:

```bash
python C:/Users/Administrator/.codex/skills/build-ppt/scripts/build_pptx.py deck_spec.json --validate-only
```

6. Generate the deck:

```bash
python C:/Users/Administrator/.codex/skills/build-ppt/scripts/build_pptx.py deck_spec.json --out output.pptx
```

## Deck Spec Rules

- Keep `slides[].layout` explicit. Never invent layout IDs.
- Keep page titles short, left-aligned, and anchored near the top-left content axis unless the chosen layout is a split or statement layout.
- Put visible slide text in PowerPoint text boxes, not images.
- Put diagrams in editable shapes/lines whenever possible. Use images only for photos, screenshots, or supplied visuals.
- Image paths are resolved relative to the JSON file location. Missing images render as editable placeholders during generation, and fail `--validate-only`.
- Speaker notes may be supplied with `notes`; the script writes them into PowerPoint notes.

## Visual Rules

Read `references/swiss-style.md` before creating a deck. The essentials:

- Use paper white `#fafaf8`, ink black `#0a0a0a`, shared greys, and one accent color. Prefer `lime` when the user wants the original neon AIGC course look.
- Use sparse chrome, bottom pagination dots, editable micro texture on cover pages, huge light-weight titles, strong grid alignment, and generous whitespace.
- Do not use gradients, shadows, rounded cards, mixed accent colors, decorative blobs, or center-aligned top titles.
- Treat images as evidence blocks. Use `S22` for one hero image and `S15`/`S16` for image grids.

## Resources

- `scripts/build_pptx.py`: CLI renderer and validator.
- `references/deck-spec.schema.json`: JSON contract for deck specs.
- `references/swiss-style.md`: theme, typography, spacing, and composition rules.
- `references/layouts-swiss-pptx.md`: layout intent and field mapping for `S01`-`S22`.
