---
name: build-ppt
description: Generate editable Swiss International Style PowerPoint decks (.pptx) from a structured deck_spec.json. Use when Codex needs to create a real PPTX file, convert an outline into a Swiss-style presentation, render editable slides with text boxes/shapes/images, or produce a PowerPoint alternative to HTML slide decks.
---

# Build PPT Skill

## Overview

Create editable PowerPoint decks in the Swiss International Style and output native `.pptx` files instead of HTML. The primary workflow is: clarify the deck, write `deck_spec.json`, run `scripts/build_pptx.py`, inspect or iterate.

## Workflow

1. Choose only the Swiss style for v1. Do not use the electronic magazine style, WebGL backgrounds, HTML transitions, or custom colors.
2. Use one theme for the whole deck: `brutalist_tech`, `swiss_classic`, `corporate_chic`, `lime`, `ikb`, `lemon`, `lemon-green`, or `safety-orange`. Default to `brutalist_tech` for AI/technology decks and `lime` when the user asks for the original neon AIGC course look.
3. Draft a `deck_spec.json` that follows `references/deck-spec.schema.json`. Prefer abstract `layout_type` values from `references/robustness.md`; direct `S01`-`S22` layout IDs are also supported.
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
- Do not use emoji in JSON. The renderer strips emoji and warns, but agents should avoid them entirely.
- Use `meta.mode` or top-level `mode` for `light` / `dark` theme inversion.
- Long `body` text is automatically split into continuation slides; never ask the renderer to shrink text below the design system.
- Prefer `layout_type + content` over raw coordinates. The Python component registry owns placement.

## Visual Rules

Read `references/editorial-art-direction.md` and `references/swiss-style.md` before creating a deck. The essentials:

- Think in absolute PowerPoint canvas coordinates, not HTML flow layout.
- Use a 12-column grid, 8% safe margins, sparse chrome, bottom pagination dots, editable micro texture on cover pages, huge titles, and generous whitespace.
- Use `Ting` as the locked font family. Hero/title text must be around 80-88pt; chrome must be 10pt or smaller.
- Do not use gradients, shadows, rounded cards, mixed accent colors, decorative blobs, or center-aligned multi-line titles.
- Treat images as evidence blocks. Use `S22` for one hero image and `S15`/`S16` for image grids.

## Resources

- `scripts/build_pptx.py`: CLI renderer and validator.
- `references/deck-spec.schema.json`: JSON contract for deck specs.
- `references/editorial-art-direction.md`: System-level aesthetic rules for agents.
- `references/robustness.md`: Overflow, emoji, dark mode, and component registry rules.
- `references/swiss-style.md`: theme, typography, spacing, and composition rules.
- `references/layouts-swiss-pptx.md`: layout intent and field mapping for `S01`-`S22`.
