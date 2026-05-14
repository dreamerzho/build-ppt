---
name: build-ppt
description: Generate editable Swiss International Style PowerPoint decks (.pptx) from deck_spec.json. Use for PPT, PPTX, PowerPoint, editable slides, Swiss-style decks, course decks, pitch decks, reports, make a pptx, create slides, and HTML-to-PPTX alternatives.
---

# Build PPT Skill

Create native editable `.pptx` decks in a Swiss editorial style. Main flow: align intent -> write `deck_spec.json` -> review outline -> validate -> render -> iterate.

## Phase 1: Intent

Before writing the spec, identify:

- Topic, audience, use case, approximate slide count, and required assets.
- Mode: `light` for projectors/print, `dark` for LED screens or online presentations.
- Theme: choose from `references/swiss-style.md`. Default to `brutalist_tech` for AI/technology; use `lime` for the original neon AIGC course look.
- Scope boundary: this skill outputs native PPTX, not HTML, WebGL, video, or arbitrary custom color systems.

## Phase 2: Spec

Write `deck_spec.json` using `references/deck-spec.schema.json`.

- Prefer abstract `layout_type` values from `references/robustness.md`; direct `S01`-`S22` layout IDs are also supported.
- Prefer polished layouts: `S01`, `S02`, `S03`, `S04`, `S05`, `S08`, `S11`, `S15`, `S16`, `S19`, `S20`, `S21`, `S22`.
- Keep each slide to one idea. Long `body` text is auto-paginated; do not shrink typography to force dense content onto one page.
- Do not use emoji. The renderer strips emoji and warns, but agents should avoid them entirely.
- Use `meta.mode` or top-level `mode` for `light` / `dark` theme inversion.
- Put visible text in PowerPoint text boxes. Use editable shapes/lines for diagrams. Use images only for photos, screenshots, or supplied visuals.
- Image paths are relative to the JSON file. Missing images render as editable placeholders during generation and fail `--validate-only`.

Visual requirements:

- Read `references/editorial-art-direction.md` before authoring.
- Use absolute PowerPoint canvas thinking, not HTML flow layout.
- Use `Ting`, 12-column grid, 8% safe margins, huge 80-88pt titles, tiny chrome, 0.5pt hairlines, sharp geometry, one spot accent color.

## Phase 2.5: Checkpoint

Before rendering, show the user a concise outline unless they explicitly asked for direct or batch generation.

Include:

- Deck title, theme, mode, and approximate slide count.
- Per-slide list: title plus `layout_type` or `Sxx` layout.
- Required image list and which pages will use placeholders.
- Prompt: "Confirm this outline or send edits; after confirmation I will validate and generate the PPTX."

Skip this checkpoint only when the user clearly asks to generate immediately, provides a finalized `deck_spec.json`, or the workflow is automated/non-interactive.

## Phase 3: Validate and Render

Run validation first:

```bash
python scripts/build_pptx.py deck_spec.json --validate-only
```

Fix all validation errors and rerun until it passes.

Render:

```bash
python scripts/build_pptx.py deck_spec.json --out output.pptx
```

The output must be a native editable `.pptx`, not screenshots or HTML.

## Phase 4: Iterate

After delivery, iterate by editing `deck_spec.json`, validating again, and rerendering. Keep the same theme and layout system unless the user asks for a new art direction.

## Recovery

| Problem | Recovery |
|---|---|
| `ModuleNotFoundError: pptx` | Install `python-pptx`. |
| `ModuleNotFoundError: PIL` or image handling failure | Install `Pillow`. |
| Font or text layout issues | Ensure `Ting` is available; install `fonttools` if font inspection is needed. |
| `FileNotFoundError: deck_spec.json` | Check the JSON path relative to the working directory. |
| Validation reports bad layout or missing field | Fix `layout`, `layout_type`, `title`, `theme`, `mode`, or image paths, then rerun validation. |
| Image missing during generation | Placeholder is rendered; replace the image path or swap the placeholder in PowerPoint. |
| Script path not found | Run from the skill directory or use relative path `python scripts/build_pptx.py ...`. |
| Python too old | Use Python 3.10+ for best compatibility. |

## Resources

- `scripts/build_pptx.py`: renderer and validator.
- `references/deck-spec.schema.json`: JSON contract.
- `references/editorial-art-direction.md`: aesthetic rules.
- `references/robustness.md`: overflow, emoji, dark mode, component registry.
- `references/swiss-style.md`: themes, typography, spacing, composition.
- `references/layouts-swiss-pptx.md`: `S01`-`S22` layout mapping.
- `references/sample_deck.json`: minimal example.
