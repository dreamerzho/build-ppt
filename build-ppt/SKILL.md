---
name: build-ppt
description: Generate editable Swiss International Style PowerPoint decks (.pptx) from deck_spec.json. Wallpaper-grade editorial art direction. Use for PPT, PPTX, PowerPoint, make a pptx, create slides, Swiss-style decks, course decks, pitch decks, reports, editable slides, HTML-to-PPTX alternatives.
---

# Build PPT Skill

Generate native editable `.pptx` decks in a Swiss editorial style.

**You are an editorial art director and PPTX architect.** Convert text into structured layout intent (JSON), not hardcoded canvas coordinates. The downstream Python renderer owns all x/y placement.

---

## Aesthetic Core (Wallpaper-Grade Rules)

These are non-negotiable constraints that every slide must satisfy:

1. **Absolute grid & negative space**: 12-column grid, strict 8% safe margins. Every element breathing. Whitespace is intentional negative space, not accidental emptiness.
2. **Extreme type scale**: Hero titles at about 80-88pt for graphic impact. Kicker/metadata at 10pt or smaller for industrial-strategy feel. No mediocre middle ground.
3. **Flush-left faith**: All multi-line text MUST be flush-left with ragged right. NEVER justify. This preserves natural character breathing.
4. **Sharp 90-degree geometry**: Zero border-radius. All cards, rects, shapes are perfect right angles. No rounded corners, no shadows, no gradients.
5. **Spot color discipline**: Most of the page is black, white, grey. ONE accent color highlights key metrics or bolded terms. NEVER use multiple accents.
6. **Zero emoji**: Never output emoji in JSON. Use minimal geometric bullets (`-` or square bullets) or Swiss hairlines.

---

## Phase 1: Intent

Before writing the spec, confirm:

- **Topic, audience, use case, approximate slide count,** and required assets.
- **Mode**: `light` for projectors/print, `dark` for LED screens or online presentations.
- **Theme**: choose from `references/swiss-style.md`. Default to `swiss_classic` when the user gives no art direction; use `lime` only for the original neon AIGC course look.
- **Scope boundary**: this skill outputs native PPTX, not HTML, WebGL, video, or arbitrary custom color systems.

---

## Phase 2: Spec

If the user provides an Obsidian Markdown brief, convert it first:

```bash
python scripts/md_to_deck_spec.py source.md --out deck_spec.json
```

Then write or refine `deck_spec.json` using `references/deck-spec.schema.json`.

### Layout Intent (Your Only Job)

You **MUST** output abstract layout intent + structured content. You **MUST NOT** output raw x/y/left/top coordinates; the renderer owns placement.

Approved layouts:

| layout_type | S-ID | Best for |
|---|---|---|
| `cover` | `S01` | Cover page - big title |
| `split_statement` | `S03` | Left-dark / right-light statement |
| `six_cells` | `S04` | Six-cell grid - **requires `icon` field** |
| `three_layers` | `S05` | Three stacked cards - **requires `icon` field** |
| `duo_compare` | `S08` | Two-column side-by-side |
| `the_pause` | `S09` | Giant centered statement (chapter break) |
| `timeline` | `S11` | Horizontal timeline |
| `image_hero` | `S22` | Full-bleed single image with text overlay |

Prefer polished layouts: `S01`, `S02`, `S03`, `S04`, `S05`, `S08`, `S11`, `S15`, `S16`, `S19`, `S20`, `S21`, `S22`.

### Vector Icon System

When using **`S04` (six_cells)** or **`S05` (three_layers)**, you **MUST** add an `icon` field to every object in the `items` array. This gives each card a scalable, font-based vector icon, with no external images needed.

- `icon` value: a single **lowercase English noun** (FontAwesome semantic keyword).
- **Recommended icon vocabulary**:

| Category | Icon keywords |
|---|---|
| Tech/System | `microchip`, `laptop`, `server`, `code` |
| Data/Goals | `chart-line`, `bullseye`, `arrow-trend-up` |
| People/Roles | `user`, `users`, `id-badge` |
| Time/Process | `calendar`, `clock`, `list-check` |
| Creative/Design | `pen-nib`, `wand-magic-sparkles`, `layer-group` |

- **JSON example**:
```json
"items": [
  {
    "icon": "microchip",
    "title": "Technology and gadgets",
    "body": "Phones, earbuds, and wearables that emphasize precision, texture, and future-facing technology."
  }
]
```

- The renderer automatically maps `icon` to Font Awesome Unicode and renders it as a vector text box above the card's title.

### Text Discipline

- **One idea per slide**. If body text exceeds ~150 characters, you MUST split it into two consecutive slide objects.
- **Do NOT shrink typography** to cram dense content onto one page.
- Put all visible text in PowerPoint text boxes, not images.
- Images only for photos, screenshots, or supplied visuals.
- Image paths are relative to the JSON file. Obsidian image syntax such as `![[image.png|512]]` is supported by the Markdown converter and by the renderer path normalizer. Missing images become editable placeholders.
- Use `fit: "contain"` for product/reference images that must not be cropped, and `fit: "cover"` only for intentional hero/background crops.
- **No emoji in JSON.** The renderer strips emoji and warns.

### Visual References

- Use Noto Sans SC for titles, Noto Sans SC Black for cover impact, and Noto Sans SC Light for body/chrome/captions.
- Read `references/editorial-art-direction.md` before authoring.
- Use absolute canvas thinking, not HTML flow layout.
- 12-column grid, 8% safe margins, 0.5pt hairlines, sharp geometry.
- `meta.mode` or top-level `mode` for `light` / `dark` theme inversion.

---

## Phase 2.5: Checkpoint

Before rendering, show the user a concise slide outline unless they explicitly asked for direct or batch generation.

Include:
- Deck title, theme, mode, approximate slide count.
- Per-slide: title + `layout_type` or `Sxx` layout.
- Required image list and which pages will use placeholders.
- "Confirm the outline or request changes. After confirmation, I will validate and generate the PPTX."

Skip this checkpoint only when the user explicitly requests immediate generation, provides a finalized `deck_spec.json`, or the workflow is non-interactive.

---

## Phase 3: Validate and Render

```bash
python scripts/build_pptx.py deck_spec.json --validate-only
```

Fix all validation errors and rerun until it passes.

```bash
python scripts/build_pptx.py deck_spec.json --out output.pptx
```

Output is a native editable `.pptx`, not screenshots or HTML.

---

## Phase 4: Iterate

After delivery, iterate by editing `deck_spec.json`, validating, and rerendering. Keep the same theme and layout system unless the user requests new art direction.

---

## Recovery

| Problem | Recovery |
|---|---|
| `ModuleNotFoundError: pptx` | Install `python-pptx`. |
| `ModuleNotFoundError: PIL` | Install `Pillow`. |
| Font/text layout issues | Ensure Noto Sans SC / Noto Sans SC Light / Noto Sans SC Black are available; install `fonttools` if needed. |
| Icon(s) missing | People not seeing the tiny icon glyphs? Make sure "Font Awesome 7 Free Solid" font is installed on the machine generating the PPTX. |
| `FileNotFoundError: deck_spec.json` | Check the JSON path relative to the working directory. |
| Validation reports bad layout or missing field | Fix `layout`, `layout_type`, `title`, `theme`, `mode`, or image paths, rerun validation. |
| Image missing during generation | Placeholder rendered; replace in PowerPoint or fix the image path. |
| Script path not found | Run from the skill directory or use `python scripts/build_pptx.py ...`. |
| Python too old | Use Python 3.10+ for best compatibility. |

## Resources

- `scripts/build_pptx.py`: renderer and validator.
- `scripts/md_to_deck_spec.py`: Obsidian Markdown to `deck_spec.json` converter.
- `references/deck-spec.schema.json`: JSON contract.
- `references/editorial-art-direction.md`: aesthetic rules.
- `references/robustness.md`: overflow, emoji, dark mode, component registry.
- `references/swiss-style.md`: themes, typography, spacing, composition.
- `references/layouts-swiss-pptx.md`: `S01`-`S22` layout mapping.
- `references/sample_deck.json`: minimal example.
