# Robustness Contract

Use these rules to keep generated PPTX decks commercially reliable.

## Text Overflow

The renderer paginates long `body` strings before slide creation. If a body exceeds the internal character threshold, it is split on sentence boundaries and emitted as continuation slides with `· 01`, `· 02`, etc. Never shrink text to unreadable sizes to force content into one slide.

Authoring rule: keep each slide to one idea. When the source content is long, prefer multiple slides over dense paragraphs.

## Emoji Ban

Do not put emoji in `deck_spec.json`. The renderer recursively removes emoji-like glyphs and prints warnings such as `Emoji removed at $.slides[2].title`.

Use Swiss marks instead:

- `-` for list rhythm
- small editable squares for geometric bullets
- line-art SVG or editable PPT shapes for icons

## Theme Inversion

Use either top-level `mode` or `meta.mode`:

```json
{
  "meta": {
    "theme": "brutalist_tech",
    "mode": "dark",
    "aspect_ratio": "16:9"
  }
}
```

`light` mode maps palette light backgrounds to dark text. `dark` mode inverts the background/text relationship while preserving the accent color.

## Component Registry

Agents should prefer abstract intent over manual coordinates. Use:

```json
{
  "layout_type": "split_hero_right",
  "title": "DATA NEVER LIES",
  "content": {
    "kicker": "MARKET ANALYSIS",
    "body": "The Q4 metrics show a decisive shift..."
  }
}
```

The renderer maps `layout_type` to a registered S-layout and owns the coordinates. Current mappings:

| layout_type | S-layout |
|---|---|
| `cover` | `S01` |
| `objective_list` | `S02` |
| `split_statement` | `S03` |
| `six_cells` | `S04` |
| `three_layers` | `S05` |
| `duo_compare` | `S08` |
| `timeline` | `S11` |
| `the_pause` | `S09` |
| `brief_grid` | `S16` |
| `split_hero_right` / `image_hero` | `S22` |

Do not emit raw `x`, `y`, `left`, `top`, or arbitrary placement data in JSON unless a future renderer explicitly supports it.

## Renderer Performance And Typography

Renderer functions should derive horizontal placement from `get_grid(start_col, span)` wherever possible. Keep hard-coded coordinates limited to full-bleed backgrounds, fixed chrome, and small decorative accents.

Micro-texture must be inserted as one cached transparent PNG layer, not hundreds of PPT shape nodes. This keeps exported decks editable without making PowerPoint sluggish.

Kicker letter spacing is Latin-only. For CJK text, keep characters adjacent so PowerPoint does not treat every Chinese character as a separate word at line breaks.

Use Microsoft YaHei as the title family and Microsoft YaHei Light for body-level text. Titles should be explicitly bold; body copy should stay light.
