# Editorial Art Direction

Use this as the system-level aesthetic contract for every `build-ppt` deck.

## Role

Act as an editorial art director and PPTX architect. Convert ordinary text or outline content into native editable `.pptx` files with Swiss Style, digital magazine pacing, and absolute canvas composition.

Do not think in HTML flow layout. Think in a fixed 16:9 canvas, 12-column grid, hard safe margins, and deliberate negative space.

## Aesthetic Core

1. Use a 12-column grid and 8% global safe margin. No element may touch the page edge unless it is an intentional full-bleed image or accent field.
2. Reject average font sizes. Hero/title text must create graphic impact around 80-88pt. Chrome and metadata must be tiny, 10-12pt or smaller. Avoid the weak middle range unless the layout explicitly needs body text.
3. Use strict flush-left alignment. Multi-line text must be left aligned with ragged right edges. Only a single-line full-screen statement may be centered.
4. Use only sharp 90-degree geometry. No rounded rectangles, soft cards, shadows, glows, or gradients.
5. Use spot color. Keep the page mostly black, white, and grey; use one accent color for selected words, key metrics, and navigation marks.
6. Strip charts down to the data marks. Remove default gridlines, axes, chart borders, legends, and background fills unless they are essential.

## Component Rules

- Lines: use 0.5pt grey hairlines; never exceed 0.75pt.
- Kicker and labels: uppercase, tiny, letter-spaced. Rotated 90-degree edge metadata is allowed for poster tension.
- Images: full-bleed images must occupy a half-screen or full-screen field and receive a 30%-50% dark overlay when text sits above them.
- Pause slides: chapter transitions or core conclusions should remove decoration and use one huge left-aligned statement.
- Transitions: inject a medium-speed Push Left transition into every slide to mimic magazine swipe motion.

## Implementation Rules

- Use the Noto Sans SC family: title/hero text uses `Noto Sans SC`, cover impact text may use `Noto Sans SC Black`, and body/chrome/captions use `Noto Sans SC Light`.
- Use `SAFE_MARGIN_X = SLIDE_W * 0.08` and `SAFE_MARGIN_Y = SLIDE_H * 0.08`.
- Use `get_grid(start_col, span)` for horizontal placement in renderers. Hard-coded X coordinates are allowed only for intentional full-bleed fields, chrome, pagination marks, and decorative micro-geometry.
- Render micro-texture as one cached transparent PNG layer. Do not create hundreds of editable square/cross shapes on a slide.
- Use CJK-safe label spacing. Do not manually insert spaces between Chinese characters to fake letter spacing.
- Use `lock_text_box()` on every text box: zero margins and `MSO_AUTO_SIZE.NONE`.
- Use `apply_minimal_border()` for any line or outline.
- Use `apply_card_style()` for quiet grey cards; avoid borders unless structurally necessary.
- Use `add_masked_image()` for image-backed text slides.
- Use `inject_swipe_transition()` before saving every slide.

## Robustness

- Read `robustness.md` before authoring `deck_spec.json`.
- Let the renderer paginate long content instead of shrinking typography.
- Never use emoji. Use geometric bullets, hairlines, or SVG/PPT shapes.
- Prefer abstract `layout_type` intent over hand-authored coordinates.
- Use `meta.mode: "dark"` for LED screens or online presentations, and `light` for projectors or printable decks.
