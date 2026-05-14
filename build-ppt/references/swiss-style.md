# Swiss PPTX Style

This skill renders native editable PowerPoint objects in a Swiss International Style. It does not attempt to reproduce HTML, canvas, WebGL, or motion.

## Themes

Use one accent color per deck.

| Theme | Accent | Accent Text | Use |
|---|---:|---:|---|
| `brutalist_tech` | `#CCFF00` | `#0F172A` | default, AI, technology, industrial, cyber editorial |
| `swiss_classic` | `#E6321E` | `#ffffff` | art, design, brand, classic Swiss editorial |
| `corporate_chic` | `#00D4FF` | `#0A2540` | finance, consulting, real estate, business decks |
| `lime` | `#B8F000` | `#0a0a0a` | AIGC course decks, AI image generation, original neon look |
| `ikb` | `#002FA7` | `#ffffff` | AI, technology, design, business |
| `lemon` | `#FFD500` | `#0a0a0a` | youth, activity, consumer, retail |
| `lemon-green` | `#C5E803` | `#0a0a0a` | ecology, future, health, emerging tech |
| `safety-orange` | `#FF6B35` | `#ffffff` | industry, urgency, warning, automotive |

Shared colors:

- Paper: `#fafaf8`
- Ink: `#0a0a0a`
- Grey 1: `#f0f0ee`
- Grey 2: `#d4d4d2`
- Grey 3: `#737373`

## Typography

- Use the Microsoft YaHei family as the locked font system: titles use `Microsoft YaHei` bold, body/chrome/captions use `Microsoft YaHei Light`.
- Main titles: very large, around 80-88pt, mostly black; use the accent color only for selected words on white pages.
- Kicker/meta: small uppercase mono-like labels where possible.
- Body: concise, low-to-medium density. Avoid long paragraphs.
- Numbers: large, tabular-looking, often paired with tiny labels.

## Layout

- Default slide size is 16:9 widescreen.
- Use a 12-column absolute coordinate grid with 8% safe margins.
- Keep a strong left/top content axis.
- Use sparse top chrome, tiny metadata, bottom pagination dots, and generous empty space.
- Cover slides should use full accent background plus editable dot/cross texture.
- Use straight rectangles, editable lines, and editable text.
- Use 1 px hairlines for structure.
- Avoid rounded cards, shadows, gradients, transparency-heavy overlays, and decorative effects.
- Keep footer/page-number areas clear.

## Images

- Use image blocks as evidence: product screenshots, field photos, UI details, or concrete examples.
- Missing images should remain visible as editable placeholders, not break generation.
- `S22` is the default single-image hero layout.
- `S15` and `S16` may be used for multi-image grids.
