# Task: edit — rebuild a deck into the Every Consulting brand

Edit quality under the hardest real-world ask: take a finished deck built in a *different*
visual language and **rebuild it into a specific, documented brand identity** — not a color
swap, a full visual-system transformation (backgrounds, texture, typography, layout language,
art treatment, footers, list/pill styling) — while preserving the deck's content, message,
and slide count. This is the realistic "make our deck match the brand" job a consulting team
faces, judged against real brand guidelines.

Input: each arm gets the same finished **15-slide** deck at `{{WORKDIR}}/deck.pptx` (a
"Compound Engineering" talk, currently in a non-Every visual style; canvas 20×11.25in, 16:9).
This is a **same-deck** setup — both arms rebuild copies of the identical input.

Brand reference + assets are in `{{WORKDIR}}/brand/`:
- `every-consulting-brand-guidelines.md` — the visual identity to rebuild into (read it in full)
- `stipple-texture-overlay.png` — the teal background grain (1920×1920)
- `every-logo-white.png` — footer logo
- `art/` — classical black-and-white illustrations for atmospheric placement

Placeholders: `{{TOOLCHAIN_BLOCK}}`, `{{WORKDIR}}`.

---

You are a presentation designer. A finished 15-slide deck needs to be rebuilt into the **Every
Consulting** visual identity. The deck is at `{{WORKDIR}}/deck.pptx`; the brand is fully
specified in `{{WORKDIR}}/brand/every-consulting-brand-guidelines.md` (read it first, in full).

## The rebuild

Transform **every** slide into the Every Consulting (Consulting/Teal) brand as the guidelines
define it. That means, at minimum:

- **Backgrounds:** deep teal `#0F5258` carrying the stipple texture's contour sweep — use the
  provided `stipple-texture-overlay.png` per the guidelines' pan/flip + ~10%-strength rules.
  No flat teal, no leftover original backgrounds.
- **Typography:** editorial serif (Georgia) for headlines and body; the sans only for small
  labels/footers. Sentence case. No all-caps headlines, no sans-serif body.
- **Color:** warm-white text `#FFFEFB`; ALL secondary text (eyebrows, captions, footers) at the
  readability floor `#C3DCDE`, nothing dimmer; accents used selectively as gold/green category
  pills and highlights — not gratuitous fills.
- **Layout language:** the brand's vocabulary — 50/50 content-and-illustration splits, rounded
  frosted cards, pill badges for categories, organic curves, generous margins, one focal point.
- **Art:** place the bundled black-and-white classical illustrations as atmospheric companions
  where they fit the content; preserve aspect ratios. Don't cram or decorate corners.
- **Footer:** `EVERY` + italic *Consulting* left (the white logo), section/act indicator right.

**Preserve:** the deck's content, message, the Act/section narrative, and the slide count (15).
You are restyling the deck into the brand, not rewriting the talk or adding/removing slides.

Honor the brand's avoid-list (no corporate blue/gray, no geometric/angular accents, no
gradients-as-backgrounds, no clutter). The bar: each slide should look like it could *only* be
an Every Consulting slide.

## Your toolchain (use ONLY this)

{{TOOLCHAIN_BLOCK}}

## Process

1. Inspect the deck first — understand each slide's structure, content, and current styling.
2. Read the brand guidelines in full before designing.
3. Rebuild the slides into the brand.
4. Review: render every slide, look at the images, and check against the guidelines — texture
   present and quiet, serif everywhere, secondary text at/above the floor, pills/cards on-brand,
   art placed well, footers consistent, nothing from the avoid-list, content intact. Fix what
   you find. At least one genuine review-and-fix round.

## Deliverables (in `{{WORKDIR}}`)

- `final.pptx` — the rebuilt 15-slide deck, fully in the Every Consulting brand
- `img/` — rendered JPGs of all final slides
- `contact-sheet.jpg` — thumbnail grid

Your final message: a report of what you changed and how (approach, ops/iterations, what review
caught), how you handled the texture and art, and any limitations. Do not ask questions; decide
everything yourself. This is unattended.
