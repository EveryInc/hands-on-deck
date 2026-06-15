# html2patch — design specification

**Status**: v1 spec, written before implementation (clean-room design record).

## What it is

`html2patch.py` compiles an HTML file into a **deck.py patch** (`{"ops": [...]}`).
The agent writes a slide as HTML/CSS — the layout language LLMs are best at — and
a headless browser (Playwright/Chromium) is used purely as a *measuring engine*:
every rendered element's box and computed style is read back and translated into
`add-shape` / `add-picture` / `add-table` / `set-text` ops. The patch is then applied
by deck.py like any other patch: validated atomically, linted, fixable, renderable.

It deliberately does NOT write a .pptx. One writer (deck.py / python-pptx) owns all
output, so created slides get the same shape ids, lint coverage, `fix` loop, and
`diff`/`render` verification as edited ones — and creation can target slides in an
**existing branded deck** (`add-slide` with a template layout), not just blank files.

## Provenance

The general idea — render HTML in a browser and read back element geometry to
position PowerPoint shapes — is a known public technique (used by Anthropic's
proprietary html2pptx, which emits a .pptx via PptxGenJS). We studied that
implementation for *behavioral learnings* (see below) but share no code with it,
and the architecture differs: Python not Node, patch-JSON output not a generated
file, deck.py as the single writer, template-aware creation, native gradient and
table support.

Learnings adopted (rediscovered behaviors, reimplemented):
- PowerPoint renders text slightly wider than the browser measures — single-line
  text boxes get ~2% extra width, distributed by alignment (center: both sides,
  right: leftward, left: rightward).
- `textContent` does not reflect CSS `text-transform`; apply uppercase/lowercase/
  capitalize manually during extraction.
- Bold on single-weight fonts (e.g. Impact) triggers PowerPoint faux-bold, which
  widens text — suppress bold for those families.
- When an inline run's font is larger than the block font, absolute line spacing
  must be rescaled by the largest run size (browser line boxes grow; Pt spacing
  doesn't).
- PPTX rotation spins the *pre-rotation* box about its center: for ±90° swap w/h
  around the center; arbitrary angles come from computed `matrix(a,b,...)` via
  atan2(b, a); `writing-mode: vertical-rl` ≡ 90°, `vertical-lr` ≡ 270°.
- Non-uniform borders (accent bars) are emitted as separate line shapes inset by
  half the border width so the stroke centers on the box edge.
- `text-align: start` must be normalized to left.
- Overflow check: body scrollWidth/Height vs CSS size with 1px tolerance.

## Coordinate model

- The HTML `<body>` is the slide. Its CSS size must be the slide size at **96 px/in**:
  16:9 → `width:1280px; height:720px` (13.333 × 7.5 in).
- All geometry comes from `getBoundingClientRect()` in CSS pixels; px / 96 = inches.
- Conversion constants: 96 px = 1 in = 72 pt → font px × 0.75 = pt.
- `--deck deck.pptx` reads the true slide size (and slide count) from a real deck;
  mismatch between body size and slide size is a hard error listing both.

## Element classification (document order = z-order, back to front)

Walk the DOM in document order. Each element matches the FIRST rule below; matched
elements are emitted and (except BOX) their subtrees are not descended further.

1. **IMG** — `<img>` → `add-picture` at its border-box rect with exact `size`.
   `src` must resolve to a local file (file:// or relative to the HTML file);
   `data:` URIs are materialized to temp files. Remote http(s) sources are an error
   (agents should download first — keeps the tool offline and deterministic).
2. **TABLE** — `<table>` → `add-table` at its rect with `rows` of plain cell text,
   followed by `set-text` ops (`"cell":[r,c]`) for cells whose resolved style
   differs from the table default. Column widths/row heights: v1 uses PowerPoint's
   even split; cell geometry fidelity is out of scope for v1.
3. **TEXT BLOCK** — `<p>`, `<h1>`–`<h6>` → one `add-shape kind=textbox` per block.
   A `<ul>`/`<ol>` is ONE textbox; each `<li>` is a paragraph with `bullet:true`
   (nested lists → `level` = nesting depth). Box = border-box rect; CSS padding maps
   to textbox `insets`. See "Text model" below.
4. **BOX** — any element painting a background (non-transparent background-color or
   a linear-gradient background-image) or a visible border → `add-shape` rect
   (`roundRect` when border-radius > 0) with `fill`/`gradient`/`line_*`. Children
   are still walked (a card div emits its box, then its text children as separate
   textboxes on top). A BOX never carries text itself — text always lives in its
   own textbox so geometry stays measured, not guessed.
5. Everything else (layout-only wrappers, flex containers) emits nothing; children
   are walked.

`<body>` background, when painted, becomes a full-slide BOX first (the back layer).
Elements that are invisible (`display:none`, `visibility:hidden`, zero-area,
`opacity:0`) emit nothing.

## Text model

Per text block:
- **Vertical anchor** is measured, not guessed: the rendered text ink box (a Range
  over the element's contents) is compared to the element's content box. A box that
  hugs its text → top (the PPTX default, no `anchor` emitted); text floating with
  roughly equal slack above and below → `anchor: MIDDLE`; text pinned to the bottom
  → `anchor: BOTTOM`. This captures flex/grid centering, explicit heights, and any
  other layout the browser resolved, with no per-property CSS rules to maintain.
- **Paragraph properties** from computed style: `alignment` (text-align left/center/
  right/justify → LEFT/CENTER/RIGHT/JUSTIFY), `line_spacing` = computed line-height
  px × 0.75 (Pt), `space_before`/`space_after` = 0 for v1 (inter-block spacing is
  already encoded in measured box positions; one box per block).
- **Runs**: walk inline descendants; each text node's effective style is resolved
  from its ancestors — `bold` (font-weight ≥ 600), `italic`, `underline`
  (text-decoration), `color` (computed rgb → hex), `font_size` (px × 0.75),
  `font_name` (first concrete family in the stack; generic fallbacks map
  sans-serif→Arial, serif→Georgia, monospace→Courier New). Adjacent runs with equal
  style are merged. A block whose runs are uniform emits plain `text` + shape-level
  font keys instead of `runs` (smaller, more readable patches).
- `<br>` keeps run flow within one paragraph (PPTX has no soft break in this model;
  v1 splits into separate paragraphs — same visual result given measured box width).
- Whitespace is collapsed the way the browser collapsed it (read from rendered text
  content, not source HTML).
- Textbox `insets` = computed padding (px→in) per side; `word_wrap` stays on.
  PPTX re-wraps text itself: same box width + same font ≈ same wrap; residual drift
  is caught by deck.py's post-apply lint (overflow detection), which is the
  designed safety net, not an accident.

## Style mapping

| CSS (computed) | patch key |
|---|---|
| background-color (α>0) | `fill` (hex; alpha dropped, α<1 emits a warning); border-only boxes get `fill:"none"` (hollow) |
| background-image: linear-gradient(...) | `gradient {colors, positions, angle}`; CSS deg → pptx angle = (90 − cssdeg) mod 360 (CSS grows clockwise from north; python-pptx counterclockwise from east) |
| border (uniform) | `line_color`, `line_width` (px×0.75 pt); dashed/dotted → `line_dash` |
| border (non-uniform, e.g. accent bar) | one `add-shape kind=line` per painted side, inset by half the width |
| border-radius > 0 | shape kind `rounded_rect` + `adjustments:[radius / min(w,h)]` (true radius, clamped to 0.5) |
| transform: rotate(θ) / writing-mode vertical | `rotation` with pre-rotation box math (see learnings) |
| flex/grid centering, explicit height (text floats in a taller box) | `anchor` MIDDLE/BOTTOM, measured from the rendered text's position in its box |
| background-image: url(...) (body or any box) | `add-picture` under the element's children; `background-size: cover/contain` honored |
| object-fit: cover / contain on `<img>` | `crop:[l,t,r,b]` source fractions (cover) or letterboxed target rect (contain), computed from the file with PIL |
| `<ol>` | paragraphs with `bullet:"number"` (a:buAutoNum); `<ul>` → `bullet:true` |
| table cell backgrounds | uniform → `fill`; mixed → `fills` row-major grid; column widths from the first row's measured rects → `col_widths` |
| theme artifacts | every box/picture gets `shadow:false` (browsers don't draw PPT's theme shadow); tables get `first_row:false, banding:false` |
| box-shadow, filters, letter-spacing | out of scope v1 (warning for shadows) |

## deck.py additions required

Two new style keys on `add-shape`/`set-style` (shared `_apply_style_keys`):
- `insets: [l, t, r, b]` (inches) → text-frame internal margins. CSS padding maps
  here 1:1; html2patch emits explicit insets on every textbox (PowerPoint's
  defaults are 0.1"/0.05", not zero).
- `adjustments: [f, ...]` → shape adjustment values (e.g. roundRect corner radius
  as a fraction of min dimension).

## Output

```
{"ops": [ ... ]}
```

- `--slide N`: ops target existing slide N.
- Default (with `--deck`): a leading `{"op":"add-slide","layout":<--layout>}` per
  input file; subsequent ops target the appended index (deck.py's validator already
  supports referencing slides created earlier in the same patch). Multiple HTML
  files in one invocation → one combined atomic patch, consecutive new slides.
- Every emitted shape gets a deterministic `name` (`h2p-<file#>-<seq>` by default,
  `--prefix` to override) so follow-up patches and humans can target them.
- Warnings (unsupported CSS, dropped alpha, remote images) go to stderr; `--strict`
  turns them into errors. The patch itself is written to `-o` or stdout.

## CLI

```bash
# create new slides in an existing (template) deck
python html2patch.py slide1.html slide2.html --deck deck.pptx --layout Blank -o patch.json
python deck.py deck.pptx apply patch.json -o out.pptx --fix --render img/

# decorate an existing slide
python html2patch.py overlay.html --deck deck.pptx --slide 3 -o patch.json

# no deck handy: assert geometry explicitly (inches)
python html2patch.py slide.html --size 13.333x7.5 --slide 0 -o patch.json
```

## Dependencies

`playwright` (Python) + its Chromium, strictly optional to hands-on-deck core. `--deck`
additionally uses python-pptx (already a core dep). No Node, no PptxGenJS, no sharp.

## Verification contract

html2patch's own job ends at a *geometrically faithful patch*. End-to-end fidelity
is verified the hands-on-deck way: `apply --fix --render`, then look. The test suite
asserts (a) measured geometry round-trips within 0.02 in, (b) run styling
round-trips through `inspect`, (c) bullets/levels, gradients, borders, insets,
tables land as specified, (d) a patch applied to a non-blank template deck composes
with `add-slide`.
