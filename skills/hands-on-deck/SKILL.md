---
name: hands-on-deck
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes reading, analyzing, or extracting content from presentations; editing text, images, styles, or layout of existing decks; creating new slides, shapes, tables, or pictures; reordering, duplicating, or merging slides across decks; and verifying decks visually. Trigger whenever the user mentions a deck, slides, a presentation, or a .pptx filename."
---

# hands-on-deck — agent-native PPTX manipulation

One tool does everything: `scripts/deck.py`. You write a JSON *patch* describing your edits; deck.py validates and executes it atomically, then lints the result. You never need to open slide XML for routine work.

**Golden rules** (the tool enforces and repeats these):
- ALL slide indices are 0-based, everywhere. Shapes are addressed by stable native ids (`s12`) or unique shape name.
- Positions and sizes are inches from the slide's top-left.
- Patches are atomic: every op is pre-validated (all errors reported at once, with the slide's real shape listing), then applied all-or-nothing.

**Full reference**: `python scripts/deck.py docs` — prints every op with semantics and recipes; needs no file. Read it before writing your first patch.

## The edit loop

```bash
# 1. READ — cheap orientation first, full JSON only when about to write a patch
python scripts/deck.py deck.pptx inspect --slide 3 --brief   # one line per shape
python scripts/deck.py deck.pptx inspect --slide 3           # full JSON: text+formatting, image rIds, geometry, issues
python scripts/deck.py deck.pptx inspect --issues            # only shapes with geometric problems

# 2. WRITE — one patch, many ops, atomic; chain repair + render of touched slides
python scripts/deck.py deck.pptx apply patch.json -o out.pptx --fix --render img/

# 3. VERIFY — look at what you changed
#    (render names files slide-<0-based-index>.jpg to match inspect)
python scripts/deck.py deck.pptx diff out.pptx               # structural changelog, no rendering
```

`--fix` runs deterministic geometry repair on the slides you touched (grows/shrinks/nudges overflowing text; never auto-moves pictures). Its *residue* report lists what still needs judgment, with a suggested op. Always look at the rendered image before declaring success.

## Patch ops at a glance

```json
{"ops": [
  {"op":"set-text",    "slide":3, "shape":"s12", "text":["Title", "Subtitle"]},
  {"op":"swap-image",  "slide":3, "shape":"s9",  "image":"/abs/new.png"},
  {"op":"swap-image",  "media":"image13.png",    "image":"/abs/logo.png"},
  {"op":"replace-text","scope":"deck", "from":"Old Name", "to":"New Name"},
  {"op":"replace-color","scope":"deck", "from":"E8A33D", "to":"F8DE6E"},
  {"op":"set-theme",   "colors":{"accent1":"BB7B19"}, "fonts":{"major":"Georgia"}},
  {"op":"set-props",   "title":"Q3 Review", "author":"Acme"},
  {"op":"set-slide",   "slide":3, "hidden":true, "background":"0F5258"},
  {"op":"set-slide",   "slide":3, "transition":{"type":"fade","speed":"med"}},
  {"op":"set-notes",   "slide":3, "notes":"speaker notes"},
  {"op":"move",        "slide":3, "shape":"s12", "to":[1.0,2.5]},
  {"op":"resize",      "slide":3, "shape":"s12", "size":[4.0,1.5]},
  {"op":"set-style",   "slide":3, "shape":"s12", "font_size":18, "fill":"0B3D3A", "rotation":6, "anchor":"MIDDLE"},
  {"op":"delete",      "slide":3, "shape":"s12"},
  {"op":"duplicate",   "slide":3, "shape":"s12", "offset":[0,1.2], "text":["Fourth pillar"]},
  {"op":"copy-shape",  "from_slide":8, "shape":"s12", "slide":3, "at":[1.0,2.0]},
  {"op":"reorder",     "slide":3, "shape":"s12", "z":"back"},
  {"op":"add-shape",   "slide":3, "kind":"textbox", "at":[1,2], "size":[4,1.5], "text":["…"], "name":"card"},
  {"op":"add-picture", "slide":3, "image":"/abs/img.png", "at":[1,2], "width":4},
  {"op":"add-table",   "slide":3, "at":[1,2], "size":[8,3], "rows":[["A","B"],["1","2"]]},
  {"op":"add-slide",   "layout":"Blank", "at":5},
  {"op":"add-row",     "slide":3, "shape":"s12", "cells":["a","b"]},
  {"op":"delete-col",  "slide":3, "shape":"s12", "col":1}
]}
```

Key semantics (details in `docs`):
- **Vertical centering is `anchor`, not `alignment`**: `alignment` (set-text) is HORIZONTAL; to center text in a box taller than the text, set `"anchor":"MIDDLE"` (TOP|MIDDLE|BOTTOM) on set-style or set-text. `inspect` reports it when it's not the default top. html2patch sets it automatically from the slide's CSS (flex/grid centering).
- **set-text inherits formatting**: new paragraph *i* inherits ALL formatting of old paragraph *i* — pass plain strings for routine replacement. Pass objects to override (`{"text":"Big","font_size":28}`), or `"runs"` for mixed in-paragraph formatting. Table cells: add `"cell":[row,col]`.
- **Prefer duplicate/copy-shape over add-shape** when a styled donor exists — new shapes start from PowerPoint defaults, not the deck's design language.
- **add-shape `"name"`** lets later ops in the same patch target the shape it creates.
- **Transitions only on request**: never add slide transitions unless the user explicitly asks — applied uninvited or inconsistently they're annoying, not polish. When asked, default to one subtle type (fade) across the whole deck.
- **Alignment is linted, advisory**: `inspect`/`apply` report a `misaligned` issue when a shape's load-bearing edge sits a hair (0.03"–0.15") off a gridline ≥3 other shapes share — the near-miss that reads as sloppy. Act on it (move/resize the edge to the named coordinate) or ignore it: intentional asymmetry is real, so it never fails the build, and apply reports only a *newly introduced* near-miss.
- Keep new text comparable in length to the old, or let `fix` repair the overflow.

## Creating slides from HTML (html2patch)

When a slide should be DESIGNED from scratch — free-form layout, no styled
donor to duplicate — write it as HTML/CSS and compile it into a patch. The
browser is used as a measuring engine; the output is ordinary deck.py ops.

**Before writing any slide HTML, read [designing-slides.md](designing-slides.md)**
— how to design for this medium: subject-derived palettes, refusing the
default AI-deck looks, the token plan, projection type sizes, and what
survives compilation. The pipeline below is mechanical; that file is taste.

```bash
python scripts/html2patch.py slide.html --deck deck.pptx --layout Blank -o patch.json
python scripts/deck.py deck.pptx apply patch.json -o out.pptx --render img/
```

- The `<body>` is the slide at 96px/inch: 16:9 → `width:1280px; height:720px`.
  Content overflowing the body is a compile error.
- Text lives in `<p>`, `<h1>`–`<h6>`, `<ul>`/`<ol>` (numbered), `<table>`, or
  any element with inline-only content (figcaption, blockquote…). Divs with
  background / border / border-radius / linear-gradient become styled rects;
  `<img>` becomes a picture (local files only; `object-fit: cover` becomes a
  real crop); `<table>` becomes a real table with per-cell fills and measured
  column widths.
- Inline `<b>`/`<i>`/`<u>`/`<span style>` become formatted runs; `<a href>`
  becomes a real hyperlink. CSS padding
  maps to text insets; `text-transform`, `transform:rotate`, and vertical
  `writing-mode` are honored. Use installed fonts (Arial, Georgia, …).
- By default each HTML file appends a new slide (`add-slide`, `--layout` picks
  the template layout); `--slide N` targets an existing slide instead. Multiple
  files compile into ONE atomic patch.
- Apply compiled patches WITHOUT `--fix` first — geometry is browser-measured,
  so look at the render before repairing; run `fix` only if the render shows
  a real problem.
- NEVER dismiss an overflow or `covered_by` flag on display-size text without
  zooming that exact shape: `render --slide N --crop l,t,w,h --scale 2`.
  Serif faces wrap differently in PowerPoint than in the browser — a clipped
  last line is invisible at thumbnail size, especially when it falls behind
  a picture.
- Extra dependency: `pip install playwright && playwright install chromium`.

## Deck structure (subcommands, not patch ops)

```bash
python scripts/deck.py deck.pptx slides 0,3,3,5 -o out.pptx        # keep these, in this order; repeat = duplicate
python scripts/deck.py deck.pptx merge module.pptx --slides 0,2 --at 12 -o out.pptx
python scripts/deck.py deck.pptx merge --list-layouts              # choose a layout for imported slides
```

## Building a deck from a template (the human workflow)

1. `slides` the template down to the target slide sequence (duplicating repeated layouts).
2. `merge` in any reusable modules.
3. One `replace-text` patch for global renames (scope `master` catches footers).
4. Per-slide patches: `inspect --slide N` for shape ids, then `set-text` / `swap-image` only what you name — nothing else is touched.
5. `apply --fix --render img/` and look at every touched slide.

## Verification

- `render -o img/ --slide 3,7` — JPGs named `slide-<index>.jpg`; `--crop l,t,w,h --scale 2` zooms a region (inches, same coordinates as inspect). Hidden slides render only when explicitly listed.
- `diff other.pptx` — text/geometry/media/notes changelog without rendering.
- Thumbnail grids for whole-deck review: `python scripts/thumbnail.py deck.pptx --cols 4` (optional second arg = output filename prefix).
- **Doubled-logo gate** (run on every built deck): `python scripts/check_logos.py deck.pptx` — flags any slide where a full-bleed / footer-covering picture is stacked on a logo/footer the master or layout already draws. Structural and deterministic (no rendering), so it's safe in CI and as a post-build workflow step. Exit 0 = clean, 1 = risk (prints slide + fix), 2 = error. `--json` for machine output; `--band`/`--coverage` tune the zone. This catches the defect image-only QA can't see — the logo baked into a PNG *and* drawn by the master. See "Never bake the logo/footer into a full-bleed image" in `designing-slides.md`.

### Reviewing more than a slide or two — fan out, don't accumulate

Rendered images and full `inspect` dumps are the heaviest things you put in context, and they never leave it: one agent that builds the deck and then reviews every slide re-reads every image and dump it has ever seen on every later turn, so cost grows with the square of the work. The fix is context discipline, not fewer checks.

**When more than a slide or two needs visual review, AND you have a sub-agent / Task tool available, fan the review out — one sub-agent per slide (or per small batch):**

- Give each sub-agent a **fresh, minimal context**: the deck path + that one slide's render + its `inspect --slide N`. NOT your build history, the other slides, or their renders. This is the whole point — each review context stays small, and they run in parallel.
- Each sub-agent returns a **verdict only** — `pass`, or the specific defects with shape ids and the `fix`/patch ops to apply. It does not hand the image back to you.
- You apply the fixes from the verdicts, then (if needed) fan out one more round on only the slides that changed.

**Even during the build, don't hoard renders in your own thread.** Look at a render, act on it, and move on — don't carry a growing pile of slide JPGs forward. If you're visually checking many slides, delegate the looking so the images live in the sub-agents' contexts, not yours.

No sub-agent tool available? Render and review in small batches and drop earlier slides' images from context once you've judged them — same principle, manual.

## Escape hatch

When no op expresses the change (animations, exotic effects):

```bash
python scripts/deck.py deck.pptx xml get --slide 5 -o slide5.xml   # pretty-printed, editable
python scripts/deck.py deck.pptx xml set slide5.xml --slide 5 -o out.pptx   # parse-checked, lint-checked
```

Out of scope by design (use the escape hatch or PowerPoint itself): creating native charts, shape animations (slide transitions ARE covered — `set-slide` `"transition"`; verify them with `inspect`/`diff`, since renders are static), embedded video/OLE, merged table cells.

## Dependencies

- Python 3.9+ with `python-pptx` and `Pillow` (`pip install python-pptx Pillow`)
- For `render` and thumbnails: LibreOffice (`soffice`) and Poppler (`pdftoppm`)
