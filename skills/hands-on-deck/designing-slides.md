# Designing slides worth presenting

Read this before writing create-path HTML (`html2patch`). It is about the
design, not the tooling — SKILL.md covers the pipeline.

Work as if you are the deck designer a team hires after firing their template.
They didn't pay for "professional"; they paid for a deck that could only be
about *this* subject. Every slide you ship is a choice someone could disagree
with — if no one could, you haven't made one.

## Start from the subject, not from style

Before any visual decision, pin three things: what this deck is about (one
concrete subject), who is in the room, and what the deck must accomplish.
Then mine the subject's own world for the design: its materials, instruments,
diagrams, vocabulary, era, texture. A deck about shipping logistics has
manifests, container codes, route lines; a deck about a poetry archive does
not. Distinctiveness comes from the subject — style applied from outside is
what templates do.

## Know the default look, then refuse it

Unprompted AI decks cluster into a few recognizable looks: white slides with a
blue header bar and bullet lists; a near-black deck with one neon-gradient
accent, a giant statistic, and three rounded cards; cream-paper editorial
serif with a terracotta accent. Each is fine *when chosen*; none should be
where you land by gravity. If the brief specifies a direction, follow it
exactly. Where the brief is silent, your test is: would this design survive
having a different subject poured into it? If yes, it isn't designed yet.

The same applies to structure. Numbered section markers, eyebrow labels, and
divider slides should encode something true (a real sequence, a real taxonomy)
— not perform organization the content doesn't have.

## Plan in tokens before writing any HTML

First pass, before any slide file: write a compact design plan.

- **Palette** — 4–6 named hex values, derived from the subject. One color
  dominates (60–70% of visual weight); one is the sharp accent, spent
  sparingly. Name what each is *for*, not just what it is.
- **Type** — two or three roles: a display face with real character (used
  with restraint, at scale), a quiet text face, optionally a utility/mono
  face for labels, code, and data. PPTX has no webfonts — choose from faces
  that exist on the target machine (Georgia, Palatino, Arial, Helvetica,
  Verdana, Trebuchet MS, Courier New, Impact…) and let size, weight, case,
  and color do the differentiating work.
- **Grid** — the margins, eyebrow row, and footer line every slide will
  share. Decide the left margin ONCE, in pixels, and put it in a shared CSS
  block. Slides are viewed in sequence: a margin that jumps between slides
  reads as a mistake even when each slide is fine alone. `inspect`/`apply`
  enforce this by measurement: a `misaligned` issue means an edge landed a
  hair off a gridline its siblings share — the near-miss that reads as sloppy.
  Snap it to the reported coordinate.
- **Signature** — the one element this deck will be remembered by: a
  recurring motif, an unusual headline treatment, a diagram language, a
  full-bleed texture. One. Everything else stays quiet so it can speak.

Second pass: critique the plan before building. For each token ask whether
it's a choice for this brief or the thing you'd produce for any similar
prompt. Rework what fails. Only then write slide HTML, deriving every value
from the plan.

## A slide is not a webpage

- **One idea per slide.** A page scrolls; a slide is read in about six
  seconds from across a room. The headline should carry the idea by itself;
  everything else is evidence.
- **Type sizes are projection sizes.** Body text below ~16px (12pt) is
  invisible from the back of the room; captions/labels bottom out around
  13px. Headlines start where webpages stop — 40px is modest, 90px is a
  statement. When in doubt, cut words, not point size.
- **The canvas is fixed and overflow is a compile error.** Design *to* the
  1280×720 box: generous margins, deliberate empty space. Vertical balance
  matters — content that ends two-thirds up the slide leaves a dead zone
  that reads as unfinished.
- **The deck's "motion" is the page turn.** There is no scroll, no hover, no
  reveal. Rhythm comes from layout variation against a constant grid: a
  full-bleed moment after three structured slides, a near-empty statement
  slide after a dense one. Vary the layout, never the system.
- **Copy is design material.** Headlines in the deck's voice, labels that
  label, no filler. Slides force the discipline webpages let you skip: if a
  sentence doesn't help the room understand, it's decoration — cut it.

## Design with the compiler's grain

html2patch translates faithfully — but only what PPTX can hold. Spend your
craft on what survives:

- **Survives**: flex/grid/absolute layout, linear gradients, solid fills,
  uniform and partial borders (accent bars!), true corner radii, real tables
  with per-cell fills and measured column widths, `object-fit: cover` crops,
  `transform: rotate`, `text-transform`, numbered and bulleted lists, padding
  (becomes text insets), per-run bold/italic/color/size mixing, vertical
  centering of text in a taller box (flex/grid/explicit-height → an `anchor`).
- **Doesn't**: box-shadow, letter-spacing, text gradients, blend modes,
  filters, custom webfonts, animation of any kind. If a treatment depends on
  these, it will quietly vanish — design something that doesn't need it.
- Decorative depth that PPTX can't fake natively (textures, soft shadows,
  grain) can be **baked into images** and placed as full-bleed or cropped
  pictures. The compiler carries pixels faithfully; use that.

## Critique with your eyes, then remove one thing

The render loop is the mirror: compile, `--render`, and LOOK at every slide
image — never declare a deck done from the patch alone. Check it as a
sequence, not as nine separate compositions: do the margins hold? does the
eyebrow row sit at the same height? does the accent color appear on every
slide or pool on one? Heed the linter on overflows and collisions — and
before you call any flag a false positive, zoom the exact shape
(`render --crop l,t,w,h --scale 2`). Serif display headlines are where
browser and PowerPoint metrics diverge most, and a re-wrapped last line
hides at thumbnail size, especially when it falls behind a picture. Only a
clean zoom earns the words "false positive".

Then the last pass: find the least necessary element on each slide — the
extra divider, the third accent, the label nobody needs — and take it off.
A deck that survives that pass was designed; one that collapses was
decorated. Small geometry corrections (a 0.25" alignment drift, one run that
should be mono) are cheaper as deck.py ops on the compiled file than as a
recompile — that's what the patch engine is for.
