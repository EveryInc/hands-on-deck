# Task: edit — review-and-fix a deck's render-only defects

The realistic "this deck looks broken, find out why and fix it" job. A 12-slide deck was
built create-path (HTML → patch) and looked correct under browser measurement, but five
visual defects survived into the PPTX render. This task gives an arm the *buggy* deck with
**no list of what's wrong** and asks it to find and fix the defects itself. It measures the
one thing a build pipeline can't fake: does an agent holding this tool actually **render,
look, and repair** — or does it declare success from a clean-looking patch?

The defects are deliberately **graded by detectability** (some the linter catches, some it
half-catches, some have no machine signal at all and only a human-style render-and-look will
surface), so a strong arm separates from a weak one on the slides where the tool stays silent.

Input: each arm gets the same 12-slide deck at `{{WORKDIR}}/deck.pptx` (a "Claude Tag"
companion guide; canvas 13.33×7.5in, 16:9). This is a **same-deck** setup — both arms review
copies of the identical buggy input. The committed source is
`evals/assets/claude-tag-review-loop/claude-tag-v1-buggy.pptx`; the orchestrator copies it to
each arm's `{{WORKDIR}}/deck.pptx`.

Placeholders: `{{TOOLCHAIN_BLOCK}}`, `{{WORKDIR}}`.

---

You are a presentation designer doing a visual QA pass. A finished 12-slide deck is at
`{{WORKDIR}}/deck.pptx`. It was built from HTML and looked right in the browser, but the
PPTX render has visual defects. Your job is to **find every visual defect and fix it**, then
prove your fixes landed.

## The job

1. **Look at the whole deck.** Render every slide to an image and actually examine each one at
   a readable resolution. Do not assume a slide is fine because its source looked fine.
2. **Find the defects.** They include things like: a hero number that wraps when it shouldn't,
   text that is invisible because it sits on a same-colored fill, a header bar or row of icons
   that stops short of the slide edge, a headline clipped behind another element, and text that
   spills past the slide boundary. That list is illustrative, not exhaustive and not a
   checklist — discover what is actually wrong by looking.
3. **Use every signal the toolchain gives you.** If your tooling emits layout warnings
   (overflow, coverage, alignment), act on them — but understand that some defects produce **no
   warning at all** and are only visible in the render. Zoom in on anything ambiguous before
   you decide it is fine or dismiss a flag.
4. **Fix each defect** with the smallest correct change, preserving the deck's content,
   message, and slide count (12). You are repairing layout, not redesigning the deck.
5. **Verify.** Re-render and re-examine. Critically: verify a positional fix with the **same
   crop/zoom window you used to diagnose it** — re-cropping to a different window can hide
   whether the element actually moved. A fix is not done until you have seen it correct in the
   render.

If a defect is real but you cannot fully fix it with the tools available, say so explicitly
rather than papering over it.

## Your toolchain (use ONLY this)

{{TOOLCHAIN_BLOCK}}

## Process

1. Inspect the deck first — understand each slide's shapes, sizes, and positions.
2. Render and look at every slide. Run any linter/validator your toolchain provides.
3. Fix the defects you find.
4. Review: re-render every slide, re-examine, confirm each fix with a matched-window check, and
   confirm you introduced no new defects. At least one genuine review-and-fix round.

## Deliverables (in `{{WORKDIR}}`)

- `final.pptx` — the repaired 12-slide deck
- `img/` — rendered JPGs of all 12 final slides
- `contact-sheet.jpg` — thumbnail grid of the final deck
- `defects.json` — what you found and did, one object per defect:
  `{"slide": <0-based index>, "shape": "<id or description>", "defect": "<what was wrong>",
  "fix": "<the op(s) you applied>", "verified": "<how you confirmed it in the render>"}`
  Use **0-based** slide indices. If you also report human-facing (1-based) numbers, label them.

Your final message: a report of the defects you found, how you diagnosed each (linter vs.
render-and-look), the fix applied, how you verified it, and anything you judged broken but
could not fully repair. Do not ask questions; decide everything yourself. This is unattended.
