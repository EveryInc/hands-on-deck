# Task: edit — apply five reviewer comments to a finished deck (the indexing trap)

The realistic "the deck is done, now go apply the review" job. A reviewer flipped through a
finished 12-slide "Claude Tag" companion deck — styled as the **Slack desktop app** — and left
**five terse comments**, each naming a slide by the number a human counts (1-based). This task
hands an arm the deck plus those five comments and asks it to apply **every** requested revision
precisely, build one new slide from a vague creative ask, then prove each fix in the render. The
deck ends at **13 slides** (one comment asks to add a slide).

What it stresses that the other edit tasks don't: this task is about **precise execution against
an explicit checklist**, not discovery. Its sibling `tasks/edit-claude-tag-slack-fidelity.md`
hands the *same deck family* with **no checklist** and tests whether an arm can *discover* what
makes a UI mock read as fake. **THIS** task gives the arm an exact list of five reviewer comments
and tests whether it can (a) map each human 1-based slide reference to the correct 0-based index —
**the central trap** — (b) edit the right targets faithfully, (c) build a genuinely visual new
slide from a one-line creative brief, (d) carry a deck-wide consistency change, and (e) verify
each fix in the render rather than declaring success from a clean patch. The discriminator is
correctness + indexing + the inserted slide's quality + verification rigor — NOT breadth of
discovery.

Input: each arm gets the same 12-slide deck at `{{WORKDIR}}/deck.pptx` (a "Claude Tag" companion
guide styled as the Slack desktop app; canvas 13.33×7.5in, 16:9). This is a **same-deck** setup —
both arms edit copies of the identical input. The committed source is
`evals/assets/claude-tag-comment-revisions/comment-revisions-input.pptx`; the orchestrator copies
it to each arm's `{{WORKDIR}}/deck.pptx`.

Placeholders: `{{TOOLCHAIN_BLOCK}}`, `{{WORKDIR}}`.

---

You are a presentation designer. The 12-slide deck at `{{WORKDIR}}/deck.pptx` is a finished
"Claude Tag" companion guide, styled as the **Slack desktop app**. A reviewer flipped through it
and left five comments. Your job: **apply all five revisions**, preserve every other piece of
content and each slide's message, end at exactly **13 slides**, and **prove each fix in the
render.**

## The reviewer's comments (apply all five)

The reviewer refers to slides by the number a human counts while flipping — i.e. the **first
slide is "slide 1."** These are exactly as written:

1. **"can we have some organisation logos in the side bar in all slides?"**
2. On **slide 4**: **"fix alignment of text in avatar shapes"**
3. On **slide 6**: **"2 msgs should be aligned at the same level"**
4. On **slide 8**: **"add a new slide after this on how this unlocks multiplayer ai mode — make
   it visual"**
5. On **slide 11**: **"make this slide in the slack dark mode aesthetic"**

## How to read these

- **The slide numbers are how a human counts (1-based).** You are responsible for mapping each
  one to the slide you actually edit. If your tooling indexes slides from 0, "slide 4" is the
  *fourth* slide, not the slide at index 4. Get this right for every comment — editing the
  slide next to the intended one is a failure even if the words of the change are correct. For
  comment 4, "add a new slide **after** slide 8" means the new slide lands *between* the current
  slide 8 and the current slide 9.
- **Comment 1 is deck-wide.** "all slides" means the change must be visible and consistent on
  **every** slide that shows the sidebar — not just the first one. The far-left workspace rail
  currently shows generic `#` placeholder tiles; give them real organisation/workspace logo
  tiles. Leave the active top tile intact.
- **Comment 4 is open-ended and creative.** "make it visual" means design a real slide — a
  visual that conveys how Claude Tag unlocks **multiplayer AI** (AI leaving the private 1:1 chat
  and moving into the shared team channel: the whole team plus Claude in one thread, one context,
  one source of truth). It must be genuinely visual, not a bullet list, and on-brand with the
  rest of the deck. This is the only comment that adds a slide; the deck goes 12 → 13.
- **Comment 5 is a restyle, not a rewrite.** Convert the named slide into the deck's dark
  aesthetic (dark canvas, near-white text, coral accents) while keeping all of its content.
- Preserve everything else: do not drop copy, do not reorder unrelated slides, do not change
  slides the reviewer didn't mention (beyond the deck-wide comment 1).

## Your toolchain (use ONLY this)

{{TOOLCHAIN_BLOCK}}

## Process

1. **Inspect the deck first.** Understand each slide's shapes, the shared sidebar/rail
   component, the avatar shapes, the two message cards, and the slide whose aesthetic must flip.
   Build an explicit map from each reviewer comment (1-based) to the slide index you will edit.
2. **Render and look at every slide.** The render is ground truth — confirm which slide each
   comment actually refers to by looking, not by guessing from the number alone.
3. **Apply all five revisions.** For comment 1, fix the shared rail component so the logos land
   consistently on every slide. For comments 2 and 3, make the alignment actually read centered /
   level. For comment 4, build the new multiplayer slide and insert it in the correct position.
   For comment 5, transfer the dark aesthetic while preserving content.
4. **Review round.** Re-render every slide, re-examine, and confirm each fix is visible and
   correct in the rendered image — for the positional fixes (comments 2 and 3) re-confirm with
   the same crop/zoom window you used to diagnose them. Confirm the deck is exactly 13 slides, the
   new slide is in the right place, and nothing else regressed (no lost content, no overflow, no
   reordered slides). At least one genuine review-and-fix round.

If a revision is genuinely impossible with the tools available, say so explicitly rather than
papering over it.

## Deliverables (in `{{WORKDIR}}`)

- `final.pptx` — the edited deck, exactly **13 slides**
- `img/` — rendered JPGs of all 13 final slides
- `contact-sheet.jpg` — thumbnail grid of the final deck
- `changes.json` — what you did, **one object per revision**:
  `{"comment": "<verbatim comment>", "reviewer_slide": <1-based number as given>,
  "target_index": <0-based index you acted on>, "change": "<what you did>",
  "verified": "<how you confirmed it in the render>"}`
  Use **0-based** indices for `target_index`. For the deck-wide comment, list the affected slides
  or `"all"`. For the inserted slide, report the index the new slide ended up at.

Your final message: report each comment, the slide you mapped it to (call out 1-based-as-given
vs. the 0-based index you acted on), the change you made, and how you verified it in the render.
Do not ask questions; decide everything yourself. This is unattended.
