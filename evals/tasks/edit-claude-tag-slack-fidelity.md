# Task: edit — make a Slack-UI mock actually look like Slack

The realistic "this is supposed to look like our app — but anyone can tell it's a fake" job. A
12-slide deck mimics the **Slack desktop client** using only native PowerPoint shapes (no
screenshots), but in its current form it does not convincingly read as Slack: a whole class of
labels is missing, the reaction pills are wrong, the header chrome is off, the avatars are the
wrong shape, and several views are missing the composer. The deck is a *Slack-shaped wireframe
with the giveaways still in it.* This task hands an arm that wireframe with **no list of what's
wrong** and asks it to make it read as the real current Slack desktop UI, then prove the fixes
in the render.

What it stresses that the other edit tasks don't: this is a **component-level fidelity
redesign against a real-world reference** the arm must reconstruct from memory and from looking.
The defects are not arbitrary layout bugs — they are the specific things that make a UI mock
read as "not Slack" to anyone who uses Slack. It measures two distinct abilities: (a)
**discovery** — can a tool-holding agent figure out *what* gives a UI mock away, with no
checklist? — and (b) **execution + verification** — can it then carry a multi-slide,
shared-component fidelity pass and confirm it landed in the PPTX render, rather than declaring
success from a clean-looking patch? Some giveaways are textual content that silently vanished
at compile (visible only by rendering and looking); others are styling that's merely *wrong*
versus real Slack. A strong arm separates from a weak one on whether it actually renders, looks,
recognizes "that's not how Slack looks," and repairs — across all 12 slides consistently.

Input: each arm gets the same 12-slide deck at `{{WORKDIR}}/deck.pptx` (a "Claude Tag"
companion guide styled as the Slack desktop app; canvas 13.33×7.5in, 16:9). This is a
**same-deck** setup — both arms edit copies of the identical wireframe input. The committed
source is `evals/assets/claude-tag-slack-fidelity/slack-wireframe-input.pptx`; the orchestrator
copies it to each arm's `{{WORKDIR}}/deck.pptx`.

This is distinct from `tasks/edit-claude-tag-review-loop.md`: that task hands a *different*
buggy deck and tests find-and-fix of five generic render-only layout defects; **this** task
hands a Slack-UI wireframe and tests whether the arm can make a mock of a specific, well-known
real application read as the genuine article.

Placeholders: `{{TOOLCHAIN_BLOCK}}`, `{{WORKDIR}}`.

---

You are a presentation designer. The 12-slide deck at `{{WORKDIR}}/deck.pptx` is a mockup of
the **Slack desktop app**, built entirely from native shapes (workspace rail, channel sidebar,
channel header, message list, reactions, composer). It is meant to look like a real Slack
screen — but it does not yet. Anyone who uses Slack would immediately see it's a fake.

Your job: **make it read as the real, current Slack desktop UI.** Fix everything that gives it
away — at the component level, consistently across every slide — while preserving the deck's
content, message, and slide count (12). Then **prove your fixes in the render.**

## The job

1. **Know your reference.** You are matching the *current Slack desktop client*. Recall what a
   real Slack window actually looks like — the aubergine workspace rail and channel sidebar with
   readable channel and DM names, the workspace name at the top, a solid channel header with
   `#` + channel name, message rows with small rounded-square avatars, neutral reaction pills
   that show an emoji and a count, and a message composer pinned to the bottom of every channel
   view. Hold the deck to that.
2. **Look at the whole deck.** Render every slide to an image and examine each one at a
   readable resolution. Do not assume a slide is fine because its source looked fine — some of
   what's wrong is text that is simply **not present in the render** even though it's in the
   source. The render is ground truth.
3. **Find what gives it away.** Discover the tells by looking and by knowing Slack — there is no
   checklist. They range from missing/dropped text content, to colors and shapes that are merely
   *wrong* versus real Slack, to whole UI elements that are absent. Some are glaring on the
   first true "Slack view" slide; some are subtle chrome details. Treat anything that wouldn't
   appear in a real Slack screenshot as a tell.
4. **Fix them at the component level.** These are shared components reused across many slides
   (sidebar, header, reaction pill, avatar, composer). Fix the component everywhere it appears,
   not slide-by-slide one-offs — the result must be consistent across all 12 slides. Make the
   smallest correct change per tell; you are raising fidelity, not redesigning the deck. Keep
   all content, the message of each slide, and the 12-slide count.
5. **Verify in the render.** Re-render and re-examine every slide. A recovered label, a
   recolored pill, a resized header, an added composer — confirm each is actually visible and
   correct in the rendered image, zooming in where ambiguous. A fix is not done until you have
   seen it correct in the render. Re-confirm a positional or sizing fix with the same crop/zoom
   window you used to diagnose it.

If a tell is real but you cannot fully fix it with the tools available, say so explicitly
rather than papering over it.

## Your toolchain (use ONLY this)

{{TOOLCHAIN_BLOCK}}

## Process

1. Inspect the deck first — understand each slide's shapes, the shared components, and their
   sizes/positions/colors.
2. Render and look at every slide. Run any linter/validator your toolchain provides — but note
   that several of the most damaging tells produce **no warning at all** and are only visible by
   rendering and comparing against what real Slack looks like.
3. Fix the tells you find, at the component level so they land consistently across slides.
4. Review: re-render every slide, re-examine, confirm each fix is visible and correct in the
   render with a matched-window check, and confirm you introduced no new defects (no lost
   content, no overflow, no broken slide). At least one genuine review-and-fix round.

## Deliverables (in `{{WORKDIR}}`)

- `final.pptx` — the edited 12-slide deck
- `img/` — rendered JPGs of all 12 final slides
- `contact-sheet.jpg` — thumbnail grid of the final deck
- `changes.json` — what you found and did, one object per fidelity fix:
  `{"slide": <0-based index or "all chrome">, "component": "<sidebar | header | reaction-pill |
  avatar | composer | rail-tile | app-badge | …>", "tell": "<what gave it away vs real Slack>",
  "fix": "<the op(s) you applied>", "verified": "<how you confirmed it in the render>"}`
  Use **0-based** slide indices. For shared-component fixes, list the slides affected (or note
  "all chrome slides"). If you also report human-facing (1-based) numbers, label them.

Your final message: a report of the tells you found, how you diagnosed each (linter vs.
render-and-look vs. Slack knowledge), the fix applied, how you verified it in the render, and
anything you judged wrong but could not fully repair. Do not ask questions; decide everything
yourself. This is unattended.
