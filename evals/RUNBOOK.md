# RUNBOOK — orchestrating a hands-on-deck eval

You are the orchestrating agent. You launch builder subagents, verify their
deliverables, blind the renders, launch judge subagents, unblind, tally, and write
the results log. You never build or judge anything yourself.

Two invariants make the whole thing valid; protect them above all:

1. **Symmetry.** Both arms get the same brief text, the same model, the same phase
   instructions, the same deliverable list. The ONLY difference is the toolchain
   block. If you nudge one arm mid-run (a retry, a clarification), give the other
   arm the identical nudge.
2. **Blindness.** Judges see only the blinded judging directory. Never put arm
   names, tool names, build reports, working paths, or the `.assignment` file in a
   judge's prompt — and do not read `.assignment` yourself until every verdict is in.

## Phase 0 — setup

1. Pick a run root, e.g. `/tmp/hands-on-deck-eval-<task>-<date>/`, with one private
   subdirectory per arm: `<root>/hands-on-deck/` and `<root>/<baseline>/`, plus
   `<root>/judging/`.
2. Toolchains on disk:
   - hands-on-deck arm: this repo's `skills/hands-on-deck/` (use the checkout under test).
   - baseline arm: e.g. Anthropic's pptx skill —
     `git clone --depth 1 --filter=blob:none --sparse https://github.com/anthropics/skills /tmp/anthropic-skills && cd /tmp/anthropic-skills && git sparse-checkout set skills/pptx`
3. Dependencies: LibreOffice (`soffice`) + Poppler (`pdftoppm`) for rendering;
   hands-on-deck arm needs `python-pptx Pillow` and (for the HTML create path)
   `playwright` + Chromium; the baseline installs whatever its skill prescribes.
4. For the **edit** task only: place each arm's input deck at `<armdir>/deck.pptx`
   and render the originals now (you'll need pre-edit images for the judges).
   Record whether this is an own-deck or same-deck setup (see the task file).
5. Choose and record the builder model — the same for both arms. Judges should run
   on a strong model regardless of what the builders run on.

## Phase 1 — builders

Take the task brief from `tasks/`, fill `{{WORKDIR}}` with the arm's directory, and
fill `{{TOOLCHAIN_BLOCK}}` per arm:

**hands-on-deck arm:**

> Use the **hands-on-deck** skill at `<path to skills/hands-on-deck>`. Start by reading its
> `SKILL.md` in full and follow everything it tells you to read and do (including
> any design guidance it references). All pptx work goes through this skill's
> scripts. Do not use any other pptx skill, library, or hand-written OOXML beyond
> what the skill prescribes.

**baseline arm (Anthropic pptx skill):**

> Use the **pptx** skill at `<path to skills/pptx>`. Start by reading its
> `SKILL.md` in full and follow everything it tells you to read and do (including
> the guides it references). All pptx work goes through this skill's prescribed
> tooling and workflow. Do not use any other pptx skill or toolchain beyond what
> the skill prescribes. You may install anything the skill's workflow requires
> (e.g. `npm install pptxgenjs` in your working directory). Do not look in any
> hands-on-deck directory.

Launch both builders **in parallel, in the background**. Expect 5–15 minutes each.
If a builder stalls past your watchdog, check its working directory before calling
it failed — in past rounds a stalled agent had already written every deliverable.

## Phase 2 — verify deliverables

For each arm: `final.pptx` exists, slide count is exactly what the brief demands
(check with python-pptx or `deck.py inspect`; do NOT open the deck in anything that
could rewrite it), and renders exist. Then **re-render both decks yourself with one
neutral command** so judges compare renderers-equal images:

```bash
soffice --headless --convert-to pdf --outdir <armdir>/neutral <armdir>/final.pptx
pdftoppm -jpeg -r 110 <armdir>/neutral/final.pdf <armdir>/neutral/slide
```

If an arm missed a hard deliverable (wrong slide count, no real downloaded asset
where required), note it in the results log — do not fix it for them.

## Phase 3 — blind

```bash
python evals/scripts/blind.py setup <root>/judging \
    --arm hands-on-deck=<root>/hands-on-deck/neutral \
    --arm <baseline>=<root>/<baseline>/neutral
# edit task: add --orig hands-on-deck=... --orig <baseline>=... for the pre-edit renders
```

From here until tally, refer to the decks only as A and B.

## Phase 4 — judges

All judges launch in parallel. Fill the placeholders from the judge file headers;
every judge in a panel gets identical text except the viewing-order block.

**Create task — three copies of `judges/create-judge.md`:**

- Judge 1: `View Deck A first, then Deck B. Do not favor whichever you saw first or second.`
- Judge 2: `View Deck B first, then Deck A. Do not favor whichever you saw first or second.`
- Judge 3: `View them INTERLEAVED pairwise: A/01 then B/01, A/02 then B/02, … comparing slide-for-slide as you go. (The two decks may sequence content differently — that's fine; judge each as its own sequence too.)`

`{{BRIEF_SUMMARY}}` for `tasks/create-dynamic-workflows.md`:

> A 15-slide deck about Dynamic Workflows in Claude Code (announced at
> claude.com/blog/introducing-dynamic-workflows-in-claude-code), for an audience
> that already uses skills in Claude Code daily (no basics), required to download
> and use at least one genuine image from the blog post, quality bar: "a deck so
> good you can't stop flipping."

**Edit task — two copies of `judges/edit-judge.md` plus one `judges/edit-auditor.md`:**

- Judge 1: `View team A first (original then edited, comparing slide by slide), then team B. View EVERY image. Do not favor either viewing order.`
- Judge 2: `View team B first (original then edited, comparing slide by slide), then team A. View EVERY image. Do not favor either viewing order.`
- Auditor: no order rotation; it works pairwise by definition.

`{{EDIT_BRIEF_SUMMARY}}` is the two numbered edits from
`tasks/edit-retheme-and-insert.md` (palette table compressed to one sentence with
all eight hexes and the C3DCDE floor; insertion message with its prediction
framing; final slide count). `{{PALETTE_LINE}}` for the auditor:

> backgrounds #0F5258 / #4999A0; text #FFFEFB / #FFFFFF; secondary text #C3DCDE
> (floor — nothing dimmer allowed); accents #BB7B19, #F8DE6E, #2E8D23

**For `tasks/edit-two-world-reskin.md`** use `judges/edit-auditor-two-world.md`
instead of the generic auditor (it needs `file-A.pptx`/`file-B.pptx` copied into
the judging dir, assignment-matched). `{{EDIT_BRIEF_SUMMARY}}` compresses the task's
four parts: the slide-by-slide world map with both palettes and floors, the
keep-dark card rule + two-tone thread continuity, the three content edits on slides
14–17, speaker notes on all 18. Phase 0 note: the input deck is committed at
`evals/assets/two-world-reskin-input.pptx` — copy it to each arm's
`{{WORKDIR}}/deck.pptx`; this task is same-deck by construction. Add this line to
EVERY judge prompt for this task (the deck's content describes one of the
toolchains; both arms received identical content):

> The deck's subject matter is irrelevant and identical across both teams. Judge
> ONLY the quality of the edits — never the deck's claims or copy.

**For `tasks/edit-claude-tag-comment-revisions.md`** the scoring is answer-key-based, not
taste-based, so it does NOT use the design judges. Run only
`judges/edit-auditor-claude-tag-comment-revisions.md` (one copy per arm, or one copy covering
both — it holds the five-comment answer key and scores each arm independently). This auditor
needs, per arm, the edited renders (`deck-A/*.jpg`, **13**), the *original input* renders
(`orig-A/*.jpg`, 12), the arm's `final.pptx` (copied in as `file-A.pptx`), and the arm's
`changes.json` (copied in as `changes-A.json`) — all assignment-matched, arm names hidden.
Phase 0 note: this is a **same-deck** task — copy the committed input
`evals/assets/claude-tag-comment-revisions/comment-revisions-input.pptx` to each arm's
`{{WORKDIR}}/deck.pptx` and render the originals now (the auditor needs the 12 pre-edit images).
The final deck is **13 slides** (one comment adds a slide). Because there is a fixed answer key,
blinding/order-rotation matters less here, but still hide arm names so the write-up stays neutral.
This task also runs fine **single-arm** as a pure capability probe (one arm, score the C1–C5
scorecard). The headline discriminator is the **1-based → 0-based indexing trap** (the reviewer
named slides 4/6/8/11; the arm must act on indices 3/5/7/10, and insert C4's new slide *after
index 7*) — weight correct targeting heaviest, then execution fidelity, the new slide's quality,
and verification rigor; do NOT weight discovery breadth (there is a checklist). C2/C3 are the two
revisions the 3.1.0 alignment linter can *partially* assist on; C1/C4/C5 are render-and-look only.

## Phase 5 — unblind and tally

Only after every judge's final verdict is in:

```bash
python evals/scripts/blind.py reveal <root>/judging
```

- **Create:** overall winner = majority of the three judges' overall verdicts.
  Report each judge's winner WITH margin (DECISIVE / CLEAR / NARROW) — "2–1 with
  two CLEAR" is a different result than "2–1 with three NARROW".
- **Edit:** report the two design verdicts and the auditor's fidelity grades as
  separate results — they measure different things and have disagreed before.
  Do not average them into one number.

## Phase 6 — findings become machinery

This phase is the reason the eval exists. For every defect any judge found in the
**hands-on-deck arm**, classify it:

- **tool gap** — no op/flag could express what the builder needed → new patch op or
  flag (never a new subcommand for a little thing)
- **lint gap** — the defect was visible in the output but nothing flagged it → new
  lint with exact geometry and a suggested fix
- **guidance gap** — the tool could do it and the lint fired, but the builder
  misused or dismissed it → SKILL.md / designing-slides.md rule
- **builder-model noise** — not reproducible, no action

Ship each fix with a regression test. A finding fixed in prompt wording alone will
recur next round; a finding turned into machinery never has.

## Phase 7 — record

Write `evals/results/<YYYY-MM-DD>-<task>.md`:

```markdown
# <task> — <date>
- hands-on-deck version: <version + commit>  |  baseline: <skill + commit>
- builder model: <model>  |  judge model: <model>  |  setup notes: <e.g. own-deck edit>
## Verdicts
<per-judge: winner, margin, one-line rationale; auditor grades if edit>
## Findings → actions
<defect → classification → issue/commit that shipped the fix>
```

Commit the results log. Heavy artifacts (decks, renders, transcripts) stay local.
