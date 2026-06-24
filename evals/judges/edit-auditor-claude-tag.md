# Auditor: review-and-fix scoring against a known defect manifest (forensic)

Task-specific auditor for `tasks/edit-claude-tag-review-loop.md`. Unlike the design judges,
this auditor **holds the answer key**: the input deck has exactly five known defects with
ground-truth fixes. Its job is not taste — it is to determine, per arm, which defects were
**found**, **fixed**, and **verified**, and to grade the arm on detection and repair. It runs
on one arm or both; with two arms, score each independently against the same manifest, then
name the cleaner review-and-fix pass.

The eval's whole point is that the defects are **graded by detectability**. Weight your read
accordingly: catching a linter-flagged defect is table stakes; catching a defect the tool is
silent on is the real signal of a render-and-look pass.

Placeholders: `{{JUDGING_DIR}}`. The auditor additionally needs each arm's `final.pptx` copied
into the judging dir (`{{JUDGING_DIR}}/file-A.pptx`, and `file-B.pptx` if two arms), the
arm's `defects.json`, and renders of the original buggy deck (`orig-A/*.jpg`) and the edited
deck (`deck-A/*.jpg`), assignment-matched by the orchestrator. The auditor must not learn arm
names.

> Indices below are **0-based** (skill convention); the parenthetical is the **1-based** number
> a human would say. Do not penalize an arm for using either as long as it maps them correctly
> — see the indexing trap.

## The answer key (the five known defects)

| # | slide | defect | root cause | linter signal? | ground-truth fix |
|---|---|---|---|---|---|
| D1 | 9 (10th) | hero stat **"65" wraps to two lines** | browser-measured box (3.5") narrower than PPTX renders "65" at 210pt → wraps | ❌ none (it wrapped *inside* the box; no overflow flag) | widen the box / prevent wrap — `resize s16 → [5.6, 2.48]` |
| D2 | 6 (7th) | **timeline titles invisible** (white text on white card) | `rgba(255,255,255,.06)` card flattened to opaque white on compile | ❌ none | darken the card — `set-style s18 fill 34153A` |
| D3 | 4 & 7 (5th, 8th) | **header bar + member icons stop ~2.8" short of the edge** | `.chead{width:934px}` hardcoded for a sidebar layout, not the full-width (1204px) slides | ❌ none today (this is the alignment-linter gap; the 3.1.0 near-miss linter only emits weak advisories here) | resize header bar to the slide edge (`s14 → [12.54, 0.67]`) and move the member pill to the right margin (the `fixheaders` moves of `s19–s25`) |
| D4 | 0 (1st) | headline **"team" clipped behind the chat card** | headline right edge (9.63") overlaps the card corner (8.19", 3.03") | ⚠️ partial — `covered_by` flag fires | move the card cluster `s18–s34` down ~0.55" (so the card top clears the headline) |
| D5 | 11 (12th) | closer text **overflows the right edge by 0.14"** | centered box 13.6" wide at x=-0.13 → right edge 13.47 > 13.33 | ✅ yes — `slide_overflow_right` fires | recenter to the canvas — move to x=0 and resize width to 13.33 (`s15`, `s16`) |

Detectability ladder — read each arm's result through this lens:
- **D5** the linter catches outright → tests *does the arm act on linter output at all?* (floor)
- **D4** the linter half-catches (`covered_by`) → tests *does the arm take an ambiguous flag seriously and zoom in?*
- **D1, D2** no linter signal → tests *did the arm actually render and look, not declare success from a clean patch?* (the discriminating slides)
- **D3** needs a capability that barely exists (alignment/contrast linter) → **stretch/upper-bound**; a strong arm finds it by eye. Score it, but treat a miss as expected, not a failure.

## Two traps to score explicitly (observed in the original session)

- **Indexing.** A request that says "slide 8" means 0-based index 7. Check the arm's
  `defects.json` and edits: did it operate on the correct slide, or did it off-by-one? An arm
  that fixed slide index 6 when it meant the 8th slide has an indexing bug even if the words
  look right.
- **Verification framing.** A positional fix (D4) re-rendered with a *different* crop window
  than the diagnosis can look fixed without being fixed. Inspect the arm's reported
  verification method (`defects.json` `verified` field + its final message): did it verify D4
  with a **matched window**, or did it re-crop and possibly fool itself? Score the
  *verification method*, not just the final pixels.

## Materials

- Arm A: original buggy renders `{{JUDGING_DIR}}/orig-A/*.jpg` (12), edited renders
  `{{JUDGING_DIR}}/deck-A/*.jpg` (12), `{{JUDGING_DIR}}/file-A.pptx`, `{{JUDGING_DIR}}/defects-A.json`
- Arm B (if present): same set with `-B`.

## What to do (per arm)

1. **Verify scope.** Confirm `file-?.pptx` still has exactly 12 slides and the deck's content
   is intact — flag any content edit, deletion, or slide-count change as collateral damage.
2. **Score each of D1–D5** as one of: **fixed** (the defect is gone in the edited render and the
   geometry matches the ground-truth intent — exact ground-truth coordinates are not required,
   the defect simply must be genuinely resolved), **attempted-but-broken** (touched the right
   slide/shape but the defect remains or a new one appeared), or **missed** (no action). Use the
   edited renders as primary evidence; open `file-?.pptx` with python-pptx / `deck.py inspect`
   to confirm the actual geometry where the render is ambiguous (e.g. measure that D1's stat box
   is now wide enough to hold "65" on one line, that D2's card fill is dark, that D5's box right
   edge ≤ 13.33"). For D3, measure whether the header bar/icons now reach the slide edge on both
   slides 4 and 7.
3. **Detection vs. fix.** Cross-check `defects.json` against your findings: did the arm *report*
   a defect it failed to fix, or *fix* one it never reported? Detection without repair still
   earns partial credit (it saw the problem); silent correct fixes count as found.
4. **Score the two traps** (indexing, verification framing) explicitly as above.
5. **Regressions.** Diff each original/edited render pair (luminance-only structural delta:
   grayscale, resize ~160×90, mean absolute difference). Only the five defect slides should
   change meaningfully; flag any other slide that moved.

## Output (final message)

A markdown audit report:
- A per-arm **defect scorecard**: D1–D5 → {fixed / attempted-but-broken / missed} → one-line
  evidence (which render, which measurement). Compute a headline score = defects fixed / 5 with
  partial credit noted, and call out separately how the arm did on the **silent** defects
  (D1, D2) vs the **flagged** ones (D4, D5), since that is the eval's real discriminator.
- The **trap verdicts**: indexing correct? verification method sound (matched-window for D4)?
- **Regression** findings (any non-defect slide that changed) with the structural-delta numbers.
- A final per-arm grade (A–F) with one paragraph. With two arms, name which did the cleaner
  review-and-fix pass on the merits (detection breadth, fix correctness, verification rigor),
  and note whether the difference came from the silent defects or the flagged ones.

Be exact and cite slide indices (note 0- vs 1-based each time).
