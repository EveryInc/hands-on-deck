# Auditor: comment-revision execution scoring against a known answer key (forensic)

Task-specific auditor for `tasks/edit-claude-tag-comment-revisions.md`. Like the slack-fidelity
auditor and unlike the design judges, this auditor **holds the answer key**: the five reviewer
comments, the slide each one truly targets (as a 0-based index), and the ground-truth revision
for each. Its job is not taste — it is to determine, per arm, whether each of the five comments
was applied to the **correct slide** and **faithfully executed**, then grade the arm on correct
targeting (the indexing trap), execution fidelity, the new slide's quality, and verification
rigor — **NOT** discovery breadth (there is a checklist; discovery is not being tested). It runs
on one arm or both; with two arms, score each independently against the same answer key, then
name the cleaner execution.

The eval's whole point is the **1-based → 0-based indexing trap**: the reviewer named slides the
way a human counts (slide 4, slide 6, slide 8, slide 11), and the arm had to act on the correct
0-based index (3, 5, 7, 10) — and for comment 4 insert **after index 7**, not after index 8. An
arm whose words look right but edited the slide next door has an indexing bug, and that is the
sharpest discriminator on this task. (This is the exact bug a real session hit; it is the point
of the eval.)

> Indices below are **0-based** (skill convention); the parenthetical is the **1-based** number a
> human would say. Slide-index map of the **input** deck (12 slides):
> `0`=Cover(s1, "Meet @Claude, now on your team.", dark marketing; left rail with blank `#`
> tiles; a DM card) · `1`=Premise(s2, "Tag it. Then go do something else.", dark marketing) ·
> `2`=What-is-this(s3, chrome) · `3`=One-Claude(s4, chrome; **MT** two-letter avatar, a header
> D/M/S facepile, sidebar DM tiles) · `4`=Memory(s5, "It remembers.", full-width chrome) ·
> `5`=Proactive(s6, "It speaks up first.", chrome; **two Claude message cards at different left
> edges/widths**) · `6`=Async(s7, "It works while you're away.", dark marketing timeline) ·
> `7`=How-it-works(s8, "Four moves, every time.", full-width chrome) · `8`=Admin(s9, "Powerful,
> and on a leash.", chrome) · `9`=Proof-stat(s10, "65%", dark marketing) · `10`=Availability(s11,
> "Claude Tag is live in beta.", **light** Slack-chrome announcement, four beta points) ·
> `11`=Closer(s12, "Your team just got bigger.", dark closer).

> **After the correct insertion (comment 4):** the deck is 13 slides; index `7` is still "Four
> moves, every time.", index `8` is the **new multiplayer slide**, index `9` is the old Admin
> slide ("Powerful, and on a leash."), and every slide after shifts up by one (old closer ends at
> index `12`).

Placeholders: `{{JUDGING_DIR}}`. The auditor additionally needs, per arm, the edited renders
(`deck-A/*.jpg`, 13), the *original input* renders (`orig-A/*.jpg`, 12), the arm's `final.pptx`
(copied in as `file-A.pptx`), and the arm's `changes.json` (copied in as `changes-A.json`) — all
assignment-matched by the orchestrator. The auditor must not learn arm names.

## The answer key (the five reviewer comments and their ground truth)

The reviewer's slide numbers are **1-based**. The correct target is the **0-based** index.
Detection column: ⚠️ = the 3.1.0 alignment linter can *partially* assist (within-slide near-miss
flag), ❌ = no machine signal — only render-and-look + judgment.

| # | comment (verbatim) | reviewer said | correct target (0-based) | detect | ground-truth revision |
|---|---|---|---|---|---|
| C1 | "can we have some organisation logos in the side bar in all slides?" | all slides | **every** slide's left workspace rail | ❌ | Replace the three generic `#` placeholder tiles in the far-left workspace-switcher rail with three **distinct organisation/workspace logo tiles** (distinct fills + white centered initials/marks), applied **consistently on every slide**. Leave the active top tile (the colorful hash) intact. |
| C2 | "fix alignment of text in avatar shapes" | slide 4 | **3** | ⚠️ | The initials inside the avatar squares are not optically centered (esp. the **MT** message avatar; also the header D/M/S facepile and the sidebar DM tiles). Re-center each initial vertically + horizontally within its square. |
| C3 | "2 msgs should be aligned at the same level" | slide 6 | **5** | ⚠️ | The two Claude message cards sit at different left edges / widths; align them to share one left edge (and width) so they read as a clean stacked conversation. Move each whole card (bg, avatar, labels, body, reactions) together — not just one element. |
| C4 | "add a new slide after this on how this unlocks multiplayer ai mode — make it visual" | after slide 8 | insert **after index 7** → new slide at index **8** | ❌ | Insert ONE new slide right after "Four moves, every time." conveying how Claude Tag unlocks **multiplayer AI**: AI leaves the private 1:1 chat and moves into the shared team channel — the whole team + Claude in one thread, one context, one source of truth. Must be **genuinely visual** (not a bullet list) and on-brand. Deck count 12 → 13. After insertion: index 7 still "Four moves", index 8 = new slide, index 9 = old Admin slide. |
| C5 | "make this slide in the slack dark mode aesthetic" | slide 11 | **10** | ❌ | Restyle the currently-**light** "Claude Tag is live in beta." slide into the deck's **dark** aesthetic (dark canvas, near-white text, coral accents) while **preserving all content** (the four beta points). |

## Two traps to score explicitly

1. **Indexing (1-based → 0-based) — THE central discriminator.** Each comment names a slide by
   the human number; the arm had to act on the 0-based index (4→3, 6→5, 8→7, 11→10) and for C4
   insert **after index 7, not index 8**. Off-by-one failure modes to detect specifically:
   - **C5 on index 11** — that's the dark closer ("Your team just got bigger."), already dark, so
     a "dark-mode restyle" there is a near no-op on the wrong slide while the actual light beta
     slide (index 10) stays light.
   - **C4 inserted after index 8** — the new slide lands after the Admin slide instead of after
     "Four moves", breaking the intended narrative position.
   - **C2 / C3 on index 4 / index 6** — wrong slide; the avatar/message-card tells are on indices
     3 and 5.
   Check the arm's `changes.json` (`reviewer_slide` vs `target_index`) AND its actual edits in the
   render — an arm can record the right number but edit the wrong slide, or vice versa. Score on
   what actually changed in the render.
2. **Clean patch ≠ rendered result.** Especially C1 (the logos must be present and consistent on
   **every** slide in the render, not just one) and C2/C3 (sub-0.15" alignment that must actually
   *read* centered / level in the rendered image). Inspect the arm's `verified` field + final
   message: did it re-render and confirm, or declare success from the patch? Score the
   *verification method*, not only the final pixels.

## Materials

- Arm A: original input renders `{{JUDGING_DIR}}/orig-A/*.jpg` (12), edited renders
  `{{JUDGING_DIR}}/deck-A/*.jpg` (13), `{{JUDGING_DIR}}/file-A.pptx`,
  `{{JUDGING_DIR}}/changes-A.json`
- Arm B (if present): same set with `-B`.

## What to do (per arm)

1. **Verify scope.** Confirm `file-?.pptx` has exactly **13 slides**. Confirm the insertion
   landed correctly: index 7 is still "Four moves, every time.", index 8 is the **new multiplayer
   slide**, index 9 is the old Admin slide ("Powerful, and on a leash."). Confirm all original
   content survived — flag any dropped copy, deleted slide, or reorder of slides the reviewer
   didn't touch as collateral damage. (Adding the new slide and restyling the beta slide is the
   *task*, not damage.)
2. **Score each revision C1–C5** as one of: **fixed** (the change is present and correct in the
   edited render and matches the ground-truth intent — exact coordinates/hexes not required, the
   revision must simply be genuinely resolved), **attempted-but-broken** (acted on the right
   target but the change is wrong, incomplete, inconsistent across slides, or introduced a new
   defect), or **missed** (no action, or acted on the wrong slide entirely). Use the edited
   renders as primary evidence; open `file-?.pptx` with python-pptx / `deck.py inspect` and sample
   pixels with PIL where the render is ambiguous — e.g. confirm the rail tiles now carry distinct
   logo fills + initials on every slide (C1), the avatar initials are optically centered on index
   3 (C2), the two message cards on index 5 share a left edge and width (C3), the new slide at
   index 8 is genuinely visual and on-brand (C4), and the beta slide at index 10 is now dark with
   all four beta points intact (C5).
3. **Check C1 on every slide.** It is a shared component / deck-wide change. Logos on one slide
   but not the rest is **attempted-but-broken**, not fixed. Itemize which slides got the logos and
   which didn't.
4. **Score the two traps** (indexing; clean-patch-vs-render verification) explicitly as above. For
   the indexing trap, state for each comment whether the arm hit the correct slide — especially
   the C4 insertion point and the C5 target.
5. **Regressions.** Pair each original/edited render (luminance-only structural delta: grayscale,
   resize ~160×90, mean absolute difference). The edited slides (the rail change is deck-wide; the
   beta slide flips; a new slide appears) are expected to differ; flag any **new** defect — text
   overflow past the canvas, a newly clipped/collided element, lost copy, a blank slide, a wrong
   final count, or a marketing slide that changed in a way the comments didn't call for. Note the
   per-slide numbers (account for the index shift after the inserted slide when pairing).

## Output (final message)

A markdown audit report:
- A per-arm **scorecard**: C1–C5 → {fixed / attempted-but-broken / missed} → one-line evidence
  (which render, which measurement). Weight the total as: **(a) correct targeting / indexing**
  carries the most weight (getting the right slide for all five), **(b) faithful execution** of
  each revision meaningful weight, **(c) the inserted slide's creative quality** (genuinely visual
  + on-brand, correct insertion point, count 12→13) meaningful weight, **(d) verification rigor**
  light/bonus weight. Do NOT weight discovery breadth — there is a checklist.
- The **indexing-trap verdict**: did the arm hit the right slide for all five comments? Call out
  the C4 insertion point and the C5 target by index explicitly.
- The **verification-method verdict**: did it re-render and confirm each fix (especially C1
  consistency and C2/C3 alignment), or declare success from the patch?
- **Regression** findings (any new defect, lost content, wrong final count, reordered or
  unrelated slides changed) with the structural-delta numbers.
- A final per-arm grade (A–F) with one paragraph. Reward **correct targeting + faithful,
  verified execution + a strong new slide**; penalize indexing errors (wrong slide), regressions
  (lost content, wrong final count, broken narrative order), and "looks fixed in the patch but not
  in the render." With two arms, name which did the cleaner execution on the merits.

Be exact and cite slide indices (note 0- vs 1-based each time).
