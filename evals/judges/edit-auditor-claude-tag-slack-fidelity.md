# Auditor: Slack-fidelity redesign scoring against a known tell manifest (forensic)

Task-specific auditor for `tasks/edit-claude-tag-slack-fidelity.md`. Like the review-loop
auditor and unlike the design judges, this auditor **holds the answer key**: a Slack-UI
wireframe deck with a known, prioritized list of the things that make it read as "not Slack,"
each with a ground-truth fix derived from a real-Slack-vs-deck audit. Its job is not taste — it
is to determine, per arm, which tells were **found**, **fixed**, and **verified**, and to grade
the arm on discovery breadth, fix correctness, and verification rigor. It runs on one arm or
both; with two arms, score each independently against the same manifest, then name the cleaner
fidelity pass.

The eval's whole point is that some tells are **silently dropped text** — content present in the
source that never renders, visible only by rendering and looking — while others are styling that
is merely *wrong* versus real Slack. Weight your read accordingly: recovering dropped content is
the highest-leverage signal that the arm actually rendered and looked (rather than trusting a
clean patch); chrome restyling shows whether it knows what real Slack looks like.

> Indices below are **0-based** (skill convention); the parenthetical is the **1-based** number a
> human would say. Do not penalize an arm for using either as long as it maps them correctly —
> see the indexing trap. Slide-index map: `0`=Cover(s1, marketing) · `1`=Premise(s2, marketing)
> · `2`=What-it-is(s3) · `3`=Multiplayer(s4) · `4`=Context/Memory(s5) · `5`=Proactive(s6) ·
> `6`=Async(s7, marketing) · `7`=How-it-works(s8) · `8`=Admin(s9) · `9`=Proof-stat(s10,
> marketing) · `10`=Availability(s11) · `11`=Closer(s12, marketing). The **chrome slides**
> (literal Slack views, where most tells live) are **2, 3, 4, 5, 7, 8, 10** (and 3's
> banner-in-composer-slot); the rest are marketing layouts with only floating-card / rail tells.

Placeholders: `{{JUDGING_DIR}}`. The auditor additionally needs, per arm, the edited renders
(`deck-A/*.jpg`, 12), the *original wireframe* renders (`orig-A/*.jpg`, 12), the arm's
`final.pptx` (copied in as `file-A.pptx`), and the arm's `changes.json` (copied in as
`changes-A.json`) — all assignment-matched by the orchestrator. The auditor must not learn arm
names.

## The answer key (the tells that give the mock away)

Each row is a real-Slack-vs-deck discrepancy. **Priority** is the leverage weight: a P0 tell
breaks the illusion outright; P1 tells are clearly-wrong chrome; P2 tells are polish. Detection
column: ⚠️ = a layout linter might half-flag it, ❌ = no machine signal, only render-and-look +
Slack knowledge.

### P0 — break the illusion (a Slack user spots these instantly)

| # | tell | slides | real Slack | wireframe (wrong) | detect | ground-truth fix |
|---|---|---|---|---|---|---|
| G1 | **Whole class of labels silently DROPPED at compile** | chrome: 2,3,4,5,7,8,10 (+ counts on 0,9) | channel names, DM names, workspace name, header title, member count, and reaction counts all render | bare text nodes inside flex containers vanish → sidebar shows only `#` with **no channel names** (general/launch/claude-tag/eng-backend/design), DM rows show avatars with **no names** (Claude/Dana R./Mike T.), workspace name **"Anthropic" missing**, channel-header titles missing (claude-tag/launch/context-and-memory/how-it-works/admin-settings/announcements), member count **"4" missing**, reaction **counts missing** | ❌ none (text is simply absent in the render) | recover every dropped label so it renders — channel names, DM names, "Anthropic", all header titles, member count "4", and all reaction counts are visible. **This is THE giveaway; nothing else matters until names and header titles are visible.** |
| G2 | **Reaction pills wrong: blue, colored dot instead of emoji, no count** | 0,2,3,5,9 | neutral light-gray pill (`~#F4F4F4` fill, `~#E2E2E2` border, r≈12), an **emoji glyph + a count**; blue tint only on pills *you* reacted to | every pill is blue, shows a flat colored **dot** where the emoji belongs, and the count is dropped (G1) → "a blue lozenge with a dot" | ❌ none | make pills neutral gray, put a real **emoji** in them, and show the **count**; reserve blue for emphasis only |

### P1 — clearly-wrong chrome (looks off even to a casual viewer)

| # | tell | slides | real Slack | wireframe (wrong) | detect | ground-truth fix |
|---|---|---|---|---|---|---|
| G3 | **Channel header see-through and too tall** | chrome: 2,3,4,5,7,8,10 | **solid** white bar, ~49px tall, hairline bottom border | `rgba(255,255,255,.86)` translucent (the paper background's colored blobs bleed *through* it) and **64px** tall | ❌ none | solid white header ~50px tall, keep the bottom hairline; title ~18px |
| G4 | **Message avatars too large and too round** | chrome + cards (0,2,3,5,8,9,…) | small **rounded square**, ~36px, radius ~4–8px (clearly not a circle/squircle) | 44px, radius 11px (oversized soft squircle); sidebar/facepile avatars a hair round too | ❌ none | ~40px square, radius ~8px; tighten facepile/sidebar avatar radius (~4px) |
| G5 | **No composer on most chrome slides** | chrome: 2,4,5,7,8,10 (+ banner-in-slot on 3) | every channel view ends in a message composer pinned to the bottom (rounded box, formatting row, send) | composer present on only the marketing slides; chrome slides have an empty white gutter — and slide **3** fills the composer slot with an aubergine "Multiplayer by default" banner instead | ❌ none | add a composer to the bottom of every chrome slide; remove/relocate the banner occupying slide 3's composer slot |

### P2 — polish (subtle tells; a strong arm gets these)

| # | tell | slides | real Slack | wireframe (wrong) | detect | ground-truth fix |
|---|---|---|---|---|---|---|
| G6 | Workspace-rail tiles render as bright solid-white squares | all (rail visible everywhere) | muted logo/letter tiles on dark aubergine; selected has a white left-bar | low-alpha white flattened to near-opaque → three glaring white chips down the rail | ❌ | give rail tiles explicit muted fills (e.g. `~#552E56`) with a faint glyph; selected tile lighter + inset left-bar |
| G7 | APP badge recolored per-app | chrome with bot messages | the "APP" badge is always the **same neutral gray** | Claude's APP badge is peach while others are gray | ❌ | use one neutral gray APP badge everywhere (brand color lives in the avatar, not the badge) |
| G8 | Channel-row sizing / no unread styling | chrome sidebars | rows ~28px, name ~15px; **unread channel = bold white name** | rows oversized (~17px font); once G1 lands, `#launch` has only a red "2" pill, no bold name | ❌ | tighten channel rows (~15px); make the unread channel's name bold white |
| G9 | Paper watermark bleeds into the message pane | 2,3,5,8,10 | channel body is plain white — no decorative shapes behind messages | faint multicolor blobs from the paper background show behind every message list | ❌ | put the message area on a solid white panel (or drop the blobs) so the watermark stays in the outer margins only |
| G10 | Placeholder/duplicate glyph icons | 8, 10 | recognizable, distinct icons | slide 8's four admin tiles all use the same ☆/◎ glyph (placeholder); slide 10's last checklist row uses ☆ instead of a check | ❌ | distinct admin icons (identities/people, scoped/lock, spend/gauge, audit/checklist); consistent ✓ on the last checklist row |

Detectability ladder — read every tell through this lens. **Every tell here is ❌ no-machine-signal**: none of them throw an overflow/coverage flag, so *all* of them require the arm to render, look, and know what real Slack looks like. There is no "linter caught it for free" floor on this task. That makes G1 (recovering dropped text — pure render-and-look) and G2 the sharpest discriminators of whether an arm did a real fidelity pass.

## Two traps to score explicitly

- **Indexing.** Phrasings like "slide 8" mean 0-based index 7. Check the arm's `changes.json` and
  its actual edits: did it operate on the correct slides, or off-by-one? An arm that restyled the
  wrong slide has an indexing bug even if its words look right.
- **"Clean patch ≠ fixed render."** A dropped-text fix (G1) is the canonical trap: an arm can add
  a `<span>` / a text run and report success from the source, but the label must actually be
  **visible in the rendered image**. Inspect the arm's reported verification (`changes.json`
  `verified` field + final message): did it re-render and confirm the label/pill/header is
  present and correct in the *render*, or did it declare success from the patch? Score the
  *verification method*, not just the final pixels.

## Materials

- Arm A: original wireframe renders `{{JUDGING_DIR}}/orig-A/*.jpg` (12), edited renders
  `{{JUDGING_DIR}}/deck-A/*.jpg` (12), `{{JUDGING_DIR}}/file-A.pptx`,
  `{{JUDGING_DIR}}/changes-A.json`
- Arm B (if present): same set with `-B`.

## What to do (per arm)

1. **Verify scope.** Confirm `file-?.pptx` still has exactly 12 slides and the deck's content is
   intact — flag any content edit, deletion, slide reorder, or slide-count change as collateral
   damage. (Recovering dropped labels and adding composers is the *task*, not damage; removing or
   altering the slides' actual message/copy is.)
2. **Score each tell G1–G10** as one of: **fixed** (gone in the edited render and the result
   matches the ground-truth intent — exact coordinates/hexes not required, the tell must simply be
   genuinely resolved and read as real Slack), **attempted-but-broken** (touched the right
   component but the tell remains, is inconsistent across slides, or a new one appeared), or
   **missed** (no action). Use the edited renders as primary evidence; open `file-?.pptx` with
   python-pptx / `deck.py inspect`, and sample pixels with PIL, where the render is ambiguous —
   e.g. confirm channel/DM/header text strings are present (G1), the header fill is opaque white
   and ~50px not ~64px (G3), reaction pills are neutral not blue and carry a glyph+count (G2),
   avatars are ~40px square (G4), a composer shape exists at the bottom of each chrome slide (G5),
   rail tiles are no longer near-white (G6). **G1 is the heaviest single line** — itemize which
   label classes (channel names / DM names / workspace name / header titles / member count /
   reaction counts) came back and which are still missing.
3. **Consistency across slides.** These are shared components. A fix that lands on one chrome
   slide but not the others is **attempted-but-broken**, not fixed. Check each P0/P1 tell on every
   slide it appears on (G1/G3/G5 across 2,3,4,5,7,8,10; G2 across 0,2,3,5,9).
4. **Discovery vs. fix.** Cross-check `changes.json` against your findings: did the arm *report* a
   tell it failed to fix, or *fix* one it never reported? Discovery without repair still earns
   partial credit (it saw the problem); silent correct fixes count as found. Because there is no
   checklist and no linter signal, **breadth of discovery is itself a graded dimension** — how
   many of the 10 tells did the arm even identify?
5. **Score the two traps** (indexing; clean-patch-vs-render verification) explicitly as above.
6. **Regressions.** Diff each original/edited render pair (luminance-only structural delta:
   grayscale, resize ~160×90, mean absolute difference). The chrome slides are expected to change
   substantially (that's the redesign); flag any **new** defect — text overflow past the canvas, a
   newly clipped/collided element, lost copy, a blank slide — and any marketing slide that changed
   in a way the task didn't call for. Note the per-slide numbers.

## Output (final message)

A markdown audit report:
- A per-arm **tell scorecard**: G1–G10 → {fixed / attempted-but-broken / missed} → one-line
  evidence (which render, which measurement). For G1, break out the six label classes. Compute a
  **headline fidelity score** as a weighted total — **P0 tells (G1, G2) carry the most weight, P1
  (G3–G5) meaningful weight, P2 (G6–G10) light/bonus weight** — and state it as
  e.g. "P0 2/2, P1 3/3, P2 3/5". Call out separately how broadly the arm **discovered** tells
  (count identified / 10) versus how many it **fixed correctly and consistently**, since the
  discovery-with-no-checklist is the eval's real discriminator.
- The **trap verdicts**: indexing correct? verification method sound (re-rendered and confirmed
  recovered text/restyled chrome is actually visible, not declared from the patch)?
- **Regression** findings (any new defect or unwarranted change) with the structural-delta numbers.
- A final per-arm grade (A–F) with one paragraph. Reward **discovery breadth + correct,
  consistent, render-verified execution**; penalize regressions (lost content, broken slide count,
  new overflow) and "looks fixed in the patch but not in the render." With two arms, name which
  did the cleaner fidelity pass on the merits (how many tells found, how correctly and consistently
  fixed across all 12 slides, verification rigor), and note whether the difference came from the
  P0 content-recovery tells or the chrome-restyle tells.

Be exact and cite slide indices (note 0- vs 1-based each time).
