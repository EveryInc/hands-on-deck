# hands-on-deck

**Agent-native PowerPoint manipulation.** One CLI — `deck.py` — lets AI agents inspect, edit, create, and verify `.pptx` files with the fidelity of a human operator: atomic JSON patches in, linted decks out.

PowerPoint, Keynote, and Google Slides are the deck apps for humans; hands-on-deck is the same category of application, built for agents. It targets `.pptx` because that's the format all three share — decks made in Keynote or Google Slides work too, since both import and export it cleanly.

**→ [everyinc.github.io/hands-on-deck](https://everyinc.github.io/hands-on-deck/)** · **[demo video](https://x.com/nityeshaga/status/2069900616986759195)**

Packaged as an [Agent Skill](https://www.anthropic.com/news/skills), so it drops into Claude Code, claude.ai, and any other agent platform that supports the skills format — and because the tool itself is just a CLI, *any* agent that can run a shell command can use it.

```bash
# the whole edit loop, in four commands
python deck.py deck.pptx inspect --slide 3 --brief         # what's there (one line per shape)
python deck.py deck.pptx apply patch.json -o out.pptx --fix --render img/
python deck.py out.pptx diff deck.pptx                     # what changed
python deck.py docs                                        # the full reference, no file needed
```

## Why this exists

A `.pptx` is a zip of XML. Agents that edit it directly hand-write OOXML — fragile, token-hungry, and one namespace typo from a corrupt file. Agents that regenerate decks from scratch lose everything a template encodes: brand, layout craft, image treatments.

hands-on-deck takes a third path: **the agent writes a declarative patch; the tool executes it.**

```json
{"ops": [
  {"op": "replace-text", "scope": "master", "from": "Globex", "to": "Acme"},
  {"op": "set-text",  "slide": 3, "shape": "s12", "text": ["Q3 results", "Tokens down 84%"]},
  {"op": "swap-image", "slide": 4, "shape": "s9", "image": "screenshot.png"},
  {"op": "duplicate", "slide": 5, "shape": "s31", "offset": [0, 1.2], "text": ["Fourth pillar"]}
]}
```

New text inherits the old text's formatting automatically. Image swaps keep aspect ratio. The duplicate keeps every bit of styling and gets fresh ids. And if any op is invalid, *nothing* is written.

## Built around how agents actually fail

The interesting part isn't that it's a CLI — it's that every design choice targets a known LLM failure mode:

**Errors teach instead of scold.** Reference a shape that doesn't exist and the error includes the slide's real shape inventory — ids, types, geometry, text previews — so the agent can correct *without another round trip*:

```
PATCH REJECTED — 2 validation error(s), nothing was modified:
  - op[0] set-text: shape 's9999' not found on slide 0.
shapes on slide 0:
  s16    PICTURE      [-1.25,-0.91 15.0x8.44in]  (image image3.png)
  s18    AUTO_SHAPE   [7.00,5.17 2.6x0.25in]  Session Management
  s19    TEXT_BOX     [0.60,1.00 4.0x0.4in]  USING CLAUDE CODE
  ...
  - op[1] add-slide: layout 'Nonexistent' not found — available: 'DEFAULT', 'Blank'
```

**All errors at once, atomically.** Every op is pre-validated; a 9-op patch with 9 mistakes returns 9 actionable errors and writes zero bytes. No partially-edited decks, ever — runtime failures abort the whole patch too.

**The linter watches the agent's hands.** After every apply, the deck is re-measured and only *new or worsened* geometry problems are reported — text overflowing its box, shapes off the slide, text-on-text overlaps, text trapped *under* a picture (it renders clipped — the defect a thumbnail never shows), and edges that *almost* line up — each with exact inch values and the exact `fix` command to run.

**It catches the near-miss, not the design.** Alignment bugs are a hair off, never a mile: nobody means "3px short of the card edge." The linter discovers the deck's implicit grid (clusters the edges several shapes actually share) and flags the lone edge sitting in the uncanny valley beside a gridline — `s14 right=10.52" — 0.14" short of the cluster at 10.66" (4 shapes)`. It never needs to know what the design *should* be; it only needs to know that an edge is trying to align and failing. Purely advisory — intentional asymmetry is real, so it ranks by suspicion and never fails the build.

**Repair is honest.** `fix` deterministically grows boxes, shrinks fonts (with a readability floor), and nudges shapes back on-slide — then *re-measures*. Anything still broken is reported as *residue* with a suggested op, not claimed as fixed. Pictures bleeding off-slide are never auto-moved (it might be intentional design).

**Tokens are a budget.** `inspect --brief` gives one line per shape for orientation; full JSON only when writing a patch. `docs` prints the complete op reference so agents never read source. `render --slide 3 --crop 1,2,6,1.5 --scale 2` zooms into the exact region under suspicion instead of re-rendering everything. `diff` verifies edits with no rendering at all.

**Verification is visual.** Slides render to `slide-<index>.jpg` (0-based, matching every other index in the tool) so the agent can *look at* what it changed — the same way a human would check their work.

## What it covers

| | |
|---|---|
| **Read** | `inspect` — shape ids, geometry (inches), text + formatting + per-run breakdowns (links included), image rIds + media names, table contents, fills/gradients/borders, rotation, alt text, speaker notes, document properties, hidden slides, detected issues; `--master` for masters/layouts |
| **Edit** | `set-text` (formatting-inheriting, per-run hyperlinks), `replace-text` (deck/master/slide scope), `replace-color` (the re-theme primitive — one op per palette mapping), `set-theme` (scheme colors + major/minor fonts — how template decks rebrand), `swap-image` (per-slide or deck-wide via media bytes), `set-style` (fonts, solid/gradient fills, borders, rotation, alt text), `set-slide` (hide/unhide, slide background: solid/gradient/image, slide transitions: fade/push/wipe/… with speed and auto-advance), `set-props` (document metadata), `move`, `resize`, `delete` (animation-reference-safe), `set-notes` |
| **Create** | `add-slide` (by layout), `add-shape` (textbox, autoshapes, any MSO_SHAPE name, lines), `add-picture` (aspect-preserving), `add-table` (style-neutralized), `duplicate`, `copy-shape` (across slides, relationships re-homed) |
| **Create from HTML** | `html2patch.py` — write a slide as HTML/CSS, get a deck.py patch back: measured boxes, formatted runs, hyperlinks, gradients, rounded corners, bullets, tables, images, rotation |
| **Structure** | `reorder` (z-order), `add-row`/`delete-row`/`add-col`/`delete-col` (formatting-inheriting, width-rescaling, merged-cell guard), `slides` (reorder/duplicate/delete), `merge` (pull slides from another deck) |
| **Verify** | `render` (JPGs, crop + zoom), `diff` (structural changelog), post-apply lint, `fix` (deterministic repair) |
| **Escape hatch** | `xml get`/`xml set` — pretty-printed part XML, parse-checked and lint-checked on write-back |

**Out of scope by design** (escape hatch or PowerPoint): creating native charts, shape animations, embedded video/OLE, merged-cell table surgery.

## Design slides in HTML, keep one writer

Free-form slide design is the one place agents beat templates — and HTML/CSS is the layout language agents are best at. `html2patch.py` uses a headless browser purely as a *measuring engine*: it renders your HTML, reads back every element's box and computed style, and compiles a **deck.py patch** — not a .pptx.

```bash
python html2patch.py slide.html --deck deck.pptx --layout Blank -o patch.json
python deck.py deck.pptx apply patch.json -o out.pptx --render img/
```

Emitting a patch instead of a file is the whole trick:

- **One writer.** Created slides get the same shape ids, lint coverage, `fix` loop, and `diff`/`render` verification as edited ones. No second engine with its own quirks.
- **Creation *into* templates.** The patch can `add-slide` with a layout from your branded master and place the HTML-measured shapes onto it — free-form layout inside an existing deck, which a generate-a-new-file architecture can't do.
- **Reflow drift is caught, not hoped away.** Browser and PowerPoint wrap text slightly differently; the post-apply lint re-measures the real deck and reports any overflow with the exact fix. The safety net covers the create path with zero new code.
- **Inspectable intermediate.** The patch is readable JSON — tweak one op by hand, or skip HTML entirely for simple slides. HTML is a frontend that compiles to the same IR every other edit uses.

Text becomes formatted runs (inline `<b>`/`<i>`/`<span>` included); styled divs become rects with gradients, borders, and true corner radii; tables keep per-cell fills and measured column widths; `<ol>` numbers, `<ul>` bullets; `object-fit: cover` becomes a real picture crop; CSS padding maps to text insets; `transform: rotate` and `text-transform` are honored. Needs `pip install playwright && playwright install chromium` — optional, the core tool doesn't.

And because a create path is only as good as what it creates, the skill ships with [designing-slides.md](skills/hands-on-deck/designing-slides.md) — an opinionated, subject-first design guide for agents: how to refuse the default AI-deck looks, plan a token system before writing HTML, size type for a projector instead of a browser, and design with the compiler's grain. The pipeline is mechanical; that file is taste.

## The engine under your own deck skill

Teams don't just need deck edits — they need edits *their* way. hands-on-deck is designed to be composed: keep it as the engine, and layer a thin skill of your own on top that encodes your firm's rules — which starter templates to use, how titles are written, what every data slide must cite. The client skill just says "use hands-on-deck for the actual .pptx work"; the engine does the mechanics.

```
your firm's skill      # a page of rules: templates, voice, constraints
  └── hands-on-deck    # the engine: inspect, patch, lint, render
        └── your .pptx files
```

This is how we deploy it for clients: the rules change per team, the engine never does — and with marketplace auto-update, every engine improvement arrives without touching the layer above.

## Install

**Claude Code** (as a plugin — enable marketplace auto-update to receive every improvement automatically):

```
/plugin marketplace add EveryInc/hands-on-deck
/plugin install hands-on-deck@hands-on-deck
```

**claude.ai / ChatGPT / other apps that support Agent Skills**: zip `skills/hands-on-deck/` and upload it as a skill.

**Codex or any other agent, any platform**: clone the repo and put the output of `deck.py docs` in front of your agent. It's just a CLI.

```bash
git clone https://github.com/EveryInc/hands-on-deck
pip install python-pptx Pillow
python hands-on-deck/skills/hands-on-deck/scripts/deck.py docs
```

## Requirements

- Python 3.9+, `python-pptx`, `Pillow` (lxml, used by the xml escape hatch, ships with python-pptx)
- For `html2patch` (create slides from HTML): `pip install playwright && playwright install chromium`
- For `render` and thumbnail grids: LibreOffice (`soffice`) and Poppler (`pdftoppm`)
  - macOS: `brew install --cask libreoffice && brew install poppler`
  - Debian/Ubuntu: `apt-get install libreoffice-impress poppler-utils`

## Development

The test suite drives `deck.py` end-to-end through its CLI — including the adversarial cases (atomic rejection, runtime aborts, merged-cell guards, deck-wide media swaps):

```bash
pip install python-pptx Pillow pytest
pytest tests/ -v
```

No binary fixtures: tests generate their decks with python-pptx on the fly.

## Benchmarked, not just claimed

We raced an agent on hands-on-deck against the same agent on Anthropic's pptx skill — same briefs, same model, three from-scratch decks plus a heavy re-theme-and-insert edit, every round blind-judged by three independent judges. Every finding the judges produced became machinery in the tool (the text-under-picture lint, serif re-wrap margins, the `<br>` table fix, the `replace-color` op) — and that failure class never recurred, while the other toolchain's defects repeated every round. Tools learn; prompts don't. The full story is on [the landing page](https://everyinc.github.io/hands-on-deck/#benchmark).

The benchmark is a committed, repeatable eval suite, not a one-off: [`evals/`](evals/) holds the builder briefs, judge prompts, blinding script, the agent-executable runbook, and the results log of every round. Change the tool, rerun the eval, compare.

## Who built this

hands-on-deck is open-sourced from real work by [Every Consulting](https://every.to/consulting). We built it to make our own decks — every training we run ships with a branded deck, and our agents build them with hands-on-deck. The tool distills the lessons from building **hundreds of decks with Claude**: every way an agent fumbled a deck along the way became a design decision in this CLI. The failure modes it guards against aren't theory; they're field notes.

If you want your team's work automated like this — decks or anything else — [that's literally what we do](https://every.to/consulting).

## License

MIT
