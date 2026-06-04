---
name: snu-slides
description: Use when creating, generating, converting, or restyling presentation slides/decks in a Korean-university academic style (SKKU navy-gold, YU blue, JNU green) — making new slides from scratch (to HTML or PowerPoint), cloning an existing .pptx to HTML, exporting to .pptx, or matching the deck look. Triggers - 슬라이드 만들어, 발표자료 만들어, deck 만들어, ppt 만들어, ppt를 html로, html을 ppt로, pptx to html, 같은 스타일로, 슬라이드 스타일, SKKU/성균관대 슬라이드, 영남대/YU 슬라이드, 전남대/JNU 슬라이드, navy gold 슬라이드.
---

# Uni Slides (SKKU · YU · JNU)

Versatile presentation toolkit for **Korean-university academic decks**. One content
model (**slidespec**) drives every direction: author once, emit HTML and/or PowerPoint;
or faithfully clone an existing `.pptx` to HTML.

**Themes (set `"theme"` in the spec; default `skku`):**

| theme | school | palette | components |
|---|---|---|---|
| `skku` | 성균관대 | navy `#0F0F70` / gold `#C5A86F` | **OOXML clone** — PPT shapes are the original reference-deck `<p:sp>` elements (chip/number-block/cards), so PowerPoint renders them faithfully. |
| `yu` | 영남대 | blue `#153974` / cyan `#00AACA` | **primitive** — drawn from tokens (no clone library). |
| `jnu` | 전남대 | green `#007A33` / navy `#003594` | **primitive** — token keys `navy`/`gold` hold the school's PRIMARY/ACCENT. |

- **Design tokens:** `themes/<theme>.json` (machine) + `DESIGN_SYSTEM.md` (the SKKU reference, human). Each `yu`/`jnu` json carries an `official_ci` block (exact colors, logo URLs, fonts).
- **PPT components:** `skku` clones `themes/skku_components.xml` (see `scripts/clone.py`); themes without a `<theme>_components.xml` fall back to primitive drawing automatically.
- **Logos:** `assets/logos/<theme>.png` (transparent). The cover/title page places it automatically.
- **New school:** copy `themes/jnu.json` → `themes/<school>.json`, edit colors + `official_ci` + `logo`; drop a transparent `assets/logos/<school>.png`. Primitive theme works immediately; add a clone library only if you have a reference `.pptx`. Ask the user for the palette first.
- Slide canvas is 1280×720 (16:9). Python via `conda run -n base python` (needs `lxml`, `python-pptx`; PPT→HTML clone also needs the fonts/Chrome setup below).

## Modes (pick by what the user has and wants)

| # | From → To | Use |
|---|---|---|
| 1 | **nothing → HTML** | author a slidespec → `scripts/generate.py spec.json out.html` |
| 2 | **nothing → PPT** | author a slidespec → `scripts/export_pptx.py spec.json out.pptx` |
| 3 | **PPT → HTML** (faithful clone) | `scripts/build_deck.py <pptdir> <outdir> <N>` (full reconstruction; preserves the original look, not a theme template) |
| 4 | **HTML/PPT → PPT or restyle into a theme** | read the source's content → **author a slidespec** in the chosen theme → modes 1+2 |

**The slidespec is the hub.** For modes 1/2/4 you (Claude) write a JSON slidespec, then run the generator(s). Mode 3 is a different tool: a pixel-faithful OOXML→HTML converter (it does NOT re-skin into a theme template — use it to port an existing deck to a portable HTML).

## Authoring a slidespec (modes 1, 2, 4)

Write a JSON file. Full schema + every block type is documented at the top of
`scripts/generate.py` (read it before authoring). Minimal shape:

```json
{ "theme": "skku", "title": "Deck",
  "slides": [
    {"layout":"cover","title":"...","subtitle":"...","presenter":"정석현 (Seokhyun Choung)","date":"2026. 06. 04."},
    {"layout":"section","chapter":"01","kicker":"Part 1","title":"..."},
    {"layout":"content","chapter":"02","section":"2.1","kicker":"Overview","title":"...",
     "body":[{"type":"p","text":"**bold**, *gold emphasis*"},
             {"type":"bullets","check":true,"items":["a","b"]}],
     "footer":"line with *orange keyword*"},
    {"layout":"cards","chapter":"03","title":"...","cols":3,
     "cards":[{"head":"A","accent":"navy","text":"..."}]},
    {"layout":"closing","title":"Thank You","subtitle":"..."}
  ] }
```

Layouts: `cover · section · content · cards · closing`. Body blocks: `h · p · bullets(check) · cards · cols · kpi · table · image`. Inline: `**bold**`→primary, `*x*`→accent. Start from `templates/sample.slidespec.json`.

**Cover = title page.** It auto-places the theme's logo (top-left), the title (large, primary), an accent rule, the subtitle, and a primary bottom band with `presenter` (left) + `date` (right). Always give a real `presenter` and `date`; `meta` is a legacy single-line fallback. Card/sub-heading `accent` ∈ `navy`/`gold`/`gray` (recolored per theme).

Then emit one or both:
```bash
conda run -n base python scripts/generate.py spec.json deck.html      # HTML
conda run -n base python scripts/export_pptx.py spec.json deck.pptx   # PowerPoint
```

**Showcase landing (optional):** to present several schools' example decks on one
hero/gallery page (logo + palette swatches + links per school):
```bash
conda run -n base python scripts/showcase.py OUT_DIR DECKS_DIR skku yu jnu
```
It reads each theme json + logo and copies `<theme>_deck.html/.pptx` from `DECKS_DIR` into `OUT_DIR/index.html`.

## PPT → HTML faithful clone (mode 3)

Reconstructs an existing .pptx into a navigable, font-embedded HTML deck (text/
shapes/gradients/images/tables/preset-shapes/animations preserved). Setup + caveats
(font install, ground-truth render, Chrome screenshot, EMU units, animation b0/bmax
verification) are in `CONVERTER.md`. Run:
```bash
conda run -n base python scripts/build_deck.py /tmp/ppt_full/ppt /tmp/deck 30
```

## MANDATORY verification (do NOT skip, do NOT show the user first)

After generating, you MUST audit before claiming done. Two non-negotiable rules:

**Rule 1 — the main agent NEVER reads slide PNGs.** Reading deck screenshots
saturates your context and the image API starts rejecting reads. ALL image
inspection is delegated to subagents that return text-only reports.

**Rule 2 — always run the audit, fix, re-audit until clean.**

```bash
conda run -n base python scripts/audit.py deck.html deck.pptx --out /tmp/_audit
```
This renders safe-sized (<=1280px) `html_NN.png` + `ppt_NN.png` per slide and
prints a manifest. Then dispatch **one verification subagent** (or a few in
parallel for big decks) with the manifest paths. The subagent checks BOTH:

1. **HTML layout** — text/box overlap, overflow past the 1280×720 canvas, text exceeding its box, sub-heading box clipping.
2. **HTML↔PPT parity** — for each slide, are these the SAME in both: cover/title page (logo, title, accent rule, presenter+date band), banner (number block with faded leading digit / themed header band / accent pill chip / title), sub-heading bars, bullet markers+spacing, KPI alignment (left-grouped, one line each), card fill/border color(per accent)/radius/height, footer caption + emphasis color, theme colors correct (no leftover navy/gold on yu/jnu), font sizes (4-size scale) and family (Arial)?

Known soffice-only artifacts (NOT real defects — verify in real PowerPoint): extra spaces injected around Korean/Latin boundaries and punctuation ("LLM 의", "요약 · 분류"); the HTML viewer nav pill (bottom-right) is absent in PPT by design.

The subagent returns: per-slide HTML defects + per-slide HTML↔PPT mismatches + top fixes. **Fix the generator/exporter, regenerate, re-run audit, repeat until the subagent reports clean parity.** Only then show the user.

HTML is the reference: it almost always renders closer to intent than PPT.
Common PPT-only gaps to watch (PPT has no autosize callback, so it's estimated):
text wrapping/overlap (widen boxes or `wrap=False`), missing shadows (`add_shadow`),
missing gradients (`add_hgradient`), default-font fallback (`apply_fonts` forces Arial+Malgun).

## Common mistakes

- Putting `display` on a layout class so it overrides `.slide{display:none}` → all slides stack. Display belongs only on `.slide.active` (and `.slide.<layout>.active` for flex layouts).
- Editing tokens inline in scripts. Tokens live in `themes/*.json`; scripts read them.
- Using mode 3 (clone) when the user wanted a theme restyle — clone keeps the *source* look. Restyle = mode 4 (author a fresh slidespec).
- Hand-writing PPT shapes for `skku` instead of cloning — the gold pill chip drawn as a python-pptx primitive renders invisibly in PowerPoint; clone the real `<p:sp>` via `scripts/clone.py` (see DESIGN_SYSTEM §2/§5).
- Forgetting EA font: keep `Malgun Gothic` in the font stack for Korean.
