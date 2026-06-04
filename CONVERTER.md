# PPT → HTML faithful clone (mode 3) — setup & caveats

`scripts/convert.py` + `scripts/build_deck.py` reconstruct an existing `.pptx` into a
navigable, font-embedded HTML deck (not a re-skin — it preserves the *source* look).

## What it handles
Text + inheritance (lstStyle/defRPr/style fontRef), solid/gradient fill (multi-stop,
alpha), outline, image srcRect crop, master/layout, groups (with rotation), connectors/
freeform → SVG, shadows, **tables (graphicFrame)**, **preset shapes** (arrows, triangle,
trapezoid, parallelogram, brackets, cube, chevron, magneticDisk), **bullets** (buChar →
Wingdings✓ mapped to Unicode), **image recolor** (duotone/grayscale → feColorMatrix),
**grpFill** inheritance, **AlternateContent** unwrapping, and a click-build **animation**
engine. Scope tuned/verified on slides 1–30 of the reference deck.

## Setup
- **Fonts**: install `malgun.ttf/bd` + `arial.ttf` (from `/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/`) into `~/Library/Fonts/`. `build_deck.py` subsets them to woff2 and embeds (USB-portable, browser-independent line breaks).
- **Extract pptx**: `unzip -o deck.pptx -d /tmp/ppt_full` → use `/tmp/ppt_full/ppt`.
- **Units**: EMU/px = 9525; pt→px = 96/72; slide 12192000×6858000 EMU = 1280×720px.
- **Python**: `conda run -n base python` (lxml, PIL, fonttools). No heredocs; write `.py` files. No compound bash (`&&`/`;`).

## Run
```bash
conda run -n base python scripts/build_deck.py /tmp/ppt_full/ppt /tmp/deck 30   # → /tmp/deck/deck.html
conda run -n base python scripts/convert.py    /tmp/ppt_full/ppt /tmp/out 3 5 9 # single slides
```

## Ground truth + verification
- **Ground truth** (only if you need to diff): `soffice --headless --convert-to pdf --outdir D deck.pptx` → `pdftoppm -r 144 -png D/deck.pdf /tmp/refall/p` (fonts must be installed for correct wrapping).
- **Screenshot**: `chrome-headless-shell --headless --disable-gpu --hide-scrollbars --force-device-scale-factor=1.5 --window-size=1280,720 --screenshot=OUT "file://HTML"`.
- **Animation b0/bmax**: copy deck.html, inject before `</body>`:
  `<script>addEventListener("DOMContentLoaded",()=>{show(IDX);var m=maxB();for(var j=0;j<m;j++)next();});</script>`, screenshot build-0 vs full-build (IDX = 0-based slide index).

## Root-cause gotchas (do not regress)
1. Animation `data-build` must be injected into the **slide part only** — master/layout reuse the same cNvPr ids (collision). `slide_body_built()` handles this.
2. Group's own rotation rotates children rigidly about the group center (`render_grp`).
3. `is_placeholder` uses the **P** namespace `<p:ph>`.
4. Child `<a:grpFill/>` inherits the group's fill (GRPFILL stack).
5. Preset arrows: leftArrow/downArrow are mirrored geometry, not flipped rightArrow.

See the reference project's `HANDOVER.md` for the full defect/fix history.
