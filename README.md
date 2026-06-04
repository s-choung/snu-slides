# snu-slides

Korean-university academic slide toolkit. One content model (**slidespec** JSON)
drives every output: author once, emit **HTML** and/or **PowerPoint** — or
faithfully clone an existing `.pptx` to HTML.

> Folder name is legacy `snu-slides`; the actual themes are **SKKU · YU · JNU**
> (not Seoul National University).

## Themes

| theme | school | palette | components |
|---|---|---|---|
| `skku` | 성균관대 | navy `#072B61` / green `#8DC63F` | OOXML clone — faithful PowerPoint shapes |
| `yu`   | 영남대 | blue `#153974` / cyan `#00AACA` | primitive (token-drawn) |
| `jnu`  | 전남대 | green `#007A33` / navy `#003594` | primitive (token-drawn) |

## Quick start

```bash
# slidespec(JSON) -> HTML deck
python scripts/generate.py spec.json deck.html

# slidespec(JSON) -> PowerPoint
python scripts/export_pptx.py spec.json deck.pptx

# a deck HTML -> contact-sheet gallery (NxM thumbnails)
python scripts/gallery.py deck.html gallery.html --cols 2

# render per-slide PNGs for verification
python scripts/audit.py deck.html deck.pptx --out /tmp/_audit
```

Needs Python with `lxml` + `python-pptx`. PPT→HTML clone and audit also need
`soffice` + a Chromium headless shell (see `CONVERTER.md`).

## Benchmark

`samples/benchmark_snu.html` — a 2×3 gallery of a 6-slide SKKU deck on
*"계산화학을 통한 촉매 연구"* (computational chemistry for catalysis), generated
from `samples/benchmark_catalysis.slidespec.json`. Open it to see the theme,
layouts (cover · content · cards · KPI · closing), and the gallery view.

## Docs

- `SKILL.md` — entry point: themes, slidespec schema, modes, verification rules.
- `DESIGN_SYSTEM.md` — measured SKKU reference tokens (single source of truth).
- `CONVERTER.md` — PPT→HTML faithful-clone setup and caveats.
