#!/usr/bin/env python3
"""Render a generated deck HTML as a contact-sheet GALLERY (all slides at once).

Reads a deck produced by generate.py, extracts its <style> and every
<section class="slide ...">, and lays them out as scaled thumbnails in a grid
(default 2 columns). Each thumbnail is the real slide markup (not an image), so
the page stays self-contained and crisp at any zoom. Clicking a thumbnail opens
the full deck.

Usage:
  python gallery.py deck.html out.html [--cols N] [--title "..."] [--deck-href deck.html]
"""
import os, sys, re, html, argparse

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck"); ap.add_argument("out")
    ap.add_argument("--cols", type=int, default=2)
    ap.add_argument("--title", default=None)
    ap.add_argument("--deck-href", default=None,
                    help="link target for thumbnails (default: basename of deck)")
    a = ap.parse_args()

    src = open(a.deck, encoding="utf-8").read()
    m = re.search(r"<style>(.*?)</style>", src, re.S)
    css = m.group(1) if m else ""
    title_m = re.search(r"<title>(.*?)</title>", src, re.S)
    title = a.title or (title_m.group(1) if title_m else "Slide Gallery")
    sections = re.findall(r'<section class="slide.*?</section>', src, re.S)
    if not sections:
        sys.exit("no slides found in " + a.deck)

    href = a.deck_href or os.path.basename(a.deck)
    cells = []
    for n, sec in enumerate(sections, 1):
        # force the slide visible inside its own frame
        sec = sec.replace('class="slide', 'class="slide active', 1)
        cells.append(
            f'<a class="cell" href="{html.escape(href)}" title="슬라이드 {n} 열기">'
            f'<div class="frame">{sec}</div>'
            f'<div class="cap">{n:02d}</div></a>')

    page = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)} — Gallery</title>
<style>
{css}
/* ---- gallery overrides ---- */
html,body{{height:auto;overflow:auto;background:#0d1117;color:#e6edf3;
 font-family:Arial,'Malgun Gothic',sans-serif}}
.wrap{{max-width:1280px;margin:0 auto;padding:40px 28px 64px}}
.head h1{{font-size:30px;font-weight:800;letter-spacing:-.01em;margin-bottom:8px}}
.head p{{color:#9aa7b4;font-size:15px;line-height:1.5;max-width:760px}}
.head .rule{{width:72px;height:4px;background:#8DC63F;margin:16px 0 30px;border-radius:2px}}
.gal{{display:grid;grid-template-columns:repeat({a.cols},1fr);gap:22px}}
.cell{{display:block;position:relative;border-radius:12px;overflow:hidden;
 background:#fff;border:1px solid #283242;text-decoration:none;
 box-shadow:0 6px 22px rgba(0,0,0,.45);transition:transform .15s,box-shadow .15s}}
.cell:hover{{transform:translateY(-3px);box-shadow:0 12px 32px rgba(0,0,0,.6);
 border-color:#8DC63F}}
.frame{{width:1280px;height:720px;transform-origin:top left}}
/* the extracted slide fills its 1280x720 frame */
.cell .slide{{position:absolute;inset:0;display:block}}
.cell .slide.section,.cell .slide.closing{{display:flex}}
.cap{{position:absolute;left:10px;bottom:8px;z-index:5;color:#fff;
 background:rgba(7,43,97,.85);font-size:12px;font-weight:700;
 padding:2px 9px;border-radius:999px;letter-spacing:.05em}}
.foot{{margin-top:34px;color:#6b7785;font-size:13px}}
.foot a{{color:#8DC63F;text-decoration:none}}
</style></head><body>
<div class="wrap">
<div class="head">
 <h1>{html.escape(title)}</h1>
 <p>SNU-slides 벤치마크 · 전체 {len(sections)}장을 {a.cols}열 컨택트시트로. 썸네일을 누르면 전체 덱이 열립니다.</p>
 <div class="rule"></div>
</div>
<div class="gal" id="gal">
{''.join(cells)}
</div>
<div class="foot">생성: <code>scripts/gallery.py</code> · 원본 덱: <a href="{html.escape(href)}">{html.escape(href)}</a></div>
</div>
<script>
function fit(){{
 document.querySelectorAll('.cell').forEach(c=>{{
  const w=c.clientWidth, s=w/1280;
  const f=c.querySelector('.frame');
  f.style.transform='scale('+s+')';
  c.style.height=(720*s)+'px';
 }});
}}
addEventListener('resize',fit); fit();
</script>
</body></html>"""
    open(a.out, "w", encoding="utf-8").write(page)
    print(f"wrote {a.out} ({len(sections)} slides, {a.cols} cols)")

if __name__ == "__main__":
    main()
