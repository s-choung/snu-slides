#!/usr/bin/env python3
"""Build a hero/gallery landing page showcasing per-school example decks.

For each theme it reads themes/<theme>.json (palette + logo), copies the matching
<theme>_deck.html / <theme>_deck.pptx from the decks dir into the output dir, and
emits index.html with a hero header + one card per school (logo, name, color
swatches, "open HTML deck" + "download PPTX" links).

Usage:
  python showcase.py OUT_DIR DECKS_DIR theme1 theme2 ...
  e.g. python showcase.py /tmp/showcase /tmp skku yu jnu
"""
import os, sys, json, base64, shutil, html

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
THEME_DIR = os.path.join(ROOT, "themes")
LOGO_DIR = os.path.join(ROOT, "assets", "logos")

def esc(s): return html.escape(str(s))

def data_uri(path):
    if not path or not os.path.exists(path): return ""
    return "data:image/png;base64," + base64.b64encode(open(path, "rb").read()).decode()

def load(theme):
    return json.load(open(os.path.join(THEME_DIR, f"{theme}.json"), encoding="utf-8"))

SCHOOL_KO = {"skku": "성균관대학교", "yu": "영남대학교", "jnu": "전남대학교"}

def card(theme, out_dir, decks_dir):
    t = load(theme); c = t["colors"]
    name_ko = SCHOOL_KO.get(theme, t.get("name", theme.upper()))
    logo = data_uri(os.path.join(LOGO_DIR, t["logo"])) if t.get("logo") else ""
    # copy deck files in
    links = []
    for ext, label in (("html", "HTML 덱 열기"), ("pptx", "PPTX 다운로드")):
        src = os.path.join(decks_dir, f"{theme}_deck.{ext}")
        if os.path.exists(src):
            dst = f"{theme}_deck.{ext}"
            shutil.copy(src, os.path.join(out_dir, dst))
            links.append(f'<a class="btn" href="{dst}">{label}</a>')
    swatch = "".join(
        f'<span class="sw" style="background:{c[k]}" title="{c[k]}"></span>'
        for k in ("navy", "gold", "gray") if k in c)
    logo_html = f'<img class="logo" src="{logo}">' if logo else f'<div class="logo ph">{esc(name_ko)}</div>'
    return f"""
    <article class="card" style="--accent:{c['navy']};--accent2:{c['gold']}">
      <div class="card-top">{logo_html}</div>
      <div class="card-body">
        <h2>{esc(name_ko)} <span class="en">{esc(t.get('name',''))}</span></h2>
        <div class="swatches">{swatch}</div>
        <div class="btns">{''.join(links)}</div>
      </div>
    </article>"""

def build(out_dir, decks_dir, themes):
    os.makedirs(out_dir, exist_ok=True)
    cards = "\n".join(card(th, out_dir, decks_dir) for th in themes)
    page = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>한국 대학 학술 슬라이드 — 디자인 시스템</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Malgun Gothic','Apple SD Gothic Neo',Arial,sans-serif;color:#1a1a1a;background:#f4f5f7}}
.hero{{background:linear-gradient(135deg,#0f0f70 0%,#222 100%);color:#fff;padding:64px 40px 72px}}
.hero h1{{font-size:34px;font-weight:800;letter-spacing:-.5px}}
.hero p{{margin-top:12px;font-size:16px;opacity:.85;max-width:760px;line-height:1.6}}
.wrap{{max-width:1080px;margin:-40px auto 60px;padding:0 24px}}
.grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:24px}}
@media(max-width:820px){{.grid{{grid-template-columns:1fr}}}}
.card{{background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 8px 28px rgba(0,0,0,.10);
 border-top:5px solid var(--accent);display:flex;flex-direction:column}}
.card-top{{height:120px;display:flex;align-items:center;justify-content:center;padding:22px;background:#fff;border-bottom:1px solid #eee}}
.logo{{max-height:64px;max-width:84%;width:auto;object-fit:contain}}
.logo.ph{{font-size:18px;font-weight:700;color:var(--accent)}}
.card-body{{padding:20px 22px 24px}}
.card-body h2{{font-size:19px;font-weight:800}}
.card-body .en{{font-size:12px;font-weight:600;color:#888;margin-left:4px}}
.swatches{{display:flex;gap:8px;margin:14px 0 18px}}
.sw{{width:26px;height:26px;border-radius:7px;border:1px solid rgba(0,0,0,.08)}}
.btns{{display:flex;gap:10px;flex-wrap:wrap}}
.btn{{flex:1;text-align:center;padding:10px 12px;border-radius:9px;font-size:13px;font-weight:700;
 text-decoration:none;color:#fff;background:var(--accent)}}
.btn:nth-child(2){{background:var(--accent2);color:#1a1a1a}}
footer{{text-align:center;color:#999;font-size:12px;padding:0 0 50px}}
</style></head><body>
<div class="hero"><h1>한국 대학 학술 슬라이드 — 디자인 시스템</h1>
<p>한 벌의 slidespec으로 학교별 정체성에 맞춰 HTML·PowerPoint 덱을 동시에 생성합니다.
SKKU는 원본 발표자료의 OOXML 컴포넌트를 그대로 복제하고, YU·JNU는 공식 CI 색상으로 렌더링합니다.</p></div>
<div class="wrap"><div class="grid">{cards}</div></div>
<footer>snu-slides skill · navy/gold · blue · green · 1280×720</footer>
</body></html>"""
    open(os.path.join(out_dir, "index.html"), "w", encoding="utf-8").write(page)
    print("wrote", os.path.join(out_dir, "index.html"), "for", ", ".join(themes))

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("usage: showcase.py OUT_DIR DECKS_DIR theme1 [theme2 ...]"); sys.exit(1)
    build(sys.argv[1], sys.argv[2], sys.argv[3:])
