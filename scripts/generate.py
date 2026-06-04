#!/usr/bin/env python3
"""SNU slide generator: slidespec(JSON) -> styled, navigable HTML deck.

slidespec schema (JSON):
{
  "theme": "skku",                     # theme name under themes/<name>.json (skku|yu|jnu)
  "title": "Deck title",               # optional
  "slides": [
    { "layout": "cover", "title": "...", "subtitle": "...", "meta": "2026 · Name" },
    { "layout": "section", "chapter": "01", "title": "Section title", "kicker": "Part" },
    { "layout": "content", "chapter": "02", "section": "2.1", "kicker": "Overview",
      "title": "Slide title", "body": [ <block>, ... ], "footer": "highlight line" },
    { "layout": "cards", "chapter": "03", "title": "...", "cards": [
        {"head":"Card A","text":"...","accent":"navy"}, ... ], "cols": 3 },
    { "layout": "closing", "title": "Thank you", "subtitle": "..." }
  ]
}
block types (used in "body"):
  {"type":"h",     "text":"Sub-heading"}
  {"type":"p",     "text":"paragraph (supports **bold** and *gold*)"}
  {"type":"bullets","items":["a","b"], "check": true}     # check=true -> ✓ marker
  {"type":"cards", "cols":3, "items":[{"head":"..","text":".."}]}
  {"type":"cols",  "items":[ [<block>...], [<block>...] ]}  # side-by-side columns
  {"type":"image", "src":"path.png", "h": 240}
  {"type":"kpi",   "items":[{"value":"150M","label":"Total"}]}

Usage: python generate.py spec.json out.html [--theme-dir DIR]
"""
import os, sys, json, html, re, argparse, base64

HERE = os.path.dirname(os.path.abspath(__file__))
THEME_DIR = os.path.join(os.path.dirname(HERE), "themes")
LOGO_DIR = os.path.join(os.path.dirname(HERE), "assets", "logos")

def logo_data_uri(t):
    """Embed the theme's logo as a base64 data URI so the HTML is self-contained."""
    fn = t.get("logo")
    if not fn: return ""
    path = fn if os.path.isabs(fn) else os.path.join(LOGO_DIR, fn)
    if not os.path.exists(path): return ""
    b64 = base64.b64encode(open(path, "rb").read()).decode()
    return f"data:image/png;base64,{b64}"

def load_theme(name, theme_dir):
    p = os.path.join(theme_dir, f"{name}.json")
    with open(p, encoding="utf-8") as f:
        return json.load(f)

def esc(t): return html.escape(str(t))

def inline(t):
    """**bold** -> <b>, *gold* -> gold emphasis span."""
    t = esc(t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"\*(.+?)\*", r'<span class="em">\1</span>', t)
    return t

# ---------- block renderers ----------
def block(b):
    t = b.get("type")
    if t == "h":
        acc = b.get("accent", "navy")
        return f'<div class="subbar {acc}">{inline(b["text"])}</div>'
    if t == "p":   return f'<p class="bp">{inline(b["text"])}</p>'
    if t == "bullets":
        mark = "check" if b.get("check") else "dot"
        lis = "".join(f'<li>{inline(x)}</li>' for x in b["items"])
        return f'<ul class="bl {mark}">{lis}</ul>'
    if t == "cards":
        cols = b.get("cols", len(b["items"]))
        cs = "".join(card(c) for c in b["items"])
        # body-embedded cards size to content (height:auto) so following blocks
        # (e.g. a kpi row) stay visible; the slide-level "cards" LAYOUT fills 100%.
        return (f'<div class="cardgrid" style="grid-template-columns:repeat({cols},1fr);'
                f'height:auto">{cs}</div>')
    if t == "cols":
        cells = "".join(f'<div class="col">{"".join(block(x) for x in col)}</div>' for col in b["items"])
        return (f'<div class="cols" style="grid-template-columns:repeat({len(b["items"])},1fr);'
                f'height:auto">{cells}</div>')
    if t == "image":
        h = f'height:{b["h"]}px;' if b.get("h") else ""
        return f'<div class="imgwrap"><img src="{esc(b["src"])}" style="{h}"></div>'
    if t == "kpi":
        ks = "".join(f'<div class="kpi"><div class="kv">{inline(k["value"])}</div>'
                     f'<div class="kl">{inline(k["label"])}</div></div>' for k in b["items"])
        return f'<div class="kpirow">{ks}</div>'
    if t == "table":
        head = "".join(f"<th>{inline(c)}</th>" for c in b.get("head", []))
        rows = "".join("<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r) + "</tr>" for r in b["rows"])
        h = f"<thead><tr>{head}</tr></thead>" if head else ""
        return f'<table class="tbl">{h}<tbody>{rows}</tbody></table>'
    return ""

def card(c):
    accent = c.get("accent", "navy")
    head = f'<div class="card-h {accent}">{inline(c["head"])}</div>' if c.get("head") else ""
    body = f'<div class="card-b">{inline(c.get("text",""))}</div>' if c.get("text") else ""
    sub = "".join(block(x) for x in c.get("body", []))
    return f'<div class="card {accent}">{head}{body}{sub}</div>'

# ---------- header band ----------
def header(s):
    chap = s.get("chapter")
    num = ""
    if chap:
        cs = esc(chap)
        # faded leading char(s) + solid trailing digit, matching the reference block
        inner = (f'<span class="nl">{esc(str(chap)[:-1])}</span>{esc(str(chap)[-1])}'
                 if len(str(chap)) > 1 else cs)
        num = f'<div class="numblock">{inner}</div>'
    chip = f'<span class="chip">{esc(s["section"])}</span>' if s.get("section") else ""
    kick = f'<div class="kicker">{esc(s["kicker"])}</div>' if s.get("kicker") else ""
    title = f'<div class="title"><span>{inline(s["title"])}</span></div>' if s.get("title") else ""
    # original stacked layout: label (top) + gold chip (below) at x=129, title at x=206
    return (f'<div class="banner"><div class="bannerbg"></div>{num}'
            f'{kick}{chip}{title}</div>')

# ---------- slide layouts ----------
def slide(s, t=None):
    layout = s.get("layout", "content")
    if layout == "cover":
        uri = logo_data_uri(t) if t else ""
        logo = f'<img class="cover-logo" src="{uri}">' if uri else ""
        presenter = s.get("presenter", "")
        date = s.get("date", "")
        meta = s.get("meta", "")          # legacy single-line fallback
        if not (presenter or date) and meta:
            presenter = meta
        band = (f'<div class="cover-band"><span class="cb-pres">{inline(presenter)}</span>'
                f'<span class="cb-date">{inline(date)}</span></div>'
                if (presenter or date) else "")
        sub = f'<div class="cover-sub">{inline(s["subtitle"])}</div>' if s.get("subtitle") else ""
        return (f'<section class="slide cover">'
                f'<div class="cover-bar"></div>{logo}'
                f'<div class="cover-mid"><h1>{inline(s.get("title",""))}</h1>'
                f'<div class="cover-rule"></div>{sub}</div>{band}</section>')
    if layout == "section":
        return (f'<section class="slide section">'
                f'<div class="sec-num">{esc(s.get("chapter",""))}</div>'
                f'<div class="sec-body"><div class="kicker">{esc(s.get("kicker",""))}</div>'
                f'<h2>{inline(s.get("title",""))}</h2></div></section>')
    if layout == "closing":
        return (f'<section class="slide closing">'
                f'<h1>{inline(s.get("title",""))}</h1>'
                f'<div class="cl-sub">{inline(s.get("subtitle",""))}</div></section>')
    if layout == "cards":
        cols = s.get("cols", len(s.get("cards", [])) or 3)
        cs = "".join(card(c) for c in s.get("cards", []))
        foot = f'<div class="footer"><span>{inline(s["footer"])}</span></div>' if s.get("footer") else ""
        return (f'<section class="slide">{header(s)}'
                f'<div class="body"><div class="cardgrid" '
                f'style="grid-template-columns:repeat({cols},1fr)">{cs}</div></div>{foot}</section>')
    # content (default)
    body = "".join(block(b) for b in s.get("body", []))
    foot = f'<div class="footer"><span>{inline(s["footer"])}</span></div>' if s.get("footer") else ""
    return f'<section class="slide">{header(s)}<div class="body">{body}</div>{foot}</section>'

# ---------- CSS from theme ----------
def css(t):
    c = t["colors"]; ty = t["type_pt"]; bx = t["box"]; ly = t["layout"]
    gr = t.get("gradients", {})
    PT = 96/72
    def px(pt): return f"{pt*PT:.1f}px"
    fonts = f"'{t['fonts']['latin']}','{t['fonts']['ea']}',sans-serif"
    return f"""
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{height:100%;background:#000;overflow:hidden;font-family:{fonts};color:{c['body']}}}
#stage{{position:absolute;left:50%;top:50%;width:{ly['w']}px;height:{ly['h']}px;
 transform:translate(-50%,-50%) scale(var(--s,1));transform-origin:center}}
.slide{{position:absolute;inset:0;width:{ly['w']}px;height:{ly['h']}px;background:{c['page_bg']};
 overflow:hidden;display:none}}
.slide.active{{display:block}}
.em{{color:{c.get('em_color', c['gold'])};font-weight:700}}
b{{color:{c.get('bold_color', c['navy'])};font-weight:700}}
/* banner */
.banner{{position:absolute;top:0;left:0;right:0;height:{ly['banner_h']}px}}
.bannerbg{{position:absolute;inset:0;background:{gr.get('banner','#fff')}}}
.numblock{{position:absolute;left:0;top:0;width:{ly['banner_h']}px;height:{ly['banner_h']}px;
 background:{gr.get('numblock', c['navy'])};color:#fff;font-size:{px(ty['number'])};font-weight:700;
 display:flex;align-items:center;justify-content:center;letter-spacing:1px}}
.numblock .nl{{opacity:.42}}
.kicker{{position:absolute;left:129px;top:6px;color:{c['muted']};font-size:{px(ty['small'])};font-weight:700}}
.chip{{position:absolute;left:129px;top:40px;background:{c['gold']};color:{c.get('chip_text','#fff')};
 font-size:{px(ty['small'])};font-weight:700;border-radius:999px;padding:3px 14px;
 display:inline-flex;align-items:center;line-height:1}}
.title{{position:absolute;left:206px;top:0;height:{ly['banner_h']}px;right:{ly['margin_x']}px;
 display:flex;align-items:center;color:{c['ink']};font-size:{px(ty['title'])};font-weight:400}}
/* body */
.body{{position:absolute;left:{ly['margin_x']}px;right:{ly['margin_x']}px;
 top:{ly['body_top']}px;bottom:48px;overflow:hidden}}
/* colored sub-heading box (SNU card-header style) */
.subbar{{display:inline-block;color:#fff;font-size:{px(ty['subtitle'])};font-weight:700;
 padding:9px 26px;border-radius:10px;margin:14px 0 10px;box-shadow:{bx['shadow']}}}
.subbar.navy{{background:{c['navy']}}} .subbar.gold{{background:{c['gold']}}}
.subbar.gray{{background:{c['gray']}}}
.bp{{font-size:{px(ty['body'])};line-height:1.45;margin:4px 0}}
.bl{{list-style:none;margin:6px 0;font-size:{px(ty['body'])};line-height:1.5}}
.bl li{{padding-left:22px;position:relative;margin:3px 0}}
.bl.dot li::before{{content:'•';position:absolute;left:4px;color:{c['navy']}}}
.bl.check li::before{{content:'✓';position:absolute;left:2px;color:{c.get('check_color', c['gold'])};font-weight:700}}
/* cards */
.cardgrid,.cols{{display:grid;gap:18px;height:100%}}
.card{{background:#fff;border:{bx['card_border_w']}px solid {c['navy']};
 border-radius:{bx['radius_px']}px;padding:{bx['pad_px']}px;overflow:hidden}}
.card.navy{{border-color:{c['navy']}}} .card.gold{{border-color:{c['gold']}}}
.card.gray{{border-color:{c['gray']}}}
.card-h{{font-size:{px(ty['subtitle'])};font-weight:700;margin-bottom:8px}}
.card-h.navy{{color:{c['navy']}}} .card-h.gold{{color:{c.get('em_color', c['gold'])}}}
.card-h.gray{{color:{c['gray']}}}
.card-b{{font-size:{px(ty['body'])};line-height:1.45}}
.col>*{{margin-bottom:8px}}
/* kpi */
.kpirow{{display:flex;gap:24px;margin:10px 0}}
.kpi{{text-align:center}}
.kv{{color:{c['navy']};font-size:{px(ty['number']*0.7)};font-weight:700}}
.kl{{color:{c['muted']};font-size:{px(ty['caption'])}}}
/* table */
.tbl{{border-collapse:collapse;width:100%;font-size:{px(ty['small'])}}}
.tbl th{{background:{c['navy']};color:#fff;padding:6px 10px;text-align:left}}
.tbl td{{border:1px solid {c['line']};padding:5px 10px}}
.imgwrap{{text-align:center;margin:8px 0}} .imgwrap img{{max-width:100%}}
/* footer: plain caption line, no navy band (box removed) */
.footer{{position:absolute;left:{ly['margin_x']}px;right:{ly['margin_x']}px;bottom:14px;
 color:{c['muted']};font-size:{px(ty['small'])};font-style:italic}}
.footer .em{{color:{c['accent_orange']};font-style:normal}}
/* cover / section / closing */
.cover{{background:{c['page_bg']}}}
.cover-bar{{position:absolute;left:0;top:0;width:12px;height:100%;background:{c['navy']}}}
.cover-logo{{position:absolute;left:90px;top:60px;height:62px;width:auto}}
.cover-mid{{position:absolute;left:90px;top:46%;transform:translateY(-50%);right:90px}}
.cover h1{{color:{c['navy']};font-size:{px(ty['cover_title'])};font-weight:700;line-height:1.15}}
.cover-rule{{width:84px;height:4px;background:{c['gold']};margin:18px 0}}
.cover-sub{{color:{c['ink']};font-size:{px(ty['subtitle'])}}}
.cover-band{{position:absolute;left:0;right:0;bottom:0;height:84px;background:{c['navy']};
 display:flex;align-items:center;justify-content:space-between;padding:0 90px;color:#fff}}
.cb-pres{{font-size:{px(ty['body'])};font-weight:700}}
.cb-date{{font-size:{px(ty['body'])};opacity:.9}}
.section{{background:{c['navy']};color:#fff;align-items:center;gap:40px;padding:0 90px}}
.slide.section.active{{display:flex}}
.sec-num{{font-size:160px;font-weight:800;color:{c['gold']};line-height:1}}
.section h2{{font-size:{px(ty['cover_title'])};font-weight:700}}
.section .kicker{{color:{c['gold']}}}
.closing{{background:{c['navy']};color:#fff;flex-direction:column;
 align-items:center;justify-content:center;gap:18px}}
.slide.closing.active{{display:flex}}
.closing h1{{font-size:{px(ty['cover_title'])};font-weight:700}}
.cl-sub{{font-size:{px(ty['subtitle'])};opacity:.85}}
#hud{{position:fixed;right:14px;bottom:10px;color:#999;font:12px Arial;
 background:rgba(0,0,0,.4);padding:3px 8px;border-radius:10px;z-index:99}}
"""

SHELL_JS = """
const slides=[...document.querySelectorAll('.slide')];let i=0;
function fit(){const s=Math.min(innerWidth/1280,innerHeight/720);
 document.getElementById('stage').style.setProperty('--s',s);}
function show(n){i=Math.max(0,Math.min(slides.length-1,n));
 slides.forEach((s,k)=>s.classList.toggle('active',k===i));
 document.getElementById('cur').textContent=(i+1)+' / '+slides.length;}
addEventListener('keydown',e=>{
 if(['ArrowRight',' ','PageDown'].includes(e.key)){show(i+1);e.preventDefault();}
 else if(['ArrowLeft','PageUp'].includes(e.key))show(i-1);
 else if(e.key==='Home')show(0); else if(e.key==='End')show(slides.length-1);
 else if(e.key.toLowerCase()==='f'){document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen();}});
addEventListener('resize',fit);
addEventListener('click',e=>show(e.clientX>innerWidth*0.5?i+1:i-1));
fit();show(0);
"""

def build(spec, theme_dir=THEME_DIR):
    theme = load_theme(spec.get("theme", "skku"), theme_dir)
    body = "\n".join(slide(s, theme) for s in spec["slides"])
    title = esc(spec.get("title", "SNU Deck"))
    return f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<title>{title}</title><style>{css(theme)}</style></head><body>
<div id="stage">{body}</div>
<div id="hud"><span id="cur">1</span> · ← → · F</div>
<script>{SHELL_JS}</script></body></html>"""

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("spec"); ap.add_argument("out")
    ap.add_argument("--theme-dir", default=THEME_DIR)
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    open(a.out, "w", encoding="utf-8").write(build(spec, a.theme_dir))
    print("wrote", a.out, f"({len(spec['slides'])} slides)")
