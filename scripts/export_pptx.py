#!/usr/bin/env python3
"""SNU slide exporter: slidespec(JSON) -> PowerPoint .pptx (python-pptx).

Same slidespec schema as generate.py (see that file). Produces a 16:9 deck in
the SNU navy/gold style: navy number block, gold section chip, light-gray title
banner, navy footer band, cards, KPI row, tables, bullets.

Usage: python export_pptx.py spec.json out.pptx [--theme-dir DIR]

Gradient banner is approximated with a solid light-gray bar (python-pptx has no
first-class gradient API); everything else is faithful to the theme tokens.
"""
import os, sys, json, re, argparse
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
THEME_DIR = os.path.join(os.path.dirname(HERE), "themes")
LOGO_DIR = os.path.join(os.path.dirname(HERE), "assets", "logos")
EMU = 9525  # per px

sys.path.insert(0, HERE)
import clone  # cloning canonical OOXML component shapes (gold chip, cards, ...)

def load_theme(name, theme_dir):
    with open(os.path.join(theme_dir, f"{name}.json"), encoding="utf-8") as f:
        return json.load(f)

def hexc(s): return RGBColor.from_string(s.lstrip("#"))
def px(v): return Emu(int(v*EMU))

def wrap_lines(text, w_px, pt):
    """Estimate how many wrapped lines `text` takes in a box of width w_px at
    `pt` (Arial). PPT has no autosize callback, so block height is estimated."""
    plain = re.sub(r"\*+", "", str(text))
    cpl = max(8, int(w_px / (pt * 0.56)))   # ~chars per line for Arial
    return max(1, (len(plain) + cpl - 1) // cpl)

# ---- inline markup: **bold** / *gold* -> runs ----
def add_runs(p, text, theme, base_color, em_hex=None):
    em = em_hex or theme["colors"].get("em_color") or theme["colors"]["gold"]
    parts = re.split(r"(\*\*.+?\*\*|\*.+?\*)", text)
    for seg in parts:
        if not seg: continue
        if seg.startswith("**") and seg.endswith("**"):
            r = p.add_run(); r.text = seg[2:-2]; r.font.bold = True
            r.font.color.rgb = hexc(theme["colors"].get("bold_color", theme["colors"]["navy"]))
        elif seg.startswith("*") and seg.endswith("*"):
            r = p.add_run(); r.text = seg[1:-1]; r.font.bold = True
            r.font.color.rgb = hexc(em)
        else:
            r = p.add_run(); r.text = seg
            if base_color: r.font.color.rgb = base_color

def _set_ea(run, ea):
    """Set East-Asian (Korean) typeface on a run via OOXML rPr."""
    rPr = run._r.get_or_add_rPr()
    el = rPr.find(qn('a:ea'))
    if el is None:
        el = rPr.makeelement(qn('a:ea'), {}); rPr.append(el)
    el.set('typeface', ea)

def apply_fonts(prs, lat, ea):
    """Force the deck's Latin + EA fonts on every run (textboxes + tables).
    Without this PPT falls back to the theme default (Calibri), not Arial."""
    def fix(tf):
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.name = lat; _set_ea(r, ea)
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame: fix(sh.text_frame)
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells: fix(cell.text_frame)

def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2); tf.margin_top = tf.margin_bottom = Pt(2)
    return tf

def add_shadow(shape):
    """Add an outer drop-shadow matching the HTML token 0 4px 9.3px rgba(0,0,0,.4)."""
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:effectLst')): spPr.remove(e)
    eff = spPr.makeelement(qn('a:effectLst'), {})
    sh = eff.makeelement(qn('a:outerShdw'),
                         {'blurRad': '88900', 'dist': '38100', 'dir': '5400000', 'rotWithShape': '0'})
    clr = sh.makeelement(qn('a:srgbClr'), {'val': '000000'})
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '40000'}))
    sh.append(clr); eff.append(sh); spPr.append(eff)

def add_hgradient(shape, stops):
    """Left->right linear gradient. stops: list of (pos_pct, hex, alpha_pct)."""
    spPr = shape._element.spPr
    for tag in ('a:solidFill', 'a:gradFill', 'a:noFill', 'a:blipFill', 'a:pattFill'):
        for e in spPr.findall(qn(tag)): spPr.remove(e)
    grad = spPr.makeelement(qn('a:gradFill'), {})
    gsLst = grad.makeelement(qn('a:gsLst'), {})
    for pos, hx, al in stops:
        gs = gsLst.makeelement(qn('a:gs'), {'pos': str(int(pos*1000))})
        clr = gs.makeelement(qn('a:srgbClr'), {'val': hx})
        clr.append(clr.makeelement(qn('a:alpha'), {'val': str(int(al*1000))}))
        gs.append(clr); gsLst.append(gs)
    grad.append(gsLst)
    grad.append(grad.makeelement(qn('a:lin'), {'ang': '0', 'scaled': '1'}))
    ln = spPr.find(qn('a:ln'))
    if ln is not None: ln.addprevious(grad)
    else: spPr.append(grad)

def rect(slide, x, y, w, h, fill_hex, line_hex=None, line_w=0.75, round_=False,
         shadow=False, radius_px=11):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
        px(x), px(y), px(w), px(h))
    if round_:
        try: shp.adjustments[0] = max(0.0, min(0.5, radius_px / max(1, min(w, h))))
        except Exception: pass
    if fill_hex: shp.fill.solid(); shp.fill.fore_color.rgb = hexc(fill_hex)
    else: shp.fill.background()
    if line_hex: shp.line.color.rgb = hexc(line_hex); shp.line.width = Pt(line_w)
    else: shp.line.fill.background()
    shp.shadow.inherit = False
    if shadow: add_shadow(shp)
    return shp

def set_para(p, text, size_pt, color, theme, bold=False, align=PP_ALIGN.LEFT, bullet=None):
    p.alignment = align
    if bullet:
        r = p.add_run(); r.text = bullet + "  "
        r.font.size = Pt(size_pt); r.font.bold = True
        r.font.color.rgb = hexc(theme["colors"]["gold" if bullet == "✓" else "navy"])
    add_runs(p, text, theme, color)
    for r in p.runs:
        if r.font.size is None: r.font.size = Pt(size_pt)
        if r.font.bold is None: r.font.bold = bold
    return p

class Deck:
    def __init__(self, theme, comps_path=None):
        self.t = theme; self.ly = theme["layout"]; self.c = theme["colors"]
        self.ty = theme["type_pt"]
        self.comps = clone.Components(comps_path) if comps_path and os.path.exists(comps_path) else None
        # remap the reference-deck component hexes -> active theme colours
        def H(k, fb): return self.c.get(k, fb).lstrip("#").upper()
        self.recolor_map = {
            "0F0F70": H("navy", "#0F0F70"), "161570": H("navy_deep", "#161570"),
            "060634": H("navy_grad0", "#060634"), "C5A86F": H("gold", "#C5A86F"),
            "888888": H("gray", "#888888"),
        }
        self.prs = Presentation()
        self.prs.slide_width = px(self.ly["w"]); self.prs.slide_height = px(self.ly["h"])
        self.blank = self.prs.slide_layouts[6]

    def new(self): return self.prs.slides.add_slide(self.blank)

    def _add(self, slide, sp):
        slide.shapes._spTree.append(sp); return sp

    def _clone(self, key):
        """Clone a component and recolor it to the active theme."""
        sp = self.comps.clone(key); clone.recolor(sp, self.recolor_map); return sp

    def _logo_path(self):
        fn = self.t.get("logo")
        if not fn: return None
        p = fn if os.path.isabs(fn) else os.path.join(LOGO_DIR, fn)
        return p if os.path.exists(p) else None

    def banner(self, slide, s):
        """Header band built from CLONED reference components (number block, gold
        pill chip, section label, title) so they render identically in PowerPoint.
        Original stacked layout: label on top, gold chip below it (x=129), title to
        the right (x=206), all over the SNU light-gray header band."""
        ly, c, ty = self.ly, self.c, self.ty
        bh = ly["banner_h"]
        band = rect(slide, 0, 0, ly["w"], bh, c["banner_g1"])    # themed header band
        g1 = c["banner_g1"].lstrip("#")
        add_hgradient(band, [(0, g1, 100), (60, g1, 35), (100, "FFFFFF", 100)])
        if self.comps is None:
            return self._banner_primitive(slide, s)
        em = c.get("em_color", c["gold"])
        if s.get("chapter"):
            nb = self._clone("numblock"); clone.set_xfrm(nb, 0, 0); self._add(slide, nb)
            nt = self._clone("numtext"); clone.set_xfrm(nt, 6.3, 4.7)
            clone.set_number(nt, s["chapter"]); clone.set_nowrap(nt); self._add(slide, nt)
        if s.get("kicker"):
            lbl = self._clone("seclabel"); clone.set_xfrm(lbl, 129, 0)
            clone.set_text(lbl, s["kicker"]); self._add(slide, lbl)
        if s.get("section"):
            chip = self._clone("chip"); clone.set_xfrm(chip, 129, 39.8)
            clone.set_text(chip, str(s["section"]))
            clone.set_run_color(chip, c.get("chip_text", "#FFFFFF")); self._add(slide, chip)
        if s.get("title"):
            titlex = 206.3
            ttl = self._clone("title")
            clone.set_xfrm(ttl, titlex, 26.7, ly["w"]-titlex-ly["margin_x"], 52.3)
            clone.set_markup(ttl, s["title"], c["navy"], em); self._add(slide, ttl)

    def _banner_primitive(self, slide, s):
        """Fallback for themes without a component library: draw with primitives."""
        ly, c, ty = self.ly, self.c, self.ty
        bh = ly["banner_h"]
        rect(slide, 0, 0, bh, bh, c["navy"])
        if s.get("chapter"):
            tf = textbox(slide, 0, 0, bh, bh, MSO_ANCHOR.MIDDLE)
            set_para(tf.paragraphs[0], s["chapter"], ty["number"], RGBColor(255,255,255),
                     self.t, bold=True, align=PP_ALIGN.CENTER)
        if s.get("kicker"):
            tf = textbox(slide, 129, 4, ly["w"]-129-ly["margin_x"], 20)
            set_para(tf.paragraphs[0], s["kicker"], ty["small"], hexc(c["muted"]), self.t, bold=True)
        if s.get("section"):
            cw = max(64, 18 + len(str(s["section"]))*ty["small"]*0.95)
            rect(slide, 129, 39.8, cw, 30, c["gold"], round_=True, radius_px=15)
            ctf = textbox(slide, 129, 39.8, cw, 30, MSO_ANCHOR.MIDDLE)
            set_para(ctf.paragraphs[0], str(s["section"]), ty["small"], RGBColor(255,255,255),
                     self.t, bold=True, align=PP_ALIGN.CENTER)
        tf = textbox(slide, 206.3, 26.7, ly["w"]-206.3-ly["margin_x"], 52.3, MSO_ANCHOR.MIDDLE)
        set_para(tf.paragraphs[0], s.get("title",""), ty["title"], hexc(c["ink"]), self.t, bold=True)

    def footer(self, slide, text):
        # plain caption line, no navy band (box removed per design)
        ly, c = self.ly, self.c
        tf = textbox(slide, ly["margin_x"], ly["h"]-34, ly["w"]-2*ly["margin_x"], 26, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        add_runs(p, text, self.t, hexc(c["muted"]), em_hex=c["accent_orange"])
        for r in p.runs:
            if r.font.size is None: r.font.size = Pt(self.ty["small"])
            r.font.italic = True

    # ---- body blocks (returns next y) ----
    def blocks(self, slide, body, x, y, w):
        ty, c = self.ty, self.c
        for b in body:
            t = b.get("type")
            if t == "h":
                acc = b.get("accent", "navy")
                col = c["gold"] if acc == "gold" else (c["gray"] if acc == "gray" else c["navy"])
                bw = min(w, 56 + len(b["text"]) * ty["subtitle"] * 0.95)
                if self.comps:           # cloned header bar (white text + drop-shadow)
                    hb = self._clone("cardhdr"); clone.set_xfrm(hb, x, y, bw, 40)
                    clone.set_solid_fill(hb, col); clone.set_text(hb, b["text"]); self._add(slide, hb)
                else:
                    rect(slide, x, y, bw, 40, col, round_=True, shadow=True, radius_px=10)
                    tf = textbox(slide, x, y, bw, 40, MSO_ANCHOR.MIDDLE, wrap=False)
                    set_para(tf.paragraphs[0], b["text"], ty["subtitle"], RGBColor(255,255,255),
                             self.t, bold=True, align=PP_ALIGN.CENTER)
                y += 54
            elif t == "p":
                n = wrap_lines(b["text"], w, ty["body"])
                lh = ty["body"]*96/72*1.5
                tf = textbox(slide, x, y, w, lh*n+6)
                set_para(tf.paragraphs[0], b["text"], ty["body"], hexc(c["body"]), self.t)
                y += lh*n + 8
            elif t == "bullets":
                mark = "✓" if b.get("check") else "•"
                lh = ty["body"]*96/72*1.5
                rows = sum(wrap_lines(it, w-22, ty["body"]) for it in b["items"])
                tf = textbox(slide, x, y, w, (lh+7)*rows+6)
                for i, it in enumerate(b["items"]):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    set_para(p, it, ty["body"], hexc(c["body"]), self.t, bullet=mark)
                    p.space_after = Pt(7)
                y += (lh+7)*rows + 10
            elif t == "kpi":
                # left-grouped with gap (matches HTML .kpirow flex), not full-width spread
                vpt = int(ty["number"]*0.7); cx2 = x; gap = 24
                for k in b["items"]:
                    iw = max(100, len(str(k["value"]))*vpt*0.75, len(str(k["label"]))*ty["caption"]*1.0)
                    tf = textbox(slide, cx2, y, iw, 42, MSO_ANCHOR.TOP, wrap=False)
                    set_para(tf.paragraphs[0], k["value"], vpt, hexc(c["navy"]),
                             self.t, bold=True, align=PP_ALIGN.CENTER)
                    tf2 = textbox(slide, cx2, y+42, iw, 22, MSO_ANCHOR.TOP, wrap=False)
                    set_para(tf2.paragraphs[0], k["label"], ty["caption"], hexc(c["muted"]),
                             self.t, align=PP_ALIGN.CENTER)
                    cx2 += iw + gap
                y += 76
            elif t == "table":
                self.table(slide, b, x, y, w); y += 30*(len(b["rows"])+1) + 10
            elif t == "cards":
                self.cards(slide, b["items"], x, y, w, b.get("cols", len(b["items"])), 220)
                y += 232
            elif t == "cols":
                # side-by-side columns; each is a list of blocks rendered recursively
                ncol = len(b["items"]); gap = 18
                cw = (w - gap*(ncol-1))/ncol
                maxy = y
                for ci, col in enumerate(b["items"]):
                    cx = x + ci*(cw+gap)
                    maxy = max(maxy, self.blocks(slide, col, cx, y, cw))
                y = maxy
        return y

    def table(self, slide, b, x, y, w):
        head = b.get("head", []); rows = b["rows"]
        ncol = len(head) or len(rows[0]); nrow = len(rows) + (1 if head else 0)
        gt = slide.shapes.add_table(nrow, ncol, px(x), px(y), px(w), px(30*nrow)).table
        ri = 0
        if head:
            for ci, h in enumerate(head):
                cell = gt.cell(0, ci); cell.text = ""
                p = cell.text_frame.paragraphs[0]; r = p.add_run(); r.text = str(h)
                r.font.size = Pt(self.ty["small"]); r.font.bold = True
                r.font.color.rgb = RGBColor(255,255,255)
                cell.fill.solid(); cell.fill.fore_color.rgb = hexc(self.c["navy"])
            ri = 1
        for r_i, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = gt.cell(ri+r_i, ci); cell.text = ""
                p = cell.text_frame.paragraphs[0]; run = p.add_run(); run.text = str(val)
                run.font.size = Pt(self.ty["small"]); run.font.color.rgb = hexc(self.c["ink"])
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255,255,255)

    def cards(self, slide, items, x, y, w, cols, h):
        gap = 18; cw = (w - gap*(cols-1))/cols
        keymap = {"gold": "card_gold", "gray": "card_gray", "navy": "card_navy"}
        pad = self.t["box"]["pad_px"]
        for i, cd in enumerate(items):
            cx = x + (i % cols)*(cw+gap); cy = y + (i//cols)*(h+gap)
            acc = cd.get("accent", "navy")
            if self.comps:               # cloned card frame: white fill + 3px accent border + r~11
                frame = self._clone(keymap.get(acc, "card_navy"))
                clone.set_xfrm(frame, cx, cy, cw, h); clone.clear_text(frame); self._add(slide, frame)
            else:
                rect(slide, cx, cy, cw, h, None, self.c["navy"], line_w=2.25,
                     round_=True, shadow=True, radius_px=self.t["box"]["radius_px"])
            tf = textbox(slide, cx+pad, cy+12, cw-2*pad, h-24)   # text on top of the frame
            gold_text = self.c.get("em_color", self.c["gold"])   # readable head text (lime fails on white)
            accent = gold_text if acc == "gold" else (self.c["gray"] if acc == "gray" else self.c["navy"])
            if cd.get("head"):
                set_para(tf.paragraphs[0], cd["head"], self.ty["subtitle"], hexc(accent), self.t, bold=True)
                set_para(tf.add_paragraph(), cd.get("text",""), self.ty["body"], hexc(self.c["body"]), self.t)
            else:
                set_para(tf.paragraphs[0], cd.get("text",""), self.ty["body"], hexc(self.c["body"]), self.t)

    # ---- layouts ----
    def render(self, s):
        layout = s.get("layout", "content"); slide = self.new(); ly, c, ty = self.ly, self.c, self.ty
        if layout == "cover":
            rect(slide, 0, 0, ly["w"], ly["h"], c["page_bg"])
            rect(slide, 0, 0, 12, ly["h"], c["navy"])              # left accent bar
            logo = self._logo_path()
            if logo:
                slide.shapes.add_picture(logo, px(90), px(60), height=px(62))  # width auto
            tw = ly["w"] - 180
            title = s.get("title", "")
            t_lines = wrap_lines(title, tw, ty["cover_title"])
            t_h = t_lines * ty["cover_title"] * 96/72 * 1.18
            sub = s.get("subtitle", "")
            s_h = (wrap_lines(sub, tw, ty["subtitle"]) * ty["subtitle"] * 96/72 * 1.3) if sub else 0
            block_h = t_h + 14 + 4 + (16 + s_h if sub else 0)
            y0 = max(150, (ly["h"] - 84 - block_h)/2 - 10)         # vertically center (above band)
            tf = textbox(slide, 90, y0, tw, t_h+8)
            set_para(tf.paragraphs[0], title, ty["cover_title"], hexc(c["navy"]), self.t, bold=True)
            rect(slide, 90, y0 + t_h + 12, 84, 4, c["gold"])        # accent rule
            if sub:
                tfs = textbox(slide, 90, y0 + t_h + 12 + 14, tw, s_h+8)
                set_para(tfs.paragraphs[0], sub, ty["subtitle"], hexc(c["ink"]), self.t)
            presenter = s.get("presenter") or s.get("meta", "")
            date = s.get("date", "")
            if presenter or date:                                  # bottom brand band
                rect(slide, 0, ly["h"]-84, ly["w"], 84, c["navy"])
                if presenter:
                    tfp = textbox(slide, 90, ly["h"]-84, (ly["w"]-180)*0.62, 84, MSO_ANCHOR.MIDDLE)
                    set_para(tfp.paragraphs[0], presenter, ty["body"], RGBColor(255,255,255), self.t, bold=True)
                if date:
                    tfd = textbox(slide, ly["w"]-90-360, ly["h"]-84, 360, 84, MSO_ANCHOR.MIDDLE)
                    set_para(tfd.paragraphs[0], date, ty["body"], RGBColor(255,255,255), self.t, align=PP_ALIGN.RIGHT)
        elif layout == "section":
            rect(slide, 0, 0, ly["w"], ly["h"], c["navy"])
            tf = textbox(slide, 90, 200, 300, 320, MSO_ANCHOR.MIDDLE)
            set_para(tf.paragraphs[0], s.get("chapter",""), 120, hexc(c["gold"]), self.t, bold=True)
            tf2 = textbox(slide, 420, 200, ly["w"]-500, 320, MSO_ANCHOR.MIDDLE)
            if s.get("kicker"):
                set_para(tf2.paragraphs[0], s["kicker"], ty["subtitle"], hexc(c["gold"]), self.t, bold=True)
                set_para(tf2.add_paragraph(), s.get("title",""), ty["cover_title"], RGBColor(255,255,255), self.t, bold=True)
            else:
                set_para(tf2.paragraphs[0], s.get("title",""), ty["cover_title"], RGBColor(255,255,255), self.t, bold=True)
        elif layout == "closing":
            rect(slide, 0, 0, ly["w"], ly["h"], c["navy"])
            tf = textbox(slide, 0, 250, ly["w"], 200, MSO_ANCHOR.MIDDLE)
            set_para(tf.paragraphs[0], s.get("title",""), ty["cover_title"], RGBColor(255,255,255), self.t, bold=True, align=PP_ALIGN.CENTER)
            if s.get("subtitle"):
                set_para(tf.add_paragraph(), s["subtitle"], ty["subtitle"], RGBColor(220,220,230), self.t, align=PP_ALIGN.CENTER)
        elif layout == "cards":
            self.banner(slide, s)
            ch = ly["h"] - ly["body_top"] - 50   # fill body height (match HTML)
            self.cards(slide, s.get("cards", []), ly["margin_x"], ly["body_top"],
                       ly["w"]-2*ly["margin_x"], s.get("cols", len(s.get("cards", [])) or 3), ch)
            if s.get("footer"): self.footer(slide, s["footer"])
        else:  # content
            self.banner(slide, s)
            self.blocks(slide, s.get("body", []), ly["margin_x"], ly["body_top"],
                        ly["w"]-2*ly["margin_x"])
            if s.get("footer"): self.footer(slide, s["footer"])

    def build(self, spec):
        for s in spec["slides"]: self.render(s)
        if self.comps:
            for slide in self.prs.slides: clone.dedupe_ids(slide)
        apply_fonts(self.prs, self.t["fonts"]["latin"], self.t["fonts"]["ea"])
        return self.prs

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("spec"); ap.add_argument("out")
    ap.add_argument("--theme-dir", default=THEME_DIR)
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    name = spec.get("theme", "skku")
    theme = load_theme(name, a.theme_dir)
    comps_path = os.path.join(a.theme_dir, f"{name}_components.xml")
    Deck(theme, comps_path).build(spec).save(a.out)
    print("wrote", a.out, f"({len(spec['slides'])} slides)")
