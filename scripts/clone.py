#!/usr/bin/env python3
"""Clone canonical SNU component shapes into a python-pptx slide.

The shapes in themes/<theme>_components.xml are the *original* OOXML <p:sp>
elements lifted from the reference deck. Cloning them (instead of re-drawing with
python-pptx primitives) means fills, borders, corner radius, gradients and
shadows render identically in real PowerPoint — which primitive ROUNDED_RECTANGLE
adjustments do NOT (e.g. the gold pill chip went invisible).

Usage from the exporter:
    comps = Components(".../themes/skku_components.xml")
    sp = comps.clone("chip"); set_xfrm(sp, 129, 39.8); set_text(sp, "1.2")
    slide.shapes._spTree.append(sp)
    ...
    dedupe_ids(slide)   # once per slide, after all shapes are added
"""
import copy
from lxml import etree
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

EMU = 9525  # per px


class Components:
    """Loads the component library once; clone(key) returns a fresh CT_Shape."""
    def __init__(self, path):
        tree = etree.parse(path)
        self.xml = {c.get("key"): etree.tostring(c[0], encoding="unicode")
                    for c in tree.getroot()}

    def has(self, key):
        return key in self.xml

    def clone(self, key):
        return parse_xml(self.xml[key])


# ---- geometry ----
def _spPr(sp):   return sp.find(qn("p:spPr"))
def _xfrm(sp):   return _spPr(sp).find(qn("a:xfrm"))

def set_xfrm(sp, x_px, y_px, w_px=None, h_px=None):
    """Move (and optionally resize) a cloned shape. px -> EMU."""
    xf = _xfrm(sp)
    off = xf.find(qn("a:off")); ext = xf.find(qn("a:ext"))
    off.set("x", str(int(round(x_px * EMU))))
    off.set("y", str(int(round(y_px * EMU))))
    if w_px is not None: ext.set("cx", str(int(round(w_px * EMU))))
    if h_px is not None: ext.set("cy", str(int(round(h_px * EMU))))


# ---- text ----
def _first_p(sp):
    tb = sp.find(qn("p:txBody"))
    return tb.find(qn("a:p")) if tb is not None else None

def _runs(p):  return p.findall(qn("a:r"))

def set_text(sp, text):
    """Collapse the first paragraph to a single run carrying `text`, keeping the
    first run's formatting (rPr). Falls back to building a run from endParaRPr."""
    p = _first_p(sp)
    if p is None: return
    runs = _runs(p)
    if runs:
        t = runs[0].find(qn("a:t"))
        if t is None:
            t = runs[0].makeelement(qn("a:t"), {}); runs[0].append(t)
        t.text = text
        for r in runs[1:]: p.remove(r)
    else:
        epr = p.find(qn("a:endParaRPr"))
        r = p.makeelement(qn("a:r"), {})
        if epr is not None:
            rpr = copy.deepcopy(epr); rpr.tag = qn("a:rPr"); r.append(rpr)
        t = r.makeelement(qn("a:t"), {}); t.text = text; r.append(t)
        p.insert(0, r)

def clear_text(sp):
    """Remove all runs from the first paragraph (e.g. a card used as a frame)."""
    p = _first_p(sp)
    if p is None: return
    for r in _runs(p): p.remove(r)

def set_nowrap(sp):
    """Disable text wrapping + autofit (single centered line). Needed for the
    number block: soffice otherwise wraps '01' to two lines and clips the 2nd
    digit below the box."""
    tb = sp.find(qn("p:txBody"))
    bodyPr = tb.find(qn("a:bodyPr")) if tb is not None else None
    if bodyPr is None: return
    bodyPr.set("wrap", "none")
    for tag in ("a:spAutoFit", "a:normAutofit"):
        e = bodyPr.find(qn(tag))
        if e is not None: bodyPr.remove(e)
    if bodyPr.find(qn("a:noAutofit")) is None:
        bodyPr.append(bodyPr.makeelement(qn("a:noAutofit"), {}))

def set_number(sp, chapter):
    """Number block: keep the faded leading run + solid trailing digit.
    '02' -> faded '0' + solid '2'. Single char -> all solid.

    Also strips each run's transparent text outline (<a:ln alpha=0>): it is
    invisible anyway, and it makes soffice mis-measure the text width and drop
    the 2nd run (only '0' renders). Removing it renders correctly everywhere."""
    p = _first_p(sp)
    if p is None: return
    runs = _runs(p)
    for r in runs:
        rpr = r.find(qn("a:rPr"))
        if rpr is not None:
            for ln in rpr.findall(qn("a:ln")): rpr.remove(ln)
    chapter = str(chapter)
    if len(runs) >= 2:
        head, tail = (chapter[:-1], chapter[-1]) if len(chapter) > 1 else ("", chapter)
        runs[0].find(qn("a:t")).text = head
        runs[1].find(qn("a:t")).text = tail
        for r in runs[2:]: p.remove(r)
    else:
        set_text(sp, chapter)


def set_markup(sp, text, navy_hex, gold_hex):
    """Rebuild the first paragraph's runs from inline markup, using the first
    run's rPr as the base style. **bold** -> navy bold, *x* -> gold bold."""
    import re
    p = _first_p(sp)
    if p is None: return
    runs = _runs(p)
    base = runs[0].find(qn("a:rPr")) if runs else None
    for r in runs: p.remove(r)

    def make(seg, color=None, bold=False):
        r = p.makeelement(qn("a:r"), {})
        rpr = copy.deepcopy(base) if base is not None else r.makeelement(qn("a:rPr"), {})
        rpr.tag = qn("a:rPr")
        if bold: rpr.set("b", "1")
        if color:
            for f in rpr.findall(qn("a:solidFill")): rpr.remove(f)
            sf = rpr.makeelement(qn("a:solidFill"), {})
            cl = sf.makeelement(qn("a:srgbClr"), {"val": color.lstrip("#")})
            sf.append(cl)
            # solidFill must precede a:latin etc.; insert after a:ln if present
            ln = rpr.find(qn("a:ln"))
            (ln.addnext(sf) if ln is not None else rpr.insert(0, sf))
        r.append(rpr)
        t = r.makeelement(qn("a:t"), {}); t.text = seg; r.append(t)
        return r

    epr = p.find(qn("a:endParaRPr"))
    for seg in re.split(r"(\*\*.+?\*\*|\*.+?\*)", text):
        if not seg: continue
        if seg.startswith("**") and seg.endswith("**"):
            r = make(seg[2:-2], navy_hex, True)
        elif seg.startswith("*") and seg.endswith("*"):
            r = make(seg[1:-1], gold_hex, True)
        else:
            r = make(seg)
        (epr.addprevious(r) if epr is not None else p.append(r))


def set_runs(sp, segments):
    """Rebuild runs from explicit (text, em_bool) segments, alternating between
    the shape's first two run rPr templates (e.g. footer: white / orange)."""
    p = _first_p(sp)
    if p is None: return
    runs = _runs(p)
    plain_rpr = copy.deepcopy(runs[0].find(qn("a:rPr"))) if runs else None
    em_rpr = (copy.deepcopy(runs[1].find(qn("a:rPr")))
              if len(runs) > 1 else copy.deepcopy(plain_rpr))
    for r in runs: p.remove(r)
    epr = p.find(qn("a:endParaRPr"))
    for text, em in segments:
        r = p.makeelement(qn("a:r"), {})
        tmpl = em_rpr if em else plain_rpr
        if tmpl is not None: r.append(copy.deepcopy(tmpl))
        t = r.makeelement(qn("a:t"), {}); t.text = text; r.append(t)
        (epr.addprevious(r) if epr is not None else p.append(r))


# ---- fill recolor (for accent variants of a single source shape) ----
def set_solid_fill(sp, hex_):
    """Replace the shape's solid fill colour (keeps everything else)."""
    spPr = _spPr(sp)
    for tag in ("a:solidFill", "a:gradFill", "a:noFill"):
        e = spPr.find(qn(tag))
        if e is not None: spPr.remove(e)
    sf = spPr.makeelement(qn("a:solidFill"), {})
    sf.append(sf.makeelement(qn("a:srgbClr"), {"val": hex_.lstrip("#")}))
    # fill goes before a:ln
    ln = spPr.find(qn("a:ln"))
    (ln.addprevious(sf) if ln is not None else spPr.append(sf))

def set_line_fill(sp, hex_):
    """Replace the border colour of a shape that has <a:ln><a:solidFill>."""
    ln = _spPr(sp).find(qn("a:ln"))
    if ln is None: return
    sf = ln.find(qn("a:solidFill"))
    if sf is None: return
    c = sf.find(qn("a:srgbClr"))
    if c is not None: c.set("val", hex_.lstrip("#"))


def recolor(sp, mapping):
    """Remap every srgbClr val in the shape per `mapping` (UPPERCASE hex, no '#').
    Lets one component library be re-themed: the reference deck's navy/gold/gray
    hexes become the active theme's colours (fills, borders, gradient stops)."""
    if not mapping: return
    for el in sp.iter(qn("a:srgbClr")):
        v = el.get("val")
        if v and v.upper() in mapping:
            el.set("val", mapping[v.upper()])


def set_run_color(sp, hex_):
    """Force the text colour of every run in the first paragraph (e.g. chip label
    needs dark text on a light accent fill). srgbClr must precede a:latin in rPr."""
    p = _first_p(sp)
    if p is None: return
    for r in _runs(p):
        rpr = r.find(qn("a:rPr"))
        if rpr is None:
            rpr = r.makeelement(qn("a:rPr"), {}); r.insert(0, rpr)
        for f in rpr.findall(qn("a:solidFill")): rpr.remove(f)
        sf = rpr.makeelement(qn("a:solidFill"), {})
        sf.append(sf.makeelement(qn("a:srgbClr"), {"val": hex_.lstrip("#")}))
        ln = rpr.find(qn("a:ln"))
        (ln.addnext(sf) if ln is not None else rpr.insert(0, sf))


# ---- id hygiene ----
def dedupe_ids(slide):
    """Assign unique cNvPr ids across the slide and drop cloned creationId GUIDs
    (duplicate creationIds across clones can confuse PowerPoint)."""
    spTree = slide.shapes._spTree
    cid = 1
    for cNvPr in spTree.iter(qn("p:cNvPr")):
        cNvPr.set("id", str(cid)); cid += 1
        for ext in cNvPr.findall(qn("a:extLst")):
            cNvPr.remove(ext)
