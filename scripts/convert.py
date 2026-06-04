#!/usr/bin/env python3
"""PPTX -> faithful HTML converter (reconstruction approach).
Walks OOXML spTree, emits absolutely-positioned HTML/CSS per slide.
PoC scope: shapes (solid/gradient fill, outline), text (runs w/ font/size/
color/align/anchor), images, video, groups. Theme color resolution.
"""
import os, re, sys, shutil, html, math
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}
EMU = 9525.0  # EMU per CSS px (914400/96)
PT = 96.0 / 72.0  # pt -> px

def q(tag): return f"{{{A}}}{tag}"  # drawingml
def qp(tag): return f"{{{P}}}{tag}"

# ---------- theme ----------
THEME = {}
def load_theme(theme_path):
    x = etree.parse(theme_path).getroot()
    cs = x.find(f".//{q('clrScheme')}")
    name2key = {"dk1":"dk1","lt1":"lt1","dk2":"dk2","lt2":"lt2","accent1":"accent1",
                "accent2":"accent2","accent3":"accent3","accent4":"accent4",
                "accent5":"accent5","accent6":"accent6","hlink":"hlink","folHlink":"folHlink"}
    for child in cs:
        key = etree.QName(child).localname
        sub = child[0]
        ln = etree.QName(sub).localname
        if ln == "srgbClr":
            THEME[key] = sub.get("val")
        elif ln == "sysClr":
            THEME[key] = sub.get("lastClr")
    # clrMap (standard): bg1->lt1, tx1->dk1, bg2->lt2, tx2->dk2
    THEME["bg1"]=THEME.get("lt1","FFFFFF"); THEME["tx1"]=THEME.get("dk1","000000")
    THEME["bg2"]=THEME.get("lt2","E7E6E6"); THEME["tx2"]=THEME.get("dk2","44546A")
    THEME["phClr"]=None

# ---------- color ----------
def hls_adjust(rgb, lummod=None, lumoff=None, shade=None, tint=None):
    import colorsys
    r,g,b = [c/255.0 for c in rgb]
    h,l,s = colorsys.rgb_to_hls(r,g,b)
    if lummod is not None: l *= lummod
    if lumoff is not None: l += lumoff
    if shade is not None:
        r,g,b = colorsys.hls_to_rgb(h,l,s);
        return tuple(int(max(0,min(255,c*shade*255))) for c in (r,g,b))
    if tint is not None:
        l = l*tint + (1-tint)  # blend toward white
    l = max(0,min(1,l))
    r,g,b = colorsys.hls_to_rgb(h,l,s)
    return tuple(int(max(0,min(255,c*255))) for c in (r,g,b))

PRESET = {
 "black":"000000","white":"FFFFFF","red":"FF0000","green":"008000","blue":"0000FF",
 "yellow":"FFFF00","cyan":"00FFFF","magenta":"FF00FF","gray":"808080","grey":"808080",
 "darkGray":"A9A9A9","dkGray":"A9A9A9","lightGray":"D3D3D3","ltGray":"D3D3D3",
 "orange":"FFA500","purple":"800080","dkBlue":"00008B","ltBlue":"ADD8E6",
 "dkGreen":"006400","ltGreen":"90EE90","dkRed":"8B0000","silver":"C0C0C0",
 "navy":"000080","teal":"008080","lime":"00FF00","maroon":"800000","olive":"808000",
}
def color_from(elem):
    """elem is a color container child (srgbClr/schemeClr/sysClr/prstClr). Returns (hex, alpha)."""
    ln = etree.QName(elem).localname
    if ln == "srgbClr":
        hexv = elem.get("val")
    elif ln == "schemeClr":
        v = elem.get("val")
        hexv = THEME.get(v, THEME.get({"bg1":"lt1","tx1":"dk1","bg2":"lt2","tx2":"dk2"}.get(v,v), None))
    elif ln == "sysClr":
        hexv = elem.get("lastClr","000000")
    elif ln == "prstClr":
        hexv = PRESET.get(elem.get("val"))
    else:
        return None, 1.0
    if not hexv or len(hexv)!=6:
        return None, 1.0
    rgb = tuple(int(hexv[i:i+2],16) for i in (0,2,4))
    alpha = 1.0
    lummod=lumoff=shade=tint=None
    for c in elem:
        cl = etree.QName(c).localname; val = c.get("val")
        if cl=="alpha": alpha = int(val)/100000.0
        elif cl=="lumMod": lummod = int(val)/100000.0
        elif cl=="lumOff": lumoff = int(val)/100000.0
        elif cl=="shade": shade = int(val)/100000.0
        elif cl=="tint": tint = int(val)/100000.0
    if any(v is not None for v in (lummod,lumoff,shade,tint)):
        rgb = hls_adjust(rgb, lummod, lumoff, shade, tint)
    return "%02X%02X%02X" % rgb, alpha

def css_rgba(hexv, alpha):
    r,g,b = (int(hexv[i:i+2],16) for i in (0,2,4))
    if alpha>=0.999: return f"#{hexv}"
    return f"rgba({r},{g},{b},{alpha:.3f})"

# ---------- fill ----------
# Stack of group fill containers (the <a:solidFill>/<a:gradFill> child of a
# group's grpSpPr). A child shape with <a:grpFill/> inherits the nearest one.
GRPFILL = []

def _grp_fill_elem(spPr):
    """If spPr requests grpFill, return the inherited group fill container."""
    if spPr is None or spPr.find(q("grpFill")) is None:
        return None
    for f in reversed(GRPFILL):
        if f is not None:
            return f
    return None

def _grpfill_solid_hexa(spPr):
    """Resolve a grpFill that is a solidFill -> (hex, alpha) or None."""
    f = _grp_fill_elem(spPr)
    if f is None: return None
    if etree.QName(f).localname == "solidFill" and len(f):
        return color_from(f[0])
    return None

def fill_css(spPr):
    """Return a CSS 'background' value or None."""
    gfe = _grp_fill_elem(spPr)
    if gfe is not None:
        ln = etree.QName(gfe).localname
        if ln == "solidFill" and len(gfe):
            hexv, a = color_from(gfe[0])
            if hexv: return css_rgba(hexv, a)
        if ln == "gradFill":
            spPr = gfe.getparent()  # so the gradFill branch below finds it
    sf = spPr.find(q("solidFill"))
    if sf is not None and len(sf):
        hexv, a = color_from(sf[0])
        if hexv: return css_rgba(hexv, a)
    gf = spPr.find(q("gradFill"))
    if gf is not None:
        stops = []
        gsLst = gf.find(q("gsLst"))
        if gsLst is not None:
            for gs in gsLst.findall(q("gs")):
                pos = int(gs.get("pos","0"))/1000.0
                if len(gs):
                    hexv,a = color_from(gs[0])
                    if hexv: stops.append((pos, css_rgba(hexv,a)))
        stops.sort(key=lambda s:s[0])
        lin = gf.find(q("lin"))
        ang = 90
        if lin is not None:
            ang = (int(lin.get("ang","0"))/60000.0 + 90) % 360
        if stops:
            sc = ", ".join(f"{c} {p:.2f}%" for p,c in stops)
            return f"linear-gradient({ang:.1f}deg, {sc})"
    if spPr.find(q("noFill")) is not None:
        return "transparent"
    return None

def outline_css(spPr):
    ln = spPr.find(q("ln"))
    if ln is None: return None
    if ln.find(q("noFill")) is not None: return None
    sf = ln.find(q("solidFill"))
    # empty <a:ln/> or ln without an explicit fill -> no visible border
    if sf is None or not len(sf): return None
    hexv,a = color_from(sf[0])
    if not hexv: return None
    w = ln.get("w")
    wpx = (int(w)/EMU) if w else 1.0
    return f"{wpx:.2f}px solid {css_rgba(hexv,a)}"

# ---------- geometry / transform ----------
# A transform tf=(ax,ay,sx,sy,fs) maps EMU coords in the current coordinate
# space to absolute CSS px:  px_x = ax + emu_x*sx ;  px_w = emu_w*sx.
# fs = accumulated vertical font scale (font glyphs are NOT stretched
# horizontally, matching PowerPoint group behavior).
TOP_TF = (0.0, 0.0, 1.0/EMU, 1.0/EMU, 1.0)

def shadow_parts(spPr):
    """Return (dx,dy,blur,color) for outerShdw, else None."""
    eff = spPr.find(q("effectLst"))
    if eff is None: return None
    sh = eff.find(q("outerShdw"))
    if sh is None: return None
    blur = int(sh.get("blurRad","0"))/EMU
    dist = int(sh.get("dist","0"))/EMU
    dir_ = int(sh.get("dir","0"))/60000.0
    dx = dist*math.cos(math.radians(dir_)); dy = dist*math.sin(math.radians(dir_))
    col="rgba(0,0,0,0.4)"
    if len(sh):
        hexv,a=color_from(sh[0])
        if hexv:
            r,g,b=(int(hexv[i:i+2],16) for i in (0,2,4))
            col=f"rgba({r},{g},{b},{a:.3f})"
    return (dx,dy,blur,col)

def shadow_box_css(spPr):
    s=shadow_parts(spPr)
    if not s: return ""
    dx,dy,blur,col=s
    return f"box-shadow:{dx:.1f}px {dy:.1f}px {blur:.1f}px {col};"

def shadow_filter_css(spPr):
    s=shadow_parts(spPr)
    if not s: return ""
    dx,dy,blur,col=s
    return f"filter:drop-shadow({dx:.1f}px {dy:.1f}px {blur:.1f}px {col});"

def parse_xfrm_emu(spPr):
    xf = spPr.find(q("xfrm"))
    if xf is None: return None
    off = xf.find(q("off")); ext = xf.find(q("ext"))
    if off is None or ext is None: return None
    return dict(x=int(off.get("x")), y=int(off.get("y")),
                w=int(ext.get("cx")), h=int(ext.get("cy")),
                rot=int(xf.get("rot","0"))/60000.0,
                flipH=xf.get("flipH")=="1", flipV=xf.get("flipV")=="1")

def place(tf, e):
    ax,ay,sx,sy,fs = tf
    return (ax + e["x"]*sx, ay + e["y"]*sy, e["w"]*sx, e["h"]*sy)

def compose_group(tf, ge, chOff, chExt):
    """Return child tf for a group whose own xfrm is ge (EMU in current space)."""
    ax,ay,sx,sy,fs = tf
    gx = ge["w"]/chExt["w"] if chExt["w"] else 1.0
    gy = ge["h"]/chExt["h"] if chExt["h"] else 1.0
    ax2 = ax + (ge["x"] - chOff["x"]*gx)*sx
    ay2 = ay + (ge["y"] - chOff["y"]*gy)*sy
    # PowerPoint reflows group text at its original pt within the resized box,
    # i.e. it does NOT shrink the font by the group scale -> keep fs at 1.0.
    return (ax2, ay2, sx*gx, sy*gy, 1.0)

def transform_css(g):
    parts = []
    if g.get("rot"): parts.append(f"rotate({g['rot']:.3f}deg)")
    if g.get("flipH"): parts.append("scaleX(-1)")
    if g.get("flipV"): parts.append("scaleY(-1)")
    return " ".join(parts)

# ---------- font ----------
def norm_font(name, theme_minor="Arial", theme_major="Arial"):
    if not name: return None
    if name.startswith("+mn-") or name.startswith("+mj-"):
        if name.endswith("-ea") or name.endswith("-cs"): return "Malgun Gothic"
        return theme_minor if name.startswith("+mn") else theme_major
    if "맑은" in name or "Malgun" in name: return "Malgun Gothic"
    return name

# Wingdings / Webdings bullet glyphs -> Unicode (so we don't need the font)
WINGDINGS={
 "ü":"✓",  # ü -> ✓ check
 "û":"•",  # û -> bullet
 "ý":"☑",  # ý -> ballot box with check
 "þ":"☒",  # þ -> ballot box with X
 "Ø":"●",  # Ø -> ●
 "l":"●","n":"■","u":"□","p":"❖","v":"♦",
 "§":"▪","·":"·","o":"○","q":"❑",
 "•":"•","-":"–","":"✓","":"●",
}

# ---------- text (with inheritance) ----------
def esc(t): return html.escape(t).replace("\n","<br>")

def props_from_rpr(rpr):
    """Partial props dict from an rPr/defRPr/endParaRPr element."""
    d={}
    if rpr is None: return d
    if rpr.get("sz"): d["sz"]=int(rpr.get("sz"))/100.0
    if rpr.get("b") is not None: d["b"]=(rpr.get("b")=="1")
    if rpr.get("i") is not None: d["i"]=(rpr.get("i")=="1")
    u=rpr.get("u")
    if u is not None: d["u"]=(u!="none")
    if rpr.get("spc"): d["spc"]=int(rpr.get("spc"))/100.0
    if rpr.get("baseline"): d["baseline"]=int(rpr.get("baseline"))/1000.0
    l=rpr.find(q("latin")); e=rpr.find(q("ea"))
    if l is not None and l.get("typeface"): d["latin"]=l.get("typeface")
    if e is not None and e.get("typeface"): d["ea"]=e.get("typeface")
    sf=rpr.find(q("solidFill"))
    if sf is not None and len(sf):
        hexv,a=color_from(sf[0])
        if hexv: d["color"]=css_rgba(hexv,a)
    return d

def merge(*dicts):
    m={}
    for d in dicts:
        if d: m.update({k:v for k,v in d.items() if v is not None})
    return m

def parse_liststyle(txBody):
    ls=txBody.find(q("lstStyle")); out={}
    if ls is None: return out
    for i in range(9):
        el=ls.find(q(f"lvl{i+1}pPr"))
        if el is not None:
            d=props_from_rpr(el.find(q("defRPr")))
            if el.get("algn"): d["_algn"]=el.get("algn")  # paragraph alignment default
            out[i]=d
    return out

def span_from_props(text, p, scale=1.0):
    fams=[]; nl=norm_font(p.get("latin")); ne=norm_font(p.get("ea"))
    if nl: fams.append(nl)
    if ne and ne not in fams: fams.append(ne)
    if not fams: fams=["Arial","Malgun Gothic"]
    elif "Malgun Gothic" not in fams: fams.append("Malgun Gothic")
    st=[f"font-family:{','.join(repr(f) for f in fams)}"]
    sz=p.get("sz",18.0)*scale
    st.append(f"font-size:{sz*PT:.2f}px")
    if p.get("b"): st.append("font-weight:700")
    if p.get("i"): st.append("font-style:italic")
    if p.get("u"): st.append("text-decoration:underline")
    if p.get("spc") is not None: st.append(f"letter-spacing:{p['spc']*PT:.3f}px")
    if p.get("color"): st.append(f"color:{p['color']}")
    bl=p.get("baseline")
    if bl: st.append("vertical-align:super" if bl>0 else "vertical-align:sub"); st.append("font-size:%.2fpx"%(sz*PT*0.66))
    return f'<span style="{";".join(st)}">{esc(text)}</span>'

def _spc_pts(el):
    """spcBef/spcAft/lnSpc child -> ('pct',float) or ('pts',pt) or None"""
    if el is None: return None
    pc=el.find(q("spcPct"))
    if pc is not None: return ("pct", int(pc.get("val"))/100000.0)
    pp=el.find(q("spcPts"))
    if pp is not None: return ("pts", int(pp.get("val"))/100.0)
    return None

def para_html(p, liststyle, shape_default, scale):
    ppr=p.find(q("pPr"))
    lvl=int(ppr.get("lvl","0")) if (ppr is not None and ppr.get("lvl")) else 0
    style=[]
    pdef={}
    # paragraph-alignment default inherited from the lstStyle level (if any)
    a_def=liststyle.get(lvl,{}).get("_algn")
    algn={"l":"left","ctr":"center","r":"right","just":"justify"}.get(a_def,"left")
    # default single line spacing tuned to PowerPoint (tighter than CSS 'normal')
    lh="line-height:1.0"
    if ppr is not None:
        a_=ppr.get("algn")
        if a_:
            algn={"l":"left","ctr":"center","r":"right","just":"justify"}.get(a_,"left")
        ls=_spc_pts(ppr.find(q("lnSpc")))
        if ls and ls[0]=="pct": lh=f"line-height:{ls[1]:.3f}"
        elif ls and ls[0]=="pts": lh=f"line-height:{ls[1]*PT*scale:.2f}px"
        bef=_spc_pts(ppr.find(q("spcBef")))
        aft=_spc_pts(ppr.find(q("spcAft")))
        if bef and bef[0]=="pts": style.append(f"margin-top:{bef[1]*PT*scale:.2f}px")
        if aft and aft[0]=="pts": style.append(f"margin-bottom:{aft[1]*PT*scale:.2f}px")
        pdef=props_from_rpr(ppr.find(q("defRPr")))
    style.append(lh)
    style.append(f"text-align:{algn}")
    base=merge(shape_default, liststyle.get(lvl,{}), pdef)
    # bullet / numbering (buChar/buAutoNum); buNone disables
    bullet=""
    if ppr is not None:
        bn=ppr.find(q("buNone")); bc=ppr.find(q("buChar")); ba=ppr.find(q("buAutoNum"))
        if bn is None and (bc is not None or ba is not None):
            marL=int(ppr.get("marL","0")); indent=int(ppr.get("indent","0"))
            if marL: style.append(f"padding-left:{marL/EMU:.1f}px")
            if indent: style.append(f"text-indent:{indent/EMU:.1f}px")
            if bc is not None:
                ch=bc.get("char","•"); ch=WINGDINGS.get(ch,ch)
            else:
                ch="•"
            bclr=""
            bcEl=ppr.find(q("buClr"))
            if bcEl is not None and len(bcEl):
                hx,a=color_from(bcEl[0])
                if hx: bclr=f"color:{css_rgba(hx,a)};"
            bsz=base.get("sz",18.0)*scale*PT
            bullet=f'<span style="{bclr}font-size:{bsz:.1f}px">{esc(ch)}&nbsp;</span>'
    inner=bullet
    for child in p:
        ln=etree.QName(child).localname
        if ln=="r":
            rp=merge(base, props_from_rpr(child.find(q("rPr"))))
            t_el=child.find(q("t"))
            inner+=span_from_props(t_el.text if (t_el is not None and t_el.text) else "", rp, scale)
        elif ln=="br": inner+="<br>"
        elif ln=="fld":
            t_el=child.find(q("t"))
            if t_el is not None and t_el.text: inner+=span_from_props(t_el.text, base, scale)
    if not inner: inner="<br>"
    return f'<p style="margin:0;{";".join(style)}">{inner}</p>'

def textbody_html(txBody, shape_default=None, fs=1.0):
    bpr=txBody.find(q("bodyPr"))
    anchor="t"; wrap="square"; scale=1.0
    ins=dict(l=9.6,t=4.8,r=9.6,b=4.8)
    if bpr is not None:
        anchor=bpr.get("anchor","t"); wrap=bpr.get("wrap","square")
        for k,attr in (("l","lIns"),("t","tIns"),("r","rIns"),("b","bIns")):
            v=bpr.get(attr)
            if v is not None: ins[k]=int(v)/EMU
        naf=bpr.find(q("normAutofit"))
        if naf is not None and naf.get("fontScale"):
            scale=int(naf.get("fontScale"))/100000.0
    scale *= fs
    liststyle=parse_liststyle(txBody)
    align_items={"t":"flex-start","ctr":"center","b":"flex-end"}.get(anchor,"flex-start")
    paras="".join(para_html(p, liststyle, shape_default or {}, scale) for p in txBody.findall(q("p")))
    pad=f"padding:{ins['t']:.1f}px {ins['r']:.1f}px {ins['b']:.1f}px {ins['l']:.1f}px"
    ww="white-space:nowrap;" if wrap=="none" else ""
    return (f'<div style="position:absolute;inset:0;display:flex;flex-direction:column;'
            f'justify-content:{align_items};{pad};{ww}box-sizing:border-box;">{paras}</div>')

# ---------- shapes ----------
def shape_box_style(rect, g=None, fill=None, border=None, radius=None, extra=""):
    x,y,w,h = rect
    s = (f"position:absolute;left:{x:.2f}px;top:{y:.2f}px;"
         f"width:{w:.2f}px;height:{h:.2f}px;")
    if g is not None:
        tr = transform_css(g)
        if tr: s += f"transform:{tr};"
    if fill: s += f"background:{fill};"
    if border: s += f"border:{border};box-sizing:border-box;"
    if radius: s += f"border-radius:{radius};"
    if extra: s += extra
    return s

def prst_radius(spPr, w, h):
    geom = spPr.find(q("prstGeom"))
    if geom is None: return None
    prst = geom.get("prst")
    if prst=="ellipse": return "50%"
    if prst in ("roundRect","round2SameRect","round1Rect"):
        adj=16667
        av=geom.find(q("avLst"))
        if av is not None:
            gd=av.find(q("gd"))
            if gd is not None and gd.get("fmla"):
                m=re.search(r'val\s+(\d+)', gd.get("fmla"))
                if m: adj=int(m.group(1))
        r=min(abs(w),abs(h))*adj/100000.0
        return f"{r:.1f}px"
    return None

def style_text_color(sp):
    style=sp.find(qp("style"))
    if style is None: return None
    fr=style.find(q("fontRef"))
    if fr is not None and len(fr):
        hexv,a=color_from(fr[0])
        if hexv: return css_rgba(hexv,a)
    return None

class Ctx:
    def __init__(self, rels, mediadir, outassets):
        self.rels=rels; self.mediadir=mediadir; self.outassets=outassets
        self.copied={}
    def media(self, rid):
        tgt = self.rels.get(rid)
        if not tgt: return None
        name = os.path.basename(tgt)
        if name not in self.copied:
            src = os.path.join(self.mediadir, name)
            if os.path.exists(src):
                shutil.copy(src, os.path.join(self.outassets, name))
            self.copied[name]=True
        return "assets/"+name

def crop_inner_style(blipFill):
    """Return CSS for an inner media element realizing srcRect crop + stretch."""
    l=t=r=b=0.0
    if blipFill is not None:
        sr=blipFill.find(q("srcRect"))
        if sr is not None:
            l=int(sr.get("l","0"))/100000.0; t=int(sr.get("t","0"))/100000.0
            r=int(sr.get("r","0"))/100000.0; b=int(sr.get("b","0"))/100000.0
    vw=max(1e-4,1-l-r); vh=max(1e-4,1-t-b)
    return (f"position:absolute;width:{100/vw:.4f}%;height:{100/vh:.4f}%;"
            f"left:{-l/vw*100:.4f}%;top:{-t/vh*100:.4f}%;")

_DUO=[0]
def blip_recolor(blip):
    """Return (css_filter, svg_defs) for a:blip recolor children.
    Handles a:grayscl and a:duotone (two color stops over luminance)."""
    if blip is None: return "",""
    if blip.find(q("grayscl")) is not None:
        return "filter:grayscale(1);",""
    duo = blip.find(q("duotone"))
    if duo is not None and len(duo)>=2:
        c0 = color_from(duo[0])[0]   # shadow color (luminance 0)
        c1 = color_from(duo[1])[0]   # highlight color (luminance 1)
        if c0 and c1:
            r0,g0,b0 = (int(c0[i:i+2],16)/255.0 for i in (0,2,4))
            r1,g1,b1 = (int(c1[i:i+2],16)/255.0 for i in (0,2,4))
            LR,LG,LB = 0.2126,0.7152,0.0722  # rec709 luminance
            def row(a0,a1):
                d=a1-a0
                return f"{LR*d:.5f} {LG*d:.5f} {LB*d:.5f} 0 {a0:.5f}"
            mat=(row(r0,r1)+"  "+row(g0,g1)+"  "+row(b0,b1)+"  0 0 0 1 0")
            _DUO[0]+=1; fid=f"duo{_DUO[0]}"
            defs=(f'<svg width="0" height="0" style="position:absolute">'
                  f'<filter id="{fid}" color-interpolation-filters="sRGB">'
                  f'<feColorMatrix type="matrix" values="{mat}"/></filter></svg>')
            return f"filter:url(#{fid});",defs
    return "",""

def render_pic(pic, ctx, tf):
    spPr = pic.find(qp("spPr"))
    g = parse_xfrm_emu(spPr) if spPr is not None else None
    if g is None: return ""
    rect = place(tf, g)
    radius = prst_radius(spPr, rect[2], rect[3])
    sid = spid_of(pic)
    box = shape_box_style(rect, g, radius=radius, extra="overflow:hidden;"+shadow_box_css(spPr)) + f'" data-spid="{sid}'
    nvPr = pic.find(f"{qp('nvPicPr')}/{qp('nvPr')}")
    vid = nvPr.find(q("videoFile")) if nvPr is not None else None
    blipFill = pic.find(qp("blipFill"))
    blip = blipFill.find(q("blip")) if blipFill is not None else None
    poster = None
    if blip is not None and blip.get(f"{{{R}}}embed"):
        poster = ctx.media(blip.get(f"{{{R}}}embed"))
    inner_s = crop_inner_style(blipFill)
    recolor_s, recolor_defs = blip_recolor(blip)
    if vid is not None:
        link = vid.get(f"{{{R}}}link")
        src = ctx.media(link) if link else None
        ptr = f' poster="{poster}"' if poster else ""
        if src:
            return (f'<div style="{box}"><video style="{inner_s}object-fit:fill;" '
                    f'src="{src}"{ptr} muted loop playsinline preload="metadata"></video></div>')
    if poster:
        return (f'<div style="{box}">{recolor_defs}'
                f'<img style="{inner_s}object-fit:fill;{recolor_s}" src="{poster}"></div>')
    return f'<div style="{box}"></div>'

# ---------- tables (graphicFrame / a:tbl) ----------
def _tc_border(tcPr, side):
    """side in lnL/lnR/lnT/lnB -> CSS border value or None."""
    if tcPr is None: return None
    ln = tcPr.find(q(side))
    if ln is None: return None
    if ln.find(q("noFill")) is not None: return "none"
    sf = ln.find(q("solidFill"))
    if sf is None or not len(sf): return None
    hexv,a = color_from(sf[0])
    if not hexv: return None
    w = ln.get("w"); wpx = (int(w)/EMU) if w else 1.0
    return f"{wpx:.2f}px solid {css_rgba(hexv,a)}"

def render_graphicframe(gf, ctx, tf):
    xfrm = gf.find(qp("xfrm"))
    if xfrm is None: return ""
    off = xfrm.find(q("off")); ext = xfrm.find(q("ext"))
    if off is None or ext is None: return ""
    g = dict(x=int(off.get("x")), y=int(off.get("y")),
             w=int(ext.get("cx")), h=int(ext.get("cy")), rot=0, flipH=False, flipV=False)
    tbl = gf.find(f".//{q('tbl')}")
    if tbl is None: return ""   # charts/diagrams not yet supported
    sid = spid_of(gf)
    ax,ay,sx,sy,fs = tf
    x,y,w,h = place(tf, g)
    cols = [int(gc.get("w")) for gc in tbl.findall(f"{q('tblGrid')}/{q('gridCol')}")]
    colpx = [c*sx for c in cols]
    rows = tbl.findall(q("tr"))
    out = [f'<table data-spid="{sid}" style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;'
           f'width:{w:.2f}px;height:{h:.2f}px;border-collapse:collapse;table-layout:fixed;">']
    out.append("<colgroup>" + "".join(f'<col style="width:{c:.2f}px;">' for c in colpx) + "</colgroup>")
    for tr in rows:
        rh = int(tr.get("h","0"))*sy
        out.append(f'<tr style="height:{rh:.2f}px;">')
        for tc in tr.findall(q("tc")):
            if tc.get("hMerge")=="1" or tc.get("vMerge")=="1": continue
            tcPr = tc.find(q("tcPr"))
            attrs = ""
            if tc.get("gridSpan"): attrs += f' colspan="{tc.get("gridSpan")}"'
            if tc.get("rowSpan"): attrs += f' rowspan="{tc.get("rowSpan")}"'
            st = ["overflow:hidden"]
            anchor = tcPr.get("anchor") if tcPr is not None else None
            st.append("vertical-align:" + {"ctr":"middle","b":"bottom"}.get(anchor,"top"))
            mar = dict(l=91440, r=91440, t=45720, b=45720)
            if tcPr is not None:
                for k,attr in (("l","marL"),("r","marR"),("t","marT"),("b","marB")):
                    v = tcPr.get(attr)
                    if v is not None: mar[k] = int(v)
            st.append(f"padding:{mar['t']/EMU:.2f}px {mar['r']/EMU:.2f}px "
                      f"{mar['b']/EMU:.2f}px {mar['l']/EMU:.2f}px")
            if tcPr is not None:
                sf = tcPr.find(q("solidFill"))
                if sf is not None and len(sf):
                    hexv,a = color_from(sf[0])
                    if hexv: st.append(f"background:{css_rgba(hexv,a)}")
                gf2 = tcPr.find(q("gradFill"))
                if gf2 is None:
                    pass
                for side,css in (("lnL","border-left"),("lnR","border-right"),
                                 ("lnT","border-top"),("lnB","border-bottom")):
                    b = _tc_border(tcPr, side)
                    if b and b!="none": st.append(f"{css}:{b}")
            txBody = tc.find(q("txBody"))
            inner = ""
            if txBody is not None:
                ls = parse_liststyle(txBody)
                inner = "".join(para_html(p, ls, {}, fs) for p in txBody.findall(q("p")))
            out.append(f'<td{attrs} style="{";".join(st)};">{inner}</td>')
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)

_UID=[0]
def _uid():
    _UID[0]+=1; return f"g{_UID[0]}"

def svg_paint(container, uid):
    """Return (paint, defs_html) for an spPr-like element's fill."""
    # grpFill inheritance: substitute the group's fill container
    gfe = _grp_fill_elem(container)
    if gfe is not None:
        container = gfe.getparent()
    gf = container.find(q("gradFill"))
    if gf is not None:
        stops=[]
        gsLst=gf.find(q("gsLst"))
        if gsLst is not None:
            for gs in gsLst.findall(q("gs")):
                pos=int(gs.get("pos","0"))/1000.0
                if len(gs):
                    hexv,a=color_from(gs[0])
                    if hexv: stops.append((pos,hexv,a))
        stops.sort(key=lambda s:s[0])
        lin=gf.find(q("lin")); ang=0
        if lin is not None: ang=int(lin.get("ang","0"))/60000.0
        st="".join(f'<stop offset="{p:.2f}%" stop-color="#{h}" stop-opacity="{a:.3f}"/>' for p,h,a in stops)
        defs=(f'<defs><linearGradient id="{uid}" gradientUnits="objectBoundingBox" '
              f'x1="0" y1="0" x2="1" y2="0" gradientTransform="rotate({ang:.2f} 0.5 0.5)">{st}'
              f'</linearGradient></defs>')
        return f"url(#{uid})", defs
    sf = container.find(q("solidFill"))
    if sf is not None and len(sf):
        hexv,a=color_from(sf[0])
        if hexv: return css_rgba(hexv,a), ""
    return "none", ""

def line_markers(lnEl, stroke):
    """Build SVG arrowhead markers from a:ln headEnd/tailEnd.
    Returns (defs_html, start_attr, end_attr). markerUnits='strokeWidth'."""
    if lnEl is None: return "", "", ""
    LEN={"sm":3.0,"med":4.5,"lg":6.0}; WID={"sm":3.0,"med":4.5,"lg":6.0}
    defs=[]; start_attr=""; end_attr=""
    for end,tag in (("start","headEnd"),("end","tailEnd")):
        e=lnEl.find(q(tag))
        if e is None: continue
        t=e.get("type","none")
        if t in (None,"none"): continue
        ln_=LEN.get(e.get("len","med"),4.5); wd=WID.get(e.get("w","med"),4.5)
        uid=_uid()
        if t in ("triangle","arrow","stealth","open"):
            shape=(f'<path d="M0 0L{ln_:.1f} {wd/2:.1f}L0 {wd:.1f}Z" fill="{stroke}"/>'
                   if t!="open" else
                   f'<path d="M0 0L{ln_:.1f} {wd/2:.1f}L0 {wd:.1f}" fill="none" '
                   f'stroke="{stroke}" stroke-width="1"/>')
            refx=ln_
        elif t in ("oval","diamond"):
            shape=f'<circle cx="{ln_/2:.1f}" cy="{wd/2:.1f}" r="{wd/2:.1f}" fill="{stroke}"/>'
            refx=ln_/2
        else:
            shape=f'<path d="M0 0L{ln_:.1f} {wd/2:.1f}L0 {wd:.1f}Z" fill="{stroke}"/>'; refx=ln_
        orient="auto-start-reverse" if end=="start" else "auto"
        defs.append(f'<marker id="{uid}" markerUnits="strokeWidth" '
                    f'markerWidth="{ln_+1:.1f}" markerHeight="{wd+1:.1f}" '
                    f'refX="{refx:.1f}" refY="{wd/2:.1f}" orient="{orient}">{shape}</marker>')
        if end=="start": start_attr=f' marker-start="url(#{uid})"'
        else: end_attr=f' marker-end="url(#{uid})"'
    return ("".join(defs), start_attr, end_attr)

def custgeom_svg(custGeom, rect, spPr):
    pathLst = custGeom.find(q("pathLst"))
    if pathLst is None: return None
    paths=pathLst.findall(q("path"))
    if not paths: return None
    pw=int(paths[0].get("w","0")) or 1; ph=int(paths[0].get("h","0")) or 1
    d=""
    for path in paths:
        for cmd in path:
            ln=etree.QName(cmd).localname
            pts=[(int(p.get("x")),int(p.get("y"))) for p in cmd.findall(q("pt"))]
            if ln=="moveTo" and pts: d+=f"M{pts[0][0]} {pts[0][1]} "
            elif ln=="lnTo" and pts: d+=f"L{pts[0][0]} {pts[0][1]} "
            elif ln=="cubicBezTo" and len(pts)==3:
                d+=f"C{pts[0][0]} {pts[0][1]} {pts[1][0]} {pts[1][1]} {pts[2][0]} {pts[2][1]} "
            elif ln=="quadBezTo" and len(pts)==2:
                d+=f"Q{pts[0][0]} {pts[0][1]} {pts[1][0]} {pts[1][1]} "
            elif ln=="close": d+="Z "
    if not d.strip(): return None
    uid=_uid()
    paint,defs=svg_paint(spPr,uid)
    # stroke
    stroke="none"; sw=0
    lnEl=spPr.find(q("ln"))
    if lnEl is not None and lnEl.find(q("noFill")) is None:
        sf=lnEl.find(q("solidFill"))
        if sf is not None and len(sf):
            hexv,a=color_from(sf[0])
            if hexv:
                stroke=css_rgba(hexv,a)
                w=lnEl.get("w"); sw=(int(w)/EMU) if w else 1.0
    x,y,w,h=rect
    mdefs,ms,me = line_markers(lnEl, stroke if stroke!="none" else "#000000")
    swv = sw*pw/max(w,1)
    return (f'<svg style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;'
            f'width:{w:.2f}px;height:{h:.2f}px;overflow:visible;{shadow_filter_css(spPr)}" '
            f'viewBox="0 0 {pw} {ph}" '
            f'preserveAspectRatio="none">{defs}{mdefs}<path d="{d.strip()}" fill="{paint}" '
            f'stroke="{stroke}" stroke-width="{swv:.2f}"{ms}{me}/></svg>')

def _adjs(geom):
    """Return {0:val,1:val,...} from avLst gd fmla 'val N'."""
    adj={}
    av=geom.find(q("avLst"))
    if av is None: return adj
    for i,gd in enumerate(av.findall(q("gd"))):
        m=re.search(r'val\s+(-?\d+)', gd.get("fmla",""))
        if m:
            v=int(m.group(1)); adj[i]=v
            if gd.get("name"): adj[gd.get("name")]=v
    return adj

def _poly(pts):
    return "M"+"L".join(f"{x:.1f} {y:.1f}" for x,y in pts)+"Z"

def _arrow_path(prst, w, h, a1, a2):
    ss=min(w,h); cy=h/2.0; cx=w/2.0
    if prst=="rightArrow":
        dx=min(ss*a2/100000.0, w); xn=w-dx; sh=h*a1/100000.0/2.0; yt=cy-sh; yb=cy+sh
        return _poly([(0,yt),(xn,yt),(xn,0),(w,cy),(xn,h),(xn,yb),(0,yb)])
    if prst=="leftArrow":
        dx=min(ss*a2/100000.0, w); xn=dx; sh=h*a1/100000.0/2.0; yt=cy-sh; yb=cy+sh
        return _poly([(w,yt),(xn,yt),(xn,0),(0,cy),(xn,h),(xn,yb),(w,yb)])
    if prst=="upArrow":
        dy=min(ss*a2/100000.0, h); yn=dy; sw=w*a1/100000.0/2.0; xl=cx-sw; xr=cx+sw
        return _poly([(xl,h),(xl,yn),(0,yn),(cx,0),(w,yn),(xr,yn),(xr,h)])
    if prst=="downArrow":
        dy=min(ss*a2/100000.0, h); yn=h-dy; sw=w*a1/100000.0/2.0; xl=cx-sw; xr=cx+sw
        return _poly([(xl,0),(xl,yn),(0,yn),(cx,h),(w,yn),(xr,yn),(xr,0)])
    if prst=="leftRightArrow":
        dx=min(ss*a2/100000.0, w/2.0); sh=h*a1/100000.0/2.0
        yt=cy-sh; yb=cy+sh
        d=(f"M0 {cy:.1f}L{dx:.1f} 0L{dx:.1f} {yt:.1f}L{w-dx:.1f} {yt:.1f}L{w-dx:.1f} 0"
           f"L{w:.1f} {cy:.1f}L{w-dx:.1f} {h:.1f}L{w-dx:.1f} {yb:.1f}L{dx:.1f} {yb:.1f}L{dx:.1f} {h:.1f}Z")
        return d
    return None

def prst_path(prst, w, h, adj):
    """Return SVG path 'd' (in px coords 0..w,0..h) for a preset, or None."""
    a1=adj.get("adj1", adj.get(0, 50000)); a2=adj.get("adj2", adj.get(1, 50000))
    a=adj.get("adj", adj.get(0, None))
    if prst in ("rightArrow","leftArrow","upArrow","downArrow","leftRightArrow"):
        return _arrow_path(prst,w,h,a1,a2)
    if prst=="triangle":
        ax=(a if a is not None else 50000)/100000.0
        return f"M{ax*w:.1f} 0L{w:.1f} {h:.1f}L0 {h:.1f}Z"
    if prst in ("rtTriangle",):
        return f"M0 0L0 {h:.1f}L{w:.1f} {h:.1f}Z"
    if prst=="trapezoid":
        t=(a if a is not None else 25000)/100000.0; off=t*w
        return f"M{off:.1f} 0L{w-off:.1f} 0L{w:.1f} {h:.1f}L0 {h:.1f}Z"
    if prst=="parallelogram":
        t=(a if a is not None else 25000)/100000.0; off=min(t*w, w)
        return f"M{off:.1f} 0L{w:.1f} 0L{w-off:.1f} {h:.1f}L0 {h:.1f}Z"
    if prst=="round2DiagRect":
        # top-left + bottom-right rounded (adj1), other diag square (adj2)
        r1=min(abs(w),abs(h))*((adj.get("adj1", adj.get(0,16667)))/100000.0)
        r2=min(abs(w),abs(h))*((adj.get("adj2", adj.get(1,0)))/100000.0)
        return (f"M{r1:.1f} 0L{w-r2:.1f} 0"
                + (f"A{r2:.1f} {r2:.1f} 0 0 1 {w:.1f} {r2:.1f}" if r2>0.5 else f"L{w:.1f} 0")
                + f"L{w:.1f} {h-r1:.1f}A{r1:.1f} {r1:.1f} 0 0 1 {w-r1:.1f} {h:.1f}"
                + f"L{r2:.1f} {h:.1f}"
                + (f"A{r2:.1f} {r2:.1f} 0 0 1 0 {h-r2:.1f}" if r2>0.5 else f"L0 {h:.1f}")
                + f"L0 {r1:.1f}A{r1:.1f} {r1:.1f} 0 0 1 {r1:.1f} 0Z")
    if prst in ("homePlate","chevron"):
        t=(a if a is not None else 50000)/100000.0; dx=min(t*h, w)
        if prst=="homePlate":
            return f"M0 0L{w-dx:.1f} 0L{w:.1f} {h/2:.1f}L{w-dx:.1f} {h:.1f}L0 {h:.1f}Z"
        return f"M0 0L{w-dx:.1f} 0L{w:.1f} {h/2:.1f}L{w-dx:.1f} {h:.1f}L0 {h:.1f}L{dx:.1f} {h/2:.1f}Z"
    if prst=="leftBracket":
        r=min((a if a is not None else 8333)/100000.0*min(w,h), w, h/2)
        return f"M{w:.1f} 0L{r:.1f} 0Q0 0 0 {r:.1f}L0 {h-r:.1f}Q0 {h:.1f} {r:.1f} {h:.1f}L{w:.1f} {h:.1f}"
    if prst=="rightBracket":
        r=min((a if a is not None else 8333)/100000.0*min(w,h), w, h/2)
        return f"M0 0L{w-r:.1f} 0Q{w:.1f} 0 {w:.1f} {r:.1f}L{w:.1f} {h-r:.1f}Q{w:.1f} {h:.1f} {w-r:.1f} {h:.1f}L0 {h:.1f}"
    if prst in ("brace","leftBrace","rightBrace"):
        r=min((a if a is not None else 8333)/100000.0*min(w,h), w/2, h/4)
        d=(f"M{w:.1f} 0Q{w/2:.1f} 0 {w/2:.1f} {r:.1f}L{w/2:.1f} {h/2-r:.1f}"
           f"Q{w/2:.1f} {h/2:.1f} 0 {h/2:.1f}Q{w/2:.1f} {h/2:.1f} {w/2:.1f} {h/2+r:.1f}"
           f"L{w/2:.1f} {h-r:.1f}Q{w/2:.1f} {h:.1f} {w:.1f} {h:.1f}")
        return d
    return None

def _magnetic_disk_paths(w, h):
    """flowChartMagneticDisk = upright cylinder (database symbol). ry=h/6."""
    ry = h/6.0; rx = w/2.0
    outline = (f"M0 {ry:.2f}"
               f"A{rx:.2f} {ry:.2f} 0 0 1 {w:.2f} {ry:.2f}"
               f"L{w:.2f} {h-ry:.2f}"
               f"A{rx:.2f} {ry:.2f} 0 0 1 0 {h-ry:.2f}"
               f"L0 {ry:.2f}Z")
    lines = f"M0 {ry:.2f}A{rx:.2f} {ry:.2f} 0 0 0 {w:.2f} {ry:.2f}"
    return outline, lines

def render_magnetic_disk(sp, spPr, rect, g, sid, adj):
    x,y,w,h = rect
    aw, ah = abs(w), abs(h)
    outline, lines = _magnetic_disk_paths(aw, ah)
    uid=_uid(); paint,defs=svg_paint(spPr,uid)
    stroke="none"; sw=1.0
    lnEl=spPr.find(q("ln"))
    if lnEl is not None and lnEl.find(q("noFill")) is None:
        lsf=lnEl.find(q("solidFill"))
        if lsf is not None and len(lsf):
            hx,al=color_from(lsf[0])
            if hx: stroke=css_rgba(hx,al)
        ww=lnEl.get("w"); sw=(int(ww)/EMU) if ww else 1.0
    tr=transform_css(g); trs=f"transform:{tr};" if tr else ""
    svg=(f'<svg style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;'
         f'width:{aw:.2f}px;height:{ah:.2f}px;overflow:visible;{trs}{shadow_filter_css(spPr)}" '
         f'viewBox="0 0 {aw:.2f} {ah:.2f}">{defs}'
         f'<path d="{outline}" fill="{paint}" stroke="{stroke}" stroke-width="{sw:.2f}"/>'
         f'<path d="{lines}" fill="none" stroke="{stroke}" stroke-width="{sw:.2f}"/></svg>')
    txBody=sp.find(qp("txBody"))
    overlay=""
    if txBody is not None:
        sc=style_text_color(sp); shd={"color":sc} if sc else {}
        overlay=(f'<div style="{shape_box_style(rect,g)}">{textbody_html(txBody,shd,fs=1.0)}</div>')
    return f'<div data-spid="{sid}">{svg}{overlay}</div>'

def render_cube(spPr, rect, g, sid, adj):
    x,y,w,h=rect
    d=(adj.get("adj", adj.get(0, 85000)))/100000.0
    dep=min(d*min(w,h), w*0.9, h*0.9)
    paint,defs=svg_paint(spPr,_uid())
    # base color for shading: try solidFill hex
    base=None
    sf=spPr.find(q("solidFill"))
    if sf is not None and len(sf):
        hx,al=color_from(sf[0])
        if hx: base=(hx,al)
    if base is None:
        ghx=_grpfill_solid_hexa(spPr)
        if ghx and ghx[0]: base=ghx
    def shade(hx,f):
        r,g_,b=(int(hx[i:i+2],16) for i in (0,2,4))
        return "#%02X%02X%02X"%(int(r*f),int(g_*f),int(b*f))
    front=paint; top=paint; right=paint
    if base:
        hx,al=base
        front=css_rgba(hx,al); top=css_rgba(shade(hx,1.0)[1:],al); right=css_rgba(shade(hx,0.78)[1:],al)
        top=css_rgba(("%02X%02X%02X"%tuple(min(255,int(int(hx[i:i+2],16)*1.15)) for i in (0,2,4))),al)
    # cube edge stroke (a:ln) -> draw on every face so the wireframe shows
    stroke="none"; sw=0
    lnEl=spPr.find(q("ln"))
    if lnEl is not None and lnEl.find(q("noFill")) is None:
        lsf=lnEl.find(q("solidFill"))
        if lsf is not None and len(lsf):
            hx2,al2=color_from(lsf[0])
            if hx2: stroke=css_rgba(hx2,al2); ww=lnEl.get("w"); sw=(int(ww)/EMU) if ww else 1.0
    sa=f' stroke="{stroke}" stroke-width="{sw:.2f}" stroke-linejoin="round"' if stroke!="none" else ""
    pf=f"M0 {dep:.1f}L{w-dep:.1f} {dep:.1f}L{w-dep:.1f} {h:.1f}L0 {h:.1f}Z"
    pt=f"M0 {dep:.1f}L{dep:.1f} 0L{w:.1f} 0L{w-dep:.1f} {dep:.1f}Z"
    pr=f"M{w-dep:.1f} {dep:.1f}L{w:.1f} 0L{w:.1f} {h-dep:.1f}L{w-dep:.1f} {h:.1f}Z"
    tr=transform_css(g); trs=f"transform:{tr};" if tr else ""
    return (f'<svg data-spid="{sid}" style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;'
            f'width:{w:.2f}px;height:{h:.2f}px;overflow:visible;{trs}{shadow_filter_css(spPr)}" '
            f'viewBox="0 0 {w:.2f} {h:.2f}">{defs}'
            f'<path d="{pf}" fill="{front}"{sa}/><path d="{pt}" fill="{top}"{sa}/>'
            f'<path d="{pr}" fill="{right}"{sa}/></svg>')

def prst_svg(sp, spPr, rect, g, tf):
    """Render a preset geometry as SVG (fill/stroke/text). Returns HTML or None."""
    geom=spPr.find(q("prstGeom"))
    if geom is None: return None
    prst=geom.get("prst")
    if prst is None: return None
    adj=_adjs(geom)
    sid=spid_of(sp)
    if prst=="cube":
        return render_cube(spPr, rect, g, sid, adj)
    if prst=="flowChartMagneticDisk":
        return render_magnetic_disk(sp, spPr, rect, g, sid, adj)
    x,y,w,h=rect
    d=prst_path(prst, abs(w), abs(h), adj)
    if d is None: return None
    uid=_uid(); paint,defs=svg_paint(spPr,uid)
    bracket = prst in ("leftBracket","rightBracket","brace","leftBrace","rightBrace")
    if bracket: paint="none"
    stroke="none"; sw=0
    lnEl=spPr.find(q("ln"))
    if lnEl is not None and lnEl.find(q("noFill")) is None:
        lsf=lnEl.find(q("solidFill"))
        if lsf is not None and len(lsf):
            hx,al=color_from(lsf[0])
            if hx: stroke=css_rgba(hx,al); ww=lnEl.get("w"); sw=(int(ww)/EMU) if ww else 1.0
    if bracket and stroke=="none": stroke="#000000"; sw=1.0
    tr=transform_css(g); trs=f"transform:{tr};" if tr else ""
    svg=(f'<svg style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;'
         f'width:{abs(w):.2f}px;height:{abs(h):.2f}px;overflow:visible;{trs}{shadow_filter_css(spPr)}" '
         f'viewBox="0 0 {abs(w):.2f} {abs(h):.2f}">{defs}'
         f'<path d="{d}" fill="{paint}" stroke="{stroke}" stroke-width="{sw:.2f}"/></svg>')
    txBody=sp.find(qp("txBody"))
    overlay=""
    if txBody is not None:
        sc=style_text_color(sp); shd={"color":sc} if sc else {}
        overlay=(f'<div style="{shape_box_style(rect,g)}">{textbody_html(txBody,shd,fs=tf[4])}</div>')
    return f'<div data-spid="{sid}">{svg}{overlay}</div>'

def is_line_geom(spPr):
    geom=spPr.find(q("prstGeom"))
    if geom is None: return False
    return geom.get("prst","").startswith(("line","straightConnector","bentConnector","curvedConnector"))

def render_line(sp, spPr, rect, g):
    lnEl=spPr.find(q("ln"))
    stroke="#000000"; sw=1.0; dash=""
    if lnEl is not None:
        sf=lnEl.find(q("solidFill"))
        if sf is not None and len(sf):
            hexv,a=color_from(sf[0])
            if hexv: stroke=css_rgba(hexv,a)
        w=lnEl.get("w"); sw=(int(w)/EMU) if w else 1.0
    x,y,w,h=rect
    # line runs corner-to-corner; flipV swaps y direction
    x1,y1,x2,y2=0,0,abs(w),abs(h)
    if g.get("flipV"): y1,y2=abs(h),0
    if g.get("flipH"): x1,x2=abs(w),0
    mdefs,ms,me = line_markers(lnEl, stroke)
    return (f'<svg style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;'
            f'width:{abs(w):.2f}px;height:{abs(h):.2f}px;overflow:visible">{mdefs}'
            f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{stroke}" '
            f'stroke-width="{sw:.2f}"{ms}{me}/></svg>')

def render_sp(sp, ctx, tf):
    spPr = sp.find(qp("spPr"))
    if spPr is None: return ""
    g = parse_xfrm_emu(spPr)
    if g is None: return ""
    rect = place(tf, g)
    sid = spid_of(sp)
    # connector / line
    if is_line_geom(spPr) and sp.find(qp("txBody")) is None:
        return render_line(sp, spPr, rect, g).replace("<svg ", f'<svg data-spid="{sid}" ',1)
    # custom freeform geometry -> SVG path
    custGeom = spPr.find(q("custGeom"))
    if custGeom is not None:
        svg = custgeom_svg(custGeom, rect, spPr)
        if svg:
            txBody = sp.find(qp("txBody"))
            overlay=""
            if txBody is not None:
                sc=style_text_color(sp); shd={"color":sc} if sc else {}
                overlay=(f'<div style="{shape_box_style(rect,g)}">'
                         f'{textbody_html(txBody,shd,fs=tf[4])}</div>')
            return f'<div data-spid="{sid}">{svg}{overlay}</div>'
    # preset geometry shapes (arrows / triangle / trapezoid / parallelogram / cube / bracket ...)
    if prst_radius(spPr, rect[2], rect[3]) is None:
        psvg = prst_svg(sp, spPr, rect, g, tf)
        if psvg is not None: return psvg
    fill = fill_css(spPr)
    border = outline_css(spPr)
    radius = prst_radius(spPr, rect[2], rect[3])
    style = shape_box_style(rect, g, fill, border, radius) + shadow_box_css(spPr)
    inner = ""
    txBody = sp.find(qp("txBody"))
    if txBody is not None:
        sc = style_text_color(sp)
        shape_default = {"color": sc} if sc else {}
        inner = textbody_html(txBody, shape_default, fs=tf[4])
    return f'<div data-spid="{sid}" style="{style}">{inner}</div>'

def render_grp(grp, ctx, tf):
    sid = spid_of(grp)
    grpSpPr = grp.find(qp("grpSpPr"))
    # push this group's fill so child shapes with <a:grpFill/> inherit it
    gfill = None
    if grpSpPr is not None:
        gfill = grpSpPr.find(q("solidFill"))
        if gfill is None: gfill = grpSpPr.find(q("gradFill"))
    GRPFILL.append(gfill)
    grp_tr = ""
    try:
        g = parse_xfrm_emu(grpSpPr)
        if g is None:
            inner = render_children(grp, ctx, tf)
        else:
            xf = grpSpPr.find(q("xfrm"))
            chOff = xf.find(q("chOff")); chExt = xf.find(q("chExt"))
            cO = dict(x=int(chOff.get("x")), y=int(chOff.get("y")))
            cE = dict(w=int(chExt.get("cx")), h=int(chExt.get("cy")))
            childtf = compose_group(tf, g, cO, cE)
            inner = render_children(grp, ctx, childtf)
            # the group's own rotation/flip rigidly rotates the whole child
            # coordinate system about the group center. Children are in absolute
            # slide px, so rotate the inset:0 wrapper about that center point.
            tcss = transform_css(g)
            if tcss:
                gx, gy, gw, gh = place(tf, g)
                cxp = gx + gw/2.0; cyp = gy + gh/2.0
                grp_tr = f"transform:{tcss};transform-origin:{cxp:.2f}px {cyp:.2f}px;"
    finally:
        GRPFILL.pop()
    # wrap in a non-distorting container so the group is animation-targetable
    return f'<div data-spid="{sid}" style="position:absolute;inset:0;{grp_tr}">{inner}</div>'

def _mc_pick(el):
    """For <mc:AlternateContent>, return the element whose children to render
    (first Choice, else Fallback)."""
    choice=None; fallback=None
    for c in el:
        l=etree.QName(c).localname
        if l=="Choice" and choice is None: choice=c
        elif l=="Fallback": fallback=c
    return choice if choice is not None else fallback

def render_children(tree, ctx, tf):
    out=[]
    for el in tree:
        ln = etree.QName(el).localname
        if ln=="sp": out.append(render_sp(el, ctx, tf))
        elif ln=="pic": out.append(render_pic(el, ctx, tf))
        elif ln=="grpSp": out.append(render_grp(el, ctx, tf))
        elif ln=="cxnSp": out.append(render_sp(el, ctx, tf))
        elif ln=="graphicFrame": out.append(render_graphicframe(el, ctx, tf))
        elif ln=="AlternateContent":
            pick=_mc_pick(el)
            if pick is not None: out.append(render_children(pick, ctx, tf))
    return "".join(out)

# ---------- rels ----------
def load_rels(path):
    rels={}
    if not os.path.exists(path): return rels
    x = etree.parse(path).getroot()
    for rel in x:
        rels[rel.get("Id")] = rel.get("Target").replace("../","")
    return rels

# ---------- main ----------
def rels_path_for(p):
    return os.path.join(os.path.dirname(p),"_rels",os.path.basename(p)+".rels")

def find_rel_by_type(rels_path, typesub):
    if not os.path.exists(rels_path): return None
    x=etree.parse(rels_path).getroot()
    for rel in x:
        if typesub in rel.get("Type"):
            return rel.get("Target").replace("../","")
    return None

def spid_of(el):
    for tag in ("nvSpPr","nvPicPr","nvCxnSpPr","nvGrpSpPr","nvGraphicFramePr"):
        nv=el.find(qp(tag))
        if nv is not None:
            c=nv.find(qp("cNvPr"))
            if c is not None and c.get("id"): return c.get("id")
    return ""

def is_placeholder(sp):
    nvPr = sp.find(f"{qp('nvSpPr')}/{qp('nvPr')}")
    return nvPr is not None and nvPr.find(qp("ph")) is not None

def has_field(sp, types=("datetime","slidenum","datetimeFigureOut")):
    for fld in sp.iter(q("fld")):
        t=fld.get("type","")
        if any(t.startswith(x) or t=="slidenum" or t.startswith("datetime") for x in types): return True
    return False

def render_bg(root, ctx):
    bg = root.find(f"{qp('cSld')}/{qp('bg')}")
    if bg is None: return ""
    bgPr = bg.find(qp("bgPr"))
    if bgPr is not None:
        fill = fill_css(bgPr)
        if fill: return f'<div style="position:absolute;inset:0;background:{fill};"></div>'
    return ""

def render_part(spTree, ctx, tf, skip_ph):
    out=[]
    for el in spTree:
        ln=etree.QName(el).localname
        if ln=="sp":
            if skip_ph and (is_placeholder(el) or has_field(el)): continue
            out.append(render_sp(el,ctx,tf))
        elif ln=="pic": out.append(render_pic(el,ctx,tf))
        elif ln=="grpSp": out.append(render_grp(el,ctx,tf))
        elif ln=="cxnSp": out.append(render_sp(el,ctx,tf))
        elif ln=="graphicFrame": out.append(render_graphicframe(el,ctx,tf))
        elif ln=="AlternateContent":
            pick=_mc_pick(el)
            if pick is not None: out.append(render_part(pick, ctx, tf, skip_ph))
    return "".join(out)

def parse_builds(slide_root):
    """Return ({spid: (build, type)}, nbuilds).
    type in {entr, exit, emph, media, path}. Only 'entr' hides until its build.
    Click order follows mainSeq's top-level <p:par> children (one per click)."""
    timing = slide_root.find(qp("timing"))
    if timing is None: return {}, 0
    mainSeq=None
    for cTn in timing.iter(qp("cTn")):
        if cTn.get("nodeType")=="mainSeq": mainSeq=cTn; break
    if mainSeq is None: return {}, 0
    childTn=mainSeq.find(qp("childTnLst"))
    if childTn is None: return {}, 0
    info={}; build=0
    for clickpar in childTn.findall(qp("par")):
        build+=1
        # effect cTns are those carrying a presetClass (entr/exit/emph/path/mediacall)
        for cTn in clickpar.iter(qp("cTn")):
            pc=cTn.get("presetClass")
            etype=None
            if pc=="entr": etype="entr"
            elif pc=="exit": etype="exit"
            elif pc=="emph": etype="emph"
            elif pc=="path": etype="path"
            elif pc=="mediacall": etype="media"
            if etype is None:
                # media play command without presetClass
                if cTn.find(f".//{qp('cmd')}") is not None: etype="media"
                else: continue
            for spt in cTn.iter(qp("spTgt")):
                sid=spt.get("spid")
                if sid and sid not in info:
                    info[sid]=(build, etype)
    return info, build

def _render_parts(pptdir, slide_no, outassets):
    """Render master+layout (background) and the slide part separately.
    Returns (pre_html, slide_html, slide_root). Keeping them separate lets
    data-build injection target ONLY the slide part — master/layout shapes
    can reuse the same cNvPr ids (id collision) and must not be animated."""
    os.makedirs(outassets, exist_ok=True)
    mediadir = os.path.join(pptdir,"media")
    slide_path = os.path.join(pptdir,"slides",f"slide{slide_no}.xml")
    layout_t = find_rel_by_type(rels_path_for(slide_path), "slideLayout")
    parts=[]   # (path, skip_ph) bottom-first
    if layout_t:
        layout_path = os.path.join(pptdir, layout_t)
        master_t = find_rel_by_type(rels_path_for(layout_path), "slideMaster")
        if master_t:
            parts.append((os.path.join(pptdir, master_t), True))
        parts.append((layout_path, True))
    parts.append((slide_path, False))
    pre=""; slide_html=""; slide_root=None
    for path, skip_ph in parts:
        if not os.path.exists(path): continue
        root = etree.parse(path).getroot()
        spTree = root.find(f"{qp('cSld')}/{qp('spTree')}")
        ctx = Ctx(load_rels(rels_path_for(path)), mediadir, outassets)
        seg = render_bg(root, ctx) + render_part(spTree, ctx, TOP_TF, skip_ph)
        if skip_ph:
            pre += seg
        else:
            slide_html += seg; slide_root = root
    return pre, slide_html, slide_root

def slide_body(pptdir, slide_no, outassets):
    pre, slide_html, _ = _render_parts(pptdir, slide_no, outassets)
    return pre + slide_html

def slide_body_built(pptdir, slide_no, outassets):
    """Like slide_body but injects data-build/data-anim into the SLIDE part only.
    Returns (html, nbuilds)."""
    pre, slide_html, slide_root = _render_parts(pptdir, slide_no, outassets)
    nb=0
    if slide_root is not None:
        builds, nb = parse_builds(slide_root)
        for sid,(stp,etype) in builds.items():
            slide_html = slide_html.replace(
                f'data-spid="{sid}"',
                f'data-spid="{sid}" data-build="{stp}" data-anim="{etype}"', 1)
    return pre + slide_html, nb

def convert_slide(pptdir, slide_no, outdir):
    outassets = os.path.join(outdir,"assets")
    body = slide_body(pptdir, slide_no, outassets)
    W,H = 1280,720
    htmldoc = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
@font-face{{font-family:'Malgun Gothic';src:local('Malgun Gothic');}}
*{{margin:0;padding:0;}}
html,body{{background:#333;}}
.slide{{position:relative;width:{W}px;height:{H}px;background:#fff;overflow:hidden;
 font-family:'Arial','Malgun Gothic';}}
</style></head><body>
<div class="slide">{body}</div>
</body></html>"""
    out_html = os.path.join(outdir, f"slide{slide_no}.html")
    open(out_html,"w",encoding="utf-8").write(htmldoc)
    return out_html

if __name__=="__main__":
    pptdir = sys.argv[1]      # /tmp/ppt_full/ppt
    outdir = sys.argv[2]
    theme = os.path.join(pptdir,"theme","theme1.xml")
    load_theme(theme)
    for n in sys.argv[3:]:
        p = convert_slide(pptdir, int(n), outdir)
        print("wrote", p)
