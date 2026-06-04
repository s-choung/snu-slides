#!/usr/bin/env python3
"""Assemble 52 slides into one navigable, font-embedded HTML deck."""
import os, sys, glob, re
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import convert
from fontTools.subset import Subsetter, Options
from fontTools.ttLib import TTFont

A="http://schemas.openxmlformats.org/drawingml/2006/main"
PPTDIR=sys.argv[1]; OUTDIR=sys.argv[2]; N=int(sys.argv[3]) if len(sys.argv)>3 else 52
FONTSRC="/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"

def gather_charset(pptdir, n):
    chars=set()
    for i in range(1,n+1):
        f=os.path.join(pptdir,"slides",f"slide{i}.xml")
        if not os.path.exists(f): continue
        x=open(f,encoding="utf-8").read()
        for t in re.findall(r"<a:t>(.*?)</a:t>", x, re.S):
            chars.update(t)
    # always include basic latin + common punctuation
    chars.update(chr(c) for c in range(0x20,0x7f))
    return chars

def subset_font(src, dst, chars):
    opt=Options()
    opt.flavor="woff2"; opt.desubroutinize=True
    opt.layout_features=["*"]; opt.name_IDs=["*"]; opt.notdef_outline=True
    opt.recalc_bounds=True; opt.drop_tables=[]
    f=TTFont(src)
    ss=Subsetter(options=opt)
    ss.populate(unicodes=[ord(c) for c in chars])
    ss.subset(f)
    f.save(dst)

def fontface_css(fdir):
    def rel(p): return "fonts/"+os.path.basename(p)
    faces=[]
    pairs=[("Arial","arial.woff2",400),("Arial","arialbd.woff2",700),
           ("Malgun Gothic","malgun.woff2",400),("Malgun Gothic","malgunbd.woff2",700)]
    for fam,fn,w in pairs:
        if os.path.exists(os.path.join(fdir,fn)):
            faces.append(f"@font-face{{font-family:'{fam}';font-weight:{w};"
                         f"font-style:normal;src:url('fonts/{fn}') format('woff2');}}")
    return "\n".join(faces)

def main():
    os.makedirs(OUTDIR, exist_ok=True)
    assets=os.path.join(OUTDIR,"assets"); os.makedirs(assets, exist_ok=True)
    fdir=os.path.join(OUTDIR,"fonts"); os.makedirs(fdir, exist_ok=True)
    convert.load_theme(os.path.join(PPTDIR,"theme","theme1.xml"))
    # fonts
    chars=gather_charset(PPTDIR,N)
    print(f"charset {len(chars)} chars; subsetting fonts...")
    for src,dst in [("arial.ttf","arial.woff2"),("arialbd.ttf","arialbd.woff2"),
                    ("malgun.ttf","malgun.woff2"),("malgunbd.ttf","malgunbd.woff2")]:
        s=os.path.join(FONTSRC,src)
        if os.path.exists(s):
            subset_font(s, os.path.join(fdir,dst), chars)
            print("  ",dst, os.path.getsize(os.path.join(fdir,dst))//1024,"KB")
    # slides
    sections=[]; total_builds=0
    for n in range(1,N+1):
        nb=0
        try:
            body,nb=convert.slide_body_built(PPTDIR,n,assets)
            total_builds+=nb
        except Exception as e:
            body=f'<div style="position:absolute;inset:0;color:red">ERR {e}</div>'
            print("ERR slide",n,e)
        sections.append(f'<section class="slide" data-n="{n}" data-builds="{nb}">{body}</section>')
    print("total build steps:", total_builds)
    deck="\n".join(sections)
    css=fontface_css(fdir)
    html=f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<title>Deck</title>
<style>
{css}
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{height:100%;background:#000;overflow:hidden;font-family:'Arial','Malgun Gothic'}}
#stage{{position:absolute;left:50%;top:50%;width:1280px;height:720px;
 transform:translate(-50%,-50%) scale(var(--s,1));transform-origin:center}}
.slide{{position:absolute;inset:0;width:1280px;height:720px;background:#fff;
 overflow:hidden;display:none}}
.slide.active{{display:block}}
#hud{{position:fixed;right:14px;bottom:10px;color:#999;font:12px Arial;
 background:rgba(0,0,0,.4);padding:3px 8px;border-radius:10px;z-index:99}}
</style></head><body>
<div id="stage">{deck}</div>
<div id="hud"><span id="cur">1</span> / {N}  ·  ← →  ·  F</div>
<script>
const slides=[...document.querySelectorAll('.slide')];
let i=0,b=0;
function fit(){{const s=Math.min(innerWidth/1280,innerHeight/720);
 document.getElementById('stage').style.setProperty('--s',s);}}
function maxB(){{return +slides[i].getAttribute('data-builds')||0;}}
function applyBuild(sl,bb){{sl.querySelectorAll('[data-build]').forEach(el=>{{
 const k=+el.getAttribute('data-build'); const t=el.getAttribute('data-anim');
 el.style.transition='opacity .35s ease';
 if(t==='entr') el.style.opacity=(k<=bb)?'1':'0';
 else if(t==='exit') el.style.opacity=(bb>=k)?'0':'1';
}});}}
function playVids(sl,only){{sl.querySelectorAll('video').forEach(v=>{{
 const h=v.closest('[data-build]'); const k=h?+h.getAttribute('data-build'):0;
 if(only===null?(k<=b):(k===only)){{v.currentTime=0;v.play().catch(()=>{{}});}}}});}}
function hud(){{document.getElementById('cur').textContent=(i+1)+(maxB()?(' · '+b+'/'+maxB()):'');}}
function show(n){{i=Math.max(0,Math.min(slides.length-1,n));b=0;
 slides.forEach((s,k)=>s.classList.toggle('active',k===i));
 const sl=slides[i]; sl.querySelectorAll('video').forEach(v=>v.pause());
 applyBuild(sl,0); playVids(sl,null); hud();}}
function next(){{if(b<maxB()){{b++;applyBuild(slides[i],b);playVids(slides[i],b);hud();}}else show(i+1);}}
function prev(){{if(b>0){{b--;applyBuild(slides[i],b);hud();}}
 else if(i>0){{show(i-1);b=maxB();applyBuild(slides[i],b);playVids(slides[i],null);hud();}}}}
addEventListener('keydown',e=>{{
 if(e.key==='ArrowRight'||e.key===' '||e.key==='PageDown'){{next();e.preventDefault();}}
 else if(e.key==='ArrowLeft'||e.key==='PageUp'){{prev();}}
 else if(e.key==='Home'){{show(0);}} else if(e.key==='End'){{show(slides.length-1);}}
 else if(e.key.toLowerCase()==='f'){{if(!document.fullscreenElement)document.documentElement.requestFullscreen();else document.exitFullscreen();}}
}});
addEventListener('resize',fit); addEventListener('click',e=>{{if(e.clientX>innerWidth*0.5)next();else prev();}});
fit(); show(0);
</script></body></html>"""
    out=os.path.join(OUTDIR,"deck.html")
    open(out,"w",encoding="utf-8").write(html)
    print("wrote", out, os.path.getsize(out)//1024,"KB")

if __name__=="__main__": main()
