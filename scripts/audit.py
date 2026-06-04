#!/usr/bin/env python3
"""Render a generated SNU deck to per-slide PNGs for subagent verification.

Produces SAFE-SIZED images (<=1280px) so a subagent can read several at once
without hitting the many-image pixel cap. The MAIN agent must NOT read these
PNGs (context saturation) — it dispatches subagents that read them and return
text-only reports.

Renders:
  - HTML  : each slide via the deck's show(N), Chrome headless, 1280x720
  - PPTX  : soffice -> pdf -> pdftoppm @96dpi  (~1280px), if a .pptx is given

Usage:
  python audit.py deck.html [deck.pptx] [--out DIR]
Prints a manifest: paired html_NN.png / ppt_NN.png paths for the subagent prompt.
"""
import os, sys, re, subprocess, argparse, glob

CH = os.path.expanduser("~/Library/Caches/ms-playwright/chromium_headless_shell-1223/"
                        "chrome-headless-shell-mac-arm64/chrome-headless-shell")
SOFFICE = "/opt/homebrew/bin/soffice"
PDFTOPPM = "/opt/homebrew/bin/pdftoppm"

def n_slides(html):
    return len(re.findall(r'<section class="slide', html))

def render_html(html_path, outdir):
    src = open(html_path, encoding="utf-8").read()
    n = n_slides(src)
    out = []
    for i in range(n):
        # freeze on slide i; kill transitions so the screenshot is settled
        inj = (f"<style>*{{transition:none !important}}</style>"
               f"<script>addEventListener('DOMContentLoaded',()=>show({i}))</script>")
        variant = src.replace("</body>", inj + "</body>")
        vp = os.path.join(outdir, f"_html_{i+1}.html")
        open(vp, "w", encoding="utf-8").write(variant)
        png = os.path.join(outdir, f"html_{i+1:02d}.png")
        subprocess.run([CH, "--headless", "--disable-gpu", "--hide-scrollbars",
                        "--force-device-scale-factor=1", "--window-size=1280,720",
                        f"--screenshot={png}", f"file://{vp}"],
                       stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
        out.append(png)
    return out

def render_pptx(pptx_path, outdir):
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx_path],
                   stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    if not os.path.exists(pdf): return []
    subprocess.run([PDFTOPPM, "-r", "96", "-png", pdf, os.path.join(outdir, "ppt")],
                   stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
    # pdftoppm names p-1.png ... rename to ppt_NN.png zero-padded
    pages = sorted(glob.glob(os.path.join(outdir, "ppt-*.png")))
    out = []
    for p in pages:
        m = re.search(r"ppt-(\d+)\.png$", p)
        if m:
            dst = os.path.join(outdir, f"ppt_{int(m.group(1)):02d}.png")
            os.replace(p, dst); out.append(dst)
    return out

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("html"); ap.add_argument("pptx", nargs="?")
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    outdir = a.out or os.path.join(os.path.dirname(os.path.abspath(a.html)), "_audit")
    os.makedirs(outdir, exist_ok=True)
    htmls = render_html(a.html, outdir)
    ppts = render_pptx(a.pptx, outdir) if a.pptx else []
    print(f"slides: {len(htmls)}   outdir: {outdir}")
    print("--- manifest (for subagent) ---")
    for i, h in enumerate(htmls):
        line = f"slide {i+1}: HTML={h}"
        if i < len(ppts): line += f"  PPT={ppts[i]}"
        print(line)
    if a.pptx and len(ppts) != len(htmls):
        print(f"WARNING: html slides={len(htmls)} pptx pages={len(ppts)} (mismatch)")
