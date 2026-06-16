#!/usr/bin/env python3
"""
build.py — keep the HTML viewer in sync with the PowerPoint file.

What it does:
  1. Converts presentation/Welcome_Inbet_Online.pptx -> PDF   (LibreOffice)
  2. Renders each page to slides/slide-NN.jpg                 (Poppler: pdftoppm)
  3. Optimises the images                                     (Pillow)
  4. Writes manifest.json + manifest.js (the slide list the viewer reads)

Run it whenever you change the .pptx:
    python build.py

All tools are free / open-source:
  - LibreOffice  (soffice)      https://www.libreoffice.org/
  - Poppler      (pdftoppm)     https://poppler.freedesktop.org/
  - Pillow       (pip install Pillow)
Optional, for nicer auto-titles:
  - python-pptx  (pip install python-pptx)
"""

import glob
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PPTX = os.path.join(HERE, "presentation", "Welcome_Inbet_Online.pptx")
SLIDES_DIR = os.path.join(HERE, "slides")
DPI = 150
JPEG_QUALITY = 88


def find_tool(*names):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def run(cmd):
    print("  $", " ".join(cmd))
    subprocess.run(cmd, check=True)


def existing_titles():
    """Preserve hand-edited titles from a previous manifest.json when possible."""
    path = os.path.join(HERE, "manifest.json")
    if not os.path.exists(path):
        return []
    try:
        data = json.load(open(path, encoding="utf-8"))
        return [s.get("title", "") for s in data.get("slides", [])]
    except Exception:
        return []


def titles_from_pptx(n_slides):
    """Best-effort: first text line of each slide. Falls back to 'Slide N'."""
    try:
        from pptx import Presentation
    except Exception:
        return ["Slide %d" % (i + 1) for i in range(n_slides)]
    titles = []
    prs = Presentation(PPTX)
    for i, slide in enumerate(prs.slides):
        title = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    line = txt.splitlines()[0].strip()
                    if 2 <= len(line) <= 42:
                        title = line
                        break
        titles.append(title or ("Slide %d" % (i + 1)))
    return titles


def extract_logo():
    """Best-effort: pull the wordmark logo out of the .pptx media into
    assets/inbet-logo.png. Never fatal — keeps the existing asset on failure."""
    import zipfile, io
    try:
        from PIL import Image
        os.makedirs(os.path.join(HERE, "assets"), exist_ok=True)
        best = None  # (area, name, bytes)
        with zipfile.ZipFile(PPTX) as z:
            for n in z.namelist():
                if not n.lower().startswith("ppt/media/") or not n.lower().endswith(".png"):
                    continue
                data = z.read(n)
                im = Image.open(io.BytesIO(data))
                w, h = im.size
                has_alpha = im.mode in ("RGBA", "LA") or "transparency" in im.info
                # wordmark heuristic: wide, not huge, with transparency
                if has_alpha and w > h * 1.8 and (w * h) < 250000:
                    area = w * h
                    if best is None or area < best[0]:
                        best = (area, n, data)
        if best:
            with open(os.path.join(HERE, "assets", "inbet-logo.png"), "wb") as f:
                f.write(best[2])
            print("     logo extracted from %s" % best[1])
        else:
            print("     no logo candidate found — keeping existing assets/inbet-logo.png")
    except Exception as e:
        print("     logo extraction skipped (%s) — keeping existing asset" % e)


def main():
    if not os.path.exists(PPTX):
        sys.exit("ERROR: %s not found." % PPTX)

    soffice = find_tool("soffice", "libreoffice")
    pdftoppm = find_tool("pdftoppm")
    if not soffice:
        sys.exit("ERROR: LibreOffice (soffice) not found on PATH.")
    if not pdftoppm:
        sys.exit("ERROR: Poppler (pdftoppm) not found on PATH.")
    try:
        from PIL import Image
    except Exception:
        sys.exit("ERROR: Pillow not installed.  Run: pip install Pillow")

    print("1/4  Converting PPTX -> PDF ...")
    run([soffice, "--headless", "--convert-to", "pdf", "--outdir", HERE, PPTX])
    pdf = os.path.join(HERE, "Welcome_Inbet_Online.pdf")

    print("2/4  Rendering PDF pages -> PNG ...")
    os.makedirs(SLIDES_DIR, exist_ok=True)
    for f in glob.glob(os.path.join(SLIDES_DIR, "slide-*")):
        os.remove(f)
    run([pdftoppm, "-png", "-r", str(DPI), pdf, os.path.join(SLIDES_DIR, "slide")])

    print("3/4  Optimising images -> JPEG ...")
    pngs = sorted(glob.glob(os.path.join(SLIDES_DIR, "slide-*.png")))
    for p in pngs:
        Image.open(p).convert("RGB").save(
            p[:-4] + ".jpg", "JPEG", quality=JPEG_QUALITY, optimize=True, progressive=True
        )
        os.remove(p)
    os.remove(pdf)
    jpgs = sorted(glob.glob(os.path.join(SLIDES_DIR, "slide-*.jpg")))
    n = len(jpgs)
    print("     %d slides rendered." % n)

    print("     Extracting logo ...")
    extract_logo()

    print("4/4  Writing manifest ...")
    prev = existing_titles()
    if len(prev) == n and all(prev):
        titles = prev                      # keep curated titles
        print("     reused titles from existing manifest.json")
    else:
        titles = titles_from_pptx(n)       # auto-generate
        print("     generated titles from the .pptx")

    slides = [
        {"file": "slides/%s" % os.path.basename(j), "title": titles[i]}
        for i, j in enumerate(jpgs)
    ]
    manifest = {
        "source": "presentation/Welcome_Inbet_Online.pptx",
        "generated": "build.py output — edit titles freely; re-run build.py to refresh images",
        "slides": slides,
    }
    with open(os.path.join(HERE, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    with open(os.path.join(HERE, "manifest.js"), "w", encoding="utf-8") as f:
        f.write("// Auto-generated by build.py — slide list the viewer reads.\n")
        f.write("// Loaded via <script src> so index.html works even when opened directly (file://).\n")
        f.write("window.INBET_SLIDES = " + json.dumps(slides, indent=2, ensure_ascii=False) + ";\n")

    print("\nDone. Open index.html (via Live Server / http) to view %d slides." % n)


if __name__ == "__main__":
    main()
