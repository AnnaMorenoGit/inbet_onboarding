# Welcome to Inbet — PPTX → HTML viewer

The PowerPoint `Welcome_Inbet_Online.pptx` presented as a web page, with every
slide as a clickable tab in a sidebar. Built with public, open-source tooling.

![slides: 23](https://img.shields.io/badge/slides-23-1f3bd4) ![license](https://img.shields.io/badge/libs-MIT%2FBSD-blue)

---

## Two viewers (open either one)

| File | What it is | Fidelity | Needs a server? |
|------|-----------|----------|-----------------|
| **`index.html`** | **Recommended.** Each slide rendered as an image, shown in a tabbed UI. | Pixel-perfect — identical to PowerPoint. | No (also works on file://). |
| `live.html` | Renders the `.pptx` **at runtime** in the browser via [PPTXjs](https://github.com/meshesha/PPTXjs). | Good for text & tables; **SmartArt** (the circular team diagrams) renders imperfectly. | Yes — must be served over http. |

Both read from the same source `.pptx`. `index.html` is the one to present from;
`live.html` is included because it literally re-reads the file every time it loads.

---

## How "syncing from the .pptx" works

- **`live.html`** parses `presentation/Welcome_Inbet_Online.pptx` in the browser
  on every load, so it always reflects the current file — no build step.
- **`index.html`** shows pre-rendered slide images. To refresh them after editing
  the deck, run the one-line build:

  ```bash
  python build.py
  ```

  That regenerates `slides/*.jpg` and `manifest.js` straight from the `.pptx`.
  On GitHub, the included Action does this **automatically on every push** — so
  committing a new `.pptx` re-syncs the published site with no manual step.

---

## Run it locally (VS Code)

The image viewer (`index.html`) works by just double-clicking it. To use the live
viewer or to mirror the hosted setup, serve the folder over http:

- **Live Server extension** — right-click `index.html` → *Open with Live Server*.
- **Terminal** —
  ```bash
  python -m http.server 8000
  # then open http://localhost:8000
  ```

> Browsers block reading local files over `file://`, which is why `live.html`
> (and `build.py`-free fetching) needs a small local server.

---

## Publish on GitHub Pages

1. Push this folder to a GitHub repo (default branch `main`).
2. **Settings → Pages → Build and deployment → Source: GitHub Actions.**
3. The included workflow (`.github/workflows/deploy.yml`) installs LibreOffice +
   Poppler, runs `build.py`, and deploys. Your site appears at
   `https://<user>.github.io/<repo>/`.

To re-sync after editing the deck: replace `presentation/Welcome_Inbet_Online.pptx`,
commit, and push. The Action rebuilds and redeploys.

---

## Project structure

```
inbet-presentation/
├── index.html            # image-based tabbed viewer (recommended)
├── live.html             # PPTXjs live viewer (reads the .pptx at runtime)
├── manifest.js           # slide list the viewer reads (generated)
├── manifest.json         # human-readable copy of the slide list
├── build.py              # regenerates slides + manifest from the .pptx
├── presentation/
│   └── Welcome_Inbet_Online.pptx
├── slides/               # slide-01.jpg … slide-23.jpg (generated)
├── assets/
│   └── inbet-logo.png    # exact Inbet wordmark, extracted from the .pptx
├── vendor/               # PPTXjs + jQuery/JSZip/D3 etc. (for live.html)
└── .github/workflows/
    └── deploy.yml        # auto build + deploy to GitHub Pages
```

## Editing tab titles

Tab labels live in `manifest.json` (`title` of each slide). Edit them freely —
`build.py` preserves your titles on rebuild as long as the slide count is unchanged.

## Credits / licences

- [PPTXjs](https://github.com/meshesha/PPTXjs) — MIT · jQuery — MIT · JSZip — MIT ·
  D3 — BSD · NVD3 — Apache-2.0. Vendored under `vendor/` with their headers intact.
