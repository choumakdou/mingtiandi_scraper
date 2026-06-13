# Mingtiandi Bobby Mak Quote PDF Builder (v4.2 — Header + Title + Scroll Cue)

A single tool that turns a list of Mingtiandi article URLs into a single
A4-portrait PDF — one page per article, with the Mingtiandi masthead
preserved and the paragraph where **Bobby Mak (CHFT Advisory)** is
quoted highlighted in yellow.

## What's new in v4.2

- **Story subject + publish date rendered at the top of every A4 page.**
  The tool now reads the title (and its continuation line) and the
  byline from page 1 of the PDF and renders them as bold text above
  the article body, so you always see "Hong Kong's Chinachem… — 2026/06/10 by Iris Hong" before the quote. Title is word-wrapped automatically.
- **Visible scroll bar on the right edge** of the article window. A
  thin gray track with a darker thumb whose position and height
  reflect where the visible window sits within the full article. So
  the reader immediately sees "this is a 1-page slice of a longer
  story".
- **Pale border + subtle drop shadow** around the article window,
  reinforcing the "scrolled view" metaphor.
- **Header layout is now user-editable.** New **Customize…** button
  in the Inputs section opens a dialog with a live A4 preview,
  X/Y/W/H spinboxes, a "lock aspect" checkbox, three presets
  (Full width / Left half / Right half), an Import button, and
  Apply/Cancel. Layout is saved on the GUI and threaded through to
  `make_a4_page_from_pdf`, the auto-scrape path, and the CLI
  (`--header-x / --header-y / --header-w / --header-h`).
- **Softer highlight on the Bobby Mak paragraph.** Background changed
  from `#fff59d` @ 86% opacity to `#fffacd` (lemon chiffon) @ 45%
  opacity, with a thicker 8-px gold left border. The quoted text
  stays clearly readable.
- **Fixed `poppler` not found error.** The script used `pdf2image`,
  which shells out to the `poppler-utils` system package — not
  installed by default on a fresh Windows box. Swapped to
  `PyMuPDF` (bundled C library, no system dep). The .exe now works
  out of the box.
- `pdf2image` removed from `requirements.txt`; `pymupdf` added.

## What's new in v4.1 (kept)

## What's new in v4 (kept)

- **Three modes** (radio buttons at the top of the GUI):
  - **🤖 Auto scrape** — headless Chrome tries to load every article
    (works when Cloudflare isn't blocking your IP)
  - **🖐 Guided manual** *(default)* — your default browser opens to
    each article in turn, you do `Ctrl+P` → "Save as PDF", the tool
    watches the folder and **auto-advances** when it sees the new file,
    then crops + combines everything at the end
  - **📄 Combine saved PDFs** — you already have the article PDFs,
    just stitch them into one
- **Re-combine button** — re-runs the chop pipeline on the most recent
  guided-run folder (or any folder you pick)
- **PDFs folder field** — only enabled in Combine mode
- **Live log** streams every step

## v4 is for you if Cloudflare is blocking the headless scraper

If Auto mode keeps saying "Cloudflare never cleared" because your home
IP got a low Trust Score, switch to **Guided manual** (already the
default). It bypasses automation detection entirely because *you* open
each article in your own Chrome — the script just helps you do the
filing and the cropping.

## Quick start (v4, recommended)

1. **Double-click `mingtiandi_scraper.exe`**.
2. Make sure **🖐 Guided manual** is selected (it is by default).
3. **Browse…** next to *Output folder* → pick where to save the final
   PDF. A subfolder called `bobby_pdfs\` will be created there for the
   19 intermediate files.
4. **Browse…** next to *Excel file* → pick `articles.xlsx`.
5. Click **▶ Start**. A dialog opens, your default browser pops up
   with article 1. The dialog tells you:
   - the exact filename to use (`article_01.pdf`)
   - the exact folder to save it in (with one-click Copy buttons)
   - the 7 steps to follow in your browser
6. **In Chrome**: solve any Cloudflare challenge → `Ctrl+P` → "Save as
   PDF" → save with the exact filename in the exact folder.
7. The dialog **auto-advances** the moment it sees the new file. Repeat
   for the remaining 18.
8. When all 19 are saved, the tool **automatically** runs the
   crop-highlight-combine pipeline. Watch the log and progress bar.
9. **📂 Open output folder** appears when the final PDF is ready.

If you make a mistake, use:
- **⏭ Skip** — move on without a PDF for this article
- **🔄 Reopen URL** — re-open the article in your browser
- **📁 I saved it under a different name** — pick the file you saved
  and the tool will rename + place it
- **⏹ Stop** — finish early, the tool will still combine whatever was
  saved so far

## Quick start (Combine mode)

If you already have the 19 PDFs saved somewhere:

1. Pick **📄 Combine saved PDFs** in the Mode section.
2. The **PDFs folder** field is now enabled. **Browse…** to the folder
   containing the PDFs.
3. Pick the **Output folder** and **Output filename** for the final PDF.
4. **▶ Start**. The tool runs the chop pipeline and reports progress
   in the log.

CLI equivalent:
```bash
mingtiandi_scraper.exe --mode pdf-chop --pdf-input "C:\path\to\pdfs" --output out.pdf
```

## How the cropping works (unchanged from v3)

1. `page.pdf()` (Chrome's print engine, viewport-independent) saves the
   article to a raw PDF.
2. CSS injected before printing hides ads / sidebar / footer.
3. `pdfplumber` locates the **Bobby Mak** text by coordinates.
4. The page is rendered to a 200 DPI image and cropped to one A4
   portrait page, focused around the quote.
5. All 19 A4 images are stacked into the final PDF.

## How the paywall bypass works (unchanged from v3)

- Uses your real Chrome (auto-detected) — not a stripped-down Chromium.
- Sets a normal User-Agent, real viewport, real locale.
- New `BrowserContext` per URL → fresh cookie jar → the 5-articles-per-month
  paywall meter is reset every time.

## Excel format

| Column A | Column B |
| --- | --- |
| `<title>\n<YYYY/MM/DD>` | full article URL |

Row 1 is the header. Each row is one article. The bundled `articles.xlsx`
has all 19 Bobby Mak articles pre-filled.

## Build the `.exe` yourself (no GitHub)

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --name mingtiandi_scraper mingtiandi_scraper.py
# dist/mingtiandi_scraper.exe
```

## Build via GitHub Actions

This repo includes `.github/workflows/build.yml`. Push to `main` and
the Actions tab will produce `mingtiandi_scraper-windows` and
`mingtiandi_scraper-macos` artifacts.

## Files in this repo

```
mingtiandi_scraper/
├── mingtiandi_scraper.py     # the tool (single file, GUI + CLI + core + GuidedManualDialog)
├── requirements.txt          # pip deps (no GUI deps needed — tkinter is stdlib)
├── README.md                 # this file
├── articles.xlsx             # 19 articles
└── .github/workflows/build.yml
```
