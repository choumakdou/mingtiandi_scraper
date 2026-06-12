# Mingtiandi Bobby Mak Quote PDF Builder (v3 — GUI edition)

A single tool that turns a list of Mingtiandi article URLs into a single
A4-portrait PDF — one page per article, with the Mingtiandi masthead
preserved and the paragraph where **Bobby Mak (CHFT Advisory)** is
quoted highlighted in yellow.

## What's new in v3

- **GUI mode (default)** — when you double-click the `.exe` (or run
  `python mingtiandi_scraper.py` with no args), you get a proper
  Tkinter window with:
  - File picker for the Excel
  - Folder picker for the output
  - Chrome path field with auto-detect
  - "Show the browser" checkbox
  - **Start** / **Stop** buttons
  - **Live log** streaming the progress
  - Progress bar + status line
  - "Open output folder" button at the end
- **CLI mode (unchanged)** — all previous flags still work
- **Stop button** actually cancels mid-article

## Quick start

### GUI (easiest)

1. Double-click `mingtiandi_scraper.exe` (Windows) or run `./mingtiandi_scraper` (macOS).
2. Click **Browse…** next to *Excel file* → pick `articles.xlsx`.
3. Click **Browse…** next to *Output folder* → pick where to save the PDF.
4. (Optional) tick **Show the browser** if you want to watch Chrome
   work, or click **Auto-detect** next to *Chrome path*.
5. Click **▶ Start**. Watch the log fill up, the bar advance, the
   status line tell you which article is being processed.
6. When done, click **📂 Open output folder** to grab the PDF.

### CLI (for cron / CI)

```bash
mingtiandi_scraper.exe --input articles.xlsx --output output.pdf
mingtiandi_scraper.exe --input articles.xlsx --output output.pdf \
    --chrome "C:\Program Files\Google\Chrome\Application\chrome.exe"
mingtiandi_scraper.exe --mode pdf-chop --pdf-input ./saved_pdfs/ --output out.pdf
```

## Excel format

| Column A | Column B |
| --- | --- |
| `<title>\n<YYYY/MM/DD>` | full article URL |

Row 1 is the header. Each row is one article. The bundled `articles.xlsx`
has all 19 Bobby Mak articles pre-filled.

## How it bypasses Cloudflare

- Uses your real Chrome (auto-detected) — not a stripped-down Chromium
- Sets a normal User-Agent, real viewport, real locale
- New `BrowserContext` per URL → fresh cookie jar → the 5-articles-per-month
  paywall meter is reset every time (this is the "anonymous mode" trick)
- Waits up to 120s for Cloudflare's Turnstile to clear

## How the cropping works

1. `page.pdf()` (Chrome's print engine, viewport-independent) saves the
   article to a raw PDF.
2. CSS injected before printing hides ads / sidebar / footer.
3. `pdfplumber` locates the **Bobby Mak** text by coordinates.
4. The page is rendered to a 200 DPI image and cropped to one A4
   portrait page, focused around the quote.
5. All 19 A4 images are stacked into the final PDF.

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
`mingtiandi_scraper-macos` artifacts. The current run is at:
`https://github.com/choumakdou/mingtiandi_scraper/actions`

## Files in this repo

```
mingtiandi_scraper/
├── mingtiandi_scraper.py     # the tool (single file, GUI + CLI + core)
├── requirements.txt          # pip deps (no GUI deps needed — tkinter is stdlib)
├── README.md                 # this file
├── articles.xlsx             # 19 articles
└── .github/workflows/build.yml
```
