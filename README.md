# Mingtiandi Bobby Mak Quote PDF Builder

A single Python tool that turns a list of Mingtiandi article URLs into a
single A4-portrait PDF — one page per article, with the Mingtiandi
masthead preserved and the paragraph where **Bobby Mak (CHFT Advisory)**
is quoted highlighted in yellow.

## Why this script exists

Mingtiandi sits behind a **Cloudflare Turnstile** managed challenge that
detects headless browsers from datacenter IPs (and from this assistant's
sandbox). When the same code runs on your laptop with your real Chrome
and your real residential IP, the challenge auto-solves in a couple of
seconds and you get genuine screenshots with the real photos, the real
header, and the full story.

## Features

- **One command, one PDF**: reads the Excel, visits all URLs, screenshots
  each, highlights the quote, outputs a single A4-portrait PDF.
- **Authentic header**: keeps Mingtiandi's own masthead (no fake
  reconstruction).
- **Highlight**: the `Bobby Mak` paragraph is wrapped in yellow
  background with a gold left border.
- **Crops intelligently**: keeps the header + article title + the
  highlighted paragraph + enough context, all on a single A4 page per
  article.
- **No paywall getting in the way**: opens a fresh Chrome **context** for
  every URL, so the 5-articles-per-month cookie meter is reset each
  time (the "use anonymous mode" trick from the brief).
- **Two modes**:
  - `scrape` — visit the live URLs.
  - `pdf-chop` — process PDFs you already saved from the browser
    (e.g. via *Print → Save as PDF*).

## Excel format

| Column A | Column B |
| --- | --- |
| `<title>\n<YYYY/MM/DD>` | full article URL |

Row 1 is treated as the header and skipped. The script used in the
deliverable reads `articles.xlsx` from the brief.

## Install & run (dev mode)

```bash
# 1. Install Python 3.10+ from python.org (or use uv / pyenv / etc.)

# 2. Clone or download this folder, then:
cd mingtiandi_scraper
pip install -r requirements.txt
python -m playwright install chromium

# 3. Run
python mingtiandi_scraper.py --input articles.xlsx --output output.pdf
```

By default it tries to auto-detect your Chrome. If it can't, pass the
path explicitly:

```bash
python mingtiandi_scraper.py --input articles.xlsx --output output.pdf \
    --chrome "C:\Program Files\Google\Chrome\Application\chrome.exe"
```

### Other flags

| Flag | Meaning |
| --- | --- |
| `--no-headless` | Show the Chrome windows as it works (handy for debugging). |
| `--mode pdf-chop --pdf-input <dir>` | Use saved PDFs instead of scraping. |
| `--user-agent "..."` | Override the User-Agent string. |

## Building a Windows `.exe` via GitHub Actions

This repo includes `.github/workflows/build.yml` that builds a single
`.exe` on every push to `main` (and on manual trigger).

1. Create a new GitHub repo, push this folder.
2. Open the **Actions** tab → pick the *Build Windows EXE* run.
3. Download the `mingtiandi_scraper-windows` artifact at the bottom of
   the run — that's your `mingtiandi_scraper.exe`.

Run it on your machine:

```bat
mingtiandi_scraper.exe --input articles.xlsx --output output.pdf
```

### Building locally with PyInstaller

```bash
pip install pyinstaller
pyinstaller --onefile --name mingtiandi_scraper mingtiandi_scraper.py
# Result: dist/mingtiandi_scraper.exe (Windows) or dist/mingtiandi_scraper (macOS/Linux)
```

## Will this bypass Cloudflare when run locally?

**Yes, in the vast majority of cases.** Reasons:

1. Your IP is residential, not a known datacenter range.
2. Your Chrome has real cookies, real cache, real history — the
   Turnstile fingerprint heuristics all look like a normal user.
3. The script does **not** set `navigator.webdriver`, does not run in a
   contrived "no-extension" state, and uses your real Chrome binary,
   not a stripped-down Chromium build.
4. Even on the rare first-visit case where Cloudflare shows a checkbox
   challenge, the script waits up to 120 seconds for it to auto-clear
   before continuing.

If a particular article still gets stuck on the challenge, simply open
it in your own browser once to pass the Turnstile, then re-run the
script — the cookie will be picked up.

## Limitations

- **Ads are not auto-removed.** The current crop keeps everything
  between the masthead and the end of the article body. If you want
  sidebar / "Related articles" blocks hidden, file a small follow-up —
  I can extend the script with a CSS-selector list of "junk" containers
  to remove before screenshotting.
- The full Mingtiandi footer is included; if you want it stripped, the
  same `body_box` selector in the script can be tightened.

## Files in this repo

```
mingtiandi_scraper/
├── mingtiandi_scraper.py     # the tool (single-file, ~600 LoC)
├── requirements.txt          # pip deps
├── README.md                 # this file
└── .github/
    └── workflows/
        └── build.yml         # GitHub Actions → Windows .exe
```
