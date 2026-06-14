#!/usr/bin/env python3
"""
Mingtiandi Bobby Mak Quote PDF Builder (v4.4 — title section bridges into article body)
================================================================================

Opens each Mingtiandi article in your real Chrome, exports the page to
PDF via Chrome's own print engine, then post-processes each PDF to
remove ads/sidebar/footer, find the paragraph that quotes Bobby Mak,
and crop the article to a single A4 portrait page.

MODES
-----
- **GUI (default)**: when launched with no arguments, shows a Tkinter
  window with Excel/folder pickers, a live log, progress bar and
  start/stop buttons.
- **CLI**: when launched with `--input`/`--output` etc., runs headless
  as before — useful for cron jobs and CI.

USAGE
-----
    # GUI (no args)
    mingtiandi_scraper.exe
    python mingtiandi_scraper.py

    # CLI
    mingtiandi_scraper.exe --input articles.xlsx --output output.pdf
    python mingtiandi_scraper.py --input articles.xlsx --output output.pdf

GUI WORKFLOW
------------
1. Click "Browse…" next to *Excel file* and pick the articles.xlsx
2. Click "Browse…" next to *Output folder* and pick where to save the PDF
3. (Optional) tick "Watch the browser" if you want Chrome to open
   visibly so you can see what it's doing
4. Click "Start" — the log fills with progress, the progress bar
   advances, the status line shows which article is being processed
5. When done, click "Open output folder" to see the result

The 5-articles-per-month Mingtiandi paywall is bypassed automatically
because every article gets a fresh Chrome context (= fresh cookie jar).

WHY page.pdf() INSTEAD OF SCREENSHOTS
-------------------------------------
`page.pdf()` uses Chrome's print engine, so the output is the same
regardless of window size. It also respects the page's @media print
stylesheet, so ads/sidebars often disappear on their own.

REQUIREMENTS
------------
    pip install -r requirements.txt
    python -m playwright install chromium
"""

from __future__ import annotations

import argparse
import asyncio
import io
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path
from typing import Optional

# GUI imports — kept lazy so the CLI still works on headless boxes
# without a display server.
def _import_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    return tk, ttk, filedialog, messagebox, scrolledtext

import openpyxl
import pdfplumber
from PIL import Image, ImageFont
from playwright.async_api import (
    Browser,
    Page,
    Playwright,
    async_playwright,
)

# A4 portrait in pixels at 150 DPI (for the final composite)
A4_W_PX = 1240
A4_H_PX = 1754

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

# CSS injected before printing to strip ads/footers/etc.
JUNK_CSS = """
aside, .sidebar, .related, .related-posts, .you-may-also-like,
.ad, .ads, .advertisement, [class*="advert"], [id*="advert"],
.footer, .site-footer, .post-footer, .article-footer,
.newsletter, .newsletter-signup, .subscribe,
.social-share, .share-buttons, .social,
.comments, #comments, .related-articles, .popular-posts,
.cookie-notice, .gdpr, .paywall, .meter, .paywall-meter,
.promo, .promoted, .sponsored, [class*="sponsor"],
iframe[src*="doubleclick"], iframe[src*="googletag"],
iframe[src*="ads"], iframe[id*="google_ads"],
#sidebar, #footer, #comments {
    display: none !important;
}
"""

HIGHLIGHT_CSS = """
.mak-quote-highlight {
    background-color: #fff59d !important;
    padding: 10px 14px !important;
    border-left: 3px solid #f6c026 !important;
    margin: 14px 0 !important;
    box-shadow: 0 0 0 6px #fff59d !important;
}
"""


# ===========================================================================
# Pure logic — shared by CLI and GUI
# ===========================================================================


def read_articles(excel_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    articles: list[dict] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None or row[1] is None:
            continue
        title_cell = str(row[0])
        url = str(row[1]).strip()
        if not url.lower().startswith(("http://", "https://")):
            continue
        if "\n" in title_cell:
            title, date = title_cell.split("\n", 1)
            title, date = title.strip(), date.strip()
        else:
            title, date = title_cell.strip(), ""
        articles.append({"title": title, "date": date, "url": url})
    return articles


def find_chrome_path(custom: Optional[str] = None) -> Optional[str]:
    if custom and os.path.exists(custom):
        return custom
    candidates: list[str] = []
    system = platform.system()
    if system == "Windows":
        candidates += [
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        ]
    elif system == "Darwin":
        candidates += [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ]
    else:
        candidates += [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium",
            shutil.which("google-chrome"),
            shutil.which("google-chrome-stable"),
            shutil.which("chromium"),
        ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


async def launch_browser(
    p: Playwright, chrome_path: Optional[str], headless: bool = True
) -> Browser:
    if chrome_path:
        try:
            browser = await p.chromium.launch(
                headless=headless,
                executable_path=chrome_path,
                channel="chrome" if "chrome" in chrome_path.lower() else None,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                ],
            )
            return browser
        except Exception as e:
            print(f"[!] Failed to launch {chrome_path}: {e}", file=sys.stderr)
    browser = await p.chromium.launch(
        headless=headless,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-setuid-sandbox",
        ],
    )
    return browser


async def wait_for_real_article(page: Page, timeout: int = 120) -> bool:
    deadline = asyncio.get_event_loop().time() + timeout
    last_title = ""
    while asyncio.get_event_loop().time() < deadline:
        try:
            title = (await page.title()) or ""
            if title != last_title:
                print(f"    title: {title[:60]}", flush=True)
                last_title = title
            if "Just a moment" not in title and "Verifying" not in title:
                ready = await page.evaluate(
                    """() => {
                        const t = (document.body && document.body.innerText) || '';
                        if (t.length < 800) return false;
                        if (t.includes('Bobby Mak')) return 'bobby';
                        if (t.length > 2000 && /\\d{4}\\/\\d{2}\\/\\d{2}/.test(t)) return 'likely';
                        return false;
                    }"""
                )
                if ready:
                    return True
        except Exception:
            pass
        await asyncio.sleep(2)
    return False


async def inject_print_prep(page: Page) -> bool:
    await page.add_style_tag(content=HIGHLIGHT_CSS)
    await page.add_style_tag(content=JUNK_CSS)
    return await page.evaluate(
        """() => {
            const candidates = Array.from(document.querySelectorAll('p, blockquote, li'));
            let best = null;
            let bestScore = 0;
            for (const el of candidates) {
                const txt = (el.innerText || '').trim();
                if (!txt.includes('Bobby Mak')) continue;
                if (txt.length < 20 || txt.length > 4000) continue;
                let p = el;
                let depth = 0;
                while (p && depth < 5) {
                    const tag = p.tagName.toLowerCase();
                    if (['nav','header','footer','aside'].includes(tag)) break;
                    p = p.parentElement;
                    depth++;
                }
                if (el.children.length > 3) continue;
                let score = 100;
                if (/CHFT|valuer|told Mingtiandi|valuing/i.test(txt)) score += 200;
                score -= Math.abs(300 - txt.length);
                if (score > bestScore) { bestScore = score; best = el; }
            }
            if (best) {
                best.classList.add('mak-quote-highlight');
                best.scrollIntoView({ block: 'center' });
                return true;
            }
            return false;
        }"""
    )


async def print_article_to_pdf(
    browser: Browser,
    art: dict,
    index: int,
    total: int,
    out_dir: str,
    user_agent: str,
    progress_cb=None,  # callable(float 0..1) called on progress
) -> Optional[str]:
    """Returns path to raw PDF, or None on failure."""
    print(f"[{index}/{total}] {art['title']}")
    print(f"    url: {art['url']}", flush=True)
    context = await browser.new_context(
        viewport={"width": 1366, "height": 768},
        user_agent=user_agent,
        locale="en-US",
        timezone_id="Asia/Hong_Kong",
    )
    page = await context.new_page()
    raw_pdf = os.path.join(out_dir, f"article_{index:02d}_raw.pdf")
    try:
        await page.goto(art["url"], wait_until="domcontentloaded", timeout=60000)
        print("    waiting for Cloudflare to clear…", end="", flush=True)
        if not await wait_for_real_article(page, timeout=120):
            print(" ✗ (timed out)")
            print("    [!] Cloudflare never cleared. Skipping.", flush=True)
            return None
        print(" ✓")
        ok = await inject_print_prep(page)
        if not ok:
            print("    [!] No 'Bobby Mak' paragraph found in DOM.")
        await page.wait_for_timeout(800)
        await page.pdf(
            path=raw_pdf,
            format="A4",
            print_background=True,
            margin={"top": "10mm", "right": "10mm", "bottom": "10mm", "left": "10mm"},
            prefer_css_page_size=True,
        )
        print(f"    raw PDF: {raw_pdf}", flush=True)
        if progress_cb:
            progress_cb(index / total)
        return raw_pdf
    except Exception as e:
        print(f"    [!] Error: {e}", flush=True)
        return None
    finally:
        await context.close()


def find_bobby_mak_in_pdf(pdf_path: str) -> Optional[tuple[int, float]]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                for line in page.extract_text_lines():
                    txt = (line.get("text") or "").strip()
                    if "Bobby Mak" in txt and len(txt) < 1200:
                        return page_idx, float(line.get("top", 0))
    except Exception as e:
        print(f"    [!] pdfplumber failed: {e}", flush=True)
    return None


# Path to the default Mingtiandi header image (lives next to the script)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_HEADER_PNG = os.path.join(SCRIPT_DIR, "default_header.png")

# Default header layout (used by the CLI when --header-layout isn't given)
DEFAULT_HEADER_LAYOUT = {
    "x": 0,
    "y": 0,
    "w": A4_W_PX,
    "h": None,           # auto from width + aspect
    "lock_aspect": True,
}


def pdf_to_images(pdf_path: str, dpi: int = 200) -> list[Image.Image]:
    """Render every page of a PDF to a PIL Image at the given DPI.

    Uses PyMuPDF (bundled C library — no poppler or other system
    dependency required, so the .exe works on a fresh Windows box
    without installing poppler-utils).
    """
    import pymupdf
    out: list[Image.Image] = []
    with pymupdf.open(pdf_path) as doc:
        zoom = dpi / 72.0
        mat = pymupdf.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    return out


def _find_title_and_date(pdf_path: str) -> tuple[Optional[str], Optional[str]]:
    """Look at the first page of the PDF for the article title (which
    may span 2 lines) and the publish-date / byline line that follows
    it. Returns (title, date) strings (either may be None if not found).
    """
    NAV_TOKENS = {
        "SUBSCRIBE", "LOGIN", "CAPITAL MARKETS", "EVENTS", "MTD TV",
        "PEOPLE", "LOGISTICS", "DATA CENTRES", "ASIA OUTBOUND",
        "RETAIL", "RESEARCH & POLICY", "ADVERTISE",
    }
    try:
        with pdfplumber.open(pdf_path) as ppdf:
            if not ppdf.pages:
                return None, None
            page = ppdf.pages[0]
            lines = page.extract_text_lines() or []
            title = None
            date = None
            for i, ln in enumerate(lines):
                top = float(ln.get("top", 0))
                text = (ln.get("text") or "").strip()
                if top < 250:
                    continue
                if not text or len(text) < 5:
                    continue
                if any(tok in text.upper() for tok in NAV_TOKENS) and len(text) < 60:
                    continue
                # Found the first line of the title; collect the next
                # 1-3 lines too (the title is often 2 lines) until we
                # hit the date / byline.
                title_parts = [text]
                for ln2 in lines[i + 1: i + 5]:
                    t2 = (ln2.get("text") or "").strip()
                    if re.search(r"\d{4}/\d{2}/\d{2}|BY\s+\w+", t2, re.IGNORECASE):
                        date = t2
                        break
                    # Heuristic: if this next line is short, uppercase-
                    # looking (same as title), and not a sentence, treat
                    # it as a continuation of the title.
                    if t2 and len(t2) < 70 and (t2.upper() == t2 or t2.istitle()):
                        title_parts.append(t2)
                    else:
                        break
                title = " ".join(title_parts)
                break
            return title, date
    except Exception:
        return None, None


def _load_font(size: int, weight: str = "regular"):
    """Load a sans-serif TTF font that matches Mingtiandi's look.

    weight: 'regular', 'bold', or 'black'. Mingtiandi's title is
    rendered in a heavy sans-serif, so the title uses 'black' (Arial
    Black / Segoe UI Black) when available. The byline and intro
    use 'regular' (Arial / Helvetica / Liberation Sans).
    """
    candidates: list[str] = []
    if weight in ("black", "bold"):
        # Heaviest weights first; fall back to bold variants.
        candidates += [
            "C:\\Windows\\Fonts\\ariblk.ttf",          # Arial Black (Win)
            "C:\\Windows\\Fonts\\seguiblk.ttf",        # Segoe UI Black (Win)
            "C:\\Windows\\Fonts\\arialbd.ttf",         # Arial Bold (Win)
            "C:\\Windows\\Fonts\\segoeuib.ttf",        # Segoe UI Bold (Win)
            "/System/Library/Fonts/Helvetica.ttc",     # macOS Helvetica
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
        ]
    if weight == "regular":
        candidates += [
            "C:\\Windows\\Fonts\\arial.ttf",           # Arial (Win)
            "C:\\Windows\\Fonts\\segoeui.ttf",         # Segoe UI (Win)
            "/System/Library/Fonts/Helvetica.ttc",     # macOS
            "/Library/Fonts/Arial.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/TTF/DejaVuSans.ttf",
        ]
    for c in candidates:
        if os.path.exists(c):
            try:
                return ImageFont.truetype(c, size)
            except Exception:
                continue
    return ImageFont.load_default()


def make_a4_page_from_pdf(
    pdf_path: str,
    bobby_mak: Optional[tuple[int, float]],
    header_path: Optional[str] = None,
    header_layout: Optional[dict] = None,
    section_heights: Optional[dict] = None,
) -> Optional[Image.Image]:
    """Build one A4 page from a PDF.

    Steps:
      1. Render the Bobby-Mak page of the PDF to a 200-DPI image.
      2. Find the article title (first text line below the partial
         header that print-CSS left behind). The crop window ALWAYS
         starts at the title, so the story subject and publish date
         are never skipped.
      3. Focus the crop around the Bobby Mak paragraph if the article
         is longer than one A4 page, but keep the title visible.
      4. Paste the header image at the user-configured (x, y, w, h) on
         a white A4 canvas.
      5. Draw a pale border + a vertical scroll bar around the article
         area so it's obvious the page is a "scrolled view".
      6. Highlight the Bobby Mak paragraph in light lemon chiffon with
         a thick gold left border, alpha-composited so the text stays
         readable.

    `section_heights` (optional) is a dict with keys "header", "title",
    "article" giving the pixel heights of the three A4 sections.
    Defaults: {header: 220, title: 130, article: 1404}. The article
    section is filled (not centered) so the whole A4 is used up.
    """
    from PIL import ImageDraw

    if header_layout is None:
        header_layout = {"x": 0, "y": 0, "w": A4_W_PX, "h": None, "lock_aspect": True}
    if section_heights is None:
        section_heights = {"header": 220, "title": 260, "article": 1274}
    intro_lines_target = int(section_heights.get("intro_lines", 3))
    # Default: no intro, so the source crop starts at the top of the page
    intro_end_y_px = 0

    pages = pdf_to_images(pdf_path, dpi=200)
    if not pages:
        return None
    target_idx = min(bobby_mak[0], len(pages) - 1) if bobby_mak else 0
    primary = pages[target_idx]
    pw, ph = primary.size
    a4_w, a4_h = A4_W_PX, A4_H_PX
    DPI = 200

    # ---- 1. Find the title and date by scanning the first page ----
    # (The title is on page 1 of the article, but Bobby Mak may be on a
    # later page. We extract the title and date as text and render them
    # ourselves so they're always at the top of the A4 page.)
    title_text, date_text = _find_title_and_date(pdf_path)
    if not title_text:
        title_text = "Article"
    if not date_text:
        date_text = ""

    # ---- 2. Find the paragraph end (for highlight height) ----
    para_height_px = 200
    para_end_pt = None
    if bobby_mak:
        _, bobby_y_pt = bobby_mak
        try:
            with pdfplumber.open(pdf_path) as ppdf:
                ppage = ppdf.pages[target_idx]
                lines = ppage.extract_text_lines() or []
                bobby_line = None
                for ln in lines:
                    if "Bobby Mak" in (ln.get("text") or "") and len(ln.get("text") or "") < 1200:
                        bobby_line = ln
                        break
                if bobby_line is not None:
                    bobby_top = float(bobby_line.get("top", 0))
                    post_lines = [ln for ln in lines if float(ln.get("top", 0)) >= bobby_top]
                    if len(post_lines) >= 2:
                        typical_gap = (
                            float(post_lines[1].get("top", 0))
                            - float(post_lines[0].get("top", 0))
                        )
                    else:
                        typical_gap = 18
                    para_end_pt = bobby_top + 4 * typical_gap
                    for i in range(1, len(post_lines)):
                        prev = post_lines[i - 1]
                        cur = post_lines[i]
                        if float(cur.get("top", 0)) - float(prev.get("top", 0)) > typical_gap * 1.6:
                            para_end_pt = float(prev.get("top", 0)) + typical_gap
                            break
                    else:
                        if post_lines:
                            last = post_lines[-1]
                            para_end_pt = float(last.get("bottom", last.get("top", 0) + typical_gap))
                    para_height_px = int((para_end_pt - bobby_top) / 72.0 * DPI) + 4
        except Exception:
            pass

    # ---- 3. Decide the crop window ----
    # Use the FULL Bobby-Mak page (no partial header on continuation
    # pages of Mingtiandi articles, so the whole page is article
    # body). Scale to fit the article section's height on the A4,
    # which fills the section vertically (with side margins if the
    # source aspect is narrower than the section). The article is
    # then centered horizontally. This makes the A4 fully used.
    #
    # If the source page DOES have a partial header (page 1 of the
    # article on some sites), we still want the page anyway —
    # seeing the partial header at the top of the article is a
    # small price for filling the A4.
    #
    # NOTE on Y coords: pdfplumber's "top" attribute is the distance
    # from the TOP of the page (top-origin), so converting to image
    # pixels is simply top_pt / 72 * DPI. Verified empirically with
    # an overlay against the rendered image.
    full_h_px = ph
    bobby_y_px_full = int((bobby_mak[1] if bobby_mak else 0) / 72.0 * DPI)
    # v4.5: Compute intro_end_y_px HERE (before article_img) so the
    # article section doesn't include the intro lines. The intro is
    # rendered as text in the title section; the article section
    # must start AFTER the intro to avoid duplication.
    if bobby_mak and intro_lines_target > 0 and intro_end_y_px == 0:
        try:
            with pdfplumber.open(pdf_path) as _ppdf:
                _ppage = _ppdf.pages[bobby_mak[0]]
                _body_lines = []
                for _ln in _ppage.extract_text_lines() or []:
                    _top = float(_ln.get("top", 0))
                    _text = (_ln.get("text") or "").strip()
                    if _text:
                        _body_lines.append((_top, _text))
                if len(_body_lines) > intro_lines_target:
                    intro_end_y_px = int(
                        _body_lines[intro_lines_target][0] / 72.0 * DPI
                    ) - 6
        except Exception:
            pass
    # Source crop: use the entire page (y=0 to y=ph). The article
    # section height controls the A4 height; the source aspect
    # controls the width (with side margins if narrower). If the
    # title section is showing an intro (the first N body lines),
    # the source crop starts AFTER the intro so the article section
    # doesn't duplicate the intro.
    src_crop_top = max(0, intro_end_y_px)
    src_crop_bottom = ph
    src_crop_offset_y = src_crop_top
    article_img = primary.crop((0, src_crop_top, pw, src_crop_bottom))

    # v4.5: If the "next paragraph" after Bobby Mak continues on the
    # following PDF page (e.g. a paragraph split across a page break),
    # the user can't see its tail. Render the next page, find where
    # the next paragraph ends, and stitch the tail of the next page
    # onto the bottom of article_img so the full next paragraph is
    # visible. (See: "The next paragraph of my quote one is not
    # shown" — the paragraph spanned two pages of the saved PDF.)
    if bobby_mak and target_idx + 1 < len(pages):
        try:
            next_page = pages[target_idx + 1]
            np_w, np_h = next_page.size
            # Find the end of the "next paragraph" — the one
            # immediately after Bobby's paragraph.
            with pdfplumber.open(pdf_path) as ppdf2:
                np_page = ppdf2.pages[target_idx + 1]
                np_body_lines = []
                for ln in np_page.extract_text_lines() or []:
                    top = float(ln.get("top", 0))
                    bottom = float(ln.get("bottom", top + 18))
                    text = (ln.get("text") or "").strip()
                    if text:
                        np_body_lines.append((top, bottom, text))
                # Where does the next paragraph end on this page?
                # Look for a line whose top is well below the previous
                # line's bottom (a paragraph gap > 12pt) — that's the
                # end of the next paragraph. If no such gap, the
                # paragraph extends to the last text line on the page.
                next_para_end_pt = None
                if np_body_lines:
                    last_bottom = np_body_lines[0][1]
                    for j, (top, bot, text) in enumerate(np_body_lines):
                        if j > 0 and top - last_bottom > 12:
                            # The previous paragraph (the "next
                            # paragraph" after Bobby) ended at
                            # last_bottom. Use that as the cut point.
                            next_para_end_pt = last_bottom
                            break
                        last_bottom = bot
                    if next_para_end_pt is None:
                        # No paragraph break found — use the last line
                        next_para_end_pt = np_body_lines[-1][1]
            if next_para_end_pt is not None:
                # Crop the next page from y=0 to y=next_para_end_pt
                next_crop_bottom_px = min(
                    int(next_para_end_pt / 72.0 * DPI) + 2,
                    np_h,
                )
                next_page_tail = next_page.crop(
                    (0, 0, np_w, next_crop_bottom_px)
                )
                # Vertically concatenate: article_img (bottom of
                # Bobby's page) + next_page_tail (top of next page,
                # up to the end of the next paragraph).
                new_w = pw
                new_h = article_img.height + next_page_tail.height
                combined = Image.new("RGB", (new_w, new_h), "white")
                combined.paste(article_img, (0, 0))
                combined.paste(next_page_tail, (0, article_img.height))
                article_img = combined
                # Note: src_crop_offset_y stays the same (only the
                # bottom of the source has been extended).
                full_h_px = new_h  # for the scrollbar thumb math
        except Exception as e:
            # If page combining fails for any reason, fall back to
            # the single-page article (the previous behavior).
            print(f"    [!] page-combine skipped: {e}", flush=True)

    # ---- 4. Soft yellow highlight on the Bobby Mak paragraph ----
    if bobby_mak:
        bobby_y_px_in_cropped = bobby_y_px_full - src_crop_offset_y
        if para_end_pt is not None:
            para_end_px_full = int(para_end_pt / 72.0 * DPI)
            para_h_in_cropped = para_end_px_full - src_crop_offset_y - bobby_y_px_in_cropped
        else:
            para_h_in_cropped = para_height_px
        highlight_y0 = max(0, bobby_y_px_in_cropped - 4)
        highlight_y1 = min(article_img.height, bobby_y_px_in_cropped + para_h_in_cropped)
        if highlight_y1 > highlight_y0:
            article_rgba = article_img.convert("RGBA")
            overlay = Image.new("RGBA", article_img.size, (0, 0, 0, 0))
            odraw = ImageDraw.Draw(overlay)
            # Light lemon chiffon at 45% opacity — text stays readable
            odraw.rectangle(
                [(0, highlight_y0), (article_img.width, highlight_y1)],
                fill=(255, 250, 205, 115),  # #fffacd, ~45% opacity
            )
            # Thicker gold left border (8px instead of 4)
            odraw.rectangle(
                [(0, highlight_y0), (8, highlight_y1)],
                fill=(246, 192, 38, 255),  # #f6c026, full opacity
            )
            article_rgba = Image.alpha_composite(article_rgba, overlay)
            article_img = article_rgba.convert("RGB")

    # ---- 5. Build the A4 canvas with header + title + article + scroll cue ----
    canvas = Image.new("RGB", (a4_w, a4_h), "white")

    # Resolve header geometry
    hdr_y = int(header_layout.get("y", 0))
    hdr_w = int(header_layout.get("w", a4_w))
    hdr_h_cfg = header_layout.get("h")
    if header_path and os.path.exists(header_path):
        try:
            with Image.open(header_path) as _h_img:
                src_w, src_h = _h_img.size
        except Exception:
            src_w, src_h = hdr_w, 200
    else:
        src_w, src_h = hdr_w, 200
    if hdr_h_cfg is None:
        hdr_h = int(hdr_w * src_h / max(1, src_w))
    else:
        hdr_h = int(hdr_h_cfg)

    # Paste header at the user-configured (x, y, w, h)
    if header_path and os.path.exists(header_path):
        try:
            header = Image.open(header_path).convert("RGB")
            header = header.resize((hdr_w, hdr_h), Image.LANCZOS)
            paste_x = int(header_layout.get("x", 0))
            paste_y = int(header_layout.get("y", 0))
            canvas.paste(header, (paste_x, paste_y))
        except Exception as e:
            print(f"    [!] Header paste failed: {e}", flush=True)
            hdr_h = 0

    # ---- 5b. Render the article title + publish date as text ----
    # Placed in a band between the header and the article body so the
    # story subject is always visible at the top of the A4 page.
    # Use the user-specified section heights so the layout uses the
    # whole A4 page.
    draw = ImageDraw.Draw(canvas)
    title_font = _load_font(40, "black")
    date_font = _load_font(18, "regular")
    text_left = 40
    text_right = a4_w - 60  # leave room for the scroll bar

    # Layout: header from 0 to header_section_h, title section from
    # header_section_h to header_section_h + title_section_h, article
    # section fills the rest.
    title_section_top = int(section_heights.get("header", 220))
    title_section_h = int(section_heights.get("title", 130))
    article_section_top = title_section_top + title_section_h
    article_area_h = int(section_heights.get("article", 1404))

    # Pad inside the title section
    title_pad = 12
    title_y_text = title_section_top + title_pad

    # Word-wrap the title into multiple lines that fit `text_right - text_left`
    def wrap(text, font, max_w):
        words = text.split()
        lines, cur = [], ""
        for w in words:
            test = (cur + " " + w).strip()
            bbox = draw.textbbox((0, 0), test, font=font)
            if bbox[2] - bbox[0] <= max_w or not cur:
                cur = test
            else:
                lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        return lines
    title_lines = wrap(title_text.upper(), title_font, text_right - text_left)
    line_h = 46
    # If there are too many title lines to fit the section, shrink the
    # line height so we still fit.
    max_title_lines = max(1, (title_section_h - title_pad * 2 - 24) // line_h)
    if len(title_lines) > max_title_lines:
        line_h = max(28, (title_section_h - title_pad * 2 - 24) // max(1, len(title_lines)))
        title_font = _load_font(max(20, int(line_h * 0.82)))
    for i, line in enumerate(title_lines):
        draw.text((text_left, title_y_text + i * line_h), line,
                  font=title_font, fill=(38, 38, 42), spacing=1)
    title_used_h = len(title_lines) * line_h

    # ---- 5c. Render the intro (first N lines of the article body) ----
    # The user-configured intro_lines count says how many body lines to
    # bridge from the title section into the scrolled-down section.
    # The intro_text_lines list is built in step 1 from pdfplumber.
    intro_text: list[str] = []
    intro_end_y_px = 0  # the y-position (in source image) where the
                         # intro ends and the article section begins.
                         # 0 = no intro, start at the top of the source.
    if bobby_mak and intro_lines_target > 0:
        try:
            with pdfplumber.open(pdf_path) as ppdf:
                ppage = ppdf.pages[bobby_mak[0]]
                body_lines = []  # (y_pt_top, text) tuples
                for ln in ppage.extract_text_lines() or []:
                    top_pt = float(ln.get("top", 0))
                    text = (ln.get("text") or "").strip()
                    if text:
                        body_lines.append((top_pt, text))
                # Take the first N lines as the intro
                for i in range(min(intro_lines_target, len(body_lines))):
                    intro_text.append(body_lines[i][1])
                # The intro ends at the bottom of the last intro line
                if intro_text and len(intro_text) < len(body_lines):
                    next_top_pt = body_lines[len(intro_text)][0]
                    intro_end_y_px = int(next_top_pt / 72.0 * DPI) - 6
                elif intro_text:
                    # Intro consumes all body lines on this page; the
                    # article section will still start at the top of
                    # the page (no offset).
                    last_bottom_pt = float(
                        ppage.extract_text_lines()[-1].get(
                            "bottom", body_lines[-1][0] + 18
                        )
                    )
                    intro_end_y_px = int(last_bottom_pt / 72.0 * DPI) + 6
        except Exception:
            intro_text = []
            intro_end_y_px = 0
    # If the title section is too small to fit the intro, hide it
    intro_y_text = title_y_text + title_used_h + 6
    if date_text:
        date_y_text = intro_y_text
        draw.text((text_left, date_y_text), date_text,
                  font=date_font, fill=(120, 120, 130))
        intro_y_text = date_y_text + 30  # gap below date
    if intro_text:
        intro_font = _load_font(18, "regular")
        intro_line_h = 26
        available_h = title_section_top + title_section_h - intro_y_text - 14
        max_intro_lines = max(0, available_h // intro_line_h)
        actual_intro = intro_text[:max_intro_lines]
        for i, line in enumerate(actual_intro):
            # Word-wrap if the line is too long for the section
            words = line.split()
            cur = ""
            for w in words:
                test = (cur + " " + w).strip()
                bbox = draw.textbbox((0, 0), test, font=intro_font)
                if bbox[2] - bbox[0] <= text_right - text_left or not cur:
                    cur = test
                else:
                    draw.text((text_left, intro_y_text + i * intro_line_h),
                              cur, font=intro_font, fill=(75, 75, 80))
                    i += 1
                    cur = w
            if cur:
                draw.text((text_left, intro_y_text + i * intro_line_h),
                          cur, font=intro_font, fill=(75, 75, 80))
    article_top = article_section_top

    # Fit the article into the article area — fill the section
    # vertically (preserve aspect ratio; center horizontally if the
    # source is narrower than the section).
    a_w, a_h = article_img.size
    if a_h > 0 and a_w > 0:
        # v4.5: Scale to fit the section width so the user can place
        # Bobby Mak anywhere in the section. Previously we scaled to
        # the section height (which forced Bobby Mak to the natural
        # top/bottom position); now we have a bobby_position slider.
        scale = a4_w / a_w
        new_w = a4_w
        new_h = max(1, int(a_h * scale))
        # Compute the bobby position slider value (clamped 0..1).
        bobby_position = float(section_heights.get("bobby_position", 0.5))
        bobby_position = max(0.0, min(1.0, bobby_position))
        # Compute where Bobby Mak sits in the SCALED article image.
        if bobby_mak and a_h > 0:
            bobby_y_in_scaled = (bobby_y_px_full - src_crop_offset_y) * scale
        else:
            bobby_y_in_scaled = new_h * 0.5  # middle if no Bobby Mak
        # Target Y in the article section (article_top + offset within
        # the section).
        target_y_in_section = article_area_h * bobby_position
        # If the article is shorter than the section, just center it
        # (the position slider is meaningless when the whole article
        # is already visible). If the article is taller, position
        # Bobby Mak at the target Y, cutting the top/bottom as needed.
        if new_h <= article_area_h:
            article_y_offset = (article_area_h - new_h) // 2
        else:
            article_y_offset = target_y_in_section - bobby_y_in_scaled
            if article_y_offset < 0:
                # Cut the top of the article
                cut_top_in_article = -article_y_offset / scale
                if cut_top_in_article < a_h:
                    article_img = article_img.crop(
                        (0, int(cut_top_in_article), a_w, a_h)
                    )
                    a_h = article_img.height  # v4.5 bugfix: update a_h
                    bobby_y_in_scaled -= cut_top_in_article * scale
                    new_h = max(1, int(a_h * scale))
                article_y_offset = 0
            elif article_y_offset + new_h > article_area_h:
                # Cut the bottom of the article
                overflow = (article_y_offset + new_h) - article_area_h
                cut_bottom_in_article = overflow / scale
                new_height_in_source = max(1, a_h - cut_bottom_in_article)
                if new_height_in_source > 0:
                    article_img = article_img.crop(
                        (0, 0, a_w, int(new_height_in_source))
                    )
                    a_h = article_img.height  # v4.5 bugfix: update a_h
                    new_h = max(1, int(a_h * scale))
                article_y_offset = max(0, article_area_h - new_h)
        resized = article_img.resize((new_w, new_h), Image.LANCZOS)
        # Center horizontally
        x = max(0, (a4_w - new_w) // 2)
        y = article_top + int(article_y_offset)
        border_pad = 4
        border_rect = [
            (x - border_pad, y - border_pad),
            (x + new_w + border_pad - 1, y + new_h + border_pad - 1),
        ]

        # ---- 6. Pale border + drop shadow (scroll cue) ----
        draw = ImageDraw.Draw(canvas)
        # Subtle drop shadow — a soft 2-px gray frame offset down+right.
        # Drawn BEFORE the article so it sits behind it.
        for dx, dy in [(2, 2), (3, 3), (4, 4)]:
            shadow_rect = [
                (border_rect[0][0] + dx, border_rect[0][1] + dy),
                (border_rect[1][0] + dx, border_rect[1][1] + dy),
            ]
            draw.rectangle(shadow_rect, outline=(215, 215, 215), width=1)

        # Paste the article (covers the shadow on its own footprint)
        canvas.paste(resized, (x, y))

        # Pale border around the article (drawn on top of the article edge)
        draw.rectangle(border_rect, outline=(180, 180, 180), width=1)

        # ---- 7. Vertical scroll bar on the right edge ----
        scroll_x = a4_w - 18
        scroll_w = 6
        scroll_top = y - border_pad
        scroll_bottom = y + new_h + border_pad
        scroll_h = scroll_bottom - scroll_top
        # Track
        draw.rectangle(
            [(scroll_x, scroll_top), (scroll_x + scroll_w, scroll_bottom)],
            fill=(235, 235, 235),
        )
        # Thumb — proportion of visible content vs. the full page
        # (reflects any cut at the top/bottom from the bobby_position
        # slider, not just the intro skip)
        visible_h_px = max(1, src_crop_bottom - src_crop_offset_y)
        if full_h_px > 0:
            thumb_h = max(20, int(visible_h_px / full_h_px * scroll_h))
            visible_start_in_full = src_crop_offset_y
            thumb_y = scroll_top + int(visible_start_in_full / full_h_px * scroll_h)
            thumb_y = max(scroll_top, min(scroll_bottom - thumb_h, thumb_y))
            draw.rectangle(
                [(scroll_x, thumb_y), (scroll_x + scroll_w, thumb_y + thumb_h)],
                fill=(150, 150, 150),
            )
    return canvas


def _center_on_a4(img: Image.Image, a4_w: int, a4_h: int) -> Image.Image:
    w, h = img.size
    if h == 0 or w == 0:
        return Image.new("RGB", (a4_w, a4_h), "white")
    scale = min(a4_w / w, a4_h / h)
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))
    img = img.resize((new_w, new_h), Image.LANCZOS)
    canvas = Image.new("RGB", (a4_w, a4_h), "white")
    canvas.paste(img, ((a4_w - new_w) // 2, (a4_h - new_h) // 2))
    return canvas


async def run_scrape_core(
    articles: list[dict],
    output_pdf: str,
    chrome_path: Optional[str],
    headless: bool,
    user_agent: str,
    log_cb=None,
    progress_cb=None,
    cancel_evt: Optional[threading.Event] = None,
    header_path: Optional[str] = None,
    header_layout: Optional[dict] = None,
    section_heights: Optional[dict] = None,
) -> Optional[str]:
    """End-to-end scrape -> process -> combine. log_cb(str), progress_cb(float 0..1)."""
    def log(msg: str):
        if log_cb:
            log_cb(msg)
        else:
            print(msg, flush=True)

    log(f"[*] Loaded {len(articles)} articles")
    log(f"[*] Output: {output_pdf}")
    log(f"[*] Chrome: {chrome_path or 'bundled Chromium (Cloudflare will likely block)'}")
    log(f"[*] Header image: {header_path or '(none)'}")

    work_dir = os.path.join(os.path.dirname(output_pdf) or ".", "_scratch")
    os.makedirs(work_dir, exist_ok=True)
    log(f"[*] Working dir: {work_dir}")

    async with async_playwright() as p:
        browser = await launch_browser(p, chrome_path, headless=headless)
        try:
            a4_pages: list[Image.Image] = []
            for i, art in enumerate(articles, 1):
                if cancel_evt and cancel_evt.is_set():
                    log("[!] Cancelled by user.")
                    return None
                log(f"\n=== Article {i}/{len(articles)} ===")
                raw_pdf = await print_article_to_pdf(
                    browser, art, i, len(articles), work_dir, user_agent,
                    progress_cb=lambda p: progress_cb(p * 0.9) if progress_cb else None,
                )
                if not raw_pdf:
                    continue
                bobby = find_bobby_mak_in_pdf(raw_pdf)
                if bobby is None:
                    log("    [!] Could not locate 'Bobby Mak' in raw PDF text.")
                else:
                    log(f"    Bobby Mak on page {bobby[0]+1}, y={bobby[1]:.0f}pt")
                a4 = make_a4_page_from_pdf(
                    raw_pdf, bobby,
                    header_path=header_path,
                    header_layout=header_layout,
                    section_heights=section_heights,
                )
                if a4 is not None:
                    a4_pages.append(a4)
                    png_out = os.path.join(work_dir, f"article_{i:02d}.png")
                    a4.save(png_out, "PNG", optimize=True)
                    log(f"    a4 page saved: {png_out}")
            if not a4_pages:
                log("[!] No pages to combine.")
                return None
            log("\n=== Combining pages into final PDF ===")
            a4_pages[0].save(
                output_pdf,
                save_all=True,
                append_images=a4_pages[1:],
                resolution=150.0,
                quality=88,
            )
            if progress_cb:
                progress_cb(1.0)
            log(f"\n[OK] Final PDF: {output_pdf}  ({len(a4_pages)} pages)")
            return output_pdf
        finally:
            await browser.close()


# ===========================================================================
# GUI
# ===========================================================================


class GuidedManualDialog:
    """Walks the user through saving each article as a PDF in their own
    default browser. The dialog polls the work folder for new PDFs and
    auto-advances when one appears.

    Lifecycle:
      win.wait_window()  (in caller) -> result available via .result
      .result = (saved_filenames, skipped_indices, stopped_early)
    """

    def __init__(self, parent, articles: list[dict], work_dir: str, log_cb):
        self.parent = parent
        self.articles = articles
        self.work_dir = work_dir
        self.log_cb = log_cb
        self.current_idx = 0
        self.saved_files: list[str] = []
        self.skipped: list[int] = []
        self.stopped = False
        self.result: tuple[list[str], list[int], bool] = ([], [], False)

        os.makedirs(work_dir, exist_ok=True)
        self.existing_files = self._snapshot()

        self.win = self.parent._tk.Toplevel(parent)
        self.win.title("Guided Manual Save")
        self.win.geometry("760x620")
        self.win.minsize(640, 540)
        self.win.transient(parent)
        self.win.grab_set()
        self.win.protocol("WM_DELETE_WINDOW", self._on_stop)
        # Bring to front
        self.win.lift()
        self.win.attributes("-topmost", True)
        self.win.after(500, lambda: self.win.attributes("-topmost", False))

        self._build()
        self._open_current()
        self._poll()

    # ---------------------- helpers ----------------------
    def _snapshot(self) -> set[str]:
        try:
            return set(os.listdir(self.work_dir))
        except FileNotFoundError:
            return set()

    # ---------------------- UI build ----------------------
    def _build(self):
        tk = self.parent._tk
        ttk = self.parent._ttk
        pad = {"padx": 10, "pady": 4}

        frm = ttk.Frame(self.win, padding=14)
        frm.pack(fill="both", expand=True)

        # Header
        self.header_var = tk.StringVar()
        ttk.Label(
            frm, textvariable=self.header_var, font=("Segoe UI", 13, "bold")
        ).pack(anchor="w", **pad)

        # Title
        self.title_var = tk.StringVar()
        ttk.Label(
            frm,
            textvariable=self.title_var,
            font=("Segoe UI", 10),
            wraplength=720,
            justify="left",
        ).pack(anchor="w", fill="x", **pad)

        # URL
        self.url_var = tk.StringVar()
        url_lbl = ttk.Label(
            frm,
            textvariable=self.url_var,
            foreground="#0066cc",
            wraplength=720,
            justify="left",
            cursor="hand2",
        )
        url_lbl.pack(anchor="w", fill="x", **pad)
        url_lbl.bind("<Button-1>", lambda e: self._open_current())

        # Folder + filename (the "save here" instructions)
        box = ttk.LabelFrame(frm, text="📁  Save the PDF here", padding=8)
        box.pack(fill="x", **pad)
        ttk.Label(box, text="Folder:", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=4, pady=2
        )
        self.folder_var = tk.StringVar(value=self.work_dir)
        ttk.Entry(box, textvariable=self.folder_var).grid(
            row=0, column=1, sticky="ew", padx=4, pady=2
        )
        ttk.Button(box, text="Copy", command=lambda: self._copy(self.work_dir)).grid(
            row=0, column=2, padx=4, pady=2
        )
        ttk.Label(box, text="Filename:", font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", padx=4, pady=2
        )
        self.expected_var = tk.StringVar()
        self.expected_entry = ttk.Entry(
            box, textvariable=self.expected_var, font=("Consolas", 11, "bold"),
            foreground="#0066cc",
        )
        self.expected_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(
            box, text="Copy", command=lambda: self._copy(self.expected_var.get())
        ).grid(row=1, column=2, padx=4, pady=2)
        box.columnconfigure(1, weight=1)

        # Status
        self.status_var = tk.StringVar(value="Waiting for you to save the PDF…")
        ttk.Label(
            frm,
            textvariable=self.status_var,
            font=("Segoe UI", 10, "italic"),
            foreground="#228B22",
        ).pack(anchor="w", **pad)

        # What to do
        instr = ttk.LabelFrame(frm, text="What to do (in your browser)", padding=8)
        instr.pack(fill="x", **pad)
        for i, line in enumerate([
            "1. Your default browser has opened the article above.",
            "2. If Cloudflare shows a challenge, click the checkbox once.",
            "3. Wait for the real article to load.",
            "4. Press Ctrl+P (or File \u2192 Print).",
            "5. Choose 'Save as PDF' as the destination.",
            "6. Save with the exact filename above, into the folder above.",
            "7. This dialog auto-advances when it sees the new file.",
        ]):
            ttk.Label(instr, text=line, justify="left").grid(
                row=i, column=0, sticky="w", padx=4, pady=1
            )

        # Buttons
        btn_frame = ttk.Frame(frm)
        btn_frame.pack(fill="x", **pad)
        ttk.Button(
            btn_frame, text="\u23ed  Skip", command=self._on_skip
        ).pack(side="left", padx=4)
        ttk.Button(
            btn_frame, text="\ud83d\udd04  Reopen URL", command=self._open_current
        ).pack(side="left", padx=4)
        ttk.Button(
            btn_frame,
            text="\ud83d\udcc1  I saved it under a different name",
            command=self._on_browse,
        ).pack(side="left", padx=4)
        ttk.Button(
            btn_frame, text="\u2705  Done with this one", command=self._on_done
        ).pack(side="right", padx=4)
        ttk.Button(
            btn_frame, text="\u23f9  Stop", command=self._on_stop
        ).pack(side="right", padx=4)

    # ---------------------- state updates ----------------------
    def _update_labels(self):
        art = self.articles[self.current_idx]
        n = self.current_idx + 1
        total = len(self.articles)
        self.header_var.set(f"Article {n} of {total}")
        self.title_var.set(art["title"])
        self.url_var.set(art["url"])
        self.expected_var.set(f"article_{n:02d}.pdf")
        self.status_var.set("Waiting for you to save the PDF\u2026")

    def _open_current(self):
        art = self.articles[self.current_idx]
        self._update_labels()
        try:
            webbrowser.open(art["url"])
            self.log_cb(
                f"[guided] Opened article {self.current_idx+1}/{len(self.articles)}: {art['title']}"
            )
        except Exception as e:
            self.log_cb(f"[guided] webbrowser.open failed: {e}")
        self.status_var.set(
            "Browser opened. Solve the challenge (if any), then Ctrl+P, then save."
        )

    def _copy(self, text: str):
        try:
            self.win.clipboard_clear()
            self.win.clipboard_append(text)
            self.status_var.set(f"Copied: {text}")
        except Exception:
            pass

    # ---------------------- file polling ----------------------
    def _poll(self):
        if not self._alive():
            return
        try:
            current = self._snapshot()
            new_pdfs = sorted(
                f for f in (current - self.existing_files)
                if f.lower().endswith(".pdf")
            )
            if new_pdfs:
                self._wait_for_stable(new_pdfs[0])
                return
        except Exception as e:
            self.log_cb(f"[guided] poll error: {e}")
        self.win.after(1000, self._poll)

    def _wait_for_stable(self, filename: str, prev_size: int = -1):
        if not self._alive():
            return
        path = os.path.join(self.work_dir, filename)
        if not os.path.exists(path):
            self.win.after(1000, self._poll)
            return
        size = os.path.getsize(path)
        if prev_size == -1:
            self.win.after(1500, lambda: self._wait_for_stable(filename, size))
            return
        if size == prev_size and size > 1000:
            # Stable, looks like a real PDF. Rename to the expected name
            # (in case the user saved with a different one) and advance.
            n = self.current_idx + 1
            target = f"article_{n:02d}.pdf"
            if filename != target:
                target_path = os.path.join(self.work_dir, target)
                if os.path.exists(target_path):
                    os.remove(target_path)
                try:
                    os.rename(path, target_path)
                except Exception:
                    shutil.move(path, target_path)
                filename = target
            self.status_var.set(
                f"\u2705 Got {filename} \u2014 moving to the next article\u2026"
            )
            self.log_cb(f"[guided] Auto-detected: {filename}")
            self._advance(filename)
        else:
            self.win.after(1500, lambda: self._wait_for_stable(filename, size))

    # ---------------------- user actions ----------------------
    def _on_done(self):
        expected = f"article_{self.current_idx+1:02d}.pdf"
        path = os.path.join(self.work_dir, expected)
        if os.path.exists(path) and os.path.getsize(path) > 1000:
            self._advance(expected)
        else:
            self.status_var.set(
                f"\u274c {expected} not found in folder. Save it first, then click Done."
            )

    def _on_skip(self):
        self.skipped.append(self.current_idx)
        self.log_cb(
            f"[guided] Skipped article {self.current_idx+1}/{len(self.articles)}"
        )
        self._advance(None)

    def _on_browse(self):
        path = self.parent._filedialog.askopenfilename(
            title="Pick the PDF you just saved",
            initialdir=self.work_dir,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return
        target = os.path.join(
            self.work_dir, f"article_{self.current_idx+1:02d}.pdf"
        )
        if os.path.normcase(path) != os.path.normcase(target):
            if os.path.exists(target):
                os.remove(target)
            try:
                os.rename(path, target)
            except Exception:
                shutil.move(path, target)
        self._advance(f"article_{self.current_idx+1:02d}.pdf")

    def _on_stop(self):
        self.stopped = True
        # Capture whatever was saved so the caller can still combine
        self.result = (self.saved_files, self.skipped, True)
        self.status_var.set("Stopping\u2026")
        self.win.destroy()

    def _advance(self, saved_file: Optional[str]):
        if saved_file:
            self.saved_files.append(saved_file)
        self.current_idx += 1
        if self.current_idx >= len(self.articles):
            # all done
            self.result = (self.saved_files, self.skipped, False)
            self.win.destroy()
            return
        self.existing_files = self._snapshot()
        self._open_current()
        self._poll()

    def _alive(self) -> bool:
        try:
            return bool(self.win.winfo_exists())
        except Exception:
            return False

    def wait(self):
        """Block until the dialog is closed. Returns .result."""
        try:
            self.win.wait_window()
        except Exception:
            pass
        if self.stopped and not self.saved_files and not self.skipped:
            self.result = ([], [], True)
        return self.result


class HeaderLayoutDialog:
    """Live-preview dialog that lets the user pick a header PNG, set its
    position (X, Y) and size (W, H) on the A4 canvas, and toggle
    "lock aspect ratio" so width and height stay in sync.

    The A4 canvas is 1240x1754 px. The preview is a 248x351 thumbnail
    (1/5 scale) with a 1-px orange rectangle showing where the header
    will land. All controls update the preview in real time.

    On Apply, calls on_apply(layout, path) and closes.
    On Cancel, discards changes and closes.
    """

    A4_W = 1240
    A4_H = 1754
    PREVIEW_SCALE = 5  # 1240 / 5 = 248 px wide preview

    def __init__(self, parent, header_path: str, layout: dict, on_apply):
        self.parent = parent
        self.header_path = header_path
        self.layout = dict(layout)
        self.on_apply = on_apply
        # Load the source image to read its aspect ratio
        self.src_w, self.src_h = 1141, 200  # default header size
        if header_path and os.path.exists(header_path):
            try:
                with Image.open(header_path) as im:
                    self.src_w, self.src_h = im.size
            except Exception:
                pass
        # Resolve auto height from width + aspect
        if self.layout.get("h") is None and self.src_w and self.src_h:
            self.layout["h"] = int(self.layout["w"] * self.src_h / self.src_w)

        self.win = parent._tk.Toplevel(parent.root)
        self.win.title("Customize Header Layout")
        self.win.geometry("720x540")
        self.win.minsize(640, 480)
        self.win.transient(parent.root)
        self.win.grab_set()
        self.win.lift()
        self.win.attributes("-topmost", True)
        self.win.after(500, lambda: self.win.attributes("-topmost", False))

        self._build()
        self._update_preview()

    def _build(self):
        tk = self.parent._tk
        ttk = self.parent._ttk
        pad = {"padx": 8, "pady": 4}

        # Top: file path + import
        top = ttk.Frame(self.win, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Header image:").grid(row=0, column=0, sticky="w", **pad)
        self.path_var = tk.StringVar(value=self.header_path)
        ttk.Entry(top, textvariable=self.path_var).grid(
            row=0, column=1, sticky="ew", **pad
        )
        ttk.Button(top, text="Import…", command=self._on_import).grid(
            row=0, column=2, **pad
        )
        ttk.Button(top, text="Reset", command=self._on_reset).grid(
            row=0, column=3, **pad
        )
        top.columnconfigure(1, weight=1)

        # Two columns: preview (left) and controls (right)
        body = ttk.Frame(self.win, padding=10)
        body.pack(fill="both", expand=True)

        # Left: A4 preview canvas
        left = ttk.LabelFrame(body, text="A4 preview (preview is 1/5 scale)", padding=8)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        prev_w = self.A4_W // self.PREVIEW_SCALE
        prev_h = self.A4_H // self.PREVIEW_SCALE
        self.canvas = tk.Canvas(
            left, width=prev_w, height=prev_h,
            background="#f7f7f7", highlightthickness=1, highlightbackground="#ccc",
        )
        self.canvas.pack(pady=4)
        self.canvas_text = ttk.Label(left, text="", font=("Consolas", 9))
        self.canvas_text.pack(pady=4)

        # Right: controls
        right = ttk.LabelFrame(body, text="Position & size (pixels on the A4 canvas)", padding=8)
        right.pack(side="left", fill="y", padx=(8, 0))

        self.x_var = tk.IntVar(value=int(self.layout.get("x", 0)))
        self.y_var = tk.IntVar(value=int(self.layout.get("y", 0)))
        self.w_var = tk.IntVar(value=int(self.layout.get("w", self.A4_W)))
        self.h_var = tk.IntVar(value=int(self.layout.get("h", 200)))
        self.lock_var = tk.BooleanVar(value=bool(self.layout.get("lock_aspect", True)))

        def add_spinbox(parent, label, var, frm, row, mn, mx, cmd=None):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=2)
            sb = ttk.Spinbox(
                parent, from_=mn, to=mx, textvariable=var, width=8, increment=1,
                command=cmd or (lambda: self._on_value_change()),
            )
            sb.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
            sb.bind("<KeyRelease>", lambda e: self._on_value_change())
            return sb

        add_spinbox(right, "X (from left):",   self.x_var, right, 0, -500, self.A4_W)
        add_spinbox(right, "Y (from top):",    self.y_var, right, 1, -500, self.A4_H)
        add_spinbox(right, "Width:",            self.w_var, right, 2, 50,   self.A4_W)
        add_spinbox(right, "Height:",           self.h_var, right, 3, 20,   self.A4_H + 500)

        ttk.Checkbutton(
            right, text="Lock aspect ratio (changes to W auto-update H)",
            variable=self.lock_var, command=self._on_lock_toggle,
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=4, pady=4)

        ttk.Separator(right, orient="horizontal").grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=6
        )

        # Preset buttons
        ttk.Label(right, text="Presets:").grid(row=6, column=0, sticky="w", padx=4, pady=2)
        preset_frame = ttk.Frame(right)
        preset_frame.grid(row=6, column=1, sticky="w", padx=4, pady=2)
        ttk.Button(preset_frame, text="Full width", command=self._preset_full,
                   width=10).pack(side="left", padx=2)
        ttk.Button(preset_frame, text="Left half",  command=self._preset_left,
                   width=10).pack(side="left", padx=2)
        ttk.Button(preset_frame, text="Right half", command=self._preset_right,
                   width=10).pack(side="left", padx=2)

        right.columnconfigure(1, weight=1)

        # Bottom: result summary + apply/cancel
        bottom = ttk.Frame(self.win, padding=10)
        bottom.pack(fill="x")
        self.summary_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.summary_var, font=("Consolas", 9),
                  foreground="#444").pack(side="left", padx=4)
        ttk.Button(bottom, text="Cancel", command=self._on_cancel, width=12).pack(
            side="right", padx=4
        )
        ttk.Button(bottom, text="Apply", command=self._on_apply, width=12).pack(
            side="right", padx=4
        )

        # Initial summary
        self._refresh_summary()

    # ---------------------- handlers ----------------------
    def _on_import(self):
        path = self.parent._filedialog.askopenfilename(
            title="Pick a header image",
            filetypes=[("PNG / JPG / GIF", "*.png *.jpg *.jpeg *.gif *.bmp"),
                       ("All files", "*.*")],
        )
        if not path:
            return
        self.header_path = path
        self.path_var.set(path)
        # Refresh aspect ratio
        try:
            with Image.open(path) as im:
                self.src_w, self.src_h = im.size
        except Exception:
            pass
        # If lock aspect, re-derive height from current width
        if self.lock_var.get() and self.src_w and self.src_h:
            self.h_var.set(int(self.w_var.get() * self.src_h / self.src_w))
        self._on_value_change()

    def _on_reset(self):
        self.x_var.set(0)
        self.y_var.set(0)
        self.w_var.set(self.A4_W)
        if self.src_w and self.src_h:
            self.h_var.set(int(self.A4_W * self.src_h / self.src_w))
        else:
            self.h_var.set(200)
        self.lock_var.set(True)
        self.header_path = DEFAULT_HEADER_PNG
        self.path_var.set(DEFAULT_HEADER_PNG)
        try:
            with Image.open(DEFAULT_HEADER_PNG) as im:
                self.src_w, self.src_h = im.size
        except Exception:
            pass
        self._on_value_change()

    def _on_lock_toggle(self):
        if self.lock_var.get() and self.src_w and self.src_h:
            self.h_var.set(int(self.w_var.get() * self.src_h / self.src_w))
        self._on_value_change()

    def _on_value_change(self):
        if self.lock_var.get() and self.src_w and self.src_h:
            self.h_var.set(int(self.w_var.get() * self.src_h / self.src_w))
        self._update_preview()
        self._refresh_summary()

    def _preset_full(self):
        self.x_var.set(0)
        self.w_var.set(self.A4_W)
        self._on_value_change()

    def _preset_left(self):
        self.w_var.set(self.A4_W // 2)
        if self.lock_var.get() and self.src_w and self.src_h:
            self.h_var.set(int(self.w_var.get() * self.src_h / self.src_w))
        self._on_value_change()

    def _preset_right(self):
        self.w_var.set(self.A4_W // 2)
        self.x_var.set(self.A4_W - self.w_var.get())
        if self.lock_var.get() and self.src_w and self.src_h:
            self.h_var.set(int(self.w_var.get() * self.src_h / self.src_w))
        self._on_value_change()

    def _update_preview(self):
        s = self.PREVIEW_SCALE
        c = self.canvas
        c.delete("all")
        prev_w = self.A4_W // s
        prev_h = self.A4_H // s
        # A4 page outline
        c.create_rectangle(0, 0, prev_w, prev_h, outline="#888", fill="white", width=1)
        # Header rectangle
        hx = int(self.x_var.get() / s)
        hy = int(self.y_var.get() / s)
        hw = max(1, int(self.w_var.get() / s))
        hh = max(1, int(self.h_var.get() / s))
        c.create_rectangle(
            hx, hy, hx + hw, hy + hh,
            outline="#f6c026", fill="#fff8e1", width=1,
        )
        # If the header is positioned roughly at the top, also draw a
        # thin guide line at y=0 to show "page top"
        c.create_line(0, 0, prev_w, 0, fill="#bbb", dash=(2, 2))

    def _refresh_summary(self):
        x, y = self.x_var.get(), self.y_var.get()
        w, h = self.w_var.get(), self.h_var.get()
        self.summary_var.set(
            f"Header: {w} x {h} px at (x={x}, y={y}) on a 1240 x 1754 A4 canvas"
        )

    def _on_apply(self):
        layout = {
            "x": int(self.x_var.get()),
            "y": int(self.y_var.get()),
            "w": int(self.w_var.get()),
            "h": int(self.h_var.get()),
            "lock_aspect": bool(self.lock_var.get()),
        }
        try:
            self.on_apply(layout, self.header_path)
        except Exception as e:
            self.parent._messagebox.showerror("Apply failed", str(e))
            return
        self.win.destroy()

    def _on_cancel(self):
        self.win.destroy()


class ScraperGUI:
    """Tkinter GUI. Routes print/log output through self._log."""

    def __init__(self):
        tk, ttk, filedialog, messagebox, scrolledtext = _import_gui()
        self._tk = tk
        self._ttk = ttk
        self._filedialog = filedialog
        self._messagebox = messagebox

        self.root = tk.Tk()
        self.root.title("Mingtiandi Bobby Mak PDF Builder")
        self.root.geometry("960x720")
        self.root.minsize(820, 600)

        # State
        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.cancel_evt = threading.Event()
        self.worker: Optional[threading.Thread] = None
        self.output_pdf_path: Optional[str] = None

        # Detect a sensible default Excel path
        script_dir = os.path.dirname(os.path.abspath(__file__))
        default_xlsx = os.path.join(script_dir, "articles.xlsx")
        default_xlsx = default_xlsx if os.path.exists(default_xlsx) else ""
        # Output folder: ~/Documents by default
        default_out = os.path.join(os.path.expanduser("~"), "Documents")

        self.excel_var = tk.StringVar(value=default_xlsx)
        self.output_dir_var = tk.StringVar(value=default_out)
        self.pdfs_dir_var = tk.StringVar(value=default_out)
        self.chrome_var = tk.StringVar(value=find_chrome_path() or "")
        self.filename_var = tk.StringVar(value="bobby_mak_quotes.pdf")
        self.watch_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready.")
        self.progress_var = tk.DoubleVar(value=0.0)
        # v4: mode selector
        self.mode_var = tk.StringVar(value="guided")  # default to guided (Cloudflare-friendly)
        # v4.1: header image (defaults to the bundled default_header.png)
        self.header_var = tk.StringVar(value=DEFAULT_HEADER_PNG)
        self.use_header_var = tk.BooleanVar(value=True)
        # v4.2: header layout (x, y, w, h in pixels on the 1240x1754 A4 canvas)
        # None for w/h means "auto" (full width or aspect-locked)
        self.header_layout = {
            "x": 0,
            "y": 0,
            "w": 1240,
            "h": None,           # auto from width + aspect
            "lock_aspect": True,
        }
        # v4.3: A4 section heights (header bar, title/subject, article window)
        # Sum should be ≈ A4_H_PX (1754). User can edit via the GUI.
        self.header_section_h = tk.IntVar(value=220)
        self.title_section_h = tk.IntVar(value=260)
        self.article_section_h = tk.IntVar(value=1274)
        # v4.4: how many lines of article body the title section should
        # show between the date and the scrolled-down section.
        self.intro_lines = tk.IntVar(value=3)
        # v4.5: where in the article section the Bobby Mak paragraph
        # should sit. 0.0 = top, 1.0 = bottom, 0.5 = middle. Lets the
        # user see context both above and below the quote. The default
        # is computed from the section heights so Bobby lands near the
        # middle of the A4 page (the user said "around half of the A4").
        self.bobby_position = tk.DoubleVar(value=self._default_bobby_position())
        self.height_preset_var = tk.StringVar(value="Balanced")
        # v4: where the GuidedManualDialog writes user-saved PDFs
        self.guided_pdfs_dir: Optional[str] = None

        self._build_ui()
        # Trigger initial enable/disable of PDFs folder field
        self.root.after(50, self._on_mode_change)
        self.root.after(100, self._drain_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------- UI build ----------------------
    def _build_ui(self):
        tk, ttk = self._tk, self._ttk
        root = self.root

        # Use ttk style — looks more modern on Windows
        try:
            style = ttk.Style()
            if "vista" in style.theme_names():
                style.theme_use("vista")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

        pad = {"padx": 8, "pady": 4}

        # Title bar
        title = ttk.Frame(root)
        title.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Label(
            title,
            text="Mingtiandi Bobby Mak Quote PDF Builder",
            font=("Segoe UI", 15, "bold"),
        ).pack(side="left")
        ttk.Label(
            title,
            text="v4.4 — pick a mode, then Start",
            foreground="#666",
        ).pack(side="left", padx=12)

        # v4: Mode selector
        mode_frm = ttk.LabelFrame(root, text="1.  Choose a mode", padding=10)
        mode_frm.pack(fill="x", padx=12, pady=6)
        ttk.Radiobutton(
            mode_frm,
            text="🤖  Auto scrape  — try to load each article in headless Chrome (may be blocked by Cloudflare)",
            variable=self.mode_var,
            value="auto",
            command=self._on_mode_change,
        ).pack(anchor="w", padx=4, pady=2)
        ttk.Radiobutton(
            mode_frm,
            text="🖐  Guided manual  — I open each article in MY browser, save as PDF, the tool does the rest (recommended when Cloudflare is blocking)",
            variable=self.mode_var,
            value="guided",
            command=self._on_mode_change,
        ).pack(anchor="w", padx=4, pady=2)
        ttk.Radiobutton(
            mode_frm,
            text="📄  Combine saved PDFs  — I already saved the article PDFs, just stitch them into one",
            variable=self.mode_var,
            value="chop",
            command=self._on_mode_change,
        ).pack(anchor="w", padx=4, pady=2)

        # File picker section
        frm = ttk.LabelFrame(root, text="2.  Pick your inputs", padding=10)
        frm.pack(fill="x", padx=12, pady=6)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Excel file:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.excel_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self._browse_excel).grid(
            row=0, column=2, **pad
        )

        ttk.Label(frm, text="Output folder:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_dir_var).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self._browse_output_dir).grid(
            row=1, column=2, **pad
        )

        ttk.Label(frm, text="Output filename:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.filename_var).grid(row=2, column=1, sticky="ew", **pad)

        # v4: PDFs folder (used only in "Combine" mode)
        self.pdfs_dir_label = ttk.Label(frm, text="PDFs folder:")
        self.pdfs_dir_label.grid(row=3, column=0, sticky="w", **pad)
        self.pdfs_dir_entry = ttk.Entry(frm, textvariable=self.pdfs_dir_var)
        self.pdfs_dir_entry.grid(row=3, column=1, sticky="ew", **pad)
        self.pdfs_dir_btn = ttk.Button(
            frm, text="Browse…", command=self._browse_pdfs_dir
        )
        self.pdfs_dir_btn.grid(row=3, column=2, **pad)

        # v4.1: Header image (Mingtiandi masthead to paste at the top of
        # every A4 page — Chrome's @media print CSS hides the logo on
        # Save-as-PDF, so we re-add it here)
        self.header_label = ttk.Label(frm, text="Header image:")
        self.header_label.grid(row=4, column=0, sticky="w", **pad)
        self.header_entry = ttk.Entry(frm, textvariable=self.header_var)
        self.header_entry.grid(row=4, column=1, sticky="ew", **pad)
        self.header_btn = ttk.Button(
            frm, text="Browse…", command=self._browse_header
        )
        self.header_btn.grid(row=4, column=2, **pad)
        # v4.2: header layout dialog
        self.header_customize_btn = ttk.Button(
            frm, text="Customize…", command=self._open_header_layout
        )
        self.header_customize_btn.grid(row=4, column=3, **pad)

        # Options section
        opt = ttk.LabelFrame(root, text="3.  Options", padding=10)
        opt.pack(fill="x", padx=12, pady=6)
        opt.columnconfigure(1, weight=1)

        ttk.Label(opt, text="Chrome path (optional):").grid(
            row=0, column=0, sticky="w", **pad
        )
        ttk.Entry(opt, textvariable=self.chrome_var).grid(
            row=0, column=1, sticky="ew", **pad
        )
        ttk.Button(opt, text="Auto-detect", command=self._auto_detect_chrome).grid(
            row=0, column=2, **pad
        )
        ttk.Checkbutton(
            opt,
            text="Show the browser while scraping  (un-check to run hidden)",
            variable=self.watch_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(
            opt,
            text="Add the Mingtiandi masthead to the top of every A4 page (replaces what @media print CSS hides)",
            variable=self.use_header_var,
        ).grid(row=2, column=0, columnspan=3, sticky="w", **pad)

        # v4.3: A4 section heights — user-adjustable
        heights = ttk.LabelFrame(
            root, text="4.  A4 section heights (px, sum should ≈ 1754)", padding=10
        )
        heights.pack(fill="x", padx=12, pady=6)
        heights.columnconfigure(1, weight=1)

        def add_height_row(parent, label, var, row, mn, mx, hint=""):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)
            ttk.Spinbox(
                parent, from_=mn, to=mx, textvariable=var, width=6, increment=10,
                command=self._on_height_change,
            ).grid(row=row, column=1, sticky="w", **pad)
            ttk.Label(parent, text=hint, foreground="#777").grid(
                row=row, column=2, sticky="w", **pad
            )

        add_height_row(heights, "Header bar height:",
                       self.header_section_h, 0, 100, 600,
                       "(Mingtiandi masthead)")
        add_height_row(heights, "Subject section height:",
                       self.title_section_h, 1, 60, 400,
                       "(story subject + date)")
        add_height_row(heights, "Article window height:",
                       self.article_section_h, 2, 200, 1600,
                       "(bordered scroll-down)")

        # v4.4: Intro lines in the subject section (between date and
        # the scroll-down article). Shows the first N lines of the
        # article body so the title section bridges into the story.
        ttk.Label(heights, text="Intro lines in subject:").grid(
            row=3, column=0, sticky="w", **pad
        )
        ttk.Spinbox(
            heights, from_=0, to=12, textvariable=self.intro_lines, width=6, increment=1,
            command=self._on_height_change,
        ).grid(row=3, column=1, sticky="w", **pad)
        ttk.Label(heights, text="(body lines between date and scroll-down)",
                  foreground="#777").grid(row=3, column=2, sticky="w", **pad)

        # v4.5: Bobby Mak position within the article section. Lets
        # the user put the quoted paragraph at the top / upper-third /
        # middle / lower-third / bottom so they can see context both
        # above and below the quote. 0.0 = top, 1.0 = bottom, 0.5 = middle.
        ttk.Label(heights, text="Bobby Mak position:").grid(
            row=4, column=0, sticky="w", **pad
        )
        bobby_pos_frame = ttk.Frame(heights)
        bobby_pos_frame.grid(row=4, column=1, columnspan=2, sticky="ew", **pad)
        self.bobby_position_scale = ttk.Scale(
            bobby_pos_frame, from_=0.0, to=1.0, variable=self.bobby_position,
            orient="horizontal", command=self._on_bobby_position_change,
        )
        self.bobby_position_scale.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.bobby_position_label = ttk.Label(
            bobby_pos_frame, text="50% (middle)", width=14, foreground="#0066cc",
        )
        self.bobby_position_label.pack(side="left", padx=4)
        ttk.Button(bobby_pos_frame, text="Top", width=5,
                   command=lambda: self.bobby_position.set(0.0)).pack(side="left", padx=1)
        ttk.Button(bobby_pos_frame, text="Mid", width=5,
                   command=lambda: self.bobby_position.set(0.5)).pack(side="left", padx=1)
        ttk.Button(bobby_pos_frame, text="Bot", width=5,
                   command=lambda: self.bobby_position.set(1.0)).pack(side="left", padx=1)
        ttk.Button(bobby_pos_frame, text="A4-Mid", width=7,
                   command=lambda: self.bobby_position.set(
                       self._default_bobby_position()
                   )).pack(side="left", padx=1)

        # Preset dropdown
        ttk.Label(heights, text="Preset:").grid(row=3, column=0, sticky="w", **pad)
        preset_frame = ttk.Frame(heights)
        preset_frame.grid(row=3, column=1, columnspan=2, sticky="w", **pad)
        ttk.Label(preset_frame, textvariable=self.height_preset_var,
                  foreground="#555", font=("Segoe UI", 9, "italic")).pack(side="left", padx=4)
        ttk.Button(preset_frame, text="Compact", width=8,
                   command=lambda: self._apply_preset("Compact")).pack(side="left", padx=2)
        ttk.Button(preset_frame, text="Balanced", width=8,
                   command=lambda: self._apply_preset("Balanced")).pack(side="left", padx=2)
        ttk.Button(preset_frame, text="Spacious", width=8,
                   command=lambda: self._apply_preset("Spacious")).pack(side="left", padx=2)
        ttk.Button(preset_frame, text="Fill A4", width=8,
                   command=self._fill_a4).pack(side="left", padx=2)

        # v4.3: a running total showing how the three heights add up
        self.height_total_var = tk.StringVar(value="")
        ttk.Label(heights, textvariable=self.height_total_var,
                  foreground="#0066cc", font=("Consolas", 9)).grid(
            row=5, column=0, columnspan=3, sticky="w", padx=8, pady=(4, 0)
        )
        self.root.after(50, self._on_height_change)

        # Action buttons
        act = ttk.Frame(root)
        act.pack(fill="x", padx=12, pady=6)
        self.start_btn = ttk.Button(
            act, text="▶  Start", command=self._on_start, width=14
        )
        self.start_btn.pack(side="left", padx=4)
        self.stop_btn = ttk.Button(
            act, text="■  Stop", command=self._on_stop, width=14, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=4)
        self.combine_btn = ttk.Button(
            act, text="🔁  Re-combine saved PDFs",
            command=self._on_recombine, width=28,
        )
        self.combine_btn.pack(side="left", padx=4)
        ttk.Button(act, text="📂  Open output folder", command=self._open_output).pack(
            side="right", padx=4
        )
        ttk.Button(act, text="Clear log", command=self._clear_log).pack(
            side="right", padx=4
        )

        # Progress bar
        prog = ttk.Frame(root)
        prog.pack(fill="x", padx=12, pady=(2, 4))
        ttk.Label(prog, textvariable=self.status_var, foreground="#333").pack(
            anchor="w", padx=2
        )
        self.progress = ttk.Progressbar(
            prog, variable=self.progress_var, maximum=100, mode="determinate"
        )
        self.progress.pack(fill="x", padx=2, pady=(2, 4))

        # Log area
        log_frame = ttk.LabelFrame(root, text="4.  Log", padding=4)
        log_frame.pack(fill="both", expand=True, padx=12, pady=6)
        self.log_text = self._tk.scrolledtext.ScrolledText(
            log_frame,
            wrap="word",
            font=("Consolas", 9),
            background="#1e1e1e",
            foreground="#d4d4d4",
            insertbackground="#d4d4d4",
            height=18,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.insert("end", "Ready. Pick an Excel file, then press Start.\n")
        self.log_text.configure(state="disabled")

    # ---------------------- Browse callbacks ----------------------
    def _browse_excel(self):
        path = self._filedialog.askopenfilename(
            title="Pick the articles Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.excel_var.set(path)
            self._log(f"Excel: {path}")

    def _browse_output_dir(self):
        path = self._filedialog.askdirectory(title="Pick the output folder")
        if path:
            self.output_dir_var.set(path)
            # also default the PDFs folder to the same place
            self.pdfs_dir_var.set(path)
            self._log(f"Output folder: {path}")

    def _browse_pdfs_dir(self):
        path = self._filedialog.askdirectory(title="Pick the folder containing the saved article PDFs")
        if path:
            self.pdfs_dir_var.set(path)
            self._log(f"PDFs folder: {path}")

    def _browse_header(self):
        path = self._filedialog.askopenfilename(
            title="Pick a header PNG (Mingtiandi masthead)",
            filetypes=[("PNG images", "*.png"), ("All files", "*.*")],
        )
        if path:
            self.header_var.set(path)
            self._log(f"Header image: {path}")

    def _open_header_layout(self):
        try:
            HeaderLayoutDialog(
                parent=self,
                header_path=self.header_var.get(),
                layout=dict(self.header_layout),
                on_apply=self._apply_header_layout,
            )
        except Exception as e:
            self._messagebox.showerror("Header layout", f"Could not open: {e}")

    def _apply_header_layout(self, new_layout: dict, new_path: str):
        self.header_layout = dict(new_layout)
        self.header_var.set(new_path)
        self._log(
            f"Header layout updated: x={new_layout['x']} y={new_layout['y']} "
            f"w={new_layout['w']} h={new_layout['h']} "
            f"lock_aspect={new_layout['lock_aspect']}"
        )

    # ---------------------- section heights (v4.3) ----------------------
    PRESETS = {
        "Compact":   {"header": 180, "title": 220, "article": 1354},
        "Balanced":  {"header": 220, "title": 260, "article": 1274},
        "Spacious":  {"header": 280, "title": 320, "article": 1154},
    }

    def _apply_preset(self, name: str):
        p = self.PRESETS[name]
        self.header_section_h.set(p["header"])
        self.title_section_h.set(p["title"])
        self.article_section_h.set(p["article"])
        self.height_preset_var.set(name)
        self._on_height_change()
        self._log(f"[*] Heights preset: {name} (header={p['header']} title={p['title']} article={p['article']})")

    def _fill_a4(self):
        """Set the three heights so the sum is exactly 1754 (A4 height),
        keeping the current header and article heights and adjusting
        the title section to absorb the remainder."""
        A4 = A4_W_PX, A4_H_PX  # (1240, 1754)
        header = self.header_section_h.get()
        article = self.article_section_h.get()
        title = max(60, A4_H_PX - header - article)
        self.title_section_h.set(title)
        self.height_preset_var.set("Custom")
        self._on_height_change()
        self._log(f"[*] Heights filled to A4: header={header} title={title} article={article} (sum={header+title+article})")

    def _default_bobby_position(self) -> float:
        """Compute the bobby_position value that puts Bobby Mak at the
        middle of the A4 page (y=877), based on the current section
        heights. If the article section starts at y=article_top, then
        bobby_position = (A4_mid - article_top) / article_area_h."""
        try:
            header = int(self.header_section_h.get())
            title = int(self.title_section_h.get())
            article = int(self.article_section_h.get())
        except Exception:
            return 0.5
        article_top = header + title
        a4_mid = A4_H_PX // 2  # 877
        bobby_y_in_section = max(0, a4_mid - article_top)
        # Clamp to a sensible range (0.05..0.95) so the article isn't
        # pushed to the very edge if the title section is very tall.
        return max(0.05, min(0.95, bobby_y_in_section / max(1, article)))

    def _on_bobby_position_change(self, *_args):
        pos = float(self.bobby_position.get())
        a4_mid = self._default_bobby_position()
        if pos <= 0.04:
            label = "Top"
        elif pos >= 0.96:
            label = "Bottom"
        elif abs(pos - 0.5) < 0.04:
            label = "Middle (sec)"
        elif abs(pos - a4_mid) < 0.04:
            label = "A4 Middle"
        elif pos < 0.5:
            label = f"{int(round(pos * 100))}% (upper)"
        else:
            label = f"{int(round(pos * 100))}% (lower)"
        try:
            self.bobby_position_label.configure(text=label)
        except Exception:
            pass

    def _on_height_change(self):
        h = self.header_section_h.get()
        t = self.title_section_h.get()
        a = self.article_section_h.get()
        total = h + t + a
        a4 = A4_H_PX
        if total == a4:
            msg = f"Sum = {total}  (fills the 1754-px A4 exactly)"
            color = "#228B22"  # green
        elif total < a4:
            msg = f"Sum = {total}  ({a4 - total} px of empty space below)"
            color = "#aa8800"  # amber
        else:
            msg = f"Sum = {total}  ({total - a4} px over the A4 — will overflow)"
            color = "#cc3333"  # red
        try:
            self.height_total_var.set(msg)
        except Exception:
            pass
        # Refresh the bobby position label too (the "A4 Middle" value
        # changes with the section heights).
        try:
            self._on_bobby_position_change()
        except Exception:
            pass

    def _resolve_header_path(self) -> Optional[str]:
        if not self.use_header_var.get():
            return None
        h = self.header_var.get().strip()
        if h and os.path.exists(h):
            return h
        if h:
            self._log(f"[!] Header file not found: {h} (running without header)")
        return None

    def _resolve_section_heights(self) -> dict:
        return {
            "header": int(self.header_section_h.get()),
            "title": int(self.title_section_h.get()),
            "article": int(self.article_section_h.get()),
            "intro_lines": int(self.intro_lines.get()),
            "bobby_position": float(self.bobby_position.get()),
        }

    def _on_mode_change(self):
        mode = self.mode_var.get()
        # PDFs folder row is meaningful only in "chop" mode
        state_pdfs = "normal" if mode == "chop" else "disabled"
        try:
            self.pdfs_dir_label.configure(state=state_pdfs)
            self.pdfs_dir_entry.configure(state=state_pdfs)
            self.pdfs_dir_btn.configure(state=state_pdfs)
        except Exception:
            pass
        # Chrome path + "show browser" are only used in auto mode
        # (we don't disable the widgets — users may want to know)
        # Update the title description
        descriptions = {
            "auto": "Auto scrape — headless Chrome will try to load every article",
            "guided": "Guided manual — Chrome opens for you, you Ctrl+P each, the tool combines",
            "chop": "Combine — pick a folder of PDFs you already saved and we'll stitch them",
        }
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Frame if False else self._ttk.Frame):  # ttk.Frame
                pass
        # Reset the second section's label frame text
        # (We use a static label here; the mode radio descriptions are the source of truth.)

    def _auto_detect_chrome(self):
        path = find_chrome_path()
        if path:
            self.chrome_var.set(path)
            self._log(f"Auto-detected Chrome: {path}")
        else:
            self._messagebox.showwarning(
                "Chrome not found",
                "No Chrome installation was auto-detected. "
                "On Windows, install Google Chrome and retry.",
            )

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _open_output(self):
        path = self.output_pdf_path or self._resolve_output_path()
        if not path:
            return
        if os.path.exists(path):
            webbrowser.open(os.path.dirname(path))
        else:
            self._messagebox.showinfo(
                "Output not found",
                f"Output PDF not found at:\n{path}\n"
                "Run a scrape first to generate it.",
            )

    # ---------------------- Log + status plumbing ----------------------
    def _log(self, msg: str):
        self.log_queue.put(msg)

    def _drain_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log_queue)

    def _set_status(self, status: str, pct: Optional[float] = None):
        self.status_var.set(status)
        if pct is not None:
            self.progress_var.set(pct * 100)

    # ---------------------- Start / stop ----------------------
    def _resolve_output_path(self) -> str:
        folder = self.output_dir_var.get().strip() or "."
        name = self.filename_var.get().strip() or "bobby_mak_quotes.pdf"
        if not name.lower().endswith(".pdf"):
            name = name + ".pdf"
        return os.path.join(folder, name)

    def _on_start(self):
        if self.worker and self.worker.is_alive():
            return
        mode = self.mode_var.get()
        out_folder = self.output_dir_var.get().strip()
        if not out_folder:
            self._messagebox.showerror(
                "Missing output folder", "Please pick where to save the PDF."
            )
            return
        os.makedirs(out_folder, exist_ok=True)
        output_pdf = self._resolve_output_path()

        if mode in ("auto", "guided"):
            excel = self.excel_var.get().strip()
            if not excel or not os.path.exists(excel):
                self._messagebox.showerror(
                    "Missing Excel", "Please pick the articles Excel file first."
                )
                return
            try:
                articles = read_articles(excel)
            except Exception as e:
                self._messagebox.showerror(
                    "Excel error", f"Could not read Excel:\n{e}"
                )
                return
            if not articles:
                self._messagebox.showerror(
                    "No articles", "The Excel has no articles in column B."
                )
                return
        else:  # chop
            articles = []

        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.cancel_evt.clear()
        self.progress_var.set(0)
        self.output_pdf_path = output_pdf

        if mode == "auto":
            self._set_status(f"Starting… ({len(articles)} articles)", 0.0)
            self._start_auto_scrape(articles, output_pdf)
        elif mode == "guided":
            self._set_status("Guided mode — opening Chrome to article 1", 0.0)
            self._start_guided(articles, output_pdf)
        else:  # chop
            self._set_status("Combining saved PDFs…", 0.0)
            self._start_chop_only(output_pdf)

    # ----- mode: auto -----
    def _start_auto_scrape(self, articles: list[dict], output_pdf: str):
        chrome_path = self.chrome_var.get().strip() or None
        headless = not self.watch_var.get()

        def worker():
            try:
                def gui_log(msg: str):
                    self._log(msg)

                def gui_prog(p: float):
                    self._set_status(f"Working… {int(p*100)}%", p)

                result = asyncio.run(
                    run_scrape_core(
                        articles=articles,
                        output_pdf=output_pdf,
                        chrome_path=chrome_path,
                        headless=headless,
                        user_agent=DEFAULT_USER_AGENT,
                        log_cb=gui_log,
                        progress_cb=gui_prog,
                        cancel_evt=self.cancel_evt,
                        header_path=self._resolve_header_path(),
                        header_layout=dict(self.header_layout),
                        section_heights=self._resolve_section_heights(),
                    )
                )
                if result:
                    self._set_status("Done ✓", 1.0)
                    self._log(f"\n[OK] Saved: {result}")
                    self.output_pdf_path = result
                else:
                    self._set_status(
                        "Stopped (no PDF written)",
                        self.progress_var.get() / 100.0,
                    )
            except Exception as e:
                self._log(f"\n[!] Fatal error: {e}")
                self._set_status("Error", None)
            finally:
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    # ----- mode: guided -----
    def _start_guided(self, articles: list[dict], output_pdf: str):
        # Create the folder where the modal will collect PDFs
        guided_dir = os.path.join(
            self.output_dir_var.get().strip() or ".", "bobby_pdfs"
        )
        os.makedirs(guided_dir, exist_ok=True)
        self.guided_pdfs_dir = guided_dir
        self._log(f"[*] Guided mode — PDFs will be saved into: {guided_dir}")
        # Open the modal; it walks the user through 19 articles
        dlg = GuidedManualDialog(
            parent=self.root,
            articles=articles,
            work_dir=guided_dir,
            log_cb=self._log,
        )
        self._set_status("Guided mode — waiting for you to save PDFs…", 0.0)
        # When the modal closes, run the combine step
        def on_modal_close():
            saved, skipped, stopped = dlg.result
            self._log(
                f"[guided] Done. saved={len(saved)}, skipped={len(skipped)}, stopped={stopped}"
            )
            if not saved:
                self._set_status(
                    "No PDFs saved — nothing to combine.",
                    self.progress_var.get() / 100.0,
                )
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")
                return
            # Run the chop pipeline in a worker thread
            def combine_worker():
                try:
                    self._set_status(
                        f"Combining {len(saved)} PDFs into the final PDF…",
                        0.05,
                    )
                    out = self._run_chop_pipeline(
                        guided_dir,
                        output_pdf,
                        log_cb=self._log,
                        progress_cb=self._set_status,
                        header_path=self._resolve_header_path(),
                        header_layout=dict(self.header_layout),
                        section_heights=self._resolve_section_heights(),
                    )
                    if out:
                        self._set_status("Done ✓", 1.0)
                        self._log(f"\n[OK] Final PDF: {out}")
                        self.output_pdf_path = out
                except Exception as e:
                    self._log(f"\n[!] Combine failed: {e}")
                    self._set_status("Combine error", None)
                finally:
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")

            self.worker = threading.Thread(target=combine_worker, daemon=True)
            self.worker.start()

        # wait_window is non-blocking in Tk; we poll
        def poll_modal():
            try:
                if dlg.win.winfo_exists():
                    self.root.after(200, poll_modal)
                else:
                    on_modal_close()
            except Exception:
                on_modal_close()
        self.root.after(200, poll_modal)

    # ----- mode: chop only -----
    def _start_chop_only(self, output_pdf: str):
        pdfs_dir = self.pdfs_dir_var.get().strip()
        if not pdfs_dir or not os.path.isdir(pdfs_dir):
            self._messagebox.showerror(
                "Missing PDFs folder",
                "Please pick a folder containing the saved article PDFs.",
            )
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            return

        def worker():
            try:
                out = self._run_chop_pipeline(
                    pdfs_dir,
                    output_pdf,
                    log_cb=self._log,
                    progress_cb=self._set_status,
                    header_path=self._resolve_header_path(),
                    header_layout=dict(self.header_layout),
                    section_heights=self._resolve_section_heights(),
                )
                if out:
                    self._set_status("Done ✓", 1.0)
                    self._log(f"\n[OK] Final PDF: {out}")
                    self.output_pdf_path = out
            except Exception as e:
                self._log(f"\n[!] Combine failed: {e}")
                self._set_status("Combine error", None)
            finally:
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _on_recombine(self):
        """Combine the PDFs from the most recent guided run (or pick a folder)."""
        folder = self.guided_pdfs_dir
        if not folder or not os.path.isdir(folder):
            folder = self._filedialog.askdirectory(
                title="Pick the folder of saved article PDFs"
            )
            if not folder:
                return
        output_pdf = self._resolve_output_path()
        if self.worker and self.worker.is_alive():
            self._log("[!] A job is already running.")
            return
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress_var.set(0)
        self._set_status("Re-combining…", 0.0)

        def worker():
            try:
                out = self._run_chop_pipeline(
                    folder,
                    output_pdf,
                    log_cb=self._log,
                    progress_cb=self._set_status,
                    header_path=self._resolve_header_path(),
                    header_layout=dict(self.header_layout),
                    section_heights=self._resolve_section_heights(),
                )
                if out:
                    self._set_status("Done ✓", 1.0)
                    self._log(f"\n[OK] Final PDF: {out}")
                    self.output_pdf_path = out
            except Exception as e:
                self._log(f"\n[!] Combine failed: {e}")
                self._set_status("Combine error", None)
            finally:
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    # ----- shared: chop pipeline (no Playwright, just PDFs) -----
    def _run_chop_pipeline(
        self,
        pdfs_dir: str,
        output_pdf: str,
        log_cb=None,
        progress_cb=None,
        header_path: Optional[str] = None,
        header_layout: Optional[dict] = None,
        section_heights: Optional[dict] = None,
    ) -> Optional[str]:
        def log(msg: str):
            if log_cb:
                log_cb(msg)
            else:
                print(msg, flush=True)

        if header_path and os.path.exists(header_path):
            log(f"[*] Using header image: {header_path}")
        else:
            log("[*] No header image — pages will have no Mingtiandi masthead")

        log(f"[*] Scanning {pdfs_dir} for PDFs…")
        pdf_files = sorted(
            os.path.join(pdfs_dir, f)
            for f in os.listdir(pdfs_dir)
            if f.lower().endswith(".pdf")
        )
        if not pdf_files:
            log(f"[!] No PDF files in {pdfs_dir}")
            return None
        log(f"[*] Found {len(pdf_files)} PDF(s)")

        a4_pages: list[Image.Image] = []
        for i, pdf_path in enumerate(pdf_files, 1):
            log(f"  [{i}/{len(pdf_files)}] {os.path.basename(pdf_path)}")
            bobby = find_bobby_mak_in_pdf(pdf_path)
            if bobby is None:
                log("      [!] No 'Bobby Mak' line found in text — page will still be cropped to A4")
            else:
                log(f"      Bobby Mak on page {bobby[0]+1}, y={bobby[1]:.0f}pt")
            a4 = make_a4_page_from_pdf(
                pdf_path, bobby,
                header_path=header_path,
                header_layout=header_layout,
                section_heights=section_heights,
            )
            if a4 is not None:
                a4_pages.append(a4)
            if progress_cb:
                progress_cb(0.1 + 0.85 * i / len(pdf_files))
        if not a4_pages:
            log("[!] No pages produced.")
            return None
        log(f"[*] Combining {len(a4_pages)} A4 pages into {output_pdf}…")
        a4_pages[0].save(
            output_pdf,
            save_all=True,
            append_images=a4_pages[1:],
            resolution=150.0,
            quality=88,
        )
        if progress_cb:
            progress_cb(1.0)
        log(f"[OK] Final PDF: {output_pdf}  ({len(a4_pages)} pages)")
        return output_pdf

    def _on_stop(self):
        if self.worker and self.worker.is_alive():
            self.cancel_evt.set()
            self._log("[!] Stop requested — finishing current article…")
            self.stop_btn.configure(state="disabled")
        # If a guided dialog is open, close it too
        for child in self.root.winfo_children():
            if isinstance(child, self._tk.Toplevel) and child.winfo_title() == "Guided Manual Save":
                child.destroy()

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            if not self._messagebox.askyesno(
                "Quit?",
                "A scrape is running. Quit anyway?",
            ):
                return
            self.cancel_evt.set()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ===========================================================================
# CLI
# ===========================================================================


def run_pdf_chop_mode(
    pdf_input_dir: str, output_pdf: str, articles: list[dict],
    header_path: Optional[str] = None,
    header_layout: Optional[dict] = None,
    section_heights: Optional[dict] = None,
) -> None:
    pdf_files = sorted(
        os.path.join(pdf_input_dir, f)
        for f in os.listdir(pdf_input_dir)
        if f.lower().endswith(".pdf")
    )
    if not pdf_files:
        print(f"[!] No PDF files in {pdf_input_dir}")
        return
    a4_pages: list[Image.Image] = []
    for pdf_path in pdf_files:
        print(f"[*] {os.path.basename(pdf_path)}")
        bobby = find_bobby_mak_in_pdf(pdf_path)
        a4 = make_a4_page_from_pdf(
            pdf_path, bobby,
            header_path=header_path,
            header_layout=header_layout,
            section_heights=section_heights,
        )
        if a4 is not None:
            a4_pages.append(a4)
    if not a4_pages:
        return
    a4_pages[0].save(
        output_pdf,
        save_all=True,
        append_images=a4_pages[1:],
        resolution=150.0,
        quality=88,
    )
    print(f"[OK] Final PDF: {output_pdf}  ({len(a4_pages)} pages)")


def parse_args(argv: list[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Mingtiandi Bobby Mak Quote PDF Builder (v4.4 — title section bridges into article body)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("--input", "-i", help="Excel with article list")
    ap.add_argument("--output", "-o", default="output.pdf", help="Final PDF path")
    ap.add_argument(
        "--mode",
        choices=["scrape", "pdf-chop"],
        default="scrape",
        help="scrape = visit live URLs; pdf-chop = process saved PDFs",
    )
    ap.add_argument("--pdf-input", help="Directory of saved PDFs (--mode pdf-chop)")
    ap.add_argument("--chrome", help="Path to a real Chrome executable")
    ap.add_argument("--no-headless", action="store_true", help="Show Chrome windows")
    ap.add_argument("--user-agent", default=DEFAULT_USER_AGENT)
    ap.add_argument(
        "--header",
        default=DEFAULT_HEADER_PNG,
        help=f"PNG to paste at the top of every A4 page (default: {DEFAULT_HEADER_PNG})",
    )
    ap.add_argument(
        "--no-header",
        action="store_true",
        help="Skip the header banner (use the raw page only)",
    )
    ap.add_argument(
        "--header-x", type=int, default=0,
        help="Header X offset in px (default: 0)"
    )
    ap.add_argument(
        "--header-y", type=int, default=0,
        help="Header Y offset in px (default: 0)"
    )
    ap.add_argument(
        "--header-w", type=int, default=A4_W_PX,
        help="Header width in px (default: 1240, full A4 width)"
    )
    ap.add_argument(
        "--header-h", type=int, default=0,
        help="Header height in px (default: 0 = auto from width + aspect)"
    )
    ap.add_argument(
        "--section-header", type=int, default=220,
        help="Height of the header bar section in px (default: 220)"
    )
    ap.add_argument(
        "--section-title", type=int, default=260,
        help="Height of the subject section in px (default: 260 — "
             "includes title, date and a few intro lines)"
    )
    ap.add_argument(
        "--section-article", type=int, default=1274,
        help="Height of the article window section in px (default: 1274)"
    )
    ap.add_argument(
        "--intro-lines", type=int, default=3,
        help="Number of body lines to bridge from the date into the "
             "scrolled-down section (default: 3)"
    )
    ap.add_argument(
        "--bobby-position", type=float, default=0.5,
        help="Where the Bobby Mak paragraph sits in the article "
             "section: 0.0 = top, 1.0 = bottom, 0.5 = middle (default: 0.5)"
    )
    ap.add_argument(
        "--gui",
        action="store_true",
        help="Force the GUI even when other args are given (rarely needed)",
    )
    return ap.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> None:
    # If no args at all, launch the GUI
    if argv is None:
        argv = sys.argv[1:]
    if not argv:
        try:
            ScraperGUI().run()
            return
        except Exception as e:
            print(f"[!] GUI failed to start ({e}); falling back to CLI.", file=sys.stderr)
            print("    Use --help for CLI usage.", file=sys.stderr)
            return

    args = parse_args(argv)
    if args.gui:
        ScraperGUI().run()
        return

    header_path: Optional[str] = None
    header_layout = dict(DEFAULT_HEADER_LAYOUT)
    section_heights = {
        "header": args.section_header,
        "title": args.section_title,
        "article": args.section_article,
        "intro_lines": args.intro_lines,
        "bobby_position": args.bobby_position,
    }
    if not args.no_header:
        if args.header and os.path.exists(args.header):
            header_path = args.header
            print(f"[*] Header image: {header_path}")
        elif args.header and args.header != DEFAULT_HEADER_PNG:
            print(f"[!] Header file not found: {args.header}")
        header_layout["x"] = args.header_x
        header_layout["y"] = args.header_y
        header_layout["w"] = args.header_w
        header_layout["h"] = args.header_h if args.header_h > 0 else None
        print(f"[*] Header layout: x={header_layout['x']} y={header_layout['y']} "
              f"w={header_layout['w']} h={header_layout['h']}")
    print(f"[*] Section heights: header={section_heights['header']} "
          f"title={section_heights['title']} article={section_heights['article']} "
          f"(sum={sum(section_heights.values())})")

    articles: list[dict] = []
    if args.input and os.path.exists(args.input):
        articles = read_articles(args.input)
        print(f"[*] Loaded {len(articles)} articles from {args.input}")

    if args.mode == "scrape":
        if not articles:
            print("[!] --input is required for scrape mode.", file=sys.stderr)
            sys.exit(2)
        chrome_path = find_chrome_path(args.chrome)
        asyncio.run(
            run_scrape_core(
                articles=articles,
                output_pdf=args.output,
                chrome_path=chrome_path,
                headless=not args.no_headless,
                user_agent=args.user_agent,
                header_path=header_path,
                header_layout=header_layout,
                section_heights=section_heights,
            )
        )
    else:
        if not args.pdf_input or not os.path.isdir(args.pdf_input):
            print("[!] --pdf-input is required for pdf-chop mode.", file=sys.stderr)
            sys.exit(2)
        run_pdf_chop_mode(args.pdf_input, args.output, articles, header_path=header_path, header_layout=header_layout, section_heights=section_heights)


if __name__ == "__main__":
    main()
