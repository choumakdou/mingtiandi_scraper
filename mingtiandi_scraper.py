#!/usr/bin/env python3
"""
Mingtiandi Bobby Mak Quote PDF Builder (v4.1 — GUI + header restoration)
====================================================================

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
from PIL import Image
from pdf2image import convert_from_path
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


def make_a4_page_from_pdf(
    pdf_path: str,
    bobby_mak: Optional[tuple[int, float]],
    header_path: Optional[str] = None,
) -> Optional[Image.Image]:
    """Build one A4 page from a PDF.

    Steps:
      1. Convert the Bobby-Mak page of the PDF to a 200-DPI image.
      2. Skip the partial header at the top of the PDF (the one Chrome's
         @media print CSS leaves behind — social icons + nav menu but no
         logo).
      3. Optionally focus the crop around the Bobby Mak line if the
         article is longer than one A4 page.
      4. Paste the supplied header image at the top of a white A4 canvas
         (so the page looks like the live site), and drop the article
         content below it.
    """
    pages = [img.convert("RGB") for img in convert_from_path(pdf_path, dpi=200)]
    if not pages:
        return None
    target_idx = min(bobby_mak[0], len(pages) - 1) if bobby_mak else 0
    primary = pages[target_idx]
    pw, ph = primary.size
    a4_w, a4_h = A4_W_PX, A4_H_PX

    # Heuristic: the print-CSS-bleed header (social icons + nav menu but
    # no logo) takes roughly the top 30% of an A4 page at 200 DPI.
    partial_header_h = int(ph * 0.30)
    content_y0 = partial_header_h
    content_h_total = ph - content_y0

    # Apply Bobby-Mak focus within the content area if it's taller than
    # what fits below the header.
    DPI = 200
    header_h_on_a4 = 0
    if header_path and os.path.exists(header_path):
        try:
            with Image.open(header_path) as h:
                header_h_on_a4 = int(h.height * (a4_w / h.width))
        except Exception:
            header_h_on_a4 = 0
    article_budget = a4_h - header_h_on_a4

    crop_offset_y = 0  # y in the *original* page where article_img starts
    if bobby_mak and content_h_total > article_budget:
        _, bobby_y_pt = bobby_mak
        bobby_y_px = int(bobby_y_pt / 72.0 * DPI) - content_y0
        crop_top = max(0, bobby_y_px - content_h_total // 3)
        crop_bottom = min(content_h_total, crop_top + article_budget)
        if crop_bottom - crop_top < article_budget and content_h_total >= article_budget:
            crop_bottom = content_h_total
            crop_top = max(0, crop_bottom - article_budget)
        article_img = primary.crop(
            (0, content_y0 + crop_top, pw, content_y0 + crop_bottom)
        )
        crop_offset_y = content_y0 + crop_top
    elif content_h_total <= article_budget:
        article_img = primary.crop((0, content_y0, pw, ph))
        crop_offset_y = content_y0
    else:
        # Long article, no Bobby Mak found — just take the top of the content
        article_img = primary.crop((0, content_y0, pw, content_y0 + article_budget))
        crop_offset_y = content_y0

    # Yellow highlight on the Bobby Mak paragraph (chop-mode only — the
    # auto-scrape mode injects CSS for this). Find the paragraph's start
    # in the cropped image and paint a yellow rectangle with a gold left
    # border, matching the auto-scrape CSS.
    if bobby_mak:
        from PIL import ImageDraw
        _, bobby_y_pt = bobby_mak
        bobby_y_px_in_cropped = int(bobby_y_pt / 72.0 * DPI) - crop_offset_y
        para_height_px = 200  # default fallback
        try:
            with pdfplumber.open(pdf_path) as ppdf:
                ppage = ppdf.pages[min(bobby_mak[0], len(ppdf.pages) - 1)]
                lines = ppage.extract_text_lines() or []
                # Find Bobby Mak line
                bobby_line = None
                for ln in lines:
                    if "Bobby Mak" in (ln.get("text") or "") and len(ln.get("text") or "") < 1200:
                        bobby_line = ln
                        break
                if bobby_line is not None:
                    bobby_top = float(bobby_line.get("top", 0))
                    # All lines after Bobby Mak
                    post_lines = [ln for ln in lines if float(ln.get("top", 0)) >= bobby_top]
                    # Typical line height = the gap between the first two
                    # post-Bobby lines (assumes body text is uniform)
                    if len(post_lines) >= 2:
                        typical_gap = (
                            float(post_lines[1].get("top", 0))
                            - float(post_lines[0].get("top", 0))
                        )
                    else:
                        typical_gap = 18
                    # Find the first gap > 1.6x typical_gap = paragraph end
                    para_end_pt = bobby_top + 4 * typical_gap
                    for i in range(1, len(post_lines)):
                        prev = post_lines[i - 1]
                        cur = post_lines[i]
                        prev_top = float(prev.get("top", 0))
                        cur_top = float(cur.get("top", 0))
                        if cur_top - prev_top > typical_gap * 1.6:
                            para_end_pt = prev_top + typical_gap
                            break
                    else:
                        # Never broke — paragraph goes to the end of the page
                        if post_lines:
                            last = post_lines[-1]
                            para_end_pt = float(last.get("bottom", last.get("top", 0) + typical_gap))
                    para_height_px = int((para_end_pt - bobby_top) / 72.0 * DPI) + 4
        except Exception:
            pass
        highlight_y0 = max(0, bobby_y_px_in_cropped - 6)
        highlight_y1 = min(article_img.height, bobby_y_px_in_cropped + para_height_px)
        if highlight_y1 > highlight_y0:
            # Alpha-composite the highlight so the text shows through
            article_rgba = article_img.convert("RGBA")
            overlay = Image.new("RGBA", article_img.size, (0, 0, 0, 0))
            odraw = ImageDraw.Draw(overlay)
            odraw.rectangle(
                [(0, highlight_y0), (article_img.width, highlight_y1)],
                fill=(255, 245, 157, 220),  # #fff59d, ~86% opacity
            )
            odraw.rectangle(
                [(0, highlight_y0), (4, highlight_y1)],
                fill=(246, 192, 38, 255),  # #f6c026, full opacity
            )
            article_rgba = Image.alpha_composite(article_rgba, overlay)
            article_img = article_rgba.convert("RGB")

    # Build the A4 canvas
    canvas = Image.new("RGB", (a4_w, a4_h), "white")
    if header_h_on_a4 > 0 and header_path and os.path.exists(header_path):
        try:
            header = Image.open(header_path).convert("RGB")
            header = header.resize((a4_w, header_h_on_a4), Image.LANCZOS)
            canvas.paste(header, (0, 0))
        except Exception as e:
            print(f"    [!] Header paste failed: {e}", flush=True)
            header_h_on_a4 = 0

    # Fit the article content into the space below the header
    content_h = a4_h - header_h_on_a4
    a_w, a_h = article_img.size
    if a_h > 0 and a_w > 0:
        scale = min(a4_w / a_w, content_h / a_h)
        new_w = max(1, int(a_w * scale))
        new_h = max(1, int(a_h * scale))
        resized = article_img.resize((new_w, new_h), Image.LANCZOS)
        x = (a4_w - new_w) // 2
        y = header_h_on_a4 + max(0, (content_h - new_h) // 2)
        canvas.paste(resized, (x, y))
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
                a4 = make_a4_page_from_pdf(raw_pdf, bobby, header_path=header_path)
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
            text="v4.1 — pick a mode, then Start",
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

    def _resolve_header_path(self) -> Optional[str]:
        if not self.use_header_var.get():
            return None
        h = self.header_var.get().strip()
        if h and os.path.exists(h):
            return h
        if h:
            self._log(f"[!] Header file not found: {h} (running without header)")
        return None

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
            a4 = make_a4_page_from_pdf(pdf_path, bobby, header_path=header_path)
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
        a4 = make_a4_page_from_pdf(pdf_path, bobby, header_path=header_path)
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
        description="Mingtiandi Bobby Mak Quote PDF Builder (v4.1 — GUI + header restoration)",
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
    if not args.no_header:
        if args.header and os.path.exists(args.header):
            header_path = args.header
            print(f"[*] Header image: {header_path}")
        elif args.header and args.header != DEFAULT_HEADER_PNG:
            print(f"[!] Header file not found: {args.header}")

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
            )
        )
    else:
        if not args.pdf_input or not os.path.isdir(args.pdf_input):
            print("[!] --pdf-input is required for pdf-chop mode.", file=sys.stderr)
            sys.exit(2)
        run_pdf_chop_mode(args.pdf_input, args.output, articles, header_path=header_path)


if __name__ == "__main__":
    main()
