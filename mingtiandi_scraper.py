#!/usr/bin/env python3
"""
Mingtiandi Bobby Mak Quote PDF Builder (v3 — GUI edition)
=========================================================

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


def make_a4_page_from_pdf(
    pdf_path: str, bobby_mak: Optional[tuple[int, float]]
) -> Optional[Image.Image]:
    pages = [img.convert("RGB") for img in convert_from_path(pdf_path, dpi=200)]
    if not pages:
        return None
    target_idx = min(bobby_mak[0], len(pages) - 1) if bobby_mak else 0
    primary = pages[target_idx]
    pw, ph = primary.size
    a4_w, a4_h = A4_W_PX, A4_H_PX
    if ph <= a4_h:
        return _center_on_a4(primary, a4_w, a4_h)
    DPI = 200
    if bobby_mak:
        _, bobby_y_pt = bobby_mak
        bobby_y_px = int(bobby_y_pt / 72.0 * DPI)
        crop_top = max(0, bobby_y_px - ph // 3)
        crop_bottom = min(ph, crop_top + a4_h)
        if crop_bottom - crop_top < a4_h and ph >= a4_h:
            crop_bottom = ph
            crop_top = max(0, crop_bottom - a4_h)
        cropped = primary.crop((0, crop_top, pw, crop_bottom))
        return _center_on_a4(cropped, a4_w, a4_h)
    cropped = primary.crop((0, 0, pw, min(ph, a4_h)))
    return _center_on_a4(cropped, a4_w, a4_h)


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
                a4 = make_a4_page_from_pdf(raw_pdf, bobby)
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
        self.chrome_var = tk.StringVar(value=find_chrome_path() or "")
        self.filename_var = tk.StringVar(value="bobby_mak_quotes.pdf")
        self.watch_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready.")
        self.progress_var = tk.DoubleVar(value=0.0)

        self._build_ui()
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
            text="v3 — picks articles from Excel, prints each to PDF, highlights the quote",
            foreground="#666",
        ).pack(side="left", padx=12)

        # File picker section
        frm = ttk.LabelFrame(root, text="1.  Pick your inputs", padding=10)
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

        # Options section
        opt = ttk.LabelFrame(root, text="2.  Options", padding=10)
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
        log_frame = ttk.LabelFrame(root, text="3.  Log", padding=4)
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
            self._log(f"Output folder: {path}")

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
        excel = self.excel_var.get().strip()
        out_folder = self.output_dir_var.get().strip()
        if not excel or not os.path.exists(excel):
            self._messagebox.showerror(
                "Missing Excel", "Please pick the articles Excel file first."
            )
            return
        if not out_folder:
            self._messagebox.showerror(
                "Missing output folder", "Please pick where to save the PDF."
            )
            return
        os.makedirs(out_folder, exist_ok=True)
        output_pdf = self._resolve_output_path()

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

        # Lock the UI
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.cancel_evt.clear()
        self.progress_var.set(0)
        self._set_status(f"Starting… ({len(articles)} articles)", 0.0)
        self.output_pdf_path = output_pdf

        # Snapshot of settings for the worker
        chrome_path = self.chrome_var.get().strip() or None
        headless = not self.watch_var.get()
        watch = self.watch_var.get()

        def worker():
            try:
                # If "watch" is checked, run a one-arg lambda that prints to
                # the GUI in real time. Otherwise, suppress Playwright's
                # own console by redirecting to the GUI logger.
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
                    )
                )
                if result:
                    self._set_status("Done ✓", 1.0)
                    self._log(f"\n[OK] Saved: {result}")
                    self.output_pdf_path = result
                else:
                    self._set_status("Stopped (no PDF written)", self.progress_var.get() / 100.0)
            except Exception as e:
                self._log(f"\n[!] Fatal error: {e}")
                self._set_status("Error", None)
            finally:
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _on_stop(self):
        if self.worker and self.worker.is_alive():
            self.cancel_evt.set()
            self._log("[!] Stop requested — finishing current article…")
            self.stop_btn.configure(state="disabled")

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
    pdf_input_dir: str, output_pdf: str, articles: list[dict]
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
        a4 = make_a4_page_from_pdf(pdf_path, bobby)
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
        description="Mingtiandi Bobby Mak Quote PDF Builder (v3 — GUI edition)",
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
            )
        )
    else:
        if not args.pdf_input or not os.path.isdir(args.pdf_input):
            print("[!] --pdf-input is required for pdf-chop mode.", file=sys.stderr)
            sys.exit(2)
        run_pdf_chop_mode(args.pdf_input, args.output, articles)


if __name__ == "__main__":
    main()
