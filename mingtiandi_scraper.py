#!/usr/bin/env python3
"""
Mingtiandi Bobby Mak Quote PDF Builder (print-PDF edition)
==========================================================

Opens each Mingtiandi article in your real Chrome, exports the page to
PDF via Chrome's own print engine (so output is independent of window
size and looks like a clean printable article), then post-processes each
PDF to:
  - remove ads / sidebar / footer
  - find the paragraph that quotes Bobby Mak
  - crop the article to a single A4 portrait page that contains the
    title, the highlighted Bobby Mak paragraph, and 1-2 paragraphs of
    context on each side

Everything is bundled into a single multi-page PDF — one A4 per article.

WHY page.pdf() INSTEAD OF SCREENSHOTS
-------------------------------------
A screenshot is a snapshot of pixels on screen. Different window sizes
=> different screenshots. `page.pdf()` uses Chrome's print engine, so
the output is the same regardless of viewport. It also respects the
page's @media print stylesheet, so ads/sidebars often disappear on
their own.

USAGE
-----
    # Default end-to-end: scrape + process into one PDF
    python mingtiandi_scraper.py --input articles.xlsx --output output.pdf

    # Use a specific Chrome install
    python mingtiandi_scraper.py --input articles.xlsx --output output.pdf \
        --chrome "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

    # You already saved the 19 article PDFs (Print -> Save as PDF in
    # Chrome). Skip the scrape step and just process them.
    python mingtiandi_scraper.py --mode pdf-chop \
        --pdf-input ./saved_pdfs/ --output output.pdf

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
import shutil
import sys
from pathlib import Path
from typing import Optional

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

# CSS injected into the page BEFORE printing to strip ads/footers/etc.
# These are conservative defaults. Add/remove to taste.
JUNK_CSS = """
/* Hide common ad / sidebar / footer blocks before printing */
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

# Highlight styling — applied via classList, not inline, so it survives
# page.pdf() which uses the @media print engine.
HIGHLIGHT_CSS = """
.mak-quote-highlight {
    background-color: #fff59d !important;
    padding: 10px 14px !important;
    border-left: 3px solid #f6c026 !important;
    margin: 14px 0 !important;
    box-shadow: 0 0 0 6px #fff59d !important;
}
"""

# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------


def read_articles(excel_path: str) -> list[dict]:
    """Read the Excel file. Returns a list of {title, date, url} dicts."""
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
    if not articles:
        print(f"[!] No articles found in {excel_path}", file=sys.stderr)
    return articles


# ---------------------------------------------------------------------------
# Chrome discovery
# ---------------------------------------------------------------------------


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
            print(f"[*] Using Chrome: {chrome_path}")
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
    print("[*] Using bundled Chromium (Cloudflare will likely block this).")
    return browser


# ---------------------------------------------------------------------------
# Page helpers
# ---------------------------------------------------------------------------


async def wait_for_real_article(page: Page, timeout: int = 120) -> bool:
    print("    waiting for Cloudflare to clear…", end="", flush=True)
    deadline = asyncio.get_event_loop().time() + timeout
    last_title = ""
    while asyncio.get_event_loop().time() < deadline:
        try:
            title = (await page.title()) or ""
            if title != last_title:
                print(f" '{title[:40]}'", end="", flush=True)
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
                    print(" ✓")
                    return True
        except Exception:
            pass
        await asyncio.sleep(2)
    print(" ✗ (timed out)")
    return False


async def inject_print_prep(page: Page) -> bool:
    """Inject the highlight CSS, the junk-hiding CSS, and tag the Bobby
    Mak paragraph. Returns True if the paragraph was found."""
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


# ---------------------------------------------------------------------------
# Scrape one article -> raw PDF
# ---------------------------------------------------------------------------


async def print_article_to_pdf(
    browser: Browser,
    art: dict,
    index: int,
    total: int,
    out_dir: str,
    user_agent: str,
) -> Optional[str]:
    """Open the article, wait for the challenge, hide ads, highlight
    the Bobby Mak paragraph, then print to PDF. Returns the raw PDF path."""
    print(f"\n[{index}/{total}] {art['title']}")
    print(f"    url: {art['url']}")
    # Fresh context = fresh cookies => resets the 5-articles-per-month meter
    context = await browser.new_context(
        viewport={"width": 1366, "height": 768},  # "natural" laptop size
        user_agent=user_agent,
        locale="en-US",
        timezone_id="Asia/Hong_Kong",
    )
    page = await context.new_page()
    raw_pdf = os.path.join(out_dir, f"article_{index:02d}_raw.pdf")
    try:
        await page.goto(art["url"], wait_until="domcontentloaded", timeout=60000)
        if not await wait_for_real_article(page, timeout=120):
            print("    [!] Cloudflare never cleared. Skipping.")
            return None
        ok = await inject_print_prep(page)
        if not ok:
            print("    [!] No 'Bobby Mak' paragraph found.")
        # Let CSS settle
        await page.wait_for_timeout(800)
        # Print to PDF — viewport-independent, uses Chrome's print engine
        await page.pdf(
            path=raw_pdf,
            format="A4",
            print_background=True,
            margin={"top": "10mm", "right": "10mm", "bottom": "10mm", "left": "10mm"},
            prefer_css_page_size=True,
        )
        print(f"    raw PDF: {raw_pdf}")
        return raw_pdf
    except Exception as e:
        print(f"    [!] Error: {e}")
        return None
    finally:
        await context.close()


# ---------------------------------------------------------------------------
# Process raw PDF -> single A4 portrait image of the article
# ---------------------------------------------------------------------------


def find_bobby_mak_in_pdf(pdf_path: str) -> Optional[tuple[int, float]]:
    """Return (page_index, y_top_in_points) of the line containing
    'Bobby Mak', or None if not found."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                for line in page.extract_text_lines():
                    txt = (line.get("text") or "").strip()
                    if "Bobby Mak" in txt and len(txt) < 1200:
                        return page_idx, float(line.get("top", 0))
    except Exception as e:
        print(f"    [!] pdfplumber failed on {pdf_path}: {e}")
    return None


def pdf_pages_to_images(pdf_path: str, dpi: int = 200) -> list[Image.Image]:
    """Render every page of the PDF to a PIL Image."""
    return [img.convert("RGB") for img in convert_from_path(pdf_path, dpi=dpi)]


def make_a4_page_from_pdf(
    pdf_path: str,
    bobby_mak: Optional[tuple[int, float]],
) -> Optional[Image.Image]:
    """Render the PDF and return a single A4-portrait image. If the
    article is too tall to fit one A4, the result focuses on the
    Bobby Mak paragraph plus a window of context above and below."""
    pages = pdf_pages_to_images(pdf_path, dpi=200)
    if not pages:
        return None

    # If we know the page+line for Bobby Mak, use that page as the
    # primary canvas; otherwise the first page.
    if bobby_mak:
        target_idx = min(bobby_mak[0], len(pages) - 1)
    else:
        target_idx = 0
    primary = pages[target_idx]
    pw, ph = primary.size

    # If the primary page alone already fits A4-ish, use it directly.
    a4_w, a4_h = A4_W_PX, A4_H_PX
    if ph <= a4_h:
        return _center_on_a4(primary, a4_w, a4_h)

    # Otherwise: crop around the Bobby Mak line. We use the PDF
    # points and convert to image pixels via the same DPI.
    DPI = 200
    if bobby_mak:
        _, bobby_y_pt = bobby_mak
        # pdfplumber "top" is from the top of the page in points.
        bobby_y_px = int(bobby_y_pt / 72.0 * DPI)
        # Want roughly 1/3 above, 2/3 below the quote.
        crop_top = max(0, bobby_y_px - ph // 3)
        crop_bottom = min(ph, crop_top + a4_h)
        if crop_bottom - crop_top < a4_h and ph >= a4_h:
            # Anchor at the bottom
            crop_bottom = ph
            crop_top = max(0, crop_bottom - a4_h)
        cropped = primary.crop((0, crop_top, pw, crop_bottom))
        return _center_on_a4(cropped, a4_w, a4_h)

    # No Bobby Mak found — just use the top of the article page.
    cropped = primary.crop((0, 0, pw, min(ph, a4_h)))
    return _center_on_a4(cropped, a4_w, a4_h)


def _center_on_a4(img: Image.Image, a4_w: int, a4_h: int) -> Image.Image:
    """Resize to fit inside A4, then center on a white A4 canvas."""
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


# ---------------------------------------------------------------------------
# End-to-end flow
# ---------------------------------------------------------------------------


async def run_scrape_mode(
    articles: list[dict],
    output_pdf: str,
    chrome_path: Optional[str],
    headless: bool,
    user_agent: str,
) -> None:
    work_dir = os.path.join(os.path.dirname(output_pdf) or ".", "_scratch")
    os.makedirs(work_dir, exist_ok=True)

    async with async_playwright() as p:
        browser = await launch_browser(p, chrome_path, headless=headless)
        try:
            a4_pages: list[Image.Image] = []
            for i, art in enumerate(articles, 1):
                raw_pdf = await print_article_to_pdf(
                    browser, art, i, len(articles), work_dir, user_agent
                )
                if not raw_pdf:
                    continue
                # Find Bobby Mak and turn the raw PDF into a single A4 image
                bobby = find_bobby_mak_in_pdf(raw_pdf)
                if bobby is None:
                    print("    [!] Could not locate 'Bobby Mak' in raw PDF text.")
                else:
                    print(
                        f"    Bobby Mak on page {bobby[0]+1}, y={bobby[1]:.0f}pt"
                    )
                a4 = make_a4_page_from_pdf(raw_pdf, bobby)
                if a4 is not None:
                    a4_pages.append(a4)
                    png_out = os.path.join(work_dir, f"article_{i:02d}.png")
                    a4.save(png_out, "PNG", optimize=True)
                    print(f"    a4 page: {png_out}")
            if not a4_pages:
                print("[!] No pages to combine.", file=sys.stderr)
                return
            a4_pages[0].save(
                output_pdf,
                save_all=True,
                append_images=a4_pages[1:],
                resolution=150.0,
                quality=88,
            )
            print(f"\n[✓] Final PDF: {output_pdf}  ({len(a4_pages)} pages)")
        finally:
            await browser.close()


# ---------------------------------------------------------------------------
# PDF-chop mode (you saved the 19 PDFs from Chrome's Print dialog yourself)
# ---------------------------------------------------------------------------


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
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"[*] {os.path.basename(pdf_path)}")
        bobby = find_bobby_mak_in_pdf(pdf_path)
        if bobby is None:
            print("    [!] Bobby Mak not found by text-search; using page-1.")
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
    print(f"\n[✓] Final PDF: {output_pdf}  ({len(a4_pages)} pages)")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Mingtiandi Bobby Mak Quote PDF Builder (print-PDF edition)",
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
    ap.add_argument(
        "--pdf-input",
        help="Directory of saved PDFs (--mode pdf-chop)",
    )
    ap.add_argument("--chrome", help="Path to a real Chrome executable")
    ap.add_argument("--no-headless", action="store_true", help="Show Chrome windows")
    ap.add_argument("--user-agent", default=DEFAULT_USER_AGENT)
    return ap.parse_args()


def main() -> None:
    args = parse_args()

    articles: list[dict] = []
    if args.input and os.path.exists(args.input):
        articles = read_articles(args.input)
        print(f"[*] Loaded {len(articles)} articles from {args.input}")

    if args.mode == "scrape":
        if not articles:
            print("[!] --input is required for scrape mode.", file=sys.stderr)
            sys.exit(2)
        chrome_path = find_chrome_path(args.chrome)
        if not chrome_path:
            print(
                "[!] No Chrome found. Falling back to bundled Chromium — Cloudflare "
                "will likely block this. Pass --chrome to specify one explicitly.",
                file=sys.stderr,
            )
        asyncio.run(
            run_scrape_mode(
                articles,
                args.output,
                chrome_path,
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
