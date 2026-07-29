"""
screenshot_sections.py
=======================
Renders the browser-preview newsletter HTML (output/newsletter_<Month>_<Year>.html)
and saves one 2x PNG per section into "<Month>-<Year>/Section wise images",
alongside that month's "Row Data" source-photo folder.

Reads month/year from content.json, so it always targets the current month's
output file. Skips sections that are absent or hidden this month (e.g. an
empty New Openings), and clears out any stale PNGs from a previous run first
so output/sections/ never shows a section that isn't in the current issue.

Filenames follow the naming convention from Biztech's existing "flat image
per section" Outlook-safe email format (e.g. HR-INSIDER.png, new-openings.png)
rather than this project's own section keys, so they drop straight into that
email structure.

Header and Footer are NOT exported here — they stay as real HTML in the final
email (logo/text and address/social-icons respectively), matching the
reference format; neither has complex-enough CSS to need flattening.

Marketing Highlights is a special case: instead of one flat image, it's split
into a background image (title/icon/divider only) plus one image per blog
post, so each blog card can be wrapped in its own <a href> link in the final
email — a flat image can't be partially clickable.

Usage:
    python screenshot_sections.py

Run this once the browser preview from import_content.py has been approved —
these PNGs are what actually get mailed. Then upload them to the host and run
generate_simple_email.py to build the final email.
"""

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
CONTENT_JSON = BASE_DIR / "content.json"
OUTPUT_DIR = BASE_DIR / "output"

# Per-issue output folder, e.g. "July-2026/Section wise images", alongside the
# "Row Data" folder holding that month's source photos.
SECTIONS_SUBDIR = "Section wise images"

# Milliseconds allowed for the preview page to load and settle. See the note
# where it's used — remote source photos can be very large.
LOAD_TIMEOUT_MS = 180_000


def sections_dir(month: str, year: str) -> Path:
    return BASE_DIR / f"{month}-{year}" / SECTIONS_SUBDIR

# (output filename, CSS selector in template.html) — Header/Footer excluded,
# Marketing Highlights handled separately by save_marketing_highlights().
SECTIONS = [
    ("Director-desk", "section.sec-ceo"),
    ("Whistle-Blow", "section.sec-campaign"),
    ("Newly-Added-Customers", "section.sec-new-customers"),
    ("Delivery-Insights", "section.sec-delivery"),
    ("Acknowledging-Excellence", "section.sec-excellence"),
    ("HR-INSIDER", "section.sec-hr"),
    ("Rewards-Recognitions", "section.sec-rewards"),
    ("Employee-Certifications", "section.sec-certifications"),
    ("WORK-ANNIVERSARY", "section.sec-anniversaries"),
    ("NEW-ADDITION", "section.sec-new-joinees"),
    ("new-openings", "section.sec-openings"),
    ("Upcoming-Event-Announcement", "section.sec-announcements"),
]


def save_marketing_highlights(page, out_dir):
    """
    Splits .sec-marketing into:
      - Marketing-Highlights.png — title/icon/"Blog posts" divider only
      - blog-N.png               — one image per blog card (clickable in the
        final email via each post's own href, read straight from the DOM)
    Returns (bg_path, [(image_path, href), ...]) for the blog cards, in order.
    """
    section = page.query_selector("section.sec-marketing")
    if section is None or not section.is_visible():
        return None

    # bounding_box() is viewport-relative, and so is a plain page.screenshot()
    # clip — but the per-section loop above has left the page scrolled somewhere
    # arbitrary, so the clip can fall partly (or entirely) outside the viewport
    # and get silently truncated to zero height. Scroll back to the top and pair
    # the boxes with full_page=True, which makes the clip document-relative.
    page.evaluate("window.scrollTo(0, 0)")
    section_box = section.bounding_box()
    divider = page.query_selector("section.sec-marketing .sub-divider")
    divider_box = divider.bounding_box()
    bg_path = out_dir / "Marketing-Highlights.png"
    # +8px buffer: .sub-divider uses line-height:1, so its CSS bounding box
    # (what divider_box reports) is tighter than letters with descenders
    # (e.g. the "g" in "Blog") actually need — without this the crop clips them.
    divider_bottom = divider_box["y"] + divider_box["height"] + 8
    page.screenshot(
        path=str(bg_path),
        type="png",
        full_page=True,
        clip={
            "x": section_box["x"],
            "y": section_box["y"],
            "width": section_box["width"],
            "height": divider_bottom - section_box["y"],
        },
    )

    cards = page.query_selector_all("section.sec-marketing .blog-card")
    blog_images = []
    for i, card in enumerate(cards, start=1):
        card_path = out_dir / f"blog-{i}.png"
        card.screenshot(path=str(card_path), type="png")
        href = card.get_attribute("href") or ""
        blog_images.append((card_path, href))

    return bg_path, blog_images


def main():
    data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    month, year = data["month"], data["year"]
    html_path = OUTPUT_DIR / f"newsletter_{month}_{year}.html"
    if not html_path.exists():
        raise SystemExit(
            f"{html_path} not found — run 'python generate.py' first to render the browser-preview HTML."
        )

    out_dir = sections_dir(month, year)
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("*.png"):
        stale.unlink()

    with sync_playwright() as p:
        browser = p.chromium.launch()
        # device_scale_factor=2 captures at 2x pixel density (retina-quality)
        # while the images still display at their normal 800px layout width —
        # generate_simple_email.py halves the read-back PNG dimensions to
        # compensate when setting each <img>'s width/height attributes.
        page = browser.new_page(viewport={"width": 900, "height": 1200}, device_scale_factor=2)
        # Generous timeouts: every section has to be fully painted before it's
        # captured, and photos served from GitHub Pages can be tens of MB each
        # (a single 18MB event photo takes ~20s on its own), so the default
        # 30s is not nearly enough for an image-heavy issue.
        page.goto(html_path.resolve().as_uri(), timeout=LOAD_TIMEOUT_MS)
        page.wait_for_load_state("networkidle", timeout=LOAD_TIMEOUT_MS)
        page.wait_for_timeout(500)

        saved, skipped = [], []
        for name, selector in SECTIONS:
            el = page.query_selector(selector)
            if el is None or not el.is_visible():
                skipped.append(name)
                continue
            out_path = out_dir / f"{name}.png"
            el.screenshot(path=str(out_path), type="png")
            saved.append(out_path)

        marketing_result = save_marketing_highlights(page, out_dir)
        if marketing_result is None:
            skipped.append("Marketing-Highlights")
        else:
            bg_path, blog_images = marketing_result
            saved.append(bg_path)
            saved.extend(p for p, _href in blog_images)
            links_path = out_dir / "marketing_links.json"
            links_path.write_text(
                json.dumps(
                    [{"image": p.name, "href": href} for p, href in blog_images],
                    indent=2,
                ),
                encoding="utf-8",
            )

        browser.close()

    print(f"Section images -> {out_dir}")
    print(f"  Saved ({len(saved)}): {', '.join(p.stem for p in saved)}")
    if skipped:
        print(f"  Skipped, no content this month ({len(skipped)}): {', '.join(skipped)}")


if __name__ == "__main__":
    main()
