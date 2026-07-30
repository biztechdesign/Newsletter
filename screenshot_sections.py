"""
screenshot_sections.py
=======================
Renders the browser-preview newsletter HTML (<Month>-<Year>/HTML/newsletter_<Month>_<Year>.html)
and saves one PNG per section into "<Month>-<Year>/Section wise images",
alongside that month's "Row Data" source-photo folder.

Reads month/year from content.json, so it always targets the current month's
output file. Skips sections that are absent or hidden this month (e.g. an
empty New Openings), and clears out any stale PNGs from a previous run first
so the folder never shows a section that isn't in the current issue.

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
HTML_SUBDIR = "HTML"


def html_dir(month: str, year: str) -> Path:
    """Where generate.py writes the rendered newsletter for this issue."""
    return BASE_DIR / f"{month}-{year}" / HTML_SUBDIR

# Per-issue output folder, e.g. "July-2026/Section wise images", alongside the
# "Row Data" folder holding that month's source photos.
SECTIONS_SUBDIR = "Section wise images"

# Milliseconds allowed for the preview page to load and settle. See the note
# where it's used — remote source photos can be very large.
LOAD_TIMEOUT_MS = 180_000

# Pixel density the section PNGs are captured at. 1.5x keeps them sharp on a
# retina screen at a third less weight than 2x. generate_simple_email.py does
# not copy this number — it measures the density off the PNGs themselves, so
# changing it here is enough and the two cannot disagree.
CAPTURE_SCALE_FACTOR = 1.5


def sections_dir(month: str, year: str) -> Path:
    return BASE_DIR / f"{month}-{year}" / SECTIONS_SUBDIR

# (output filename, CSS selector in template.html) — Header/Footer excluded.
# Marketing Highlights and HR Insider are handled separately, by
# save_marketing_highlights() and save_hr_parts().
SECTIONS = [
    ("Director-desk", "section.sec-ceo"),
    ("Whistle-Blow", "section.sec-campaign"),
    ("Expo-Exhibition", "section.sec-expo"),
    ("Newly-Added-Customers", "section.sec-new-customers"),
    ("Delivery-Insights", "section.sec-delivery"),
    ("Acknowledging-Excellence", "section.sec-excellence"),
    ("Rewards-Recognitions", "section.sec-rewards"),
    ("NEW-ADDITION", "section.sec-new-joinees"),
    ("WORK-ANNIVERSARY", "section.sec-anniversaries"),
    ("Employee-Certifications", "section.sec-certifications"),
    ("new-openings", "section.sec-openings"),
    ("Employee-Training", "section.sec-training"),
    ("Employee-Workshop", "section.sec-workshop"),
    ("Upcoming-Event-Announcement", "section.sec-announcements"),
]


def save_hr_parts(page, out_dir):
    """
    Cut .sec-hr into one image per sub-titled block: the Wellness & HR Corner
    heading and banner, then each HR event. Saved as HR-INSIDER-1.png,
    HR-INSIDER-2.png, … in order.

    HR Insider is the one section holding several sub-titled blocks, so as a
    single flat image it grows without limit — with a full set of event photos
    it reached 2747 CSS px and 7.2MB, over half the finished email's weight.
    Its own sub-dividers are the natural seams, so each block becomes its own
    image and the email stacks them back with no visible join.

    Returns the list of paths written, or None if the section isn't in this
    issue. A section with a single block yields a single image, so nothing is
    split needlessly.
    """
    section = page.query_selector("section.sec-hr")
    if section is None or not section.is_visible():
        return None

    # Measure from the top: bounding_box() and a plain screenshot clip are both
    # viewport-relative, so a scrolled page silently truncates the capture.
    page.evaluate("window.scrollTo(0, 0)")
    box = section.bounding_box()
    dividers = page.query_selector_all("section.sec-hr .sub-divider")

    top, bottom = box["y"], box["y"] + box["height"]
    # The first divider belongs to the opening block, under the section title;
    # every later one starts a new block. GAP keeps the divider's own line from
    # being clipped by the cut above it.
    GAP = 24
    cuts = [top]
    for divider in dividers[1:]:
        cuts.append(max(top, divider.bounding_box()["y"] - GAP))
    cuts.append(bottom)

    paths = []
    for i in range(len(cuts) - 1):
        height = cuts[i + 1] - cuts[i]
        if height <= 1:
            continue
        path = out_dir / f"HR-INSIDER-{len(paths) + 1}.png"
        page.screenshot(
            path=str(path),
            type="png",
            full_page=True,
            clip={"x": box["x"], "y": cuts[i], "width": box["width"], "height": height},
        )
        paths.append(path)
    return paths


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
    html_path = html_dir(month, year) / f"newsletter_{month}_{year}.html"
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
        # Captured at 1.5x pixel density, so the PNGs stay sharp on a retina
        # screen without the file size 2x costs. The email still displays them
        # at their true layout size (800px wide): generate_simple_email.py
        # divides the read-back pixel dimensions by the same factor when it
        # sets each <img>'s width/height, so keep the two in step.
        page = browser.new_page(
            viewport={"width": 900, "height": 1200},
            device_scale_factor=CAPTURE_SCALE_FACTOR,
        )
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

        hr_parts = save_hr_parts(page, out_dir)
        if hr_parts is None:
            skipped.append("HR-INSIDER")
        else:
            saved.extend(hr_parts)

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
