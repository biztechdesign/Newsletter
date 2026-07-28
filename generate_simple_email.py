"""
generate_simple_email.py
=========================
Builds an Outlook-safe newsletter email out of the flat per-section PNGs in
output/sections/ (produced by screenshot_sections.py), instead of the fully
CSS-styled template_email.html — which renders correctly in Gmail but breaks
in Outlook desktop (gradients, border-radius, absolute positioning, etc. are
all unsupported there). Flattening each section to an image sidesteps that
entirely; only Header, Footer, and the Marketing Highlights blog cards +
"View All" button stay as real HTML (simple enough to be Outlook-safe, and
the blog cards need to stay clickable, which a flat image can't do for just
part of itself).

IMPORTANT: this reads image dimensions from the LOCAL files in
output/sections/, but writes URLs pointing at wherever you've uploaded those
*same* files on your own server (see HOSTED_SECTIONS_BASE_URL below) — you
must upload the current output/sections/*.png files there yourself before
sending, this script does not do that upload.

Usage:
    python generate_simple_email.py

Run this after screenshot_sections.py (or just python import_content.py,
which already calls screenshot_sections.py) has produced this month's PNGs.
"""

import json
from pathlib import Path

from jinja2 import Environment, FileSystemLoader
from PIL import Image

BASE_DIR = Path(__file__).parent
CONTENT_JSON = BASE_DIR / "content.json"
OUTPUT_DIR = BASE_DIR / "output"
SECTIONS_DIR = OUTPUT_DIR / "sections"
ASSETS_DIR = BASE_DIR / "assets"

GITHUB_PAGES_URL = "https://biztechdesign.github.io/Newsletter/"  # logo + social icons only

# Where you upload output/sections/*.png each month — update if the folder
# or host ever changes. See output/sections/UPLOAD_ME.txt for the exact
# current file list to upload.
HOSTED_SECTIONS_BASE_URL = "https://w.indiaondesk.com/biztech-insider/Newsletter/"

# (filename stem in output/sections/, alt text) — canonical newsletter order,
# excluding Header/Footer (real HTML) and Marketing Highlights (special-cased).
SECTIONS_BEFORE_MARKETING = [
    ("Director-desk", "From CEO's Desk"),
    ("Whistle-Blow", "Campaign Banner"),
    ("Newly-Added-Customers", "Newly Added Customers"),
    ("Delivery-Insights", "Delivery Insights"),
    ("Acknowledging-Excellence", "Acknowledging Excellence"),
    ("HR-INSIDER", "HR Insider"),
    ("Rewards-Recognitions", "Rewards & Recognitions"),
    ("Employee-Certifications", "Employee Certifications"),
    ("WORK-ANNIVERSARY", "Work Anniversaries"),
    ("NEW-ADDITION", "New Addition to the Biztech Family"),
    ("new-openings", "New Openings"),
]
SECTIONS_AFTER_MARKETING = [
    ("Upcoming-Event-Announcement", "Upcoming Event Announcement"),
]


# Must match screenshot_sections.py's device_scale_factor — the PNGs are
# captured at 2x pixel density for retina-quality images, so their actual
# pixel dimensions need to be halved back down to CSS/layout pixels here.
CAPTURE_SCALE_FACTOR = 2


def png_size(path: Path) -> tuple[int, int]:
    """Returns (width, height) in CSS pixels — i.e. already divided back down
    from the PNG's actual (2x) pixel dimensions."""
    with Image.open(path) as im:
        w, h = im.size
        return w // CAPTURE_SCALE_FACTOR, h // CAPTURE_SCALE_FACTOR


def resolve_sections(pairs):
    resolved = []
    for stem, alt in pairs:
        path = SECTIONS_DIR / f"{stem}.png"
        if not path.exists():
            continue
        _w, h = png_size(path)
        resolved.append({
            "url": f"{HOSTED_SECTIONS_BASE_URL.rstrip('/')}/{stem}.png",
            "alt": alt,
            "height": h,
        })
    return resolved


def resolve_marketing():
    bg_path = SECTIONS_DIR / "Marketing-Highlights.png"
    if not bg_path.exists():
        return None
    _w, bg_h = png_size(bg_path)

    links_path = SECTIONS_DIR / "marketing_links.json"
    links = json.loads(links_path.read_text(encoding="utf-8")) if links_path.exists() else []
    href_by_name = {entry["image"]: entry["href"] for entry in links}

    blogs = []
    blog_paths = sorted(SECTIONS_DIR.glob("blog-*.png"), key=lambda p: p.name)
    for path in blog_paths:
        _w, h = png_size(path)
        blogs.append({
            "url": f"{HOSTED_SECTIONS_BASE_URL.rstrip('/')}/{path.name}",
            "height": h,
            "href": href_by_name.get(path.name, "#"),
        })

    data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    return {
        "bg_url": f"{HOSTED_SECTIONS_BASE_URL.rstrip('/')}/Marketing-Highlights.png",
        "bg_height": bg_h,
        "blogs": blogs,
        "view_all_url": data["marketing"]["view_all_url"],
    }


def main():
    data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    month, year = data["month"], data["year"]

    sections_before = resolve_sections(SECTIONS_BEFORE_MARKETING)
    sections_after = resolve_sections(SECTIONS_AFTER_MARKETING)
    marketing = resolve_marketing()

    env = Environment(loader=FileSystemLoader(str(BASE_DIR)))
    template = env.get_template("template_simple_email.html")
    html = template.render(
        month=month,
        year=year,
        logo_img=f"{GITHUB_PAGES_URL}assets/logo.png",
        icon_social_linkedin=f"{GITHUB_PAGES_URL}assets/icon_social_linkedin.png",
        icon_social_facebook=f"{GITHUB_PAGES_URL}assets/icon_social_facebook.png",
        icon_social_youtube=f"{GITHUB_PAGES_URL}assets/icon_social_youtube.png",
        icon_social_instagram=f"{GITHUB_PAGES_URL}assets/icon_social_instagram.png",
        icon_social_twitter=f"{GITHUB_PAGES_URL}assets/icon_social_twitter.png",
        icon_social_dribbble=f"{GITHUB_PAGES_URL}assets/icon_social_dribbble.png",
        footer=data["footer"],
        sections_before_marketing=sections_before,
        sections_after_marketing=sections_after,
        marketing=marketing,
    )

    OUTPUT_DIR.mkdir(exist_ok=True)
    out_file = OUTPUT_DIR / f"newsletter_{month}_{year}_simple_email.html"
    out_file.write_text(html, encoding="utf-8")

    all_stems = [s for s, _ in SECTIONS_BEFORE_MARKETING] + [s for s, _ in SECTIONS_AFTER_MARKETING]
    included = [s for s in all_stems if (SECTIONS_DIR / f"{s}.png").exists()]
    skipped = [s for s in all_stems if s not in included]

    print(f"\nSimple email generated: {out_file}")
    print(f"  Sections included: {' | '.join(included)}{' | Marketing Highlights' if marketing else ''}")
    if skipped or not marketing:
        missing = skipped + ([] if marketing else ["Marketing Highlights"])
        print(f"  Sections skipped (no PNG found): {' | '.join(missing)}")
    print(f"  Image host: {HOSTED_SECTIONS_BASE_URL}")
    print(f"  NOTE: Make sure output/sections/*.png are uploaded there before sending.")
    print(f"  Open in browser: file:///{out_file.as_posix()}\n")


if __name__ == "__main__":
    main()
