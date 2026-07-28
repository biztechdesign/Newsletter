"""
generate_simple_email.py
=========================
Builds the final Gmail- and Outlook-safe newsletter email out of the flat
per-section PNGs in output/sections/ (produced by screenshot_sections.py),
instead of the fully CSS-styled template_email.html — which renders correctly
in Gmail but breaks in Outlook desktop (gradients, border-radius, absolute
positioning, etc. are all unsupported there). Flattening each section to an
image sidesteps that entirely; only Header, Footer, and the Marketing
Highlights blog cards + "View All" button stay as real HTML (simple enough to
be Outlook-safe, and the blog cards need to stay clickable, which a flat image
can't do for just part of itself).

The markup follows the approved reference layout (Roboto, 800px tables,
.mobile_table responsive rules, teal footer) — see template_simple_email.html.

IMPORTANT: this reads image dimensions from the LOCAL files in
output/sections/, but writes URLs pointing at wherever you've uploaded those
*same* files on your own server (see HOSTED_SECTIONS_BASE below) — you must
upload the current output/sections/*.png files there yourself before sending,
this script does not do that upload.

Each issue gets its OWN folder on the server, e.g.
    https://w.indiaondesk.com/biztech-insider/January_2026/
so uploading a new month never breaks the previously sent emails.

Usage:
    python generate_simple_email.py

Run this after screenshot_sections.py has produced this month's PNGs.
"""

import json
from pathlib import Path

from jinja2 import Environment, FileSystemLoader
from PIL import Image

BASE_DIR = Path(__file__).parent
CONTENT_JSON = BASE_DIR / "content.json"
OUTPUT_DIR = BASE_DIR / "output"
SECTIONS_DIR = OUTPUT_DIR / "sections"

# Where you upload output/sections/*.png each month. The {month}/{year}
# placeholders give every issue its own folder, so re-uploading next month
# doesn't overwrite the images an already-sent email still points at.
HOSTED_SECTIONS_BASE = "https://w.indiaondesk.com/biztech-insider/{month}_{year}/"

# Long-lived email assets (logo, social icons, View All button) — these live on
# the CMS host and are shared by every issue, so they're not regenerated.
CMS_ASSETS = "https://cms.indiaondesk.com/biz_img/biztech-tech-frontier/"
LOGO_URL = CMS_ASSETS + "june-2024/logo.png"
VIEW_ALL_URL = CMS_ASSETS + "view_all_btn.png"
SOCIAL_ICONS = {
    "linkedin":  CMS_ASSETS + "linkedin.png",
    "facebook":  CMS_ASSETS + "fb.png",
    "youtube":   CMS_ASSETS + "youtube.png",
    "instagram": CMS_ASSETS + "insta.png",
    "twitter":   CMS_ASSETS + "Subtract.png",
    "dribbble":  CMS_ASSETS + "DRIBBBLE.png",
}

# (filename stem in output/sections/, alt text) — canonical newsletter order,
# matching template.html so the approved browser preview and the final email
# always show the sections in the same sequence. Header/Footer are excluded
# (real HTML) and Marketing Highlights is special-cased.
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

# Rendered width of a blog card in the final email: the 800px table has 50px
# side padding and each of the two cells adds 15px either side, leaving
# (800 - 100) / 2 - 30 = 320px. The captured card is wider than that, so its
# height is scaled by the same ratio to keep the aspect ratio intact.
BLOG_DISPLAY_WIDTH = 320


def png_size(path: Path) -> tuple[int, int]:
    """Returns (width, height) in CSS pixels — i.e. already divided back down
    from the PNG's actual (2x) pixel dimensions."""
    with Image.open(path) as im:
        w, h = im.size
        return w // CAPTURE_SCALE_FACTOR, h // CAPTURE_SCALE_FACTOR


def hosted_base(month: str, year: str) -> str:
    return HOSTED_SECTIONS_BASE.format(month=month, year=year).rstrip("/")


def resolve_sections(pairs, base_url):
    resolved = []
    for stem, alt in pairs:
        path = SECTIONS_DIR / f"{stem}.png"
        if not path.exists():
            continue
        _w, h = png_size(path)
        resolved.append({
            "url": f"{base_url}/{stem}.png",
            "alt": alt,
            "height": h,
        })
    return resolved


def resolve_marketing(base_url, data):
    bg_path = SECTIONS_DIR / "Marketing-Highlights.png"
    if not bg_path.exists():
        return None
    _w, bg_h = png_size(bg_path)

    links_path = SECTIONS_DIR / "marketing_links.json"
    links = json.loads(links_path.read_text(encoding="utf-8")) if links_path.exists() else []
    href_by_name = {entry["image"]: entry["href"] for entry in links}

    # Blog titles live in content.json, keyed by post URL — the same href the
    # card was captured with — so the <a title> and <img alt> stay meaningful.
    posts = data.get("marketing", {}).get("blog_posts", [])
    title_by_url = {p.get("url", ""): p.get("title", "") for p in posts}

    blogs = []
    blog_paths = sorted(SECTIONS_DIR.glob("blog-*.png"), key=lambda p: p.name)
    for i, path in enumerate(blog_paths):
        w, h = png_size(path)
        href = href_by_name.get(path.name, "#")
        # Scale the captured card down to its rendered width, keeping aspect.
        display_h = round(h * BLOG_DISPLAY_WIDTH / w) if w else h
        title = title_by_url.get(href) or (posts[i].get("title", "") if i < len(posts) else "")
        blogs.append({
            "url": f"{base_url}/{path.name}",
            "width": BLOG_DISPLAY_WIDTH,
            "height": display_h,
            "href": href,
            "title": title,
        })

    return {
        "bg_url": f"{base_url}/Marketing-Highlights.png",
        "bg_height": bg_h,
        "blogs": blogs,
        "view_all_url": data["marketing"]["view_all_url"],
    }


def main():
    data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    month, year = data["month"], data["year"]
    base_url = hosted_base(month, year)

    sections_before = resolve_sections(SECTIONS_BEFORE_MARKETING, base_url)
    sections_after = resolve_sections(SECTIONS_AFTER_MARKETING, base_url)
    marketing = resolve_marketing(base_url, data)

    env = Environment(loader=FileSystemLoader(str(BASE_DIR)))
    template = env.get_template("template_simple_email.html")
    html = template.render(
        month=month,
        year=year,
        logo_img=LOGO_URL,
        view_all_img=VIEW_ALL_URL,
        icon_social_linkedin=SOCIAL_ICONS["linkedin"],
        icon_social_facebook=SOCIAL_ICONS["facebook"],
        icon_social_youtube=SOCIAL_ICONS["youtube"],
        icon_social_instagram=SOCIAL_ICONS["instagram"],
        icon_social_twitter=SOCIAL_ICONS["twitter"],
        icon_social_dribbble=SOCIAL_ICONS["dribbble"],
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

    # Everything in output/sections/ has to exist at base_url before sending.
    upload_files = sorted(p.name for p in SECTIONS_DIR.glob("*.png"))
    upload_note = SECTIONS_DIR / "UPLOAD_ME.txt"
    upload_note.write_text(
        f"Upload these {len(upload_files)} files to:\n{base_url}/\n\n"
        + "\n".join(upload_files)
        + "\n",
        encoding="utf-8",
    )

    print(f"\nSimple email generated: {out_file}")
    print(f"  Sections included: {' | '.join(included)}{' | Marketing Highlights' if marketing else ''}")
    if skipped or not marketing:
        missing = skipped + ([] if marketing else ["Marketing Highlights"])
        print(f"  Sections skipped (no PNG found): {' | '.join(missing)}")
    print(f"  Image host: {base_url}/")
    print(f"  UPLOAD {len(upload_files)} PNG(s) from output/sections/ to that folder before sending.")
    print(f"  File list written to: {upload_note}")
    print(f"  Open in browser: file:///{out_file.as_posix()}\n")


if __name__ == "__main__":
    main()
