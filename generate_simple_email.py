"""
generate_simple_email.py
=========================
Builds the final Gmail- and Outlook-safe newsletter email out of the flat
per-section PNGs in "<Month>-<Year>/Section wise images" (produced by
screenshot_sections.py),
instead of the fully CSS-styled template_email.html — which renders correctly
in Gmail but breaks in Outlook desktop (gradients, border-radius, absolute
positioning, etc. are all unsupported there). Flattening each section to an
image sidesteps that entirely; only Header, Footer, and the Marketing
Highlights blog cards + "View All" button stay as real HTML (simple enough to
be Outlook-safe, and the blog cards need to stay clickable, which a flat image
can't do for just part of itself).

The markup follows the approved reference layout (Roboto, 800px tables,
.mobile_table responsive rules, teal footer) — see template_simple_email.html.

IMPORTANT: this reads image dimensions from the LOCAL PNGs, but writes URLs
pointing at wherever you've uploaded those *same* files on your own server
(see HOSTED_SECTIONS_BASE below) — you must upload them there yourself before
sending, this script does not do that upload. The exact file list is written
to UPLOAD_ME.txt next to the PNGs.

Each issue gets its OWN folder on the server, e.g.
    https://w.indiaondesk.com/biztech-insider/January_2026/
so uploading a new month never breaks the previously sent emails.

Usage:
    python generate_simple_email.py

Run this after screenshot_sections.py has produced this month's PNGs.
"""

import json
import os
import sys
from pathlib import Path
from urllib.parse import quote

from jinja2 import Environment, FileSystemLoader
from PIL import Image

BASE_DIR = Path(__file__).parent
CONTENT_JSON = BASE_DIR / "content.json"
HTML_SUBDIR = "HTML"


def html_dir(month: str, year: str) -> Path:
    """Everything an issue produces sits in its own month folder."""
    return BASE_DIR / f"{month}-{year}" / HTML_SUBDIR

# Per-issue folder written by screenshot_sections.py, e.g.
# "July-2026/Section wise images".
SECTIONS_SUBDIR = "Section wise images"


def sections_dir(month: str, year: str) -> Path:
    return BASE_DIR / f"{month}-{year}" / SECTIONS_SUBDIR

# Where you upload each month's section PNGs. The {month}/{year} placeholders
# give every issue its own folder, so re-uploading next month doesn't overwrite
# the images an already-sent email still points at.
#
# On the server this is  /home/wordpress/public_html/biztech-insider/<Month>_<Year>/
# so July 2026 lands in  /home/wordpress/public_html/biztech-insider/July_2026/
# Upload credentials come from Parth.
HOSTED_SECTIONS_BASE = "https://w.indiaondesk.com/biztech-insider/{month}_{year}/"
HOSTED_SERVER_PATH = "/home/wordpress/public_html/biztech-insider"

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

# (filename stem in "<Month>-<Year>/Section wise images", alt text) — canonical order,
# matching template.html so the approved browser preview and the final email
# always show the sections in the same sequence. Header/Footer are excluded
# (real HTML) and Marketing Highlights is special-cased.
SECTIONS_BEFORE_MARKETING = [
    ("Director-desk", "From CEO's Desk"),
    ("Whistle-Blow", "Campaign Banner"),
    ("Expo-Exhibition", "Expo & Exhibition"),
    ("Newly-Added-Customers", "Newly Added Customers"),
    ("Delivery-Insights", "Delivery Insights"),
    ("Acknowledging-Excellence", "Acknowledging Excellence"),
    ("HR-INSIDER", "HR Insider"),
    ("Rewards-Recognitions", "Rewards & Recognitions"),
    ("NEW-ADDITION", "New Addition to the Biztech Family"),
    ("WORK-ANNIVERSARY", "Work Anniversaries"),
    ("Employee-Certifications", "Employee Certifications"),
    ("new-openings", "New Openings"),
    ("Employee-Training", "Employee Training"),
    ("Employee-Workshop", "Employee Workshop"),
    ("Upcoming-Event-Announcement", "Upcoming Event Announcement"),
]
# Marketing Highlights is last in the master sheet, so nothing follows it.
SECTIONS_AFTER_MARKETING = []


# The width every full-width section is laid out at, and displayed at in the
# email. Section PNGs are captured at some multiple of this for sharpness on
# retina screens, so dividing a PNG's pixel width by it recovers the capture
# density — see capture_scale(). Derived rather than hardcoded so the email can
# never disagree with whatever density the PNGs were actually captured at.
SECTION_DISPLAY_WIDTH = 800
FALLBACK_SCALE = 1.5

# Rendered width of a blog card in the final email: the 800px table has 50px
# side padding and each of the two cells adds 15px either side, leaving
# (800 - 100) / 2 - 30 = 320px. The captured card is wider than that, so its
# height is scaled by the same ratio to keep the aspect ratio intact.
BLOG_DISPLAY_WIDTH = 320


def capture_scale(sec_dir: Path) -> float:
    """
    The pixel density the section PNGs were captured at, read off a full-width
    one: every such section is laid out at SECTION_DISPLAY_WIDTH, so its PNG's
    pixel width divided by that is the density.

    Measured rather than assumed. A hardcoded factor that disagreed with the
    actual capture would rescale every image in the email — by a third, if the
    capture moved between 1.5x and 2x — and the email would still look
    plausible, just subtly wrong throughout.
    """
    for path in sorted(sec_dir.glob("*.png")):
        if path.stem.startswith("blog-"):
            continue        # blog cards are not full-width
        with Image.open(path) as im:
            width = im.size[0]
        if width >= SECTION_DISPLAY_WIDTH:
            return width / SECTION_DISPLAY_WIDTH
    return FALLBACK_SCALE


def png_size(path: Path, scale: float) -> tuple[int, int]:
    """The size a PNG should be displayed at, in CSS pixels: its pixel
    dimensions divided back down by the density it was captured at."""
    with Image.open(path) as im:
        w, h = im.size
    return round(w / scale), round(h / scale)


def hosted_base(month: str, year: str) -> str:
    return HOSTED_SECTIONS_BASE.format(month=month, year=year).rstrip("/")


def expand_parts(stem: str, sec_dir: Path) -> list[str]:
    """
    The filename stems for a section, in order.

    HR Insider is cut into one image per sub-titled block (HR-INSIDER-1,
    HR-INSIDER-2, …), so it contributes several stems where every other section
    contributes one. Stacked in order they rejoin with no visible seam.
    """
    numbered = sorted(
        (p.stem for p in sec_dir.glob(f"{stem}-*.png")),
        key=lambda s: int(s.rsplit("-", 1)[1]) if s.rsplit("-", 1)[1].isdigit() else 0,
    )
    return numbered or [stem]


def resolve_sections(pairs, base_url, sec_dir, scale):
    resolved = []
    for stem_base, alt in pairs:
        for stem in expand_parts(stem_base, sec_dir):
            path = sec_dir / f"{stem}.png"
            if not path.exists():
                continue
            _w, h = png_size(path, scale)
            resolved.append({
                "url": f"{base_url}/{stem}.png",
                "alt": alt,
                "height": h,
            })
    return resolved


def resolve_marketing(base_url, data, sec_dir, scale):
    bg_path = sec_dir / "Marketing-Highlights.png"
    if not bg_path.exists():
        return None
    _w, bg_h = png_size(bg_path, scale)

    links_path = sec_dir / "marketing_links.json"
    links = json.loads(links_path.read_text(encoding="utf-8")) if links_path.exists() else []
    href_by_name = {entry["image"]: entry["href"] for entry in links}

    # Blog titles live in content.json, keyed by post URL — the same href the
    # card was captured with — so the <a title> and <img alt> stay meaningful.
    posts = data.get("marketing", {}).get("blog_posts", [])
    title_by_url = {p.get("url", ""): p.get("title", "") for p in posts}

    blogs = []
    blog_paths = sorted(sec_dir.glob("blog-*.png"), key=lambda p: p.name)
    for i, path in enumerate(blog_paths):
        w, h = png_size(path, scale)
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
    sec_dir = sections_dir(month, year)
    out_dir = html_dir(month, year)

    # --local points the <img> tags at the PNGs on this machine instead of the
    # server, so the email can be checked (or pasted into Gmail, which uploads
    # the images itself) before anything has been uploaded. Written to its own
    # file so the sendable, server-hosted one is never overwritten by it.
    local_mode = "--local" in sys.argv
    if local_mode:
        rel = os.path.relpath(sec_dir, out_dir).replace("\\", "/")
        base_url = quote(rel)
    else:
        base_url = hosted_base(month, year)
    if not sec_dir.is_dir():
        raise SystemExit(
            f"{sec_dir} not found — run 'python screenshot_sections.py' first to cut the section images."
        )

    scale = capture_scale(sec_dir)
    sections_before = resolve_sections(SECTIONS_BEFORE_MARKETING, base_url, sec_dir, scale)
    sections_after = resolve_sections(SECTIONS_AFTER_MARKETING, base_url, sec_dir, scale)
    marketing = resolve_marketing(base_url, data, sec_dir, scale)

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

    out_dir.mkdir(parents=True, exist_ok=True)
    suffix = "_simple_email_LOCAL" if local_mode else "_simple_email"
    out_file = out_dir / f"newsletter_{month}_{year}{suffix}.html"
    out_file.write_text(html, encoding="utf-8")

    all_stems = [s for s, _ in SECTIONS_BEFORE_MARKETING] + [s for s, _ in SECTIONS_AFTER_MARKETING]
    # Checked through expand_parts so a split section counts as present: HR
    # Insider ships as HR-INSIDER-1/-2/-… and has no HR-INSIDER.png of its own.
    included = [
        s for s in all_stems
        if any((sec_dir / f"{part}.png").exists() for part in expand_parts(s, sec_dir))
    ]
    skipped = [s for s in all_stems if s not in included]

    upload_files = sorted(p.name for p in sec_dir.glob("*.png"))
    if not local_mode:
        # Every PNG has to exist at base_url before sending.
        upload_note = sec_dir / "UPLOAD_ME.txt"
        upload_note.write_text(
            f"Upload these {len(upload_files)} files to:\n\n"
            f"  server : {HOSTED_SERVER_PATH}/{month}_{year}/\n"
            f"  url    : {hosted_base(month, year)}/\n\n"
            + "\n".join(upload_files)
            + "\n",
            encoding="utf-8",
        )

    print(f"\nSimple email generated: {out_file}")
    print(f"  Sections included: {' | '.join(included)}{' | Marketing Highlights' if marketing else ''}")
    if skipped or not marketing:
        missing = skipped + ([] if marketing else ["Marketing Highlights"])
        print(f"  Sections skipped (no PNG found): {' | '.join(missing)}")
    if local_mode:
        print(f"  Images: local files in '{sec_dir}' - for checking, and for")
        print(f"          pasting into Gmail, which uploads them itself.")
        print(f"  NOT for sending as raw HTML: the paths only exist on this machine.")
    else:
        print(f"  Image host: {hosted_base(month, year)}/")
        print(f"  UPLOAD {len(upload_files)} PNG(s) from '{sec_dir}' to that folder before sending.")
        print(f"  File list written to: {upload_note}")
        print(f"  Preview it before uploading with: python generate_simple_email.py --local")
    print(f"  Open in browser: file:///{out_file.as_posix()}\n")


if __name__ == "__main__":
    main()
