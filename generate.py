"""
Newsletter Generator
====================
Usage:
    python generate.py              → browser preview (relative paths)
    python generate.py --email      → Gmail-ready HTML (GitHub Pages URLs)
    python generate.py --embed-images → single-file HTML (data URIs)

Output: output/newsletter_<Month>_<Year>.html
"""

import json
import logging
import os
import sys
import base64
import mimetypes
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

# ── Config ──────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
CONTENT_FILE = BASE_DIR / "content.json"
ASSETS_DIR   = BASE_DIR / "assets"
IMAGES_DIR   = BASE_DIR / "images"
OUTPUT_DIR   = BASE_DIR / "output"

GITHUB_PAGES_URL = "https://biztechdesign.github.io/Newsletter/"

EMBED_IMAGES   = "--embed-images" in sys.argv
EMAIL_MODE_ARG = "--email" in sys.argv
IMAGE_BASE_URL = (
    GITHUB_PAGES_URL if EMAIL_MODE_ARG
    else next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--image-url=")), None)
)
EMAIL_MODE  = EMBED_IMAGES or bool(IMAGE_BASE_URL)
TEMPLATE    = "template_email.html" if EMAIL_MODE else "template.html"

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg"}

# ── Helpers ──────────────────────────────────────────────────────────────────

def to_data_uri(path: Path) -> str:
    if not path or not path.exists():
        return ""
    mime, _ = mimetypes.guess_type(str(path))
    mime = mime or "image/png"
    data = base64.b64encode(path.read_bytes()).decode()
    return f"data:{mime};base64,{data}"


def img_src(path: Path | None) -> str:
    if not path or not path.exists():
        return ""
    if EMBED_IMAGES:
        return to_data_uri(path)
    if IMAGE_BASE_URL:
        rel = path.relative_to(BASE_DIR).as_posix()
        return f"{IMAGE_BASE_URL.rstrip('/')}/{rel}"
    return f"../{path.relative_to(BASE_DIR).as_posix()}"


def first_image(folder: Path) -> Path | None:
    """Return the first image file found in *folder*, sorted by name."""
    if not folder.exists():
        return None
    files = sorted(
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS
    )
    return files[0] if files else None


def all_images(folder: Path) -> list[Path]:
    """Return all image files in *folder*, sorted by name."""
    if not folder.exists():
        return []
    return sorted(
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS
    )


def normalize_hr(hr: dict, legacy_event_photos: list[str]) -> dict:
    """
    Accept either the new multi-event shape or the old single-event shape
    (from a content.json produced before repeatable HR events were supported).
    """
    hr = hr or {}
    events = hr.get("events")
    if not isinstance(events, list):
        events = []
        legacy_title = hr.get("event_title", "")
        if legacy_title or legacy_event_photos:
            events = [{"title": legacy_title, "photos": legacy_event_photos}]
    events = [e for e in events if e.get("title") or e.get("photos")]
    return {"wellness_title": hr.get("wellness_title", ""), "events": events}


def normalize_entries(value, legacy_logo_url: str = "") -> list[dict]:
    """
    Accept either the new list-of-entries shape or the old single-dict shape
    (from a content.json produced before repeatable entries were supported),
    normalizing both into a list. Old-shape logos lived under a separate
    top-level image_urls key, passed in here as legacy_logo_url.
    """
    if isinstance(value, list):
        return value
    if isinstance(value, dict) and value:
        entry = dict(value)
        entry.setdefault("logo_url", legacy_logo_url)
        return [entry]
    return []


def build_reward_rows(awards: list[dict], award_images: list[list[str]]) -> list[dict]:
    """
    Group awards into display rows: max 2 solo awards per row (side by side),
    a duo award (2 recipients) always takes its own full-width row. A lone
    trailing solo award (odd count) gets its own full-width row too, instead
    of being left half-width with empty space beside it.
    """
    rows: list[dict] = []
    pending = None
    for award, photos in zip(awards, award_images):
        entry = {"award": award, "photos": photos, "is_duo": len(award.get("recipients", [])) > 1}
        if entry["is_duo"]:
            if pending:
                rows.append({"entries": [pending]})
                pending = None
            rows.append({"entries": [entry]})
        else:
            if pending:
                rows.append({"entries": [pending, entry]})
                pending = None
            else:
                pending = entry
    if pending:
        rows.append({"entries": [pending]})
    return rows


def parse_new_joinees_from_images(folder: Path) -> list[dict]:
    """
    Scan images/new_joinees/ for files named:
        Name--Designation.jpg
    Returns a list of {name, designation, image_url} for local preview fallback.
    """
    people = []
    for f in all_images(folder):
        parts = [p.strip() for p in f.stem.split("--")]
        if len(parts) < 2:
            continue
        people.append({"name": parts[0], "designation": parts[1], "image_url": img_src(f)})
    return people


def parse_new_customers_from_images(folder: Path) -> list[dict]:
    """
    Scan images/new_customer/ for logo image files (any filenames — no
    naming convention required). Returns a list of {name, image_url}.
    """
    return [{"name": f.stem, "image_url": img_src(f)} for f in all_images(folder)]


def parse_announcements_from_images(folder: Path) -> list[dict]:
    """
    Scan images/announcement/ for banner image files (any filenames — no
    naming convention required). Returns a list of {name, image_url}.
    """
    return [{"name": f.stem, "image_url": img_src(f)} for f in all_images(folder)]


def parse_certifications_from_images(folder: Path) -> list[dict]:
    """
    Scan images/certification/ for certificate image files (any filenames —
    no naming convention required). Returns a list of {name, image_url}.
    """
    return [{"name": f.stem, "image_url": img_src(f)} for f in all_images(folder)]


# Balloon decoration for Work Anniversary cards — cycles through the 4 colors
# and a handful of corner positions, deterministically per card (based on the
# card's own content) so re-generating the newsletter doesn't reshuffle them.
# One balloon per card, always on the top-right corner. The colour still
# varies per card; only the position is fixed. (This used to rotate through
# five positions, which scattered balloons around and across the cards.)
BALLOON_POSITIONS = [
    {"top": "-18px", "right": "14px"},
]


def assign_anniversary_balloons(anniversaries: list[dict], balloon_icons: list[str]):
    """
    Attach a {icon, style} 'balloon' dict to each anniversary group in place.

    Colours cycle by position rather than being picked from a hash of the
    names. There are only four balloon images, so with more cards than that
    some colour must repeat — but cycling guarantees neighbours never match,
    left-to-right or top-to-bottom in the three-column grid, which hashing
    could not (it happily gave two adjacent cards the same colour).
    """
    icons = [i for i in balloon_icons if i]
    for n, group in enumerate(anniversaries):
        pos = BALLOON_POSITIONS[n % len(BALLOON_POSITIONS)]
        style = ";".join(f"{k}:{v}" for k, v in pos.items())
        group["balloon"] = {
            "icon": icons[n % len(icons)] if icons else "",
            "style": style,
        }


def parse_awards_from_images(folder: Path) -> tuple[list[dict], list[list[str]]]:
    """
    Scan images/awards/ for files named:
        Award Title--Recipient Name--Designation.jpg
    or, with an explicit display-order prefix:
        01--Award Title--Recipient Name--Designation.jpg
    Returns (awards_list, award_images_list) ready for the template.
    Multiple recipients for the same award = multiple files with the same title prefix.
    """
    files = all_images(folder)
    award_map: dict[str, dict] = {}
    award_order: list[str] = []

    for f in files:
        parts = [p.strip() for p in f.stem.split("--")]
        if parts and parts[0].isdigit():
            parts = parts[1:]  # strip leading order-number segment, e.g. "01"
        if len(parts) < 2:
            continue  # skip files without the naming convention
        title = parts[0]
        name  = parts[1] if len(parts) > 1 else ""
        desig = parts[2] if len(parts) > 2 else ""

        if title not in award_map:
            award_map[title] = {"title": title, "folder": "", "recipients": [], "_photos": []}
            award_order.append(title)
        award_map[title]["recipients"].append({"name": name, "designation": desig, "image_url": ""})
        award_map[title]["_photos"].append(f)

    awards      = [award_map[t] for t in award_order]
    award_imgs  = [[img_src(p) for p in award_map[t]["_photos"]] for t in award_order]
    for a in awards:
        del a["_photos"]
    return awards, award_imgs



# ── Main ─────────────────────────────────────────────────────────────────────

def build():
    # ── Load content ────────────────────────────────────────────
    with open(CONTENT_FILE, encoding="utf-8") as fh:
        data = json.load(fh)

    month = data["month"]
    year  = data["year"]

    # ── Resolve images ───────────────────────────────────────────

    # Fixed assets
    logo_img       = img_src(ASSETS_DIR / "logo.png")
    icon_ceo       = img_src(ASSETS_DIR / "icon_ceo.png")
    icon_delivery  = img_src(ASSETS_DIR / "icon_delivery.png")
    icon_excellence= img_src(ASSETS_DIR / "icon_excellence.png")
    icon_hr        = img_src(ASSETS_DIR / "icon_hr.png")
    icon_rewards   = img_src(ASSETS_DIR / "icon_rewards.png")
    icon_anniversaries = img_src(ASSETS_DIR / "icon_anniversaries.png")
    icon_openings  = img_src(ASSETS_DIR / "icon_openings.png")
    icon_marketing = img_src(ASSETS_DIR / "icon_marketing.png")
    icon_new_joinees = img_src(ASSETS_DIR / "addition.png")
    icon_customers   = img_src(ASSETS_DIR / "icon_customers.png")
    icon_events      = img_src(ASSETS_DIR / "icon_events.png")
    icon_certifications = img_src(ASSETS_DIR / "Certificate.png")
    icon_expo           = img_src(ASSETS_DIR / "icon_expo.png")
    icon_training       = img_src(ASSETS_DIR / "icon_training.png")

    balloon_icons = [
        img_src(ASSETS_DIR / "green-Balloon.png"),
        img_src(ASSETS_DIR / "pink-Balloon.png"),
        img_src(ASSETS_DIR / "purple-Balloon.png"),
        img_src(ASSETS_DIR / "red-Balloon.png"),
    ]
    assign_anniversary_balloons(data.get("anniversaries", []), balloon_icons)

    # PNG, not SVG — Outlook desktop can't render SVG images at all
    icon_social_linkedin  = img_src(ASSETS_DIR / "icon_social_linkedin.png")
    icon_social_facebook  = img_src(ASSETS_DIR / "icon_social_facebook.png")
    icon_social_youtube   = img_src(ASSETS_DIR / "icon_social_youtube.png")
    icon_social_instagram = img_src(ASSETS_DIR / "icon_social_instagram.png")
    icon_social_twitter   = img_src(ASSETS_DIR / "icon_social_twitter.png")
    icon_social_dribbble  = img_src(ASSETS_DIR / "icon_social_dribbble.png")

    # Monthly images come from content.json, which import_content.py has
    # already resolved (this month's Row Data folder, else the sheet's links).
    #
    # There used to be a fallback here to the repo's images/ folder whenever a
    # value was missing. That silently refilled any section left empty with the
    # PREVIOUS issue's photos — so an empty folder, or a row marked NA to drop a
    # section, produced a section that looked perfectly correct and was wrong.
    # content.json is the single source of truth; empty means empty.
    img_urls = data.get("image_urls", {})

    ceo_photo          = img_urls.get("ceo_photo", "")
    campaign_banner    = img_urls.get("campaign_banner", "")
    hr_wellness_banner = img_urls.get("hr_wellness_banner", "")

    # HR events — dynamic list (up to 3+ events, each with its own title + photos).
    # Normalizes old single-event content.json shape too (photos lived under image_urls.event_photos).
    data["hr"] = normalize_hr(data.get("hr"), img_urls.get("event_photos", []))
    # An event with neither a title nor photos is nothing to show.
    data["hr"]["events"] = [
        e for e in data["hr"]["events"]
        if (e.get("title") or "").strip() or e.get("photos")
    ]

    # Delivery Insights / Acknowledging Excellence — dynamic list of entries.
    # Normalizes old single-dict content.json shape too (logo lived under image_urls).
    data["delivery_insights"] = normalize_entries(data.get("delivery_insights"), img_urls.get("delivery_logo", ""))
    data["acknowledging_excellence"] = normalize_entries(data.get("acknowledging_excellence"), img_urls.get("excellence_logo", ""))

    # Drop entries with no text and no logo — an empty row must not keep a
    # section alive, or marking it NA in the sheet would have no effect.
    data["delivery_insights"] = [
        e for e in data["delivery_insights"]
        if (e.get("description") or "").strip() or e.get("logo_url")
    ]
    data["acknowledging_excellence"] = [
        e for e in data["acknowledging_excellence"]
        if (e.get("testimonial") or e.get("description") or "").strip() or e.get("logo_url")
    ]

    award_images = [
        [r.get("image_url", "") for r in award.get("recipients", [])]
        for award in data["rewards"]["awards"]
    ]
    reward_rows = build_reward_rows(data["rewards"]["awards"], award_images)

    blog_images = [post.get("image_url", "") for post in data["marketing"]["blog_posts"]]

    openings_qr_img = img_src(ASSETS_DIR / "qr_code.png")  # optional fixed asset

    # ── Section visibility — sections with no content are skipped ─
    show = {
        "ceo":           bool(((data.get("ceo_desk") or {}).get("message") or "").strip()),
        "campaign":      bool(campaign_banner),
        "new_customers": bool(data.get("new_customers")),
        "delivery":      bool(data.get("delivery_insights")),
        "excellence":    bool(data.get("acknowledging_excellence")),
        "hr":            bool(hr_wellness_banner or data["hr"]["events"]),
        "rewards":       bool((data.get("rewards") or {}).get("awards")),
        "certifications": bool(data.get("certifications")),
        "anniversaries": bool(data.get("anniversaries")),
        "new_joinees":   bool(data.get("new_joinees")),
        "openings":      bool((data.get("new_openings") or {}).get("positions")),
        "marketing":     bool((data.get("marketing") or {}).get("blog_posts")),
        "announcements": bool(data.get("announcements")),
        "expo":          bool((data.get("expo") or {}).get("photos")),
        "training":      bool((data.get("training") or {}).get("photos")),
        "workshop":      bool((data.get("workshop") or {}).get("photos")),
    }

    # A section the sheet marks NA is dropped even when content exists for it,
    # e.g. last month's photos still sitting in its Row Data folder. "NA" is an
    # instruction to leave the section out, not merely a missing value.
    for key in data.get("_na_sections", []):
        show[key] = False

    # ── Render template ──────────────────────────────────────────
    env = Environment(loader=FileSystemLoader(str(BASE_DIR)))
    tpl = env.get_template(TEMPLATE)

    html = tpl.render(
        # Content data (passed as-is so template can access nested keys)
        **data,

        # Resolved image sources
        logo_img              = logo_img,
        icon_ceo              = icon_ceo,
        icon_delivery         = icon_delivery,
        icon_excellence       = icon_excellence,
        icon_hr               = icon_hr,
        icon_rewards          = icon_rewards,
        icon_anniversaries    = icon_anniversaries,
        icon_openings         = icon_openings,
        icon_marketing        = icon_marketing,
        icon_new_joinees      = icon_new_joinees,
        icon_customers        = icon_customers,
        icon_events           = icon_events,
        icon_certifications   = icon_certifications,
        icon_expo             = icon_expo,
        icon_training         = icon_training,

        icon_social_linkedin  = icon_social_linkedin,
        icon_social_facebook  = icon_social_facebook,
        icon_social_youtube   = icon_social_youtube,
        icon_social_instagram = icon_social_instagram,
        icon_social_twitter   = icon_social_twitter,
        icon_social_dribbble  = icon_social_dribbble,

        ceo_photo             = ceo_photo,
        campaign_banner       = campaign_banner,
        hr_wellness_banner    = hr_wellness_banner,

        award_images          = award_images,
        reward_rows           = reward_rows,

        blog_images           = blog_images,
        blog_posts_enum       = list(enumerate(data["marketing"]["blog_posts"])),
        openings_qr_img       = openings_qr_img,
        show                  = show,
    )

    # ── Inline CSS for email mode ────────────────────────────────
    if EMAIL_MODE:
        try:
            from premailer import transform
            html = transform(
                html,
                remove_classes=False,
                strip_important=False,
                allow_network=False,
                cssutils_logging_level=logging.CRITICAL,
            )
        except Exception:
            pass

    # ── Write output ─────────────────────────────────────────────
    OUTPUT_DIR.mkdir(exist_ok=True)
    if EMBED_IMAGES:
        suffix = "_embedded"
    elif IMAGE_BASE_URL:
        suffix = "_email"
    else:
        suffix = ""
    out_file = OUTPUT_DIR / f"newsletter_{month}_{year}{suffix}.html"
    out_file.write_text(html, encoding="utf-8")

    # Same order as the master sheet everyone fills in, which is also the
    # order the sections appear in template.html.
    section_labels = [
        ("ceo", "CEO Desk"),
        ("campaign", "Campaign Banner"),
        ("expo", "Expo / Exhibition"),
        ("new_customers", "Newly Added Customers"),
        ("delivery", "Delivery Insights"),
        ("excellence", "Excellence"),
        ("hr", "HR Insider / Events"),
        ("rewards", "Rewards"),
        ("new_joinees", "New Additions"),
        ("anniversaries", "Anniversaries"),
        ("certifications", "Employee Certifications"),
        ("openings", "New Openings"),
        ("training", "Employee Training"),
        ("workshop", "Employee Workshop"),
        ("announcements", "Upcoming Event Announcement"),
        ("marketing", "Marketing Highlights"),
    ]
    included = [label for key, label in section_labels if show[key]]
    skipped  = [label for key, label in section_labels if not show[key]]

    print(f"\nNewsletter generated: {out_file}")
    print(f"  Sections included: {' | '.join(included)}")
    if skipped:
        print(f"  Sections skipped (no content): {' | '.join(skipped)}")
    if EMBED_IMAGES:
        print(f"  Images: embedded as data URIs")
    elif IMAGE_BASE_URL:
        print(f"  Images: hosted at {IMAGE_BASE_URL}")
    else:
        print(f"  Images: relative paths (browser preview)")
    print(f"  Open in browser: file:///{out_file.as_posix()}\n")


if __name__ == "__main__":
    build()
