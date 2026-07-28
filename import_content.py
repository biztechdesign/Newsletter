"""
import_content.py
=================
Auto-generates content.json from two Google Sheets:
  1. Newsletter Content sheet  — text, awards, openings, blogs
  2. Work Anniversary sheet    — employee anniversary data

Both sheets must be shared as "Anyone with the link can view".

Usage:
    python import_content.py

This also renders the browser-preview HTML (generate.py) and refreshes the
per-section JPGs in output/sections/ automatically — no separate steps needed
just to see the updated content.

Update the SHEET IDs at the top when working on a new month.
Only if sending via Gmail, additionally run:
    python generate.py --email
"""

import io
import json
import re
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl

# ── Configure these each month ────────────────────────────────────────────
CONTENT_SHEET_ID = "1vZhE7L6O7NNAHYrsGOm8nlZjjwQtMzbu"
# Anniversary file + tab are now read from the 'anniversary' row in the
# content sheet (Drive Link column = xlsx file, Value column = tab name).
# ANNIV_SHEET_ID below is only used as a fallback if not specified in sheet.
ANNIV_SHEET_ID   = "19m9QWBhOqLh9pZT-u4eoaPkN9XS3hu69"
# ──────────────────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).parent
OUTPUT   = BASE_DIR / "content.json"

# Footer is static — update here if it ever changes
FOOTER = {
    "address": "C/801 Dev Aurum Commercial, Anandnagar Cross Road Prahalad Nagar, Satellite, Ahmedabad, Gujarat 380015",
    "email":   "career@biztechcs.com",
    "phone":   "+91 93276 55844",
    "social": {
        "linkedin":  "https://www.linkedin.com/company/biztech/",
        "facebook":  "https://www.facebook.com/biztech/",
        "youtube":   "https://www.youtube.com/@Biztechcs",
        "instagram": "https://www.instagram.com/biztechcs/",
        "twitter":   "https://twitter.com/biztechcs",
        "dribbble":  "https://dribbble.com/biztechcs",
    },
}


# ── Helpers ───────────────────────────────────────────────────────────────

def download_xlsx(sheet_id: str) -> openpyxl.Workbook:
    """Download a Google Sheet as xlsx (follows redirects automatically)."""
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp:
        return openpyxl.load_workbook(io.BytesIO(resp.read()))


def load_xlsx(sheet_id: str, local_fallback: str) -> openpyxl.Workbook:
    """
    Try to download from Google Sheets. If the sheet is private (401),
    fall back to reading a local xlsx file.
    """
    try:
        wb = download_xlsx(sheet_id)
        print("  Downloaded from Google Sheets.")
        return wb
    except urllib.error.HTTPError as e:
        if e.code == 401:
            local = BASE_DIR / local_fallback
            if local.exists():
                print(f"  Sheet is private — reading local file: {local_fallback}")
                return openpyxl.load_workbook(local)
            else:
                raise FileNotFoundError(
                    f"\n  The sheet is not publicly shared and no local file found.\n"
                    f"  Fix option 1: Share the sheet → 'Anyone with the link can view'\n"
                    f"  Fix option 2: Download it as .xlsx and save as: {local}\n"
                ) from e
        raise


def parse_years(tenure_str) -> int:
    """Extract numeric years from '14 years' or '1 year'."""
    m = re.search(r"(\d+)", str(tenure_str or ""))
    return int(m.group(1)) if m else 0


def cell_str(value) -> str:
    """Safe string conversion of a cell value."""
    return str(value).strip() if value is not None else ""


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg"}


def drive_file_id(drive_url: str) -> str:
    """Extract the file ID from any Google Drive URL."""
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", drive_url or "")
    return m.group(1) if m else ""


# ── GitHub helpers ────────────────────────────────────────────────────────

def parse_github_url(url: str) -> dict | None:
    """
    Parse a github.com tree/blob URL.
    Returns {owner, repo, branch, path, is_folder} or None.
    """
    m = re.match(r"https://github\.com/([^/]+)/([^/]+)/(tree|blob)/([^/]+)/(.+)", url or "")
    if m:
        return {
            "owner":     m.group(1),
            "repo":      m.group(2),
            "is_folder": m.group(3) == "tree",
            "branch":    m.group(4),
            "path":      m.group(5).rstrip("/"),
        }
    return None


def list_github_folder(owner: str, repo: str, branch: str, path: str) -> list[dict]:
    """
    List image files in a GitHub repo folder via the API.
    Returns [{name, stem, url}] sorted by name, where url is the GitHub Pages URL.
    """
    api = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}?ref={branch}"
    req = urllib.request.Request(api, headers={
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/vnd.github.v3+json",
    })
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            files = json.loads(resp.read())
            result = []
            for f in files:
                if f["type"] == "file" and Path(f["name"]).suffix.lower() in IMAGE_EXTS:
                    pages_url = f"https://{owner}.github.io/{repo}/{path}/{f['name']}"
                    result.append({"name": f["name"], "stem": Path(f["name"]).stem, "url": pages_url})
            return sorted(result, key=lambda x: x["name"])
    except Exception as e:
        print(f"  WARNING: Could not list GitHub folder '{path}': {e}")
        return []


def resolve_single_url(raw: str) -> str:
    """
    Resolve any raw URL to a direct image URL:
    - GitHub folder (tree) → first image in that folder (GitHub Pages URL)
    - GitHub file  (blob)  → GitHub Pages URL
    - Drive share link     → uc?export=view URL
    - Any other http URL   → as-is
    """
    if not raw or "YOUR_FILE_ID" in raw:
        return ""
    info = parse_github_url(raw)
    if info:
        if info["is_folder"]:
            files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
            return files[0]["url"] if files else ""
        else:
            return f"https://{info['owner']}.github.io/{info['repo']}/{info['path']}"
    if "drive.google.com" in raw:
        fid = drive_file_id(raw)
        return f"https://drive.google.com/uc?export=view&id={fid}" if fid else ""
    if raw.startswith("http"):
        return raw
    return ""


def resolve_folder_urls(raw: str) -> list[str]:
    """
    Resolve a folder URL to a list of all image URLs in that folder.
    For non-folder URLs, wraps resolve_single_url() in a list.
    """
    if not raw:
        return []
    info = parse_github_url(raw)
    if info and info["is_folder"]:
        files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
        return [f["url"] for f in files]
    single = resolve_single_url(raw)
    return [single] if single else []


def parse_awards_from_github_folder(folder_url: str) -> list[dict]:
    """
    List a GitHub folder, parse filenames as 'Title--Name--Designation.ext',
    or, with an explicit display-order prefix, '01--Title--Name--Designation.ext'.
    Return awards list with recipients and image_url per recipient.
    """
    info = parse_github_url(folder_url)
    if not info or not info["is_folder"]:
        return []
    files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
    award_map: dict[str, dict] = {}
    award_order: list[str] = []
    for f in files:
        parts = [p.strip() for p in f["stem"].split("--")]
        if parts and parts[0].isdigit():
            parts = parts[1:]  # strip leading order-number segment, e.g. "01"
        if len(parts) < 2:
            continue
        title = parts[0]
        name  = parts[1] if len(parts) > 1 else ""
        desig = parts[2] if len(parts) > 2 else ""
        if title not in award_map:
            award_map[title] = {"title": title, "folder": "", "recipients": []}
            award_order.append(title)
        award_map[title]["recipients"].append({"name": name, "designation": desig, "image_url": f["url"]})
    return [award_map[t] for t in award_order]


def parse_new_joinees_from_github_folder(folder_url: str) -> list[dict]:
    """
    List a GitHub folder, parse filenames as 'Name--Designation.ext',
    return a list of {name, designation, image_url}.
    """
    info = parse_github_url(folder_url)
    if not info or not info["is_folder"]:
        return []
    files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
    people = []
    for f in files:
        parts = [p.strip() for p in f["stem"].split("--")]
        if len(parts) < 2:
            continue
        people.append({"name": parts[0], "designation": parts[1], "image_url": f["url"]})
    return people


def parse_new_customers_from_github_folder(folder_url: str) -> list[dict]:
    """
    List a GitHub folder of customer logo images (any filenames — no naming
    convention required). Returns a list of {name, image_url}.
    """
    info = parse_github_url(folder_url)
    if not info or not info["is_folder"]:
        return []
    files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
    return [{"name": f["stem"], "image_url": f["url"]} for f in files]


def parse_announcements_from_github_folder(folder_url: str) -> list[dict]:
    """
    List a GitHub folder of event announcement banner images (any filenames —
    no naming convention required). Returns a list of {name, image_url}.
    """
    info = parse_github_url(folder_url)
    if not info or not info["is_folder"]:
        return []
    files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
    return [{"name": f["stem"], "image_url": f["url"]} for f in files]


def parse_certifications_from_github_folder(folder_url: str) -> list[dict]:
    """
    List a GitHub folder of employee certification images (any filenames —
    no naming convention required). Returns a list of {name, image_url}.
    """
    info = parse_github_url(folder_url)
    if not info or not info["is_folder"]:
        return []
    files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
    return [{"name": f["stem"], "image_url": f["url"]} for f in files]


def resolve_all_images(data: dict):
    """
    Walk all image URL fields in data and resolve folder/Drive/GitHub URLs
    to final direct image URLs. Modifies data in place.
    """
    img = data.setdefault("image_urls", {})

    print("  Resolving image URLs...")

    # Single-image sections
    for key in ["ceo_photo", "campaign_banner", "hr_wellness_banner"]:
        raw = img.get(key, "")
        if raw:
            img[key] = resolve_single_url(raw)
            print(f"    {key}: {img[key][:80] if img[key] else '(not found)'}")

    # Delivery Insights — resolve each entry's logo
    for entry in data.get("delivery_insights", []):
        if entry.get("logo_url"):
            entry["logo_url"] = resolve_single_url(entry["logo_url"])
    print(f"    delivery_insights: {len(data.get('delivery_insights', []))} entry(ies)")

    # Acknowledging Excellence — resolve each entry's logo
    for entry in data.get("acknowledging_excellence", []):
        if entry.get("logo_url"):
            entry["logo_url"] = resolve_single_url(entry["logo_url"])
    print(f"    acknowledging_excellence: {len(data.get('acknowledging_excellence', []))} entry(ies)")

    # HR events — each event's raw refs (folder or single URLs) expand into its own photo list
    for event in data.get("hr", {}).get("events", []):
        raws = event.pop("_photos_raw", [])
        photos = []
        for raw in raws:
            photos.extend(resolve_folder_urls(raw))
        event["photos"] = photos
    n_events = len(data.get("hr", {}).get("events", []))
    n_event_photos = sum(len(e["photos"]) for e in data.get("hr", {}).get("events", []))
    print(f"    hr events: {n_events} event(s), {n_event_photos} photo(s) total")

    # Awards — folder URL in sheet (no title/name in row)
    award_folders = data.pop("_award_folder_urls", [])
    if award_folders:
        for folder_url in award_folders:
            awards = parse_awards_from_github_folder(folder_url)
            if awards:
                data["rewards"]["awards"] = awards
                print(f"    awards: {len(awards)} award(s) from folder")
                break
    else:
        # Explicit award rows — resolve per-recipient image_url
        for award in data["rewards"]["awards"]:
            for r in award["recipients"]:
                if r.get("image_url"):
                    r["image_url"] = resolve_single_url(r["image_url"])

    # New joinees — folder URL in sheet (name blank) OR explicit rows
    new_joinee_folders = data.pop("_new_joinee_folder_urls", [])
    if new_joinee_folders:
        for folder_url in new_joinee_folders:
            people = parse_new_joinees_from_github_folder(folder_url)
            if people:
                data["new_joinees"] = people
                print(f"    new_joinees: {len(people)} new joinee(s) from folder")
                break
    else:
        for person in data.get("new_joinees", []):
            if person.get("image_url"):
                person["image_url"] = resolve_single_url(person["image_url"])

    # New customers — folder URL in sheet (name blank) OR explicit rows
    new_customer_folders = data.pop("_new_customer_folder_urls", [])
    if new_customer_folders:
        for folder_url in new_customer_folders:
            customers = parse_new_customers_from_github_folder(folder_url)
            if customers:
                data["new_customers"] = customers
                print(f"    new_customers: {len(customers)} customer logo(s) from folder")
                break
    else:
        for customer in data.get("new_customers", []):
            if customer.get("image_url"):
                customer["image_url"] = resolve_single_url(customer["image_url"])

    # Announcements — folder URL in sheet (name blank) OR explicit rows
    announcement_folders = data.pop("_announcement_folder_urls", [])
    if announcement_folders:
        for folder_url in announcement_folders:
            announcements = parse_announcements_from_github_folder(folder_url)
            if announcements:
                data["announcements"] = announcements
                print(f"    announcements: {len(announcements)} announcement(s) from folder")
                break
    else:
        for announcement in data.get("announcements", []):
            if announcement.get("image_url"):
                announcement["image_url"] = resolve_single_url(announcement["image_url"])

    # Certifications — folder URL in sheet (name blank) OR explicit rows
    certification_folders = data.pop("_certification_folder_urls", [])
    if certification_folders:
        for folder_url in certification_folders:
            certifications = parse_certifications_from_github_folder(folder_url)
            if certifications:
                data["certifications"] = certifications
                print(f"    certifications: {len(certifications)} certificate(s) from folder")
                break
    else:
        for certification in data.get("certifications", []):
            if certification.get("image_url"):
                certification["image_url"] = resolve_single_url(certification["image_url"])

    # Blog images — if multiple blogs share a folder URL, assign images by index
    folder_cache: dict[str, list[str]] = {}
    folder_counter: dict[str, int] = {}
    for i, post in enumerate(data["marketing"]["blog_posts"], 1):
        raw = post.get("image_url", "")
        if not raw:
            continue
        info = parse_github_url(raw)
        if info and info["is_folder"]:
            if raw not in folder_cache:
                files = list_github_folder(info["owner"], info["repo"], info["branch"], info["path"])
                folder_cache[raw] = [f["url"] for f in files]
                folder_counter[raw] = 0
            idx = folder_counter[raw]
            urls = folder_cache[raw]
            post["image_url"] = urls[idx] if idx < len(urls) else ""
            folder_counter[raw] = idx + 1
        else:
            post["image_url"] = resolve_single_url(raw)
        print(f"    blog{i}: {post['image_url'][:80] if post['image_url'] else '(not found)'}")


def download_drive_xlsx(drive_url: str) -> openpyxl.Workbook:
    """Download an Excel file from Google Drive (must be shared publicly)."""
    fid = drive_file_id(drive_url)
    if not fid:
        raise ValueError(f"Could not extract file ID from: {drive_url}")
    url = f"https://drive.google.com/uc?export=download&id={fid}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp:
        return openpyxl.load_workbook(io.BytesIO(resp.read()))


# ── Anniversary parser ────────────────────────────────────────────────────

def parse_anniversaries(wb: openpyxl.Workbook, sheet_name: str) -> list:
    """
    Find the sheet named sheet_name (e.g. 'April-2026') in wb.
    Each sheet has two tables side by side:
      - Birthday table   (columns A–G)
      - Work Anniversary table (columns I onwards)
    Row 1: table title headers ('Birthday' | 'Work Anniversary')
    Row 2: column headers (E.Id | Employee Name | ... | Tenure | ...)
    Row 3+: data rows

    Extracts Employee Name + Tenure from the Work Anniversary table only,
    groups names by years, returns list sorted by years descending.
    """
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"  Using sheet: '{sheet_name}'")
    else:
        # Partial match fallback on the month part
        month_part = sheet_name.split("-")[0].lower()
        ws = None
        for name in wb.sheetnames:
            if month_part in name.lower():
                ws = wb[name]
                print(f"  Sheet '{sheet_name}' not found — using '{name}'")
                break
        if not ws:
            print(f"  WARNING: No sheet found for '{sheet_name}'. Available: {wb.sheetnames}")
            return []

    # ── Find where 'Work Anniversary' header is (row 1) ──────────────────
    # The sheet layout: Birthday header in col A, Work Anniversary header in col I
    wa_col_start = None
    title_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    for col_idx, cell in enumerate(title_row, start=1):
        if cell and "work anniversary" in str(cell).lower():
            wa_col_start = col_idx
            break

    if not wa_col_start:
        print("  WARNING: Could not find 'Work Anniversary' header in row 1.")
        return []

    # ── Find Employee Name & Tenure columns in row 2, right of wa_col_start ─
    header_row = next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    name_col = tenure_col = None

    for col_idx, cell in enumerate(header_row, start=1):
        if col_idx < wa_col_start:
            continue          # skip Birthday table columns
        text = str(cell).lower().strip() if cell else ""
        if "employee name" in text and name_col is None:
            name_col = col_idx
        if text == "tenure" and tenure_col is None:
            tenure_col = col_idx

    if not name_col or not tenure_col:
        print(f"  WARNING: Could not locate Employee Name / Tenure in Work Anniversary table.")
        return []

    print(f"  Work Anniversary starts at col {wa_col_start} | name_col={name_col} | tenure_col={tenure_col}")

    # ── Read data rows (row 3 onwards) ────────────────────────────────────
    year_map: dict[int, list[str]] = {}

    for row in ws.iter_rows(min_row=3, values_only=True):
        name   = cell_str(row[name_col - 1])   if name_col   <= len(row) else ""
        tenure = cell_str(row[tenure_col - 1]) if tenure_col <= len(row) else ""
        if not name or not tenure:
            continue
        years = parse_years(tenure)
        if years == 0:
            continue
        year_map.setdefault(years, []).append(name)

    return [{"years": y, "names": names} for y, names in sorted(year_map.items(), reverse=True)]


# ── New Openings (external sheet) parser ────────────────────────────────────

def parse_openings_sheet(wb: openpyxl.Workbook, month: str) -> list[dict]:
    """
    Parse an external 'New Openings' sheet, columns: Month | Position | Experience | Opening.
    Row 1: title, Row 2: headers, Row 3+: data — Month is only filled on the
    first row of each month's block, so it's forward-filled down.
    Returns positions for the given month only.
    """
    ws = wb[wb.sheetnames[0]]
    positions = []
    current_month = None
    for row in ws.iter_rows(min_row=3, values_only=True):
        month_cell = cell_str(row[0]) if len(row) > 0 else ""
        if month_cell:
            current_month = month_cell
        if not current_month or current_month.strip().lower() != (month or "").strip().lower():
            continue
        title = cell_str(row[1]) if len(row) > 1 else ""
        if not title:
            continue
        experience = cell_str(row[2]) if len(row) > 2 else ""
        opening_raw = row[3] if len(row) > 3 else None
        try:
            count = int(float(opening_raw)) if opening_raw not in (None, "") else 1
        except (ValueError, TypeError):
            count = 1
        positions.append({"title": title, "openings": count, "experience": experience})
    return positions


# ── Main content sheet parser ─────────────────────────────────────────────

def parse_content_sheet(wb: openpyxl.Workbook) -> dict:
    """Parse the Newsletter Content sheet into the content.json structure."""

    sheet_name = "Newsletter Content" if "Newsletter Content" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    data = {
        "month": None,
        "year":  None,
        "ceo_desk": {},
        "delivery_insights": [],
        "acknowledging_excellence": [],
        "hr": {"events": []},
        "rewards": {"awards": []},
        "anniversaries": [],
        "new_joinees": [],
        "new_customers": [],
        "announcements": [],
        "certifications": [],
        "new_openings": {
            "positions": [],
            "view_all_url": "https://www.biztechcs.com/career/",
        },
        "marketing": {
            "blog_posts": [],
            "view_all_url": "https://www.biztechcs.com/blog/",
        },
        "image_urls": {
            "ceo_photo": "",
            "campaign_banner": "",
            "hr_wellness_banner": "",
        },
        "footer": FOOTER,
    }

    award_map: dict = {}   # award title → award dict
    _legacy_delivery: dict = {}    # merges old single-row delivery_description/delivery_logo fields
    _legacy_excellence: dict = {}  # merges old single-row excellence_client/_testimonial/_logo fields
    _legacy_hr_event_title = ""    # merges old single-event hr_event_title field
    _legacy_event_photos_raw: list = []  # merges old single-event event_photo rows

    for row in ws.iter_rows():  # Cell objects, not values_only — the Drive Link
                                 # column needs .hyperlink (see below), since a
                                 # cell inserted as "Insert link" with custom
                                 # display text stores the URL only there, not
                                 # in .value.
        field = cell_str(row[0].value) if row else ""

        # Skip blank rows, section dividers, legend row
        if not field or field.startswith("▸") or field.lower() == "field":
            continue

        val    = row[1].value if len(row) > 1 else None
        extra  = row[2].value if len(row) > 2 else None
        extra2 = row[3].value if len(row) > 3 else None
        drive_cell = row[4] if len(row) > 4 else None
        drive = drive_cell.value if drive_cell is not None else None
        if (
            not (isinstance(drive, str) and drive.strip().lower().startswith("http"))
            and drive_cell is not None
            and drive_cell.hyperlink
            and drive_cell.hyperlink.target
        ):
            # The cell's own text isn't a usable URL — fall back to the
            # hyperlink target, for a cell inserted as a rich link with a
            # custom display label (e.g. Google Sheets "Insert link").
            # Only used as a fallback: when the visible text IS already a
            # URL, trust it over the hyperlink — copy-pasting a row and
            # editing its visible URL text can leave the underlying
            # hyperlink target stale (still pointing at the original row's
            # link), which would otherwise silently override correct data.
            drive = drive_cell.hyperlink.target

        # ── General ──────────────────────────────────────────────────────
        if   field == "month":        data["month"] = cell_str(val)
        elif field == "year":         data["year"]  = str(int(float(str(val)))) if val else None
        elif field == "ceo_name":     data["ceo_desk"]["ceo_name"]  = cell_str(val)
        elif field == "ceo_title":    data["ceo_desk"]["ceo_title"] = cell_str(val)
        elif field == "ceo_message":  data["ceo_desk"]["message"]   = cell_str(val)

        # ── Image rows — store raw URL, resolved later ───────────────────
        elif field == "ceo_photo":
            data["image_urls"]["ceo_photo"] = cell_str(drive)
        elif field == "campaign_banner":
            data["image_urls"]["campaign_banner"] = cell_str(drive)
        elif field == "hr_wellness_banner":
            data["image_urls"]["hr_wellness_banner"] = cell_str(drive)
            if cell_str(val):
                data["hr"]["wellness_title"] = cell_str(val)
        elif field == "event_photo":   # legacy single-event rows
            raw = cell_str(drive)
            if raw:
                _legacy_event_photos_raw.append(raw)

        # ── Delivery Insights (repeatable: one row per entry) ─────────────
        elif field == "delivery":
            data["delivery_insights"].append({
                "description": cell_str(val),
                "logo_url":    cell_str(drive),   # resolved later
            })
        elif field == "delivery_description":   # legacy single-entry fields
            _legacy_delivery["description"] = cell_str(val)
        elif field == "delivery_logo":
            _legacy_delivery["logo_url"] = cell_str(drive)

        # ── Acknowledging Excellence (repeatable: one row per entry) ──────
        elif field == "excellence":
            data["acknowledging_excellence"].append({
                "client_label": cell_str(extra),
                "testimonial":  cell_str(val),
                "logo_url":     cell_str(drive),   # resolved later
            })
        elif field == "excellence_client":       # legacy single-entry fields
            _legacy_excellence["client_label"] = cell_str(val)
        elif field == "excellence_testimonial":
            _legacy_excellence["testimonial"] = cell_str(val)
        elif field == "excellence_logo":
            _legacy_excellence["logo_url"] = cell_str(drive)

        # ── HR Insider ────────────────────────────────────────────────────
        elif field == "hr_wellness_title":
            data["hr"]["wellness_title"] = cell_str(val)
        elif re.match(r"^hr_event\d*$", field):   # repeatable: 1 row per event (Value=title, Drive Link=photo folder); accepts hr_event, hr_event1, hr_event2, ...
            title = cell_str(val)
            raw   = cell_str(drive)
            if title or raw:
                data["hr"]["events"].append({"title": title, "_photos_raw": [raw] if raw else []})
        elif field == "hr_event_title":   # legacy single-event field
            _legacy_hr_event_title = cell_str(val)

        # ── Awards ────────────────────────────────────────────────────────
        elif field == "award":
            title = cell_str(val)
            raw   = cell_str(drive)
            if not title:
                # Folder-based: images named "Title--Name--Designation.jpg"
                if raw and raw not in data.get("_award_folder_urls", []):
                    data.setdefault("_award_folder_urls", []).append(raw)
            else:
                # Explicit: title/name/designation in columns
                name  = cell_str(extra) or None
                desig = cell_str(extra2)
                if title not in award_map:
                    award_map[title] = {"title": title, "folder": "", "recipients": []}
                if name:
                    award_map[title]["recipients"].append({
                        "name":        name,
                        "designation": desig,
                        "image_url":   raw,   # resolved later
                    })

        # ── Employee Certifications ────────────────────────────────────────
        elif field == "employee_certification":
            title = cell_str(val)
            raw   = cell_str(drive)
            if not title:
                # Folder-based: any certificate image files, no naming convention required
                if raw and raw not in data.get("_certification_folder_urls", []):
                    data.setdefault("_certification_folder_urls", []).append(raw)
            else:
                # Explicit row
                data["certifications"].append({
                    "name":      title,
                    "image_url": raw,   # resolved later
                })

        # ── New Additions (new joinees) ──────────────────────────────────
        elif field == "new_joinee":
            name = cell_str(val)
            raw  = cell_str(drive)
            if not name:
                # Folder-based: images named "Name--Designation.jpg"
                if raw and raw not in data.get("_new_joinee_folder_urls", []):
                    data.setdefault("_new_joinee_folder_urls", []).append(raw)
            else:
                # Explicit row
                data["new_joinees"].append({
                    "name":        name,
                    "designation": cell_str(extra),
                    "image_url":   raw,   # resolved later
                })

        # ── Newly Added Customers ─────────────────────────────────────────
        elif field == "new_customer":
            name = cell_str(val)
            raw  = cell_str(drive)
            if not name:
                # Folder-based: any logo image files, no naming convention required
                if raw and raw not in data.get("_new_customer_folder_urls", []):
                    data.setdefault("_new_customer_folder_urls", []).append(raw)
            else:
                # Explicit row
                data["new_customers"].append({
                    "name":      name,
                    "image_url": raw,   # resolved later
                })

        # ── Upcoming Event Announcements ──────────────────────────────────
        elif field == "new_announcement":
            title = cell_str(val)
            raw   = cell_str(drive)
            if not title:
                # Folder-based: any banner image files, no naming convention required
                if raw and raw not in data.get("_announcement_folder_urls", []):
                    data.setdefault("_announcement_folder_urls", []).append(raw)
            else:
                # Explicit row
                data["announcements"].append({
                    "name":      title,
                    "image_url": raw,   # resolved later
                })

        # ── New Openings ──────────────────────────────────────────────────
        elif field == "opening":
            if val:
                try:
                    count = int(float(str(extra))) if extra else 1
                except ValueError:
                    count = 1
                data["new_openings"]["positions"].append({
                    "title":      cell_str(val),
                    "openings":   count,
                    "experience": cell_str(extra2),
                })
        elif field == "new_openings_sheet":
            # External sheet link — Drive Link column = Google Sheets URL.
            # When present, its data replaces any "opening" rows above.
            url = cell_str(drive)
            if url and url.startswith("http"):
                data["_openings_sheet_url"] = url

        # ── Marketing Blogs (blog1, blog2, …) ────────────────────────────
        elif re.match(r"^blog\d*$", field):
            if val:
                data["marketing"]["blog_posts"].append({
                    "title":     cell_str(val),
                    "url":       cell_str(extra),
                    "image_url": cell_str(drive),   # resolved later
                })

        # ── Anniversary — tab name in Value, Drive file in Drive Link ───────
        elif field == "anniversary":
            # Value may be a date object (Excel stores "March-2026" as a date)
            if hasattr(val, "strftime"):
                tab = val.strftime("%B-%Y")   # datetime → "March-2026"
            else:
                tab = cell_str(val)
            url = cell_str(drive)             # Drive link to the xlsx file
            if tab:
                data["_anniv_tab"] = tab
            if url and url.startswith("http"):
                data["_anniv_drive_url"] = url

        # Unknown fields: silently skip

    data["rewards"]["awards"] = list(award_map.values())

    if not data["delivery_insights"] and _legacy_delivery:
        data["delivery_insights"].append(_legacy_delivery)
    if not data["acknowledging_excellence"] and _legacy_excellence:
        data["acknowledging_excellence"].append(_legacy_excellence)
    if not data["hr"]["events"] and (_legacy_hr_event_title or _legacy_event_photos_raw):
        data["hr"]["events"].append({"title": _legacy_hr_event_title, "_photos_raw": _legacy_event_photos_raw})

    return data


# ── Main ──────────────────────────────────────────────────────────────────

def main():
    # Keep this script's own prints interleaved correctly with the generate.py /
    # screenshot_sections.py subprocess output below (stdout is block-buffered,
    # not line-buffered, when piped instead of attached to a real terminal).
    sys.stdout.reconfigure(line_buffering=True)

    # ── Step 1: Newsletter content ────────────────────────────────────────
    print("Downloading Newsletter Content sheet...")
    content_wb = download_xlsx(CONTENT_SHEET_ID)
    data = parse_content_sheet(content_wb)
    print(f"  month={data['month']}, year={data['year']}")
    print(f"  Awards : {[a['title'] for a in data['rewards']['awards']]}")
    print(f"  New Joinees: {len(data['new_joinees'])}")
    print(f"  New Customers: {len(data['new_customers'])}")
    print(f"  Announcements: {len(data['announcements'])}")
    print(f"  Certifications: {len(data['certifications'])}")
    print(f"  Blogs   : {len(data['marketing']['blog_posts'])}")

    # ── Step 1b: New Openings — external sheet link overrides sheet rows ───
    openings_sheet_url = data.pop("_openings_sheet_url", None)
    if openings_sheet_url:
        print(f"Downloading New Openings sheet...")
        try:
            openings_sheet_id = drive_file_id(openings_sheet_url)
            openings_wb = download_xlsx(openings_sheet_id)
            positions = parse_openings_sheet(openings_wb, data["month"])
            if positions:
                data["new_openings"]["positions"] = positions
            else:
                print(f"  WARNING: No openings found for month '{data['month']}' in external sheet.")
        except Exception as e:
            print(f"  WARNING: Could not load New Openings sheet ({e}); using sheet-row data instead.")
    print(f"  Openings: {len(data['new_openings']['positions'])}")

    # ── Step 2: Work Anniversary data ────────────────────────────────────
    anniv_tab       = data.pop("_anniv_tab",       None)
    anniv_drive_url = data.pop("_anniv_drive_url", None)

    # Sheet name to look up — from content sheet row, else construct from month/year
    sheet_name = anniv_tab or f"{data['month']}-{data['year']}"

    print(f"\nDownloading Work Anniversary file (tab: '{sheet_name}')...")
    if anniv_drive_url:
        try:
            anniv_wb = download_drive_xlsx(anniv_drive_url)
            print("  Downloaded from Drive link in sheet.")
        except Exception as e:
            print(f"  Drive download failed ({e}), falling back to Sheet ID...")
            anniv_wb = load_xlsx(ANNIV_SHEET_ID, "anniversaries.xlsx")
    else:
        anniv_wb = load_xlsx(ANNIV_SHEET_ID, "anniversaries.xlsx")

    anniversaries = parse_anniversaries(anniv_wb, sheet_name)
    data["anniversaries"] = anniversaries

    if anniversaries:
        print(f"  {len(anniversaries)} anniversary group(s):")
        for a in anniversaries:
            label = "year" if a["years"] == 1 else "years"
            print(f"    {a['years']} {label}: {', '.join(a['names'])}")
    else:
        print("  No anniversary data found.")

    # ── Step 3: Resolve all image URLs (GitHub folders → direct URLs) ───
    print("\nResolving images...")
    resolve_all_images(data)

    # ── Step 4: Write content.json ────────────────────────────────────────
    OUTPUT.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nDone -> {OUTPUT}")

    # ── Step 5: Render browser-preview HTML + refresh per-section JPGs ─────
    print("\nRendering updated newsletter and section images...")
    subprocess.run([sys.executable, str(BASE_DIR / "generate.py")], check=True)
    subprocess.run([sys.executable, str(BASE_DIR / "screenshot_sections.py")], check=True)

    print("\nNext step (only if sending via Gmail): python generate.py --email")


if __name__ == "__main__":
    main()
