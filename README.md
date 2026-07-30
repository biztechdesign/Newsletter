# Biztech Newsletter

Builds the monthly newsletter from one spreadsheet and a folder of photos, and
produces an email that renders in both Gmail and Outlook.

---

## What's in this folder

```
Newsletter/
│
├── August-2026/              ← ONE FOLDER PER ISSUE. Everything for that month.
│   ├── content.xlsx              the sheet people fill in
│   ├── Row Data/                 the photos you drop in
│   ├── Section wise images/      generated: one PNG per section
│   └── HTML/                     generated: the newsletter + the email
│
├── assets/                   logo, section icons, balloons (rarely change)
├── images/                   older photos, still served to already-sent emails
├── templates/                the HTML layouts — edit to change the design
│
├── master_template.xlsx      the blank sheet you share with contributors
├── .env                      your upload credentials (never committed)
└── *.py                      the scripts, described below
```

**You only ever edit two things:** `content.xlsx` and the photos in `Row Data/`.
Everything under `Section wise images/` and `HTML/` is regenerated each run —
don't edit those by hand, your changes will be overwritten.

---

## Making a newsletter

### 1. Start the month

```
python create_master_template.py August-2026/content.xlsx
```

Creates the sheet and all 17 photo folders. Share the sheet with the team.

### 2. Fill it in

- Type the content into the **Content** column of `content.xlsx`
- Drop photos into the matching folder under `Row Data/`
- Nothing this month? Leave Content blank, or write **NA** — the section is
  left out of the newsletter entirely

See `Row Data/README.txt` for the filename rules (awards and new joiners take
the person's name from the filename).

### 3. Build and review

```
python import_content.py --sheet=August-2026/content.xlsx
```

Writes `August-2026/HTML/newsletter_August_2026.html`. Open it and check it.
Repeat this step until it looks right — nothing is published yet.

*Tip:* `python watch.py --sheet=August-2026/content.xlsx` rebuilds and reloads
the page in your browser every time you save the sheet or add a photo.

### 4. Cut the section images

```
python screenshot_sections.py
```

Flattens each section to a PNG in `Section wise images/`. This is what makes
the email work in Outlook, which can't render the real layout.

### 5. Upload them

```
python upload_sections.py --dry-run     check what will be sent
python upload_sections.py               send it
```

Needs `.env` — copy `.env.example` and fill in the three values.

### 6. Send

```
python generate_simple_email.py
```

Open `August-2026/HTML/newsletter_August_2026_simple_email.html`, select all,
copy, paste into Gmail.

Not uploaded yet and want to look first? `python generate_simple_email.py --local`
writes a copy that reads the images off your disk. You can paste that into
Gmail too — Gmail uploads the images itself.

---

## The scripts

| Script | What it does |
|---|---|
| `create_master_template.py` | makes the blank sheet + photo folders for a new month |
| `migrate_to_master.py` | converts an old-format sheet to the current one |
| `import_content.py` | reads the sheet + photos, renders the newsletter |
| `watch.py` | rebuilds automatically while you edit |
| `screenshot_sections.py` | cuts each section to a PNG |
| `upload_sections.py` | sends those PNGs to the image host |
| `generate_simple_email.py` | builds the final Gmail/Outlook email |
| `generate.py` | renders the HTML (called by `import_content.py`) |

---

## Things that catch people out

**iPhone photos.** HEIC won't display in email. Renaming it `.png` doesn't
convert it — the run warns and skips the file. Re-save it as PNG or JPEG.

**Empty folder = last month's photos.** If a section's folder is empty but the
sheet still links to the old GitHub folder, the previous issue's images appear.
The run prints a `NOTE:` line whenever that happens.

**Order comes from filenames.** Prefix `01 -`, `02 -` where sequence matters.

**Big photos.** An 18MB source photo makes the section PNG enormous and slow to
load. Resize to about 1600px wide before dropping it in.
