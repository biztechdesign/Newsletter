"""
watch.py
========
Rebuilds the newsletter automatically whenever you save the month's Excel file
(or add/replace a photo in its Row Data folder), and live-reloads the preview
in your browser so you can see the change without touching anything.

Usage:
    python watch.py --online                       # follow the live Google Sheet
    python watch.py                                # newest local <Month>-<Year>/content.xlsx
    python watch.py --sheet=July-2026/content.xlsx # a specific local sheet
    python watch.py --port=8080
    python watch.py --no-open                      # don't launch a browser

Stop it with Ctrl+C.

What it watches:
    --online:  the Google Sheet itself, re-fetched every 30s, so edits made in
               the browser show up with nothing to download or save
    otherwise: <Month>-<Year>/content.xlsx
    both:      <Month>-<Year>/Row Data/**   that month's source photos

On any change it re-runs import_content.py, which rewrites content.json and
re-renders the preview HTML. It deliberately does NOT re-cut the section PNGs
— those stay a separate, deliberate step once the preview is approved
(screenshot_sections.py), same as the normal flow.

The preview is served over HTTP rather than opened as a file:// page because
the rendered HTML points at assets with relative "../assets/..." paths, which
only resolve when the project root is the server root.
"""

import subprocess
import sys
import threading
import time
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

# The project root — one level up now that the scripts live in scripts/.
# Every path below (content.json, assets, month folders) hangs off it.
BASE_DIR = Path(__file__).resolve().parent.parent.resolve()
SCRIPTS_DIR = Path(__file__).resolve().parent   # this folder, for launching siblings

# Which workbook --online polls. Imported so there is one definition of it.
sys.path.insert(0, str(SCRIPTS_DIR))
from import_content import CONTENT_SHEET_ID  # noqa: E402

POLL_SECONDS = 1.0
# How often --online re-fetches the Google Sheet. Slower than the local poll on
# purpose: each check downloads the whole workbook.
ONLINE_POLL_SECONDS = 30.0
# A save is only acted on once the files have stopped changing for this long —
# Excel writes to a temp file and renames, so a rebuild fired at the wrong
# instant would read a half-written workbook.
SETTLE_SECONDS = 1.5

LIVE_RELOAD = """
<script>
(function () {
  var last = null;
  setInterval(function () {
    fetch('/__mtime?f=' + encodeURIComponent(location.pathname), { cache: 'no-store' })
      .then(function (r) { return r.text(); })
      .then(function (t) {
        if (last === null) { last = t; return; }
        if (t !== last) { location.reload(); }
      })
      .catch(function () { /* server restarting or offline — try again next tick */ });
  }, 1000);
})();
</script>
"""


# ── Arguments ──────────────────────────────────────────────────────────────

def arg_value(name: str, default=None):
    prefix = f"--{name}="
    return next((a[len(prefix):] for a in sys.argv if a.startswith(prefix)), default)


def find_sheet() -> Path:
    """The explicitly named sheet, else the most recently modified month's."""
    named = arg_value("sheet")
    if named:
        path = Path(named)
        if not path.is_absolute():
            path = BASE_DIR / path
        if not path.is_file():
            raise SystemExit(f"--sheet file not found: {path}")
        return path

    candidates = sorted(
        (p for p in BASE_DIR.glob("*/content.xlsx") if not p.name.startswith("~$")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise SystemExit(
            "No <Month>-<Year>/content.xlsx found.\n"
            "Create one (or pass --sheet=<file.xlsx>) before running watch.py."
        )
    return candidates[0]


def online_fingerprint() -> str | None:
    """
    A hash of the live sheet's *content*, or None if it can't be fetched.

    Hashes the cell values of the month tab, not the downloaded file. Google's
    xlsx export is not byte-stable — two fetches seconds apart with no edit in
    between produce different bytes — so hashing the download would rebuild
    every single poll, forever.

    A failed fetch returns None rather than a hash, so a network blip doesn't
    read as "changed" and trigger a pointless rebuild.
    """
    import hashlib
    import io
    import urllib.request

    import openpyxl

    from import_content import pick_tab

    url = (f"https://docs.google.com/spreadsheets/d/{CONTENT_SHEET_ID}"
           f"/export?format=xlsx")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read()
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb[pick_tab(wb)]
        digest = hashlib.sha256()
        for row in ws.iter_rows(values_only=True):
            digest.update(repr(row).encode("utf-8", "replace"))
        wb.close()
        return digest.hexdigest()
    except Exception as e:
        print(f"  (couldn't read the sheet: {e})")
        return None


# ── Change detection ───────────────────────────────────────────────────────

def snapshot(sheet: Path, row_data: Path) -> dict:
    """Map every watched file to (mtime, size). Excel's ~$ lock files are
    ignored — they appear and vanish on their own while a workbook is open."""
    state = {}
    # sheet is None in --online mode: the workbook lives in Google Sheets, so
    # only the photo folder exists on disk to watch.
    targets = [sheet] if sheet is not None else []
    if row_data.is_dir():
        targets.extend(p for p in row_data.rglob("*") if p.is_file())
    for p in targets:
        if p.name.startswith("~$"):
            continue
        try:
            st = p.stat()
            state[str(p)] = (st.st_mtime, st.st_size)
        except OSError:
            # Mid-write or just deleted; the next poll will pick up the result.
            continue
    return state


def wait_until_settled(sheet: Path, row_data: Path) -> dict:
    """Poll until two consecutive snapshots match, so we never read a
    workbook Excel is still writing."""
    previous = snapshot(sheet, row_data)
    while True:
        time.sleep(SETTLE_SECONDS)
        current = snapshot(sheet, row_data)
        if current == previous:
            return current
        previous = current


# ── Preview server ─────────────────────────────────────────────────────────

class PreviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def log_message(self, *args):
        pass  # keep the console readable — the watcher prints its own status

    def _resolve(self, url_path: str) -> Path | None:
        """Map a URL path to a file inside BASE_DIR, or None if it escapes."""
        candidate = (BASE_DIR / url_path.lstrip("/")).resolve()
        if not candidate.is_relative_to(BASE_DIR):
            return None
        return candidate

    def _send_bytes(self, body: bytes, content_type: str):
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)

        # Polled by the injected script to decide when to reload.
        if parsed.path == "/__mtime":
            wanted = (parse_qs(parsed.query).get("f") or [""])[0]
            target = self._resolve(wanted)
            try:
                stamp = str(target.stat().st_mtime) if target else "0"
            except OSError:
                stamp = "0"
            self._send_bytes(stamp.encode(), "text/plain; charset=utf-8")
            return

        if parsed.path.endswith(".html"):
            target = self._resolve(parsed.path)
            if target and target.is_file():
                html = target.read_text(encoding="utf-8", errors="replace")
                if "</body>" in html:
                    html = html.replace("</body>", LIVE_RELOAD + "</body>", 1)
                else:
                    html += LIVE_RELOAD
                self._send_bytes(html.encode("utf-8"), "text/html; charset=utf-8")
                return

        super().do_GET()


def serve(port: int) -> ThreadingHTTPServer:
    server = ThreadingHTTPServer(("127.0.0.1", port), PreviewHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server


# ── Build ──────────────────────────────────────────────────────────────────

def rebuild(sheet: Path | None) -> bool:
    # Without --sheet, import_content.py downloads the live workbook itself.
    cmd = [sys.executable, str(SCRIPTS_DIR / "import_content.py")]
    if sheet is not None:
        cmd.append(f"--sheet={sheet}")
    return subprocess.run(cmd, cwd=str(BASE_DIR)).returncode == 0


def preview_url(port: int) -> str | None:
    """Newest rendered preview, as a URL under the server root."""
    # Each issue writes into its own "<Month>-<Year>/HTML" folder.
    previews = sorted(
        (p for p in BASE_DIR.glob("*/HTML/newsletter_*.html") if "_email" not in p.name),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not previews:
        return None
    rel = previews[0].relative_to(BASE_DIR).as_posix()
    return f"http://127.0.0.1:{port}/{rel}"


def main():
    # Without this, this script's own status lines sit in a block buffer while
    # the rebuild subprocess writes straight to the same stream — so the output
    # arrives badly out of order and "Change detected" never shows up when it
    # actually happened.
    sys.stdout.reconfigure(line_buffering=True)

    online = "--online" in sys.argv
    sheet = None if online else find_sheet()
    port = int(arg_value("port", "8000"))

    if online:
        # No local sheet: import_content.py downloads the workbook itself, so
        # only this month's photos are watched on disk.
        month_dir = max(
            (p.parent for p in BASE_DIR.glob("*/Row Data")),
            key=lambda p: p.stat().st_mtime,
            default=BASE_DIR,
        )
        row_data = month_dir / "Row Data"
        print(f"Watching sheet : the live Google Sheet, every {int(ONLINE_POLL_SECONDS)}s")
        print(f"                 (edit it in the browser - no download needed)")
    else:
        row_data = sheet.parent / "Row Data"
        print(f"Watching sheet : {sheet}")
    print(f"Watching photos: {row_data}{'' if row_data.is_dir() else '  (not created yet)'}")
    print()

    print("Building once so there's something to show...")
    rebuild(sheet)

    try:
        server = serve(port)
    except OSError as e:
        raise SystemExit(f"Could not start the preview server on port {port}: {e}\nTry --port=8080")

    url = preview_url(port)
    if url:
        print(f"\nPreview: {url}")
        if "--no-open" not in sys.argv:
            webbrowser.open(url)
    else:
        print("\nNo preview HTML yet — fix the errors above and save the sheet again.")

    print("\nWatching for changes. Save the Excel (or drop photos into Row Data) and")
    print("the preview rebuilds and reloads by itself. Ctrl+C to stop.\n")

    state = snapshot(sheet, row_data)
    remote = online_fingerprint() if online else None
    last_online_check = time.monotonic()
    try:
        while True:
            time.sleep(POLL_SECONDS)

            changed_locally = snapshot(sheet, row_data) != state
            changed_online = False
            if online and time.monotonic() - last_online_check >= ONLINE_POLL_SECONDS:
                last_online_check = time.monotonic()
                current = online_fingerprint()
                # None means the fetch failed; only a real, different hash counts.
                if current is not None and remote is not None and current != remote:
                    changed_online = True
                if current is not None:
                    remote = current

            if not (changed_locally or changed_online):
                continue

            print("-" * 70)
            print("Sheet edited - rebuilding..." if changed_online
                  else "Change detected - rebuilding...")
            state = wait_until_settled(sheet, row_data)
            ok = rebuild(sheet)
            # Re-snapshot: a rebuild can touch watched files itself, and that
            # must not count as the next change.
            state = snapshot(sheet, row_data)
            print("Rebuilt - the browser reloads on its own." if ok
                  else "Build FAILED - fix the above, then save again.")
            print("-" * 70 + "\n")
    except KeyboardInterrupt:
        print("\nStopped.")
        server.shutdown()


if __name__ == "__main__":
    main()
