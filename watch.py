"""
watch.py
========
Rebuilds the newsletter automatically whenever you save the month's Excel file
(or add/replace a photo in its Row Data folder), and live-reloads the preview
in your browser so you can see the change without touching anything.

Usage:
    python watch.py                                # picks the newest <Month>-<Year>/content.xlsx
    python watch.py --sheet=July-2026/content.xlsx # or name it explicitly
    python watch.py --port=8080
    python watch.py --no-open                      # don't launch a browser

Stop it with Ctrl+C.

What it watches:
    <Month>-<Year>/content.xlsx     the month's content sheet
    <Month>-<Year>/Row Data/**      that month's source photos

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

BASE_DIR = Path(__file__).parent.resolve()

POLL_SECONDS = 1.0
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


# ── Change detection ───────────────────────────────────────────────────────

def snapshot(sheet: Path, row_data: Path) -> dict:
    """Map every watched file to (mtime, size). Excel's ~$ lock files are
    ignored — they appear and vanish on their own while a workbook is open."""
    state = {}
    targets = [sheet]
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

def rebuild(sheet: Path) -> bool:
    result = subprocess.run(
        [sys.executable, str(BASE_DIR / "import_content.py"), f"--sheet={sheet}"],
        cwd=str(BASE_DIR),
    )
    return result.returncode == 0


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

    sheet = find_sheet()
    row_data = sheet.parent / "Row Data"
    port = int(arg_value("port", "8000"))

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
    try:
        while True:
            time.sleep(POLL_SECONDS)
            if snapshot(sheet, row_data) == state:
                continue

            print("-" * 70)
            print("Change detected - rebuilding...")
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
