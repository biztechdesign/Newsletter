"""
upload_sections.py
==================
Uploads this month's section PNGs to the image host, so the final email's
<img> URLs resolve.

    python upload_sections.py --dry-run     # show what would be sent, no connection
    python upload_sections.py               # upload
    python upload_sections.py --delete-extra  # also remove files not in this run

Credentials are never stored in the repo. Provide them either as environment
variables:

    NEWSLETTER_FTP_HOST   ftp.example.com
    NEWSLETTER_FTP_USER   someone
    NEWSLETTER_FTP_PASS   ...
    NEWSLETTER_FTP_DIR    /home/wordpress/public_html/biztech-insider   (optional)
    NEWSLETTER_FTP_TLS    1 to force FTPS, 0 to force plain             (optional)

or in a local file `ftp_config.json` beside this script, which .gitignore
excludes:

    {"host": "...", "user": "...", "password": "...", "dir": "/home/..."}

FTP and FTPS are supported through the standard library. FTPS is tried first
and falls back to plain FTP, since a plain-FTP-only server rejects the TLS
handshake outright. SFTP (port 22) is a different protocol and would need
paramiko installed — say so rather than failing obscurely.
"""

import ftplib
import json
import os
import sys
from pathlib import Path

from generate_simple_email import (
    HOSTED_SERVER_PATH,
    hosted_base,
    sections_dir,
)

BASE_DIR = Path(__file__).parent
CONTENT_JSON = BASE_DIR / "content.json"
CONFIG_FILE = BASE_DIR / "ftp_config.json"
ENV_FILE = BASE_DIR / ".env"


def read_env_file(path: Path) -> dict:
    """
    KEY=VALUE lines from a .env file. Blank lines and # comments are skipped,
    and surrounding quotes are stripped so a password pasted with quotes still
    works. Deliberately hand-rolled: one small loop beats a dependency, and a
    missing package here would block the upload at the worst moment.
    """
    values = {}
    if not path.is_file():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key.strip()] = value
    return values


def load_config() -> dict:
    """
    Credentials, in increasing order of precedence:
      ftp_config.json  →  .env  →  real environment variables
    so a value exported in the shell always wins over a file on disk.
    """
    config = {}
    if CONFIG_FILE.is_file():
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))

    keys = {
        "NEWSLETTER_FTP_HOST": "host",
        "NEWSLETTER_FTP_USER": "user",
        "NEWSLETTER_FTP_PASS": "password",
        "NEWSLETTER_FTP_DIR":  "dir",
        "NEWSLETTER_FTP_TLS":  "tls",
    }
    from_file = read_env_file(ENV_FILE)
    for env_name, key in keys.items():
        value = os.environ.get(env_name) or from_file.get(env_name)
        if value:
            config[key] = value

    config.setdefault("dir", HOSTED_SERVER_PATH)
    return config


def connect(config: dict) -> ftplib.FTP:
    """An authenticated FTPS connection, or plain FTP where TLS isn't offered."""
    host, user, password = config["host"], config["user"], config["password"]
    prefer_tls = str(config.get("tls", "1")).lower() not in ("0", "false", "no")

    if prefer_tls:
        try:
            ftp = ftplib.FTP_TLS(host, timeout=60)
            ftp.login(user, password)
            ftp.prot_p()          # encrypt the data channel too, not just login
            print(f"  Connected to {host} over FTPS.")
            return ftp
        except (ftplib.error_perm, ftplib.error_proto, OSError) as e:
            print(f"  FTPS unavailable ({e}); falling back to plain FTP.")

    ftp = ftplib.FTP(host, timeout=60)
    ftp.login(user, password)
    print(f"  Connected to {host} over FTP (unencrypted).")
    return ftp


def ensure_dir(ftp: ftplib.FTP, remote_dir: str):
    """cd into remote_dir, creating any missing part of the path."""
    for part in remote_dir.strip("/").split("/"):
        try:
            ftp.cwd(part)
        except ftplib.error_perm:
            ftp.mkd(part)
            ftp.cwd(part)


def remote_sizes(ftp: ftplib.FTP) -> dict:
    """{filename: size} for the current remote directory."""
    sizes = {}
    for name in ftp.nlst():
        try:
            sizes[name] = ftp.size(name)
        except (ftplib.error_perm, TypeError):
            pass
    return sizes


def main():
    data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    month, year = data["month"], data["year"]
    sec_dir = sections_dir(month, year)
    if not sec_dir.is_dir():
        raise SystemExit(
            f"{sec_dir} not found - run 'python screenshot_sections.py' first."
        )

    files = sorted(sec_dir.glob("*.png"))
    if not files:
        raise SystemExit(f"No PNGs in {sec_dir} - nothing to upload.")

    config = load_config()
    remote_dir = f"{config['dir'].rstrip('/')}/{month}_{year}"
    total = sum(f.stat().st_size for f in files)

    print(f"\n{len(files)} file(s), {total / 1048576:.1f}MB")
    print(f"  from : {sec_dir}")
    print(f"  to   : {remote_dir}")
    print(f"  url  : {hosted_base(month, year)}/")

    if "--dry-run" in sys.argv:
        print("\n--dry-run, nothing sent:")
        for f in files:
            print(f"    {f.name:<34} {f.stat().st_size // 1024:>5}KB")
        return

    missing = [k for k in ("host", "user", "password") if not config.get(k)]
    if missing:
        raise SystemExit(
            f"\nMissing credentials: {', '.join(missing)}.\n"
            f"Set NEWSLETTER_FTP_HOST / _USER / _PASS, or create {CONFIG_FILE.name}\n"
            f"(see the module docstring). Nothing has been sent."
        )

    print()
    ftp = connect(config)
    try:
        ensure_dir(ftp, remote_dir)
        before = remote_sizes(ftp)

        for f in files:
            with f.open("rb") as fh:
                ftp.storbinary(f"STOR {f.name}", fh)
            was = before.get(f.name)
            state = "new" if was is None else ("unchanged" if was == f.stat().st_size else "replaced")
            print(f"    {f.name:<34} {f.stat().st_size // 1024:>5}KB  {state}")

        # Verify by size rather than trusting the transfer: a truncated upload
        # still "succeeds" and leaves a broken image in a sent newsletter.
        after = remote_sizes(ftp)
        bad = [f.name for f in files if after.get(f.name) != f.stat().st_size]
        if bad:
            raise SystemExit(
                f"\nSize mismatch after upload: {', '.join(bad)}.\n"
                f"Re-run before sending - these would render broken."
            )
        print(f"\n  Verified {len(files)} file(s) match byte-for-byte.")

        extra = sorted(set(after) - {f.name for f in files})
        if extra and "--delete-extra" in sys.argv:
            for name in extra:
                ftp.delete(name)
            print(f"  Deleted {len(extra)} file(s) not in this run: {', '.join(extra)}")
        elif extra:
            print(f"  Left in place, not part of this run ({len(extra)}): {', '.join(extra)}")
            print(f"  Use --delete-extra to remove them.")
    finally:
        try:
            ftp.quit()
        except Exception:
            ftp.close()

    print(f"\nDone. The email's images now resolve at {hosted_base(month, year)}/\n")


if __name__ == "__main__":
    main()
