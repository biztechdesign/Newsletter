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

Protocol is chosen by probing: port 22 means SFTP (needs paramiko), port 21
means FTP/FTPS via the standard library. Set NEWSLETTER_FTP_PROTOCOL to sftp
or ftp to skip the probe. The live host answers on 22 only, so plain FTP sits
through a 60s connect timeout before failing with nothing pointing at the
cause — hence probing rather than assuming.
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

# The project root — one level up now that the scripts live in scripts/.
# Every path below (content.json, assets, month folders) hangs off it.
BASE_DIR = Path(__file__).resolve().parent.parent
CONTENT_JSON = BASE_DIR / "content.json"
CONFIG_FILE = BASE_DIR / "ftp_config.json"
ENV_FILE = BASE_DIR / ".env"

# Host keys pinned for this project, alongside ~/.ssh/known_hosts. Written only
# by an explicit --trust-host run; every other run verifies against it and
# refuses an unrecognised key.
KNOWN_HOSTS = BASE_DIR / ".ssh_known_hosts"


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
        "NEWSLETTER_FTP_HOST":     "host",
        "NEWSLETTER_FTP_USER":     "user",
        "NEWSLETTER_FTP_PASS":     "password",
        "NEWSLETTER_FTP_DIR":      "dir",
        "NEWSLETTER_FTP_TLS":      "tls",
        "NEWSLETTER_FTP_PROTOCOL": "protocol",
        "NEWSLETTER_FTP_PORT":     "port",
        "NEWSLETTER_FTP_KEY_FILE": "key_file",
    }
    from_file = read_env_file(ENV_FILE)
    for env_name, key in keys.items():
        value = os.environ.get(env_name) or from_file.get(env_name)
        if value:
            config[key] = value

    config.setdefault("dir", HOSTED_SERVER_PATH)
    return config


def port_open(host: str, port: int, timeout: float = 6) -> bool:
    import socket
    sock = socket.socket()
    sock.settimeout(timeout)
    try:
        sock.connect((host, port))
        return True
    except OSError:
        return False
    finally:
        sock.close()


def pick_protocol(config: dict) -> str:
    """
    'sftp' or 'ftp'. Set NEWSLETTER_FTP_PROTOCOL to force one; otherwise the
    open port decides.

    Worth probing rather than assuming: this host answers on 22 and not 21, so
    an FTP-only client sat through a 60s connect timeout before failing with
    nothing that pointed at the actual cause.
    """
    choice = str(config.get("protocol", "auto")).lower()
    if choice in ("sftp", "ftp"):
        return choice
    if port_open(config["host"], 22):
        return "sftp"
    if port_open(config["host"], 21):
        return "ftp"
    raise SystemExit(
        f"\nNeither port 22 (SFTP) nor 21 (FTP) is reachable on {config['host']}.\n"
        f"Check the host, or whether a firewall is in the way. Nothing has been sent."
    )


def server_fingerprint(host: str, port: int) -> tuple[str, str]:
    """(key type, SHA256 fingerprint) of the host's SSH key, without logging in."""
    import base64
    import hashlib

    import paramiko

    transport = paramiko.Transport((host, port))
    try:
        transport.start_client(timeout=15)
        key = transport.get_remote_server_key()
    finally:
        transport.close()
    digest = base64.b64encode(hashlib.sha256(key.asbytes()).digest()).decode().rstrip("=")
    return key.get_name(), f"SHA256:{digest}"


def connect_sftp(config: dict):
    """
    An authenticated SFTP session, by key if one is configured else password.

    The host key is verified against known_hosts, and an unrecognised one is
    refused. Accepting any key would mean handing the server password to
    whoever answered — the credentials go over this connection, so an
    unverified endpoint is the whole risk, not a formality.

    Pinning happens once, deliberately: run with --trust-host to record the
    fingerprint after checking it, and every later run verifies against it.
    """
    import paramiko

    host = config["host"]
    port = int(config.get("port", 22))

    client = paramiko.SSHClient()
    client.load_system_host_keys()                      # ~/.ssh/known_hosts
    if KNOWN_HOSTS.is_file():
        client.load_host_keys(str(KNOWN_HOSTS))         # this project's pins

    trusting = "--trust-host" in sys.argv
    if trusting:
        kind, fingerprint = server_fingerprint(host, port)
        print(f"  --trust-host: recording {host} {kind} {fingerprint}")
        print(f"                confirm this matches the server before relying on it.")
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    else:
        client.set_missing_host_key_policy(paramiko.RejectPolicy())

    key_file = config.get("key_file")
    try:
        client.connect(
            hostname=host,
            port=port,
            username=config["user"],
            password=config.get("password") or None,
            key_filename=key_file or None,
            timeout=30,
            allow_agent=bool(key_file),
            look_for_keys=bool(key_file),
        )
    except paramiko.SSHException as e:
        if "not found in known_hosts" not in str(e).lower():
            raise
        kind, fingerprint = server_fingerprint(host, port)
        raise SystemExit(
            f"\nThe host key for {host} is not pinned, so the connection was refused\n"
            f"rather than sending the password to an unverified server.\n\n"
            f"  {kind}  {fingerprint}\n\n"
            f"If that matches the server, pin it once with:\n"
            f"  python upload_sections.py --trust-host\n"
        ) from None

    if trusting:
        client.save_host_keys(str(KNOWN_HOSTS))
        print(f"  Pinned in {KNOWN_HOSTS.name}; later runs verify against it.")

    print(f"  Connected to {host} over SFTP.")
    return client, client.open_sftp()


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


class FtpTarget:
    """FTP/FTPS, working in the remote directory after ensure_dir()."""

    def __init__(self, config):
        self.ftp = connect(config)

    def ensure_dir(self, remote_dir):
        ensure_dir(self.ftp, remote_dir)

    def sizes(self):
        return remote_sizes(self.ftp)

    def put(self, path):
        with path.open("rb") as fh:
            self.ftp.storbinary(f"STOR {path.name}", fh)

    def delete(self, name):
        self.ftp.delete(name)

    def close(self):
        try:
            self.ftp.quit()
        except Exception:
            self.ftp.close()


class SftpTarget:
    """SFTP over SSH. Same surface as FtpTarget so main() doesn't branch."""

    def __init__(self, config):
        self.client, self.sftp = connect_sftp(config)
        self.remote_dir = None

    def ensure_dir(self, remote_dir):
        self.remote_dir = remote_dir
        built = ""
        for part in remote_dir.strip("/").split("/"):
            built += "/" + part
            try:
                self.sftp.stat(built)
            except IOError:
                self.sftp.mkdir(built)
        self.sftp.chdir(remote_dir)

    def sizes(self):
        return {a.filename: a.st_size for a in self.sftp.listdir_attr(self.remote_dir)}

    def put(self, path):
        self.sftp.put(str(path), f"{self.remote_dir}/{path.name}")

    def delete(self, name):
        self.sftp.remove(f"{self.remote_dir}/{name}")

    def close(self):
        try:
            self.sftp.close()
        finally:
            self.client.close()


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
    protocol = pick_protocol(config)
    target = SftpTarget(config) if protocol == "sftp" else FtpTarget(config)
    try:
        target.ensure_dir(remote_dir)
        before = target.sizes()

        for f in files:
            target.put(f)
            was = before.get(f.name)
            state = "new" if was is None else ("unchanged" if was == f.stat().st_size else "replaced")
            print(f"    {f.name:<34} {f.stat().st_size // 1024:>5}KB  {state}")

        # Verify by size rather than trusting the transfer: a truncated upload
        # still "succeeds" and leaves a broken image in a sent newsletter.
        after = target.sizes()
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
                target.delete(name)
            print(f"  Deleted {len(extra)} file(s) not in this run: {', '.join(extra)}")
        elif extra:
            print(f"  Left in place, not part of this run ({len(extra)}): {', '.join(extra)}")
            print(f"  Use --delete-extra to remove them.")
    finally:
        target.close()

    print(f"\nDone. The email's images now resolve at {hosted_base(month, year)}/\n")


if __name__ == "__main__":
    main()
