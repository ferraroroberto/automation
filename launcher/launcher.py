"""Flask-based remote project launcher.

Serves a small password-protected web UI (intended to be reached over
Tailscale from a phone) that lists every ``*remote*.bat`` file in the
parent directory of this repo and launches the selected one in a new
visible CMD window on the host machine.

Configuration via root ``.env``:
    LAUNCHER_PASSWORD       - required, plaintext password for the login form
    LAUNCHER_SECRET_KEY     - required, Flask session signing key
    LAUNCHER_PROJECTS_DIR   - optional, override for the directory scanned
                              for ``*remote*.bat`` files. Defaults to the
                              parent of the automation repo root.
    LAUNCHER_PORT           - optional, defaults to 5050
    LAUNCHER_HOST           - optional, defaults to 0.0.0.0
    LAUNCHER_SSL_CERT       - optional, path to TLS certificate file
                              (e.g. from ``tailscale cert``). Both
                              LAUNCHER_SSL_CERT and LAUNCHER_SSL_KEY must
                              be set together; setting only one is an error.
    LAUNCHER_SSL_KEY        - optional, path to TLS private key file.
                              When both SSL vars are set the server runs
                              HTTPS only; SESSION_COOKIE_SECURE is enabled
                              automatically.
"""

from __future__ import annotations

import hmac
import logging
import os
import secrets
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from dotenv import load_dotenv
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

LAUNCHER_DIR = Path(__file__).resolve().parent
REPO_ROOT = LAUNCHER_DIR.parent
DEFAULT_PROJECTS_DIR = REPO_ROOT.parent

load_dotenv(REPO_ROOT / ".env")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("launcher")


def _require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        log.error("❌ Missing required env var %s (set it in the root .env)", name)
        sys.exit(1)
    return value


PASSWORD = _require_env("LAUNCHER_PASSWORD")
SECRET_KEY = _require_env("LAUNCHER_SECRET_KEY")
PROJECTS_DIR = Path(os.environ.get("LAUNCHER_PROJECTS_DIR") or DEFAULT_PROJECTS_DIR).resolve()
PORT = int(os.environ.get("LAUNCHER_PORT", "5050"))
HOST = os.environ.get("LAUNCHER_HOST", "0.0.0.0")

_SSL_CERT = os.environ.get("LAUNCHER_SSL_CERT")
_SSL_KEY = os.environ.get("LAUNCHER_SSL_KEY")
if bool(_SSL_CERT) != bool(_SSL_KEY):
    log.error("❌ Set both LAUNCHER_SSL_CERT and LAUNCHER_SSL_KEY, or neither.")
    sys.exit(1)
USE_HTTPS: bool = bool(_SSL_CERT and _SSL_KEY)
SSL_CONTEXT: Optional[tuple] = (_SSL_CERT, _SSL_KEY) if USE_HTTPS else None

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=USE_HTTPS,
)


def _pretty_name(bat: Path) -> str:
    stem = bat.stem
    parts = [p for p in stem.replace("-", "_").split("_") if p and p.lower() != "remote"]
    if not parts:
        parts = [stem]
    return " ".join(p.capitalize() for p in parts)


def discover_projects() -> List[dict]:
    """Return the list of remote launcher bats in PROJECTS_DIR.

    Each entry: ``{"name": str, "filename": str, "path": Path}``.
    Filenames containing ``remote`` (case-insensitive) qualify.
    """
    if not PROJECTS_DIR.is_dir():
        log.warning("⚠️ Projects dir does not exist: %s", PROJECTS_DIR)
        return []
    found = []
    for bat in sorted(PROJECTS_DIR.glob("*.bat")):
        if "remote" not in bat.stem.lower():
            continue
        found.append(
            {
                "name": _pretty_name(bat),
                "filename": bat.name,
                "path": bat,
            }
        )
    return found


def _get_csrf_token() -> str:
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def _check_csrf() -> None:
    sent = request.form.get("csrf_token", "")
    expected = session.get("csrf_token", "")
    if not expected or not hmac.compare_digest(sent, expected):
        log.warning("⚠️ CSRF token mismatch from %s", request.remote_addr)
        abort(400)


def _is_authed() -> bool:
    return bool(session.get("authed"))


def _spawn_bat(bat_path: Path) -> None:
    """Open a new visible CMD window that runs the bat and stays open."""
    cmd = ["cmd", "/c", "start", "", "cmd", "/k", str(bat_path)]
    creationflags = 0
    if sys.platform == "win32":
        creationflags = subprocess.CREATE_NEW_CONSOLE  # type: ignore[attr-defined]
    subprocess.Popen(
        cmd,
        cwd=str(bat_path.parent),
        shell=False,
        creationflags=creationflags,
        close_fds=True,
    )


@app.get("/")
def index():
    if not _is_authed():
        return redirect(url_for("login"))
    projects = discover_projects()
    return render_template(
        "index.html",
        projects=projects,
        projects_dir=str(PROJECTS_DIR),
        csrf_token=_get_csrf_token(),
    )


@app.get("/login")
def login():
    if _is_authed():
        return redirect(url_for("index"))
    return render_template("login.html", csrf_token=_get_csrf_token())


@app.post("/login")
def login_post():
    _check_csrf()
    submitted = request.form.get("password", "")
    if hmac.compare_digest(submitted, PASSWORD):
        session.clear()
        session["authed"] = True
        session["csrf_token"] = secrets.token_urlsafe(32)
        log.info("✅ Login from %s", request.remote_addr)
        return redirect(url_for("index"))
    log.warning("⚠️ Failed login from %s", request.remote_addr)
    flash("Wrong password.", "error")
    return redirect(url_for("login"))


@app.post("/logout")
def logout():
    _check_csrf()
    session.clear()
    return redirect(url_for("login"))


@app.post("/launch")
def launch():
    if not _is_authed():
        abort(401)
    _check_csrf()
    requested = request.form.get("name", "")
    project = _resolve_project(requested)
    if project is None:
        log.warning("⚠️ Launch rejected for unknown bat %r from %s", requested, request.remote_addr)
        flash(f"Unknown project: {requested}", "error")
        return redirect(url_for("index"))
    try:
        _spawn_bat(project["path"])
    except OSError as exc:
        log.exception("❌ Failed to launch %s", project["filename"])
        flash(f"Failed to launch {project['name']}: {exc}", "error")
        return redirect(url_for("index"))
    stamp = datetime.now().strftime("%H:%M:%S")
    log.info("✅ Launched %s (%s)", project["name"], project["filename"])
    flash(f"✅ Launched {project['name']} — {stamp}", "success")
    return redirect(url_for("index"))


def _resolve_project(filename: str) -> Optional[dict]:
    if not filename or "/" in filename or "\\" in filename:
        return None
    for project in discover_projects():
        if project["filename"] == filename:
            return project
    return None


def main() -> None:
    scheme = "https" if USE_HTTPS else "http"
    log.info("ℹ️ Launcher serving on %s://%s:%s", scheme, HOST, PORT)
    log.info("ℹ️ Scanning for *remote*.bat in %s", PROJECTS_DIR)
    app.run(host=HOST, port=PORT, debug=False, ssl_context=SSL_CONTEXT)


if __name__ == "__main__":
    main()
