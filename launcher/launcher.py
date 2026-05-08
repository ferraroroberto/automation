"""Flask-based remote project launcher.

Serves a web UI (optionally password-protected) intended to be reached over
Tailscale from a phone.  Two tabs:

* **Cloud Code** — every ``*remote*.bat`` file in ``LAUNCHER_PROJECTS_DIR``
  (defaults to the parent of this repo).
* **Apps** — Streamlit launchers discovered recursively under
  ``LAUNCHER_APPS_SCAN_ROOT`` (defaults to the parent of this repo, matching
  ``LAUNCHER_PROJECTS_DIR``) and persisted to ``launcher/apps_config.json``
  (gitignored). The scan is manual: hit *Scan projects* to add new entries.

Configuration via root ``.env``:
    LAUNCHER_PASSWORD       - optional, plaintext password for the login form.
                              When unset, auth is disabled and the launcher is
                              accessible without logging in.
    LAUNCHER_SECRET_KEY     - required, Flask session signing key
    LAUNCHER_PROJECTS_DIR   - optional, override for the directory scanned
                              for ``*remote*.bat`` files. Defaults to the
                              parent of the automation repo root.
    LAUNCHER_APPS_SCAN_ROOT - optional, override for the recursive scan that
                              powers the Apps tab. Defaults to the parent of
                              this repo (matches LAUNCHER_PROJECTS_DIR).
    LAUNCHER_STREAMLIT_PORT - optional, port targeted by the "Kill :PORT"
                              button on the Apps tab. Defaults to 8501.
    LAUNCHER_WEBAPP_PORT    - optional, port targeted by the second
                              "Kill :PORT" button (uvicorn / FastAPI
                              webapps). Defaults to 8443.
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
import json
import logging
import os
import re
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


PASSWORD: Optional[str] = os.environ.get("LAUNCHER_PASSWORD") or None
AUTH_ENABLED: bool = PASSWORD is not None
SECRET_KEY = _require_env("LAUNCHER_SECRET_KEY")
PROJECTS_DIR = Path(os.environ.get("LAUNCHER_PROJECTS_DIR") or DEFAULT_PROJECTS_DIR).resolve()
PORT = int(os.environ.get("LAUNCHER_PORT", "5050"))
HOST = os.environ.get("LAUNCHER_HOST", "0.0.0.0")

CONFIG_PATH = LAUNCHER_DIR / "config.json"
APPS_CONFIG_PATH = LAUNCHER_DIR / "apps_config.json"
APPS_SCAN_ROOT = Path(os.environ.get("LAUNCHER_APPS_SCAN_ROOT") or DEFAULT_PROJECTS_DIR).resolve()
APPS_SCAN_SKIP_DIRS = {".venv", "venv", "__pycache__", "node_modules", "certificates", ".git", "old"}
STREAMLIT_PORT = int(os.environ.get("LAUNCHER_STREAMLIT_PORT", "8501"))
WEBAPP_PORT = int(os.environ.get("LAUNCHER_WEBAPP_PORT", "8443"))

VALID_MODELS = ("opus", "sonnet", "haiku")
VALID_EFFORTS = ("off", "low", "medium", "high")
DEFAULT_CONFIG: dict = {
    "model": "opus",
    "effort": "high",
    "verbose": True,
    "debug": False,
}
ALWAYS_ON_FLAGS = ("--remote-control", "--dangerously-skip-permissions")

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


def _pretty_name_from_stem(stem: str) -> str:
    parts = [p for p in re.split(r"[_\-\s]+", stem) if p and p.lower() != "remote"]
    if not parts:
        parts = [stem]
    return " ".join(p.capitalize() for p in parts)


def _parse_legacy_flags(flags: str) -> dict:
    """Convert an old free-text ``claude_flags`` string into structured config."""
    tokens = flags.split()
    result = dict(DEFAULT_CONFIG)
    result["verbose"] = "--verbose" in tokens
    result["debug"] = "--debug" in tokens
    if "--model" in tokens:
        i = tokens.index("--model")
        if i + 1 < len(tokens) and tokens[i + 1] in VALID_MODELS:
            result["model"] = tokens[i + 1]
    if "--effort" in tokens:
        i = tokens.index("--effort")
        if i + 1 < len(tokens) and tokens[i + 1] in VALID_EFFORTS:
            result["effort"] = tokens[i + 1]
    else:
        result["effort"] = "off"
    return result


def load_config() -> dict:
    raw: dict = {}
    if CONFIG_PATH.exists():
        try:
            parsed = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(parsed, dict):
                raw = parsed
        except Exception:
            log.exception("⚠️ Could not parse %s — using defaults", CONFIG_PATH)
    if "claude_flags" in raw and not any(k in raw for k in DEFAULT_CONFIG):
        raw = _parse_legacy_flags(str(raw.get("claude_flags") or ""))
    merged = dict(DEFAULT_CONFIG)
    if raw.get("model") in VALID_MODELS:
        merged["model"] = raw["model"]
    if raw.get("effort") in VALID_EFFORTS:
        merged["effort"] = raw["effort"]
    merged["verbose"] = bool(raw.get("verbose", DEFAULT_CONFIG["verbose"]))
    merged["debug"] = bool(raw.get("debug", DEFAULT_CONFIG["debug"]))
    return merged


def save_config(data: dict) -> None:
    clean = {
        "model": data["model"] if data.get("model") in VALID_MODELS else DEFAULT_CONFIG["model"],
        "effort": data["effort"] if data.get("effort") in VALID_EFFORTS else DEFAULT_CONFIG["effort"],
        "verbose": bool(data.get("verbose")),
        "debug": bool(data.get("debug")),
    }
    CONFIG_PATH.write_text(json.dumps(clean, indent=2), encoding="utf-8")


def build_claude_flags(config: dict) -> str:
    parts = list(ALWAYS_ON_FLAGS)
    model = config.get("model", DEFAULT_CONFIG["model"])
    if model in VALID_MODELS:
        parts.extend(["--model", model])
    effort = config.get("effort", DEFAULT_CONFIG["effort"])
    if effort in VALID_EFFORTS and effort != "off":
        parts.extend(["--effort", effort])
    if config.get("verbose"):
        parts.append("--verbose")
    if config.get("debug"):
        parts.append("--debug")
    return " ".join(parts)


def _config_from_form(form) -> dict:
    return {
        "model": form.get("model", DEFAULT_CONFIG["model"]),
        "effort": form.get("effort", DEFAULT_CONFIG["effort"]),
        "verbose": form.get("verbose") is not None,
        "debug": form.get("debug") is not None,
    }


def discover_cloud_projects() -> List[dict]:
    """Return cloud-code projects from workspaces and orphan ``*-remote.bat`` files.

    Each entry: ``{"id": str, "name": str, "project_dir": Path, "source": str}``.
    """
    if not PROJECTS_DIR.is_dir():
        log.warning("⚠️ Projects dir does not exist: %s", PROJECTS_DIR)
        return []
    results: List[dict] = []
    workspace_stems: set[str] = set()
    for ws in sorted(PROJECTS_DIR.glob("*.code-workspace")):
        try:
            data = json.loads(ws.read_text(encoding="utf-8"))
            raw_path = data["folders"][0]["path"]
        except Exception:
            continue
        project_dir = Path(raw_path)
        if not project_dir.is_absolute():
            project_dir = (PROJECTS_DIR / raw_path).resolve()
        results.append(
            {
                "id": ws.stem,
                "name": _pretty_name_from_stem(ws.stem),
                "project_dir": project_dir,
                "source": "workspace",
            }
        )
        workspace_stems.add(ws.stem)
    for bat in sorted(PROJECTS_DIR.glob("*-remote.bat")):
        stem = bat.stem[: -len("-remote")]
        if stem in workspace_stems:
            continue
        project_dir: Optional[Path] = None
        try:
            for line in bat.read_text(encoding="utf-8", errors="ignore").splitlines():
                m = re.match(r'set\s+"PROJECT_DIR=(.+)"', line.strip())
                if m:
                    project_dir = Path(m.group(1).strip())
                    break
        except Exception:
            pass
        if project_dir is None:
            project_dir = PROJECTS_DIR / stem
        results.append(
            {
                "id": stem,
                "name": _pretty_name_from_stem(stem),
                "project_dir": project_dir,
                "source": "orphan_bat",
            }
        )
    results.sort(key=lambda x: x["name"].lower())
    return results


def discover_workspaces() -> List[dict]:
    """Return workspaces paired with their derived bat info.

    Each entry: name, project_dir (Path), bat_name (str), bat_exists (bool).
    """
    if not PROJECTS_DIR.is_dir():
        return []
    results = []
    for ws in sorted(PROJECTS_DIR.glob("*.code-workspace")):
        try:
            data = json.loads(ws.read_text(encoding="utf-8"))
            raw_path = data["folders"][0]["path"]
        except Exception:
            continue
        project_dir = Path(raw_path)
        if not project_dir.is_absolute():
            project_dir = (PROJECTS_DIR / raw_path).resolve()
        bat_name = ws.stem + "-remote.bat"
        results.append({
            "name": ws.stem,
            "project_dir": project_dir,
            "bat_name": bat_name,
            "bat_exists": (PROJECTS_DIR / bat_name).exists(),
        })
    return results


def discover_orphan_bats() -> List[dict]:
    """Return *-remote.bat files that have no matching .code-workspace.

    Each entry: name (str), project_dir (Path), bat_name (str), ws_name (str).
    """
    if not PROJECTS_DIR.is_dir():
        return []
    workspace_stems = {ws.stem for ws in PROJECTS_DIR.glob("*.code-workspace")}
    results = []
    for bat in sorted(PROJECTS_DIR.glob("*-remote.bat")):
        stem = bat.stem[: -len("-remote")]
        if stem in workspace_stems:
            continue
        project_dir: Optional[Path] = None
        try:
            for line in bat.read_text(encoding="utf-8", errors="ignore").splitlines():
                m = re.match(r'set\s+"PROJECT_DIR=(.+)"', line.strip())
                if m:
                    project_dir = Path(m.group(1).strip())
                    break
        except Exception:
            pass
        if project_dir is None:
            project_dir = PROJECTS_DIR / stem
        results.append({
            "name": stem,
            "project_dir": project_dir,
            "bat_name": bat.name,
            "ws_name": stem + ".code-workspace",
        })
    return results


def render_bat_content(project_dir: Path, flags: str) -> str:
    d = str(project_dir)
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "\r\n"
        ":: -----------------------------------------------\r\n"
        ":: launch_claude_remote.bat\r\n"
        ":: Opens Claude Code with Remote Control enabled\r\n"
        ":: -----------------------------------------------\r\n"
        "\r\n"
        f'set "PROJECT_DIR={d}"\r\n'
        "\r\n"
        ":: -----------------------------------------------\r\n"
        "\r\n"
        'if not exist "%PROJECT_DIR%" (\r\n'
        "    echo [ERROR] Folder not found: %PROJECT_DIR%\r\n"
        "    pause\r\n"
        "    exit /b 1\r\n"
        ")\r\n"
        "\r\n"
        "echo.\r\n"
        "echo  Starting Claude Code with Remote Control\r\n"
        "echo  Project: %PROJECT_DIR%\r\n"
        "echo.\r\n"
        "\r\n"
        'cd /d "%PROJECT_DIR%"\r\n'
        "\r\n"
        f'"C:\\Windows\\System32\\cmd.exe" /c claude {flags}\r\n'
        "\r\n"
        "endlocal\r\n"
    )


def render_workspace_content(project_dir: Path) -> str:
    try:
        rel = project_dir.relative_to(PROJECTS_DIR)
        path_str = str(rel).replace("\\", "/")
    except ValueError:
        path_str = str(project_dir)
    return json.dumps({"folders": [{"path": path_str}]}, indent="\t") + "\n"


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
    if not AUTH_ENABLED:
        return True
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


def _spawn_claude(project_dir: Path, flags: str) -> None:
    """Open a new visible CMD window that runs ``claude`` in ``project_dir``.

    Uses ``cmd /c`` (not ``/k``) so the window closes when claude exits —
    no double-exit. The outer Popen's ``cwd`` is inherited by ``start``.
    """
    if not project_dir.is_dir():
        raise OSError(f"Project directory not found: {project_dir}")
    cmd = ["cmd", "/c", "start", "", "cmd", "/c", f"claude {flags}"]
    creationflags = 0
    if sys.platform == "win32":
        creationflags = subprocess.CREATE_NEW_CONSOLE  # type: ignore[attr-defined]
    subprocess.Popen(
        cmd,
        cwd=str(project_dir),
        shell=False,
        creationflags=creationflags,
        close_fds=True,
    )


@app.get("/")
def index():
    if not _is_authed():
        return redirect(url_for("login"))
    projects = discover_cloud_projects()
    config = load_config()
    return render_template(
        "index.html",
        projects=projects,
        projects_dir=str(PROJECTS_DIR),
        config=config,
        flags_string=build_claude_flags(config),
        csrf_token=_get_csrf_token(),
    )


@app.post("/options")
def options_post():
    if not _is_authed():
        abort(401)
    _check_csrf()
    save_config(_config_from_form(request.form))
    flash("Options saved.", "success")
    return redirect(url_for("index"))


@app.get("/login")
def login():
    if _is_authed():
        return redirect(url_for("index"))
    return render_template("login.html", csrf_token=_get_csrf_token())


@app.post("/login")
def login_post():
    if not AUTH_ENABLED:
        return redirect(url_for("index"))
    _check_csrf()
    submitted = request.form.get("password", "")
    if hmac.compare_digest(submitted, PASSWORD or ""):
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
    requested = request.form.get("id", "")
    project = _resolve_project(requested)
    if project is None:
        log.warning("⚠️ Launch rejected for unknown project %r from %s", requested, request.remote_addr)
        flash(f"Unknown project: {requested}", "error")
        return redirect(url_for("index"))
    config = load_config()
    flags = build_claude_flags(config)
    try:
        _spawn_claude(project["project_dir"], flags)
    except OSError as exc:
        log.exception("❌ Failed to launch %s", project["name"])
        flash(f"Failed to launch {project['name']}: {exc}", "error")
        return redirect(url_for("index"))
    stamp = datetime.now().strftime("%H:%M:%S")
    log.info("✅ Launched %s (%s) with flags: %s", project["name"], project["project_dir"], flags)
    flash(f"✅ Launched {project['name']} — {stamp}", "success")
    return redirect(url_for("index"))


def _resolve_project(project_id: str) -> Optional[dict]:
    if not project_id or "/" in project_id or "\\" in project_id:
        return None
    for project in discover_cloud_projects():
        if project["id"] == project_id:
            return project
    return None


@app.get("/generate")
def generate():
    if not _is_authed():
        return redirect(url_for("login"))
    config = load_config()
    return render_template(
        "generate.html",
        workspaces=discover_workspaces(),
        orphans=discover_orphan_bats(),
        config=config,
        flags_string=build_claude_flags(config),
        csrf_token=_get_csrf_token(),
    )


@app.post("/generate/run")
def generate_run():
    if not _is_authed():
        abort(401)
    _check_csrf()

    config = _config_from_form(request.form)
    flags = build_claude_flags(config)
    overwrite_names = set(request.form.getlist("overwrite"))
    create_ws_names = set(request.form.getlist("create_ws"))

    save_config(config)

    created, overwritten, ws_created, errors = [], [], [], []

    for ws in discover_workspaces():
        bat_path = PROJECTS_DIR / ws["bat_name"]
        if ws["bat_exists"] and ws["name"] not in overwrite_names:
            continue
        try:
            bat_path.write_bytes(render_bat_content(ws["project_dir"], flags).encode("utf-8"))
            (overwritten if ws["bat_exists"] else created).append(ws["bat_name"])
            log.info("✅ Wrote %s", ws["bat_name"])
        except OSError as exc:
            errors.append(f"{ws['bat_name']}: {exc}")

    for orphan in discover_orphan_bats():
        if orphan["name"] not in create_ws_names:
            continue
        ws_path = PROJECTS_DIR / orphan["ws_name"]
        try:
            ws_path.write_text(render_workspace_content(orphan["project_dir"]), encoding="utf-8")
            ws_created.append(orphan["ws_name"])
            log.info("✅ Created workspace %s", orphan["ws_name"])
        except OSError as exc:
            errors.append(f"{orphan['ws_name']}: {exc}")

    if created:
        flash(f"Created {len(created)} new bat file(s): {', '.join(created)}", "success")
    if overwritten:
        flash(f"Overwrote {len(overwritten)} bat file(s): {', '.join(overwritten)}", "success")
    if ws_created:
        flash(f"Created {len(ws_created)} workspace file(s): {', '.join(ws_created)}", "success")
    for err in errors:
        flash(f"Error: {err}", "error")
    if not created and not overwritten and not ws_created and not errors:
        flash("Nothing was generated — no items selected.", "error")

    return redirect(url_for("generate"))


def _slugify(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "-", value).strip("-").lower()
    return cleaned or "app"


def _pretty_folder_name(folder: Path) -> str:
    parts = [p for p in re.split(r"[_\-\s]+", folder.name) if p]
    if not parts:
        parts = [folder.name]
    return " ".join(p.capitalize() for p in parts)


def _app_id_from_path(bat_path: Path) -> str:
    try:
        rel = bat_path.resolve().relative_to(APPS_SCAN_ROOT)
    except ValueError:
        rel = Path(bat_path.name)
    return _slugify(str(rel.with_suffix("")))


def _classify_bat(bat_path: Path) -> Optional[str]:
    """Return ``"streamlit"`` | ``"webapp"`` | ``"tunnel"`` | ``None``.

    Classification is mutually exclusive — the first match wins:

    * ``streamlit`` — body contains ``streamlit run``. Bats that *also*
      embed ``cloudflared tunnel`` inline (e.g. ``launch_server.bat``)
      stay in this bucket; they don't write a URL file we can surface.
    * ``tunnel`` — filename stem contains ``tunnel`` AND body references
      ``uvicorn`` / ``run_tunnel`` / ``cloudflared``. These are the
      only bats we surface a tunnel URL for.
    * ``webapp`` — body runs ``uvicorn`` (or imports ``app.webapp.server``).
    """
    try:
        text = bat_path.read_text(encoding="utf-8", errors="ignore").lower()
    except OSError:
        return None
    if "streamlit run" in text:
        return "streamlit"
    stem = bat_path.stem.lower()
    has_tunnel_signal = any(token in text for token in ("uvicorn", "run_tunnel", "cloudflared"))
    if "tunnel" in stem and has_tunnel_signal:
        return "tunnel"
    if "uvicorn" in text or "app.webapp.server" in text or "app/webapp/server" in text:
        return "webapp"
    return None


def scan_app_bats() -> List[tuple[Path, str]]:
    """Recursively scan APPS_SCAN_ROOT, returning ``(path, kind)`` pairs."""
    if not APPS_SCAN_ROOT.is_dir():
        log.warning("⚠️ Apps scan root does not exist: %s", APPS_SCAN_ROOT)
        return []
    found: List[tuple[Path, str]] = []
    for bat in APPS_SCAN_ROOT.rglob("*.bat"):
        if any(part in APPS_SCAN_SKIP_DIRS for part in bat.parts):
            continue
        kind = _classify_bat(bat)
        if kind is not None:
            found.append((bat, kind))
    found.sort(key=lambda pair: pair[0])
    return found


def _tunnel_url_for(bat_path: Path) -> Optional[str]:
    """Read ``<bat.parent>/webapp/last_tunnel_url.txt`` if non-empty.

    Returns ``None`` when the file is missing or empty — used by the UI
    to show "Tunnel not running" without a server round-trip.
    """
    url_file = bat_path.parent / "webapp" / "last_tunnel_url.txt"
    try:
        text = url_file.read_text(encoding="utf-8").strip()
    except (OSError, UnicodeDecodeError):
        return None
    return text or None


def _decorate_app(entry: dict) -> dict:
    """Add render-time fields to a saved app: ``kind`` default + tunnel URL."""
    decorated = dict(entry)
    decorated.setdefault("kind", "streamlit")
    if decorated["kind"] == "tunnel":
        decorated["tunnel_url"] = _tunnel_url_for(Path(entry["bat_path"]))
    return decorated


def load_apps() -> dict:
    if APPS_CONFIG_PATH.exists():
        try:
            data = json.loads(APPS_CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict) and isinstance(data.get("apps"), list):
                return data
        except Exception:
            log.exception("⚠️ Could not parse %s — starting fresh", APPS_CONFIG_PATH)
    return {"scan_root": str(APPS_SCAN_ROOT), "apps": []}


def save_apps(data: dict) -> None:
    APPS_CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _split_apps_by_kind(apps: List[dict]) -> tuple[List[dict], List[dict]]:
    """Return ``(streamlit_apps, web_apps)``; web_apps holds webapp+tunnel."""
    streamlit_apps: List[dict] = []
    web_apps: List[dict] = []
    for entry in apps:
        decorated = _decorate_app(entry)
        if decorated["kind"] == "streamlit":
            streamlit_apps.append(decorated)
        else:
            web_apps.append(decorated)
    return streamlit_apps, web_apps


def discover_new_apps() -> List[dict]:
    """App bats found on disk but not yet saved in apps_config.json."""
    saved_paths = {a.get("bat_path") for a in load_apps().get("apps", [])}
    candidates: List[dict] = []
    for bat, kind in scan_app_bats():
        if str(bat) in saved_paths:
            continue
        candidates.append(
            {
                "id": _app_id_from_path(bat),
                "name": _pretty_folder_name(bat.parent),
                "bat_path": str(bat),
                "kind": kind,
            }
        )
    return candidates


def _split_candidates_by_kind(candidates: List[dict]) -> tuple[List[dict], List[dict]]:
    streamlit: List[dict] = []
    web: List[dict] = []
    for cand in candidates:
        (streamlit if cand.get("kind") == "streamlit" else web).append(cand)
    return streamlit, web


def _render_apps(*, new_apps: Optional[List[dict]]) -> str:
    streamlit_apps, web_apps = _split_apps_by_kind(load_apps().get("apps", []))
    new_streamlit, new_web = _split_candidates_by_kind(new_apps or [])
    return render_template(
        "apps.html",
        streamlit_apps=streamlit_apps,
        web_apps=web_apps,
        scan_root=str(APPS_SCAN_ROOT),
        new_apps=new_apps,
        new_streamlit=new_streamlit,
        new_web=new_web,
        streamlit_port=STREAMLIT_PORT,
        webapp_port=WEBAPP_PORT,
        csrf_token=_get_csrf_token(),
    )


@app.get("/apps")
def apps_index():
    if not _is_authed():
        return redirect(url_for("login"))
    return _render_apps(new_apps=None)


@app.post("/apps/scan")
def apps_scan():
    if not _is_authed():
        abort(401)
    _check_csrf()
    new_apps = discover_new_apps()
    if not new_apps:
        flash("No new apps found.", "success")
    return _render_apps(new_apps=new_apps)


@app.post("/apps/save")
def apps_save():
    if not _is_authed():
        abort(401)
    _check_csrf()
    selected_ids = set(request.form.getlist("add"))
    if not selected_ids:
        flash("No apps selected.", "error")
        return redirect(url_for("apps_index"))
    data = load_apps()
    existing_ids = {a.get("id") for a in data["apps"]}
    candidates = {c["id"]: c for c in discover_new_apps()}
    added: List[str] = []
    now = datetime.now().isoformat(timespec="seconds")
    for app_id in selected_ids:
        if app_id in existing_ids:
            continue
        cand = candidates.get(app_id)
        if cand is None:
            continue
        data["apps"].append(
            {
                "id": cand["id"],
                "name": cand["name"],
                "bat_path": cand["bat_path"],
                "kind": cand.get("kind", "streamlit"),
                "added_at": now,
            }
        )
        added.append(cand["name"])
    data["apps"].sort(key=lambda a: a.get("name", "").lower())
    save_apps(data)
    if added:
        log.info("✅ Added %d app(s) to apps_config.json", len(added))
        flash(f"Added {len(added)} app(s): {', '.join(added)}", "success")
    else:
        flash("Nothing was added — selections were already saved or unknown.", "error")
    return redirect(url_for("apps_index"))


@app.post("/apps/remove")
def apps_remove():
    if not _is_authed():
        abort(401)
    _check_csrf()
    target_id = request.form.get("id", "")
    data = load_apps()
    before = len(data["apps"])
    removed = next((a for a in data["apps"] if a.get("id") == target_id), None)
    data["apps"] = [a for a in data["apps"] if a.get("id") != target_id]
    if len(data["apps"]) < before and removed is not None:
        save_apps(data)
        log.info("✅ Removed app %s", removed.get("name"))
        flash(f"Removed {removed.get('name')}.", "success")
    else:
        flash("App not found.", "error")
    return redirect(url_for("apps_index"))


@app.post("/apps/rename")
def apps_rename():
    if not _is_authed():
        abort(401)
    _check_csrf()
    target_id = request.form.get("id", "")
    new_name = request.form.get("name", "").strip()
    if not new_name:
        flash("Display name cannot be empty.", "error")
        return redirect(url_for("apps_index"))
    data = load_apps()
    for a in data["apps"]:
        if a.get("id") == target_id:
            old = a.get("name")
            a["name"] = new_name
            data["apps"].sort(key=lambda x: x.get("name", "").lower())
            save_apps(data)
            log.info("✅ Renamed app %s → %s", old, new_name)
            flash(f"Renamed to {new_name}.", "success")
            return redirect(url_for("apps_index"))
    flash("App not found.", "error")
    return redirect(url_for("apps_index"))


def _find_pids_on_port(port: int) -> List[int]:
    """Return PIDs of processes listening on TCP ``port`` (admin-free)."""
    import psutil  # local import — not needed at module load

    own_pid = os.getpid()
    pids: set[int] = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            for conn in proc.net_connections(kind="inet"):
                if (
                    conn.status == psutil.CONN_LISTEN
                    and conn.laddr
                    and conn.laddr.port == port
                ):
                    pid = proc.info["pid"]
                    if pid and pid != own_pid:
                        pids.add(pid)
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue
    return sorted(pids)


def _kill_pids(pids: List[int]) -> tuple[List[int], List[str]]:
    """Force-kill the given PIDs. Returns (killed_pids, error_messages)."""
    import psutil

    killed: List[int] = []
    errors: List[str] = []
    for pid in pids:
        try:
            proc = psutil.Process(pid)
            proc.kill()
            try:
                proc.wait(timeout=3)
            except psutil.TimeoutExpired:
                pass
            killed.append(pid)
        except psutil.NoSuchProcess:
            killed.append(pid)
        except (psutil.AccessDenied, OSError) as exc:
            errors.append(f"PID {pid}: {exc}")
    return killed, errors


@app.post("/apps/kill_port")
def apps_kill_port():
    if not _is_authed():
        abort(401)
    _check_csrf()
    raw = request.form.get("port", "").strip()
    allowed = {str(STREAMLIT_PORT): STREAMLIT_PORT, str(WEBAPP_PORT): WEBAPP_PORT}
    port = allowed.get(raw)
    if port is None:
        log.warning("⚠️ Kill rejected for unknown port %r from %s", raw, request.remote_addr)
        flash("Unknown port.", "error")
        return redirect(url_for("apps_index"))
    pids = _find_pids_on_port(port)
    if not pids:
        flash(f"Nothing was listening on :{port}.", "success")
        return redirect(url_for("apps_index"))
    killed, errors = _kill_pids(pids)
    if killed:
        log.info("✅ Killed PID(s) %s on :%d", killed, port)
        flash(
            f"✅ Killed {len(killed)} process(es) on :{port} (PID {', '.join(str(p) for p in killed)}).",
            "success",
        )
    for err in errors:
        log.warning("⚠️ Kill error on :%d — %s", port, err)
        flash(f"Kill error: {err}", "error")
    return redirect(url_for("apps_index"))


@app.post("/apps/launch")
def apps_launch():
    if not _is_authed():
        abort(401)
    _check_csrf()
    target_id = request.form.get("id", "")
    data = load_apps()
    target = next((a for a in data["apps"] if a.get("id") == target_id), None)
    if target is None:
        log.warning("⚠️ Launch rejected for unknown app %r from %s", target_id, request.remote_addr)
        flash("Unknown app.", "error")
        return redirect(url_for("apps_index"))
    bat_path = Path(target["bat_path"])
    if not bat_path.is_file():
        log.warning("⚠️ Missing bat file for app %s: %s", target.get("name"), bat_path)
        flash(f"BAT file not found: {bat_path}", "error")
        return redirect(url_for("apps_index"))
    try:
        _spawn_bat(bat_path)
    except OSError as exc:
        log.exception("❌ Failed to launch app %s", target.get("name"))
        flash(f"Failed to launch {target.get('name')}: {exc}", "error")
        return redirect(url_for("apps_index"))
    stamp = datetime.now().strftime("%H:%M:%S")
    log.info("✅ Launched app %s (%s)", target.get("name"), bat_path)
    flash(f"✅ Launched {target.get('name')} — {stamp}", "success")
    return redirect(url_for("apps_index"))


def main() -> None:
    scheme = "https" if USE_HTTPS else "http"
    log.info("ℹ️ Launcher serving on %s://%s:%s", scheme, HOST, PORT)
    log.info("ℹ️ Scanning for *remote*.bat in %s", PROJECTS_DIR)
    log.info("ℹ️ Apps scan root: %s", APPS_SCAN_ROOT)
    log.info("ℹ️ Streamlit kill-port target: :%d", STREAMLIT_PORT)
    if not AUTH_ENABLED:
        log.warning("⚠️ Password auth is DISABLED (LAUNCHER_PASSWORD not set)")
    app.run(host=HOST, port=PORT, debug=False, ssl_context=SSL_CONTEXT)


if __name__ == "__main__":
    main()
