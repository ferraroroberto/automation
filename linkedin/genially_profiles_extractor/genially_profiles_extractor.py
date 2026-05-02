"""Extract alumni profile cards (name, role, company, email, LinkedIn, photo)
from a public Genially presentation by reading its view-API JSON.

Genially renders each card as a group of independent widgets on a slide
(a Text block, an Image, plus two Svg icons whose ``interactivities`` map to
``htmlTooltip`` actions for the email and ``openLink`` actions for the
LinkedIn URL). This script associates them per-card by clustering widgets
into rows, then pairing them by sorted-x order within each row.

Configuration lives in ``config.json`` next to this script (use
``config.example.json`` as a template). The real ``config.json`` is
gitignored. CLI usage::

    python genially_profiles_extractor.py [--config PATH]
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import math
import re
import sys
from dataclasses import asdict, dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Optional

import requests

API_TEMPLATE = "https://view.genially.com/api/view/{genially_id}"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

logger = logging.getLogger("genially_extractor")


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #
@dataclass
class Profile:
    name: str
    role: str
    company: str
    email: Optional[str]
    linkedin: Optional[str]
    photo: Optional[str]
    photo_local: Optional[str]
    slide: int
    slide_name: str
    debug: dict[str, Any] = field(default_factory=dict)


# --------------------------------------------------------------------------- #
# HTML helpers
# --------------------------------------------------------------------------- #
class _LineCollector(HTMLParser):
    """Flatten HTML into newline-separated visible text, treating <div>, <p>,
    and <br> as line breaks."""

    BLOCKS = {"div", "p", "br", "li"}

    def __init__(self) -> None:
        super().__init__()
        self._buf: list[str] = []
        self._cur: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        if tag in self.BLOCKS:
            self._flush()

    def handle_endtag(self, tag: str) -> None:
        if tag in self.BLOCKS:
            self._flush()

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        if tag in self.BLOCKS:
            self._flush()

    def handle_data(self, data: str) -> None:
        self._cur.append(data)

    def _flush(self) -> None:
        line = "".join(self._cur).strip()
        self._cur = []
        if line:
            self._buf.append(line)

    def lines(self) -> list[str]:
        self._flush()
        return self._buf


def html_to_lines(html: str) -> list[str]:
    p = _LineCollector()
    p.feed(html or "")
    cleaned: list[str] = []
    for ln in p.lines():
        ln = ln.replace("\xa0", " ").replace("&nbsp;", " ").strip()
        if ln:
            cleaned.append(ln)
    return cleaned


def html_to_text(html: str) -> str:
    return " ".join(html_to_lines(html))


# --------------------------------------------------------------------------- #
# Genially helpers
# --------------------------------------------------------------------------- #
def parse_genially_id(value: str) -> str:
    """Accept either a bare id or a view URL."""
    value = value.strip()
    m = re.search(r"([0-9a-fA-F]{24})", value)
    if not m:
        raise ValueError(f"Could not find a Genially id in {value!r}")
    return m.group(1)


def fetch_view_json(genially_id: str, timeout: int = 30) -> dict[str, Any]:
    url = API_TEMPLATE.format(genially_id=genially_id)
    logger.info("ℹ️ Fetching %s", url)
    resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


def _px(d: dict[str, Any], key: str) -> float:
    raw = d.get(key, "0")
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip().lower().replace("px", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _pos(widget: dict[str, Any]) -> tuple[float, float]:
    p = widget.get("Position") or {}
    return _px(p, "PositionLeft"), _px(p, "PositionTop")


def _size(widget: dict[str, Any]) -> tuple[float, float]:
    s = widget.get("Size") or {}
    return _px(s, "Width"), _px(s, "Height")


def _center(widget: dict[str, Any]) -> tuple[float, float]:
    x, y = _pos(widget)
    w, h = _size(widget)
    return x + w / 2.0, y + h / 2.0


def _dist(a: tuple[float, float], b: tuple[float, float]) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])


# Cards line up in rows of equal-width caption boxes; widgets within a row
# are ordered left-to-right. Direct nearest-neighbour matching is unreliable
# because the email / LinkedIn icons sit ~300 px to the right of their
# caption's left edge, often closer in 2D to the *next* card's caption. We
# instead cluster widgets into rows by y-coordinate, then pair by sorted x.
_ROW_TOLERANCE_PX = 80.0


def _cluster_rows(items: list[Any], y_of) -> list[list[Any]]:
    """Group items into rows by y-coordinate. Sort by y, then split whenever
    the gap between consecutive items exceeds the row tolerance."""
    if not items:
        return []
    sorted_items = sorted(items, key=y_of)
    rows: list[list[Any]] = [[sorted_items[0]]]
    for it in sorted_items[1:]:
        if y_of(it) - y_of(rows[-1][-1]) > _ROW_TOLERANCE_PX:
            rows.append([it])
        else:
            rows[-1].append(it)
    for r in rows:
        r.sort(key=lambda w: _center(w)[0])
    return rows


def _row_y(row: list[Any]) -> float:
    return sum(_center(w)[1] for w in row) / len(row)


# --------------------------------------------------------------------------- #
# Card detection
# --------------------------------------------------------------------------- #
def is_portrait_image(image: dict[str, Any]) -> bool:
    """Profile photos in this template are square portraits sized ~135x135 px;
    flags are ~37x25 px and the corner logo is ~100-120 px tall but not
    square. Treat anything large-and-roughly-square as a portrait."""
    w, h = _size(image)
    if w < 80 or h < 80:
        return False
    longer = max(w, h)
    shorter = min(w, h)
    if shorter == 0:
        return False
    return (shorter / longer) > 0.85


def is_profile_text(lines: list[str]) -> bool:
    """A profile Text widget contains 2-4 short lines: name, role, company."""
    if not (2 <= len(lines) <= 4):
        return False
    if any(len(ln) > 200 for ln in lines):
        return False
    joined = " ".join(lines).lower()
    if joined.strip() in {"alumnxs", "alumnos", "alumni"}:
        return False
    # name should plausibly look like a name: at least two whitespace-split tokens
    name = lines[0]
    if len(name.split()) < 2:
        return False
    return True


def email_from_html(html: str) -> Optional[str]:
    if not html:
        return None
    text = html_to_text(html)
    m = EMAIL_RE.search(text)
    return m.group(0) if m else None


def extract_actions(view_json: dict[str, Any]) -> dict[str, dict[str, Any]]:
    actions = view_json.get("interactivityActions") or {}
    if not isinstance(actions, dict):
        return {}
    return actions


def resolve_action(
    widget: dict[str, Any], actions: dict[str, dict[str, Any]]
) -> dict[str, Any]:
    """Return the first interactivityAction referenced by a widget, or {}."""
    iv = widget.get("interactivities") or {}
    if not isinstance(iv, dict):
        return {}
    for action_id in iv.values():
        action = actions.get(action_id)
        if action:
            return action
    return {}


def widget_email(widget: dict[str, Any], actions: dict[str, dict[str, Any]]) -> Optional[str]:
    a = resolve_action(widget, actions)
    if a.get("type") != "htmlTooltip":
        return None
    return email_from_html(a.get("html") or "")


def widget_linkedin(
    widget: dict[str, Any], actions: dict[str, dict[str, Any]]
) -> Optional[str]:
    a = resolve_action(widget, actions)
    if a.get("type") != "openLink":
        return None
    link = (a.get("link") or "").strip()
    if "linkedin.com" not in link.lower():
        return None
    return link


# --------------------------------------------------------------------------- #
# Per-slide assembly
# --------------------------------------------------------------------------- #
def _pair_text_rows(
    text_rows: list[list[dict[str, Any]]],
    other_rows: list[list[Any]],
) -> dict[int, list[Any]]:
    """For each text row, pick the matching 'other_rows' row.

    Captions sit *below* their card's photo and icons, so we prefer the
    closest other-row whose mean-y is above the text row. (Picking the
    globally nearest row by absolute y-distance fails on slides where the
    top-row caption is closer to the *bottom* row's icons than to its
    own.) Falls back to the closest row overall if none is above."""
    out: dict[int, list[Any]] = {}
    if not other_rows:
        return out
    for ti, trow in enumerate(text_rows):
        ty = _row_y(trow)
        above = [r for r in other_rows if _row_y(r) < ty]
        pool = above if above else other_rows
        out[ti] = min(pool, key=lambda r: abs(_row_y(r) - ty))
    return out


def build_profiles(view_json: dict[str, Any]) -> list[Profile]:
    actions = extract_actions(view_json)
    slides = view_json.get("Slides") or []
    texts = view_json.get("Texts") or []
    images = view_json.get("Images") or []
    svgs = view_json.get("Svgs") or []

    profiles: list[Profile] = []

    for slide in slides:
        sid = slide.get("Id")
        s_name = slide.get("Name") or ""
        s_order = int(slide.get("Order") or 0)

        # Profile-shaped texts.
        s_texts: list[dict[str, Any]] = []
        for t in texts:
            if t.get("IdSlide") != sid:
                continue
            lines = html_to_lines(t.get("TextMessage") or t.get("Html") or "")
            if is_profile_text(lines):
                s_texts.append(t)

        if not s_texts:
            continue

        # Interactive svgs and portrait images on this slide.
        s_email_pairs: list[tuple[dict[str, Any], str]] = []
        s_link_pairs: list[tuple[dict[str, Any], str]] = []
        for sv in svgs:
            if sv.get("IdSlide") != sid:
                continue
            email = widget_email(sv, actions)
            if email:
                s_email_pairs.append((sv, email))
            link = widget_linkedin(sv, actions)
            if link:
                s_link_pairs.append((sv, link))
        s_photos = [
            im for im in images
            if im.get("IdSlide") == sid and is_portrait_image(im)
        ]

        email_lookup = {id(w): e for w, e in s_email_pairs}
        link_lookup = {id(w): l for w, l in s_link_pairs}

        # Cluster widgets into rows by y, sorted left-to-right within each row.
        text_rows = _cluster_rows(s_texts, lambda w: _center(w)[1])
        email_rows = _cluster_rows([w for w, _ in s_email_pairs], lambda w: _center(w)[1])
        link_rows = _cluster_rows([w for w, _ in s_link_pairs], lambda w: _center(w)[1])
        photo_rows = _cluster_rows(s_photos, lambda w: _center(w)[1])

        email_match = _pair_text_rows(text_rows, email_rows)
        link_match = _pair_text_rows(text_rows, link_rows)
        photo_match = _pair_text_rows(text_rows, photo_rows)

        for ti, trow in enumerate(text_rows):
            erow = email_match.get(ti) or []
            lrow = link_match.get(ti) or []
            prow = photo_match.get(ti) or []

            for ci, t in enumerate(trow):
                lines = html_to_lines(t.get("TextMessage") or t.get("Html") or "")
                name = lines[0] if lines else ""
                role = lines[1] if len(lines) > 1 else ""
                company = lines[2] if len(lines) > 2 else ""

                email: Optional[str] = None
                if ci < len(erow):
                    email = email_lookup.get(id(erow[ci]))
                linkedin: Optional[str] = None
                if ci < len(lrow):
                    linkedin = link_lookup.get(id(lrow[ci]))
                photo: Optional[str] = None
                if ci < len(prow):
                    photo = prow[ci].get("Source")

                profiles.append(
                    Profile(
                        name=name,
                        role=role,
                        company=company,
                        email=email,
                        linkedin=linkedin,
                        photo=photo,
                        photo_local=None,
                        slide=s_order,
                        slide_name=s_name,
                        debug={
                            "text_id": t.get("Id"),
                            "text_pos": _pos(t),
                            "row": ti,
                            "col": ci,
                        },
                    )
                )

    return profiles


def deduplicate(profiles: list[Profile]) -> list[Profile]:
    """Remove duplicates that arise from copy-of slides. Key on email, then
    LinkedIn, then name. Keep the earliest slide occurrence."""
    seen: set[str] = set()
    out: list[Profile] = []
    for p in profiles:
        key = (p.email or p.linkedin or p.name or "").strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


# --------------------------------------------------------------------------- #
# Photo download
# --------------------------------------------------------------------------- #
_SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9._-]+")


def _photo_filename(profile: Profile, url: str) -> str:
    base = (profile.email or profile.name or "profile").lower()
    base = base.split("@", 1)[0]
    base = _SAFE_NAME_RE.sub("_", base).strip("_") or "profile"
    ext = Path(url.split("?", 1)[0]).suffix.lower()
    if ext not in {".jpg", ".jpeg", ".png", ".gif", ".webp"}:
        ext = ".jpg"
    return f"{base}{ext}"


def download_photos(profiles: list[Profile], photos_dir: Path) -> None:
    """Download each profile's photo into ``photos_dir``. Mutates profiles in
    place to set ``photo_local`` to a path relative to ``photos_dir.parent``.
    Already-downloaded files are reused."""
    photos_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    for p in profiles:
        if not p.photo:
            continue
        fname = _photo_filename(p, p.photo)
        target = photos_dir / fname
        rel = f"{photos_dir.name}/{fname}"
        if target.exists() and target.stat().st_size > 0:
            p.photo_local = rel
            continue
        try:
            r = session.get(p.photo, timeout=30)
            r.raise_for_status()
            target.write_bytes(r.content)
            p.photo_local = rel
            logger.info("✅ photo %s -> %s", p.name, fname)
        except requests.RequestException as e:
            logger.warning("⚠️ failed photo for %s (%s): %s", p.name, p.photo, e)


# --------------------------------------------------------------------------- #
# HTML rendering
# --------------------------------------------------------------------------- #
HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
  :root {
    --bg: #0f1115; --panel: #171a21; --border: #262a35;
    --text: #e6e8ee; --muted: #9aa3b2; --accent: #4f8cff; --accent-2: #38d39f;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; background: var(--bg); color: var(--text);
    font: 15px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
  header { position: sticky; top: 0; z-index: 10; background: rgba(15,17,21,0.92);
    backdrop-filter: blur(8px); padding: 18px 24px; border-bottom: 1px solid var(--border); }
  header h1 { margin: 0 0 10px; font-size: 22px; font-weight: 600; letter-spacing: 0.2px; }
  .tools { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
  #q { flex: 1; min-width: 260px; max-width: 480px; padding: 10px 14px;
    background: var(--panel); color: var(--text); border: 1px solid var(--border);
    border-radius: 8px; font-size: 14px; outline: none; }
  #q:focus { border-color: var(--accent); }
  #count { color: var(--muted); font-size: 13px; }
  main { padding: 24px; display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; }
  .card { background: var(--panel); border: 1px solid var(--border); border-radius: 14px;
    padding: 18px; display: flex; flex-direction: column; gap: 12px; transition: border-color .15s; }
  .card:hover { border-color: var(--accent); }
  .photo-wrap { display: flex; justify-content: center; }
  .photo { width: 120px; height: 120px; border-radius: 50%; object-fit: cover;
    background: #2a2f3a; border: 2px solid var(--border); }
  .name { font-weight: 600; font-size: 16px; text-align: center; }
  .role { color: var(--muted); font-size: 13px; text-align: center;
    line-height: 1.35; min-height: 36px; }
  .company { font-size: 13px; text-align: center; font-weight: 500;
    color: var(--accent-2); }
  .actions { display: flex; gap: 8px; justify-content: center; margin-top: auto; }
  .btn { display: inline-flex; align-items: center; gap: 6px;
    padding: 7px 12px; border-radius: 8px; font-size: 12px; font-weight: 500;
    text-decoration: none; border: 1px solid var(--border); color: var(--text);
    background: transparent; transition: background .15s, border-color .15s; }
  .btn:hover { background: var(--border); border-color: var(--accent); }
  .btn[disabled], .btn.disabled { opacity: 0.35; pointer-events: none; }
  .btn svg { width: 14px; height: 14px; }
  .empty { grid-column: 1 / -1; text-align: center; color: var(--muted);
    padding: 60px 20px; font-size: 14px; }
</style>
</head>
<body>
<header>
  <h1>__TITLE__</h1>
  <div class="tools">
    <input id="q" type="search" placeholder="Search by name, role, or company…" autofocus>
    <span id="count"></span>
  </div>
</header>
<main id="grid"></main>

<script>
const DATA = __DATA__;

const grid = document.getElementById('grid');
const countEl = document.getElementById('count');
const q = document.getElementById('q');

const ICON_MAIL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>';
const ICON_LINKEDIN = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M20.45 20.45h-3.55v-5.57c0-1.33-.02-3.04-1.85-3.04-1.85 0-2.13 1.45-2.13 2.94v5.67H9.36V9h3.41v1.56h.05c.48-.9 1.64-1.85 3.38-1.85 3.61 0 4.28 2.38 4.28 5.47v6.27zM5.34 7.43a2.06 2.06 0 1 1 0-4.12 2.06 2.06 0 0 1 0 4.12zm1.78 13.02H3.56V9h3.56v11.45zM22.22 0H1.77C.79 0 0 .77 0 1.72v20.56C0 23.23.79 24 1.77 24h20.45c.98 0 1.78-.77 1.78-1.72V1.72C24 .77 23.2 0 22.22 0z"/></svg>';

function escapeHtml(s) {
  return (s || '').replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

function card(p) {
  const photo = p.photo_local || p.photo || '';
  const initials = (p.name || '?').split(/\\s+/).slice(0,2).map(s=>s[0]||'').join('').toUpperCase();
  const photoHtml = photo
    ? `<img class="photo" src="${escapeHtml(photo)}" alt="${escapeHtml(p.name)}" loading="lazy" onerror="this.style.display='none'">`
    : `<div class="photo" style="display:flex;align-items:center;justify-content:center;font-size:32px;color:var(--muted);">${escapeHtml(initials)}</div>`;
  const mailBtn = p.email
    ? `<a class="btn" href="mailto:${encodeURIComponent(p.email)}">${ICON_MAIL} Email</a>`
    : `<span class="btn disabled">${ICON_MAIL} Email</span>`;
  const liBtn = p.linkedin
    ? `<a class="btn" href="${escapeHtml(p.linkedin)}" target="_blank" rel="noopener">${ICON_LINKEDIN} LinkedIn</a>`
    : `<span class="btn disabled">${ICON_LINKEDIN} LinkedIn</span>`;
  return `<article class="card">
    <div class="photo-wrap">${photoHtml}</div>
    <div class="name">${escapeHtml(p.name)}</div>
    <div class="role">${escapeHtml(p.role)}</div>
    <div class="company">${escapeHtml(p.company)}</div>
    <div class="actions">${mailBtn}${liBtn}</div>
  </article>`;
}

function render(filter) {
  const f = (filter || '').trim().toLowerCase();
  const matches = !f ? DATA : DATA.filter(p =>
    (p.name||'').toLowerCase().includes(f)
    || (p.role||'').toLowerCase().includes(f)
    || (p.company||'').toLowerCase().includes(f)
    || (p.email||'').toLowerCase().includes(f)
  );
  countEl.textContent = matches.length === DATA.length
    ? `${DATA.length} profiles`
    : `${matches.length} of ${DATA.length}`;
  grid.innerHTML = matches.length
    ? matches.map(card).join('')
    : '<div class="empty">No matches.</div>';
}

q.addEventListener('input', e => render(e.target.value));
render('');
</script>
</body>
</html>
"""


def write_html(profiles: list[Profile], out_dir: Path, stem: str, title: str) -> Path:
    payload = [
        {k: v for k, v in asdict(p).items() if k != "debug"} for p in profiles
    ]
    html = HTML_TEMPLATE.replace("__TITLE__", title).replace(
        "__DATA__", json.dumps(payload, ensure_ascii=False)
    )
    html_path = out_dir / f"{stem}.html"
    html_path.write_text(html, encoding="utf-8")
    return html_path


# --------------------------------------------------------------------------- #
# I/O
# --------------------------------------------------------------------------- #
def write_outputs(profiles: list[Profile], out_dir: Path, stem: str) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    json_path = out_dir / f"{stem}.json"
    csv_path = out_dir / f"{stem}.csv"

    json_payload = [
        {k: v for k, v in asdict(p).items() if k != "debug"} for p in profiles
    ]
    json_path.write_text(
        json.dumps(json_payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    fieldnames = [
        "name", "role", "company", "email", "linkedin",
        "photo", "photo_local", "slide", "slide_name",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=fieldnames)
        w.writeheader()
        for p in profiles:
            row = {k: getattr(p, k) for k in fieldnames}
            w.writerow(row)

    return json_path, csv_path


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "config.json"


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        example = path.with_name("config.example.json")
        raise FileNotFoundError(
            f"Config not found: {path}. "
            f"Copy {example.name} to config.json and fill in your URL."
        )
    cfg = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(cfg, dict):
        raise ValueError(f"Config root must be a JSON object: {path}")
    if not cfg.get("url"):
        raise ValueError(f"Config missing required key 'url': {path}")
    if not cfg.get("output_dir"):
        raise ValueError(f"Config missing required key 'output_dir': {path}")
    cfg.setdefault("output_stem", "genially_profiles")
    cfg.setdefault("save_raw", False)
    cfg.setdefault("verbose", False)
    cfg.setdefault("download_photos", True)
    cfg.setdefault("write_html", True)
    cfg.setdefault("html_title", "Alumnxs")
    return cfg


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG_PATH),
        help="Path to config.json (default: %(default)s).",
    )
    return ap.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)

    try:
        cfg = load_config(Path(args.config))
    except (FileNotFoundError, ValueError) as e:
        logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
        logger.error("❌ %s", e)
        return 2

    logging.basicConfig(
        level=logging.DEBUG if cfg["verbose"] else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    try:
        gid = parse_genially_id(cfg["url"])
    except ValueError as e:
        logger.error("❌ %s", e)
        return 2

    try:
        view = fetch_view_json(gid)
    except requests.HTTPError as e:
        logger.error("❌ HTTP error fetching Genially view: %s", e)
        return 1
    except requests.RequestException as e:
        logger.error("❌ Network error: %s", e)
        return 1

    out_dir = Path(cfg["output_dir"])
    stem = cfg["output_stem"]
    if cfg["save_raw"]:
        out_dir.mkdir(parents=True, exist_ok=True)
        raw_path = out_dir / f"{stem}_raw.json"
        raw_path.write_text(
            json.dumps(view, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        logger.info("ℹ️ Saved raw API JSON to %s", raw_path)

    profiles = build_profiles(view)
    # Real profile cards always carry at least an email tooltip or LinkedIn link
    # on the same slide; drop candidates that matched neither (e.g. the cover
    # title text on slide 1).
    profiles = [p for p in profiles if p.email or p.linkedin]
    profiles = deduplicate(profiles)

    if not profiles:
        logger.warning("⚠️ No profiles extracted")
    else:
        logger.info("✅ Extracted %d profiles", len(profiles))

    out_dir.mkdir(parents=True, exist_ok=True)
    if cfg["download_photos"]:
        download_photos(profiles, out_dir / "photos")

    json_path, csv_path = write_outputs(profiles, out_dir, stem)
    logger.info("✅ Wrote %s", json_path)
    logger.info("✅ Wrote %s", csv_path)

    if cfg["write_html"]:
        html_path = write_html(profiles, out_dir, stem, cfg["html_title"])
        logger.info("✅ Wrote %s", html_path)

    missing = [p.name for p in profiles if not p.email or not p.linkedin or not p.photo]
    if missing:
        logger.warning(
            "⚠️ %d profile(s) missing email/linkedin/photo: %s",
            len(missing),
            ", ".join(missing),
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
