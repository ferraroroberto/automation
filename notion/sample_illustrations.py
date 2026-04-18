#!/usr/bin/env python3
"""
Inspiration Samples — One Illustration Per Visual Type

Queries the Notion illustrations / visual types / concepts databases, picks the
most-recently-created illustration per visual type, and copies the matching .png
from the Affinity Designer archive into a curated 'inspiration samples' folder.

Each sample is placed in two locations (big-picture → detail):
  - Flat:   {dest}/all/{concept_type} - {concept} - {visualtype} - {topic}.png
  - Nested: {dest}/{concept_type}/{concept}/{visualtype} - {topic}.png

Also writes an XLS index with text fields paired with their Notion IDs for
traceability.

Usage:
    python sample_illustrations.py
    python sample_illustrations.py --dry-run
    python sample_illustrations.py --config sample_illustrations.json
"""

import argparse
import json
import logging
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)

NOTION_VERSION = "2022-06-28"
WINDOWS_ILLEGAL = re.compile(r'[<>:"/\\|?*]')


def load_config(config_path: str) -> Dict[str, Any]:
    paths_to_try = [
        config_path,
        os.path.join(os.path.dirname(os.path.abspath(__file__)), os.path.basename(config_path)),
    ]
    for path in paths_to_try:
        try:
            with open(path, "r", encoding="utf-8") as f:
                config = json.load(f)
            if path != config_path:
                logging.info(f"Loaded config from fallback path: {path}")
            return config
        except FileNotFoundError:
            continue
    raise FileNotFoundError(f"Config not found at {' or '.join(paths_to_try)}")


def resolve_token(config: Dict[str, Any]) -> str:
    raw = config.get("notion_api_key", "")
    if isinstance(raw, str) and raw.startswith("${") and raw.endswith("}"):
        env_name = raw[2:-1]
        token = os.getenv(env_name)
    else:
        token = raw or os.getenv("NOTION_API_TOKEN")
    if not token:
        raise ValueError("NOTION_API_TOKEN not available (env var or config)")
    return token


def query_database(db_id: str, headers: Dict[str, str]) -> List[Dict[str, Any]]:
    """Fetch all pages from a Notion database, handling pagination."""
    url = f"https://api.notion.com/v1/databases/{db_id}/query"
    pages: List[Dict[str, Any]] = []
    body: Dict[str, Any] = {"page_size": 100}
    while True:
        r = requests.post(url, headers=headers, json=body)
        r.raise_for_status()
        data = r.json()
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        body["start_cursor"] = data["next_cursor"]
    logging.info(f"Fetched {len(pages)} pages from db {db_id[:8]}…")
    return pages


def extract_title(prop: Dict[str, Any]) -> str:
    return "".join(s.get("plain_text", "") for s in prop.get("title", [])).strip()


def extract_rich_text(prop: Dict[str, Any]) -> str:
    return "".join(s.get("plain_text", "") for s in prop.get("rich_text", [])).strip()


def extract_relation_ids(prop: Dict[str, Any]) -> List[str]:
    return [r["id"] for r in prop.get("relation", [])]


def extract_multi_select(prop: Dict[str, Any]) -> List[str]:
    return [x["name"] for x in prop.get("multi_select", [])]


def extract_select(prop: Dict[str, Any]) -> Optional[str]:
    sel = prop.get("select")
    return sel["name"] if sel else None


def build_concepts_map(pages: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for p in pages:
        props = p.get("properties", {})
        types = extract_multi_select(props.get("type", {}))
        out[p["id"]] = {
            "concept": extract_title(props.get("concept", {})),
            "types": types,
            "concept_type": "+".join(types) if types else "untyped",
        }
    return out


def build_visual_types_map(pages: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for p in pages:
        props = p.get("properties", {})
        concept_ids = extract_relation_ids(props.get("concept", {}))
        out[p["id"]] = {
            "visualtype": extract_title(props.get("visualtype", {})),
            "concept_id": concept_ids[0] if concept_ids else None,
            "illustration_ids": extract_relation_ids(props.get("illustrations", {})),
        }
    return out


def build_illustrations_map(pages: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for p in pages:
        props = p.get("properties", {})
        out[p["id"]] = {
            "title": extract_title(props.get("illustration", {})),
            "theme": extract_multi_select(props.get("theme", {})),
            "tags": extract_select(props.get("Tags", {})),
            "created": p.get("created_time") or props.get("Created", {}).get("created_time"),
        }
    return out


def sanitize(segment: str) -> str:
    cleaned = WINDOWS_ILLEGAL.sub("", segment or "").strip().rstrip(".")
    return cleaned or "_"


def extract_topic(title: str, visualtype: str) -> str:
    prefix = f"{visualtype} - "
    if title.startswith(prefix):
        return title[len(prefix):].strip()
    return title


def wipe_contents(folder: Path) -> None:
    """Remove everything inside folder (but keep the folder itself)."""
    if not folder.exists():
        folder.mkdir(parents=True, exist_ok=True)
        return
    for entry in folder.iterdir():
        if entry.is_dir():
            shutil.rmtree(entry)
        else:
            entry.unlink()


def main() -> int:
    parser = argparse.ArgumentParser(description="Sample illustrations from Notion into folders + XLS.")
    parser.add_argument("--config", default="sample_illustrations.json")
    parser.add_argument("--dry-run", action="store_true", help="Don't copy or write XLS")
    args = parser.parse_args()

    config = load_config(args.config)
    token = resolve_token(config)
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }

    source_folder = Path(config["source_folder"])
    dest_folder = Path(config["dest_folder"])
    xls_output = Path(config["xls_output"])

    if not source_folder.exists():
        logging.error(f"Source folder not found: {source_folder}")
        return 1

    logging.info("Querying Notion…")
    concepts = build_concepts_map(query_database(config["concepts_db_id"], headers))
    visual_types = build_visual_types_map(query_database(config["visual_types_db_id"], headers))
    illustrations = build_illustrations_map(query_database(config["illustrations_db_id"], headers))

    rows: List[Dict[str, Any]] = []
    skipped_no_illustrations = 0

    for vt_id, vt in visual_types.items():
        valid_ids = [i for i in vt["illustration_ids"] if i in illustrations]
        if not valid_ids:
            skipped_no_illustrations += 1
            continue

        # Most recently created first
        valid_ids.sort(key=lambda i: illustrations[i].get("created") or "", reverse=True)
        sample_id = valid_ids[0]
        sample = illustrations[sample_id]

        concept_id = vt["concept_id"]
        concept_info = concepts.get(concept_id, {}) if concept_id else {}
        concept_name = concept_info.get("concept", "")
        concept_type = concept_info.get("concept_type", "untyped")

        visualtype_name = vt["visualtype"]
        title = sample["title"]
        topic = extract_topic(title, visualtype_name)

        ct_s = sanitize(concept_type)
        c_s = sanitize(concept_name or "_no_concept_")
        vt_s = sanitize(visualtype_name)
        topic_s = sanitize(topic)

        flat_name = f"{ct_s} - {c_s} - {vt_s} - {topic_s}.png"
        nested_name = f"{vt_s} - {topic_s}.png"

        source_path = source_folder / f"{title}.png"
        dest_flat = dest_folder / "all" / flat_name
        dest_nested = dest_folder / ct_s / c_s / nested_name

        missing = not source_path.exists()

        rows.append({
            "concept_type": concept_type,
            "concept": concept_name,
            "concept_id": concept_id,
            "visual_type": visualtype_name,
            "visual_type_id": vt_id,
            "illustration_title": title,
            "illustration_id": sample_id,
            "topic": topic,
            "theme": ";".join(sample.get("theme") or []),
            "tags": sample.get("tags") or "",
            "created": sample.get("created") or "",
            "source_path": str(source_path),
            "dest_path_flat": str(dest_flat),
            "dest_path_nested": str(dest_nested),
            "n_total_in_visual_type": len(vt["illustration_ids"]),
            "missing_source_file": missing,
        })

    logging.info(f"Planned samples: {len(rows)} | Visual types with no illustrations: {skipped_no_illustrations}")
    missing_count = sum(1 for r in rows if r["missing_source_file"])
    if missing_count:
        logging.warning(f"{missing_count} sample(s) have no matching .png in source")

    if args.dry_run:
        logging.info("Dry run — no files copied, no XLS written")
        for r in rows[:5]:
            logging.info(f"  → {r['dest_path_flat']}")
        if len(rows) > 5:
            logging.info(f"  … and {len(rows) - 5} more")
        return 0

    logging.info(f"Wiping {dest_folder} …")
    wipe_contents(dest_folder)
    (dest_folder / "all").mkdir(parents=True, exist_ok=True)

    copied = 0
    for r in rows:
        if r["missing_source_file"]:
            continue
        src = Path(r["source_path"])
        flat = Path(r["dest_path_flat"])
        nested = Path(r["dest_path_nested"])
        flat.parent.mkdir(parents=True, exist_ok=True)
        nested.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, flat)
        shutil.copy2(src, nested)
        copied += 1

    logging.info(f"Copied {copied} illustration(s) to flat + nested layouts")

    xls_output.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    df.to_excel(xls_output, index=False, engine="openpyxl")
    logging.info(f"Wrote XLS: {xls_output} ({len(df)} rows)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
