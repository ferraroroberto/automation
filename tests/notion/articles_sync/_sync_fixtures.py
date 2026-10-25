"""Shared builders for the articles-sync tests: Notion property payloads and a
`NotionArticlesSync` shell without `__init__` side effects."""

import sys
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "notion" / "articles_sync"))

from notion_articles_sync import NotionArticlesSync  # noqa: E402


def prop_rich_text(value):
    return {"type": "rich_text", "rich_text": [{"plain_text": value, "text": {"content": value}}]}


def prop_title(value):
    return {"type": "title", "title": [{"plain_text": value, "text": {"content": value}}]}


def prop_formula_string(value):
    return {"type": "formula", "formula": {"type": "string", "string": value}}


def prop_select(name):
    return {"type": "select", "select": {"name": name}}


def prop_multi_select(*names):
    return {"type": "multi_select", "multi_select": [{"name": n} for n in names]}


def prop_date(start):
    return {"type": "date", "date": {"start": start}}


def prop_checkbox(value):
    return {"type": "checkbox", "checkbox": value}


def make_sync(field_mapping=None):
    """Build a NotionArticlesSync without running __init__ side effects."""
    sync = NotionArticlesSync.__new__(NotionArticlesSync)
    sync.field_mapping = field_mapping or {
        "article": "article",
        "topic": "topic",
        "rowid": "source rowid",
        "created": "created",
    }
    sync.api_call_count = 0
    sync.api_call_lock = threading.Lock()
    sync._schema_cache = {}
    return sync
