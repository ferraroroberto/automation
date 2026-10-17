# requirements: public
import logging
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import os
from zipfile import BadZipFile
import json
import requests
from dotenv import load_dotenv
from typing import Optional, Dict, List, Any, Tuple

log = logging.getLogger(__name__)


# Default location of the legacy notion params txt file (single source of truth).
# Set the NOTION_PARAMS_FILE environment variable to the path on each machine.
DEFAULT_PARAMS_FILE = os.getenv("NOTION_PARAMS_FILE", "")

# Function to read the parameters from the txt file (legacy support)
def read_params_from_txt_file(file_path: str) -> Dict[str, str]:
    params = {}
    with open(file_path, 'r') as f:
        for line in f:
            if line.strip():
                key, value = line.strip().split(" = ", 1)
                params[key.strip()] = value.strip()
    return params

# Function to load JSON configuration file
def load_json_config(config_path: str) -> Dict[str, Any]:
    """
    Load configuration from a JSON file.

    Parameters:
    - config_path: Path to the JSON configuration file

    Returns:
    - Dictionary containing configuration data
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        log.info("✅ Configuration loaded from %s", config_path)
        return config
    except FileNotFoundError:
        log.error("❌ Configuration file not found: %s", config_path)
        raise
    except json.JSONDecodeError as e:
        log.error("❌ Error parsing JSON configuration: %s", e)
        raise

# Function to load environment variables from .env file
def load_env_variables(env_path: Optional[str] = None) -> Dict[str, Optional[str]]:
    """
    Load environment variables from .env file.

    Parameters:
    - env_path: Optional path to .env file. If None, searches in current directory

    Returns:
    - Dictionary containing environment variables
    """
    if env_path:
        load_dotenv(env_path)
    else:
        load_dotenv()

    env_vars = {
        'notion_api_token': os.getenv('NOTION_API_TOKEN'),
        'notion_token_v2': os.getenv('NOTION_TOKEN_V2'),
        'notion_workspace_url': os.getenv('NOTION_WORKSPACE_URL')
    }

    log.info("✅ Environment variables loaded")
    return env_vars

# Get the column widths from the existing Excel file, initializing column_widths as an empty list first
def get_column_widths(excel_path: str) -> List[Any]:
    column_widths = []
    if not os.path.exists(excel_path):
        log.warning("No existing workbook at %s; skipping column width reuse.", excel_path)
        return column_widths

    try:
        wb_existing = openpyxl.load_workbook(excel_path)
        ws_existing = wb_existing.active

        # Find the last column with content
        last_col = ws_existing.max_column

        # Get column widths for columns with content
        column_widths = [
            ws_existing.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width
            for i in range(last_col)
        ]
    except (InvalidFileException, KeyError, BadZipFile) as exc:
        log.warning("Unable to read column widths from %s: %s", excel_path, exc)

    return column_widths

# Apply column widths to an excel file (requires the column_widths)
def apply_column_widths(excel_path: str, column_widths: List[Any]) -> None:
    if column_widths:
        wb_final = openpyxl.load_workbook(excel_path)
        ws_final = wb_final.active
        for i, width in enumerate(column_widths):
            ws_final.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = width
        wb_final.save(excel_path)


# ---------------------------------------------------------------------------
# Shared Notion API helpers (dedup: audit issue #64)
#
# These consolidate logic that was independently reimplemented across
# build_newsletter.py, normalize_names.py, normalize_url.py,
# journal_automation.py, and sample_illustrations.py.
# ---------------------------------------------------------------------------

NOTION_VERSION = "2022-06-28"


def build_notion_headers(api_key: str) -> Dict[str, str]:
    """Standard Notion API request headers."""
    return {
        'Authorization': f'Bearer {api_key}',
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
    }


def resolve_notion_credentials(config: Dict[str, Any]) -> Tuple[str, str, Dict[str, str]]:
    """Resolve (api_key, database_id, headers) from a loaded config dict / env.

    Mirrors the identical `_setup_api_credentials` previously duplicated in
    normalize_names.py and normalize_url.py. Raises ValueError if either the
    API key or the database_id can't be resolved.
    """
    api_key = os.getenv('NOTION_API_TOKEN') or config.get('notion_api_key')
    database_id = config.get('database_id')

    if not all([api_key, database_id]):
        raise ValueError("Missing required configuration values: notion_api_key and database_id")

    return api_key, database_id, build_notion_headers(api_key)


def load_json_config_with_fallback(config_path: str, caller_file: str) -> Dict[str, Any]:
    """Load a JSON config file, falling back to caller_file's own directory.

    Mirrors the "try config_path, then try alongside the calling script" pattern
    previously duplicated in build_newsletter.py, normalize_names.py,
    normalize_url.py, and sample_illustrations.py.
    """
    paths_to_try = [
        config_path,
        os.path.join(os.path.dirname(os.path.abspath(caller_file)), os.path.basename(config_path)),
    ]

    for path in paths_to_try:
        try:
            with open(path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            if path != config_path:
                log.info("📁 Loaded config from fallback path: %s", path)
            return config
        except FileNotFoundError:
            continue
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON in configuration file: {path}")

    raise FileNotFoundError(f"Configuration file not found at {' or '.join(paths_to_try)}")


def paginated_database_query(
    headers: Dict[str, str],
    database_id: str,
    query_body: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """POST a Notion `databases.query`, following `has_more`/`next_cursor` pagination.

    `query_body` may carry any combination of 'filter' / 'sorts' / 'page_size' (or
    none); it is not mutated in place, a per-page copy receives 'start_cursor'.

    Consolidates the paginated-query loop previously duplicated (with an
    identical POST / has_more / next_cursor / error-branch shape) across
    build_newsletter.py, normalize_names.py, normalize_url.py, and
    journal_automation.py.
    """
    url = f"https://api.notion.com/v1/databases/{database_id}/query"
    payload: Dict[str, Any] = dict(query_body) if query_body else {}

    all_results: List[Dict[str, Any]] = []
    has_more = True
    start_cursor = None

    while has_more:
        if start_cursor:
            payload['start_cursor'] = start_cursor

        try:
            response = requests.post(url, headers=headers, json=payload)
            response.raise_for_status()

            data = response.json()
            results = data.get('results', [])
            all_results.extend(results)
            log.debug("📥 Retrieved %s pages (total: %s)", len(results), len(all_results))

            has_more = data.get('has_more', False)
            start_cursor = data.get('next_cursor')

        except requests.exceptions.RequestException as e:
            log.error("❌ Failed to query Notion database: %s", e)
            if getattr(e, 'response', None):
                log.error("📊 Status: %s, Response: %s...", e.response.status_code, (e.response.text or '')[:200])
            raise

    log.info("📊 Total pages retrieved: %s", len(all_results))
    return all_results


# ---------------------------------------------------------------------------
# Shared Notion property-value extraction (dedup: audit issue #64)
#
# Consolidates the "title / rich_text / select / multi_select / date /
# checkbox / url" extraction reimplemented across build_newsletter.py,
# journal_automation.py, sample_illustrations.py, and
# articles_sync/notion_articles_sync.py.
# ---------------------------------------------------------------------------

def extract_title_text(prop: Dict[str, Any]) -> str:
    """Join all rich-text segments of a 'title' property into plain text."""
    return ''.join(s.get('plain_text', '') for s in prop.get('title', [])).strip()


def extract_rich_text_value(prop: Dict[str, Any]) -> str:
    """Join all rich-text segments of a 'rich_text' property into plain text."""
    return ''.join(s.get('plain_text', '') for s in prop.get('rich_text', [])).strip()


def extract_select_name(prop: Dict[str, Any]) -> str:
    """Selected option's name, or '' if unset."""
    option = prop.get('select')
    return option.get('name', '') if option else ''


def extract_multi_select_names(prop: Dict[str, Any]) -> List[str]:
    """Names of all selected multi_select options (unfiltered, unstripped)."""
    return [opt.get('name', '') for opt in prop.get('multi_select', [])]


def extract_relation_ids(prop: Dict[str, Any]) -> List[str]:
    """Page ids of all related pages in a 'relation' property."""
    return [r.get('id') for r in prop.get('relation', [])]


def extract_checkbox(prop: Dict[str, Any]) -> bool:
    """Checkbox state, defaulting to False."""
    return prop.get('checkbox', False)


def extract_date_start(prop: Dict[str, Any]) -> str:
    """Start date/time of a 'date' property, or '' if unset."""
    date_obj = prop.get('date')
    return date_obj.get('start', '') if date_obj else ''


def extract_url(prop: Dict[str, Any]) -> str:
    """URL value of a 'url' property, or '' if unset."""
    return prop.get('url') or ''


_PROPERTY_EXTRACTORS = {
    'title': extract_title_text,
    'rich_text': extract_rich_text_value,
    'select': extract_select_name,
    'multi_select': extract_multi_select_names,
    'relation': extract_relation_ids,
    'checkbox': extract_checkbox,
    'date': extract_date_start,
    'url': extract_url,
}


def extract_property(prop: Dict[str, Any]) -> Any:
    """Extract the native Python value from a Notion property dict, dispatching
    on its 'type' field. Returns None for a type this helper doesn't know about
    (callers that need a different default should check `prop.get('type')` first).
    """
    extractor = _PROPERTY_EXTRACTORS.get(prop.get('type'))
    return extractor(prop) if extractor else None
