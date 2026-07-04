#!/usr/bin/env python3
"""Shared Google OAuth + config-loading helpers.

Consolidates logic that was independently duplicated in
gmail_drive_automation.py and weekly_photo_automation.py: the config
JSON loader and the load-token / refresh / re-authenticate / save-token
dance (audit issue #64). Each caller still builds its own set of API
services from the returned credentials, since that varies by script.
"""

import json
import logging
from pathlib import Path
from typing import Any, Dict, List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

logger = logging.getLogger(__name__)


def load_config(config_path: str, script_dir: Path) -> Dict[str, Any]:
    """Load configuration from a JSON file, resolving a relative path against script_dir."""
    logger.debug("📂 Loading configuration file")

    if Path(config_path).is_absolute():
        config_full_path = Path(config_path)
    else:
        config_full_path = script_dir / config_path

    if not config_full_path.exists():
        logger.error(f"❌ Error: Configuration file not found at {config_full_path}")
        raise FileNotFoundError(f"Configuration file not found: {config_full_path}")

    try:
        with open(config_full_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info("✅ Configuration loaded successfully")
        return config
    except json.JSONDecodeError as e:
        logger.error(f"❌ Error: Invalid JSON in configuration file: {e}")
        raise


def authenticate(config: Dict[str, Any], script_dir: Path, scopes: List[str]) -> Credentials:
    """Load/refresh/create OAuth credentials for the given scopes.

    Resolves token_file/credentials_file from config['auth'] (relative paths
    resolved against script_dir), tries to load + refresh an existing token,
    and falls back to the interactive InstalledAppFlow when there's no valid
    token. Persists the (possibly new) token before returning.
    """
    creds = None

    token_path = Path(config['auth']['token_file'])
    if not token_path.is_absolute():
        token_path = script_dir / token_path

    credentials_path = Path(config['auth']['credentials_file'])
    if not credentials_path.is_absolute():
        credentials_path = script_dir / credentials_path

    # Load existing token
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        # Refresh credentials to ensure they're valid
        try:
            creds.refresh(Request())
            logger.info("🔄 Credentials refreshed successfully")
        except Exception as e:
            logger.warning(f"⚠️ Failed to refresh credentials: {e}")
            creds = None

    # If there are no (valid) credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logger.info("🔄 Refreshing expired credentials")
            creds.refresh(Request())
        else:
            logger.info("🔑 Requesting new authentication")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), scopes
            )
            creds = flow.run_local_server(
                port=0,
                access_type="offline",  # get a refresh token
                prompt="consent",  # force the consent screen, don't reuse old grant
                include_granted_scopes=False  # don't merge with an older, narrower grant
            )
            # force a fresh access token right now
            creds.refresh(Request())
            logger.info(f"Granted scopes from token: {creds.scopes}")

        # Save credentials for next run
        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    return creds
