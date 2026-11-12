#!/usr/bin/env python3
"""
Google Cloud Configuration Checker
Diagnoses authentication and project configuration issues
"""

import json
import logging
import os
import subprocess
import sys
from pathlib import Path
from typing import Optional, Tuple

log = logging.getLogger(__name__)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Suppress the console window each gcloud spawn would otherwise flash.
# Repo-root shared helper (scripts run from their own folder, so add the root).
sys.path.append(str(Path(__file__).resolve().parents[1]))
from no_window import NO_WINDOW as _NO_WINDOW  # noqa: E402

# gcloud starts a Python runtime per call and can take several seconds; 60 s means it is wedged.
_GCLOUD_TIMEOUT_S = 60

_GOOGLE_DIR = Path(__file__).resolve().parent


def _auth_files() -> Tuple[Path, Path]:
    """Credentials and token paths exactly as gmail_drive_automation resolves them.

    The live config is config.json; the sample is the fallback so the diagnostic
    still names the files a fresh setup will create.
    """
    for name in ("config.json", "config_gmail_drive.json.sample"):
        config_path = _GOOGLE_DIR / name
        if config_path.exists():
            auth = json.loads(config_path.read_text(encoding="utf-8"))["auth"]
            return _GOOGLE_DIR / auth["credentials_file"], _GOOGLE_DIR / auth["token_file"]
    raise FileNotFoundError(f"no config.json or config_gmail_drive.json.sample in {_GOOGLE_DIR}")


def check_gcloud_config():
    """Check current gcloud configuration"""
    print("🔍 Checking gcloud configuration...")
    
    # Check if we're on Windows and look for gcloud in common locations
    if os.name == 'nt':  # Windows
        gcloud_paths = [
            'gcloud',  # Try direct command first
            r'C:\Program Files (x86)\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd',
            r'C:\Program Files\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd',
            r'C:\Program Files (x86)\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.ps1',
            r'C:\Program Files\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.ps1'
        ]
        
        gcloud_cmd = None
        for path in gcloud_paths:
            try:
                if path == 'gcloud':
                    # Try direct command
                    result = subprocess.run(['gcloud', '--version'],
                                          capture_output=True, text=True, check=True,
                                          creationflags=_NO_WINDOW, timeout=_GCLOUD_TIMEOUT_S)
                    gcloud_cmd = 'gcloud'
                    break
                elif os.path.exists(path):
                    gcloud_cmd = path
                    break
            except subprocess.TimeoutExpired:
                print(f"   ⚠️ `gcloud --version` timed out after {_GCLOUD_TIMEOUT_S}s; trying known install paths.")
                continue
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
        
        if not gcloud_cmd:
            print("   ❌ gcloud CLI not found. Install Google Cloud SDK first.")
            return None
    else:
        gcloud_cmd = 'gcloud'
    
    try:
        # Check current project
        result = subprocess.run([gcloud_cmd, 'config', 'get-value', 'project'],
                              capture_output=True, text=True, check=True,
                              creationflags=_NO_WINDOW, timeout=_GCLOUD_TIMEOUT_S)
        current_project = result.stdout.strip()
        print(f"   📋 Current gcloud project: {current_project}")
        
        # Check account
        result = subprocess.run([gcloud_cmd, 'auth', 'list', '--filter=status:ACTIVE'],
                              capture_output=True, text=True, check=True,
                              creationflags=_NO_WINDOW, timeout=_GCLOUD_TIMEOUT_S)
        active_accounts = result.stdout.strip()
        print(f"   👤 Active gcloud accounts:\n{active_accounts}")

        return current_project
    except subprocess.TimeoutExpired:
        print(f"   ❌ gcloud did not answer within {_GCLOUD_TIMEOUT_S}s.")
        return None
    except subprocess.CalledProcessError as e:
        print(f"   ❌ gcloud not configured or not in PATH: {e}")
        return None
    except FileNotFoundError:
        print("   ❌ gcloud CLI not found. Install Google Cloud SDK first.")
        return None

def check_credentials_file():
    """Check the credentials file configuration"""
    print("\n🔑 Checking credentials file...")
    
    creds_path, _ = _auth_files()

    if not creds_path.exists():
        print(f"   ❌ Credentials file not found at {creds_path}")
        return None

    try:
        with open(creds_path, 'r', encoding='utf-8') as f:
            creds_data = json.load(f)
        
        project_id = creds_data.get('installed', {}).get('project_id')
        client_id = creds_data.get('installed', {}).get('client_id')
        
        print(f"   ✅ Credentials file found")
        print(f"   📋 Project ID: {project_id}")
        print(f"   🆔 Client ID: {client_id}")
        
        return project_id
    except Exception as e:
        print(f"   ❌ Error reading credentials file: {e}")
        return None

def check_token_file():
    """Check the token file and its scopes"""
    print("\n🎫 Checking token file...")
    
    _, token_path = _auth_files()

    if not token_path.exists():
        print(f"   ❌ Token file not found at {token_path}")
        return None

    try:
        with open(token_path, 'r', encoding='utf-8') as f:
            token_data = json.load(f)
        
        scopes = token_data.get('scopes', [])
        print(f"   ✅ Token file found")
        print(f"   🔐 Granted scopes ({len(scopes)}):")
        for scope in scopes:
            print(f"      • {scope}")
        
        return scopes
    except Exception as e:
        print(f"   ❌ Error reading token file: {e}")
        return None

def test_authentication():
    """Test authentication with current credentials"""
    print("\n🧪 Testing authentication...")
    
    creds_path, token_path = _auth_files()

    if not creds_path.exists():
        print("   ❌ Credentials file not found")
        return False
    
    if not token_path.exists():
        print("   ❌ Token file not found")
        return False
    
    try:
        # Load credentials
        creds = Credentials.from_authorized_user_file(str(token_path))
        
        if not creds.valid:
            if creds.expired and creds.refresh_token:
                print("   🔄 Token expired, attempting refresh...")
                creds.refresh(Request())
            else:
                print("   ❌ Credentials not valid and cannot refresh")
                return False
        
        print("   ✅ Credentials loaded successfully")
        print(f"   🔑 Token scopes: {creds.scopes}")
        
        # Test Gmail API
        try:
            gmail_service = build('gmail', 'v1', credentials=creds)
            profile = gmail_service.users().getProfile(userId='me').execute()
            print(f"   ✅ Gmail API: Authentication successful (Email: {profile.get('emailAddress', 'N/A')})")
        except Exception as e:
            print(f"   ❌ Gmail API: {e}")
        
        # Test Google Drive API
        try:
            drive_service = build('drive', 'v3', credentials=creds)
            about = drive_service.about().get(fields='user,storageQuota').execute()
            user_info = about.get('user', {})
            storage_info = about.get('storageQuota', {})
            print(f"   ✅ Drive API: Authentication successful (User: {user_info.get('displayName', 'N/A')})")
            if storage_info:
                total = storage_info.get('limit', 'N/A')
                used = storage_info.get('usage', 'N/A')
                print(f"      💾 Storage: {used} / {total} bytes")
        except Exception as e:
            print(f"   ❌ Drive API: {e}")
        
        return True
        
    except Exception as e:
        print(f"   ❌ Authentication error: {e}")
        return False

def check_project_alignment() -> Tuple[bool, Optional[str]]:
    """Check if gcloud project matches credentials project.

    Returns whether they match and the credentials file's project id.
    """
    print("\n🔗 Checking project alignment...")

    gcloud_project = check_gcloud_config()
    creds_project = check_credentials_file()

    if gcloud_project and creds_project:
        if gcloud_project == creds_project:
            print(f"   ✅ Projects match: {gcloud_project}")
            return True, creds_project
        else:
            print(f"   ❌ Project mismatch!")
            print(f"      gcloud project: {gcloud_project}")
            print(f"      Credentials project: {creds_project}")
            print(f"   💡 This could be causing your authentication issues!")
            return False, creds_project
    else:
        print("   ⚠️ Could not determine project alignment")
        return False, creds_project

def provide_solutions(creds_project: Optional[str]) -> None:
    """Provide solutions for common issues"""
    _, token_path = _auth_files()
    project = creds_project or "<the credentials file's project_id>"
    print("\n💡 Solutions:")
    print("   1. If projects don't match, switch gcloud to the credentials project:")
    print(f"      gcloud config set project {project}")
    print("   2. Re-authenticate with the correct project:")
    print("      gcloud auth login")
    print(f"   3. Delete {token_path.name} and re-run authentication to get fresh tokens")
    print(f"   4. Ensure the Google Cloud project {project} has the required APIs enabled:")
    print("      - Gmail API")
    print("      - Google Drive API")
    print("      - Google Sheets API")
    print("   5. Check that your OAuth consent screen includes the required scopes")

def main():
    """Main diagnostic function"""
    print("🚀 Google Cloud Configuration Diagnostic")
    print("=" * 50)
    
    # Check project alignment
    project_aligned, creds_project = check_project_alignment()
    
    # Check token file
    check_token_file()
    
    # Test authentication
    auth_works = test_authentication()
    
    # Provide solutions
    provide_solutions(creds_project)
    
    print("\n📊 Summary:")
    print(f"   Project alignment: {'✅' if project_aligned else '❌'}")
    print(f"   Authentication: {'✅' if auth_works else '❌'}")
    
    if not project_aligned:
        print("\n⚠️  WARNING: Your gcloud project doesn't match your credentials project!")
        print("   This is likely the root cause of your authentication issues.")
        print("   Please switch to the correct project and re-authenticate.")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    main()
