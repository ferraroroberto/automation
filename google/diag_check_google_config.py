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

log = logging.getLogger(__name__)
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

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
                                          capture_output=True, text=True, check=True)
                    gcloud_cmd = 'gcloud'
                    break
                elif os.path.exists(path):
                    gcloud_cmd = path
                    break
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
                              capture_output=True, text=True, check=True)
        current_project = result.stdout.strip()
        print(f"   📋 Current gcloud project: {current_project}")
        
        # Check account
        result = subprocess.run([gcloud_cmd, 'auth', 'list', '--filter=status:ACTIVE'], 
                              capture_output=True, text=True, check=True)
        active_accounts = result.stdout.strip()
        print(f"   👤 Active gcloud accounts:\n{active_accounts}")
        
        return current_project
    except subprocess.CalledProcessError as e:
        print(f"   ❌ gcloud not configured or not in PATH: {e}")
        return None
    except FileNotFoundError:
        print("   ❌ gcloud CLI not found. Install Google Cloud SDK first.")
        return None

def check_credentials_file():
    """Check the credentials file configuration"""
    print("\n🔑 Checking credentials file...")
    
    script_dir = Path(__file__).parent
    creds_path = script_dir / "client_secret.json"
    
    if not creds_path.exists():
        print(f"   ❌ Credentials file not found at {creds_path}")
        return None
    
    try:
        with open(creds_path, 'r') as f:
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
    
    script_dir = Path(__file__).parent
    token_path = script_dir / "token.json"
    
    if not token_path.exists():
        print(f"   ❌ Token file not found at {token_path}")
        return None
    
    try:
        with open(token_path, 'r') as f:
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
    
    script_dir = Path(__file__).parent
    creds_path = script_dir / "client_secret.json"
    token_path = script_dir / "token.json"
    
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
        
        # Test Photos API
        try:
            photos_service = build('photoslibrary', 'v1', credentials=creds, static_discovery=False)
            response = photos_service.mediaItems().list(pageSize=1).execute()
            print("   ✅ Photos API: Authentication successful")
        except Exception as e:
            print(f"   ❌ Photos API: {e}")
        
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

def check_project_alignment():
    """Check if gcloud project matches credentials project"""
    print("\n🔗 Checking project alignment...")
    
    gcloud_project = check_gcloud_config()
    creds_project = check_credentials_file()
    
    if gcloud_project and creds_project:
        if gcloud_project == creds_project:
            print(f"   ✅ Projects match: {gcloud_project}")
            return True
        else:
            print(f"   ❌ Project mismatch!")
            print(f"      gcloud project: {gcloud_project}")
            print(f"      Credentials project: {creds_project}")
            print(f"   💡 This could be causing your authentication issues!")
            return False
    else:
        print("   ⚠️ Could not determine project alignment")
        return False

def provide_solutions():
    """Provide solutions for common issues"""
    print("\n💡 Solutions:")
    print("   1. If projects don't match, switch gcloud to the correct project:")
    print(f"      gcloud config set project automation-469306")
    print("   2. Re-authenticate with the correct project:")
    print("      gcloud auth login")
    print("   3. Delete token.json and re-run authentication to get fresh tokens")
    print("   4. Ensure the Google Cloud project 'automation-469306' has the required APIs enabled:")
    print("      - Google Photos Library API")
    print("      - Gmail API")
    print("      - Google Drive API")
    print("   5. Check that your OAuth consent screen includes the required scopes")

def main():
    """Main diagnostic function"""
    print("🚀 Google Cloud Configuration Diagnostic")
    print("=" * 50)
    
    # Check project alignment
    project_aligned = check_project_alignment()
    
    # Check token file
    check_token_file()
    
    # Test authentication
    auth_works = test_authentication()
    
    # Provide solutions
    provide_solutions()
    
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
