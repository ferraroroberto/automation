#!/usr/bin/env python3
"""
Simple test script to verify Google Photos API access
"""

import json
import os
from pathlib import Path
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def test_photos_api():
    """Test basic Google Photos API access"""
    
    # Load credentials - use script directory to find token.json
    script_dir = Path(__file__).parent
    token_path = script_dir / "token.json"
    if not token_path.exists():
        print(f"❌ No token.json found at {token_path}. Run the main script first.")
        return
    
    print(f"🔍 Found token at: {token_path}")
    
    with open(token_path, 'r') as f:
        token_data = json.load(f)
    
    print(f"🔑 Token scopes: {token_data.get('scopes', 'No scopes found')}")
    
    # Create credentials
    creds = Credentials.from_authorized_user_file(str(token_path), token_data['scopes'])
    
    if not creds.valid:
        print("❌ Credentials are not valid")
        return
    
    print("✅ Credentials are valid")
    
    # Try to build the service
    try:
        photos_service = build('photoslibrary', 'v1', credentials=creds, static_discovery=False)
        print("✅ Photos API service built successfully")
        
        # Try a simple API call
        print("🔍 Testing API access...")
        response = photos_service.mediaItems().list(pageSize=1).execute()
        print(f"✅ API call successful! Found {len(response.get('mediaItems', []))} items")
        
    except HttpError as e:
        print(f"❌ HTTP Error: {e}")
        if "insufficient authentication scopes" in str(e):
            print("🔍 This suggests the API isn't enabled or there's a scope issue")
        elif "API not enabled" in str(e):
            print("🔍 The Google Photos Library API is not enabled in your project")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

if __name__ == "__main__":
    test_photos_api()
