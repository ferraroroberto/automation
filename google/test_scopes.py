#!/usr/bin/env python3
"""
Test different scope variations for Google Photos API
"""

import json
from pathlib import Path
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def test_different_scopes():
    """Test various scope combinations"""
    
    # Different scope variations to try
    scope_variations = [
        # Original scopes
        [
            'https://www.googleapis.com/auth/photoslibrary',
            'https://www.googleapis.com/auth/photoslibrary.sharing',
            'https://www.googleapis.com/auth/gmail.send'
        ],
        # Alternative scope format
        [
            'https://www.googleapis.com/auth/photoslibrary.readonly',
            'https://www.googleapis.com/auth/photoslibrary',
            'https://www.googleapis.com/auth/gmail.send'
        ],
        # Minimal scope test
        [
            'https://www.googleapis.com/auth/photoslibrary'
        ]
    ]
    
    token_path = Path("token.json")
    if not token_path.exists():
        print("❌ No token.json found")
        return
    
    with open(token_path, 'r') as f:
        token_data = json.load(f)
    
    print(f"🔑 Current token scopes: {token_data.get('scopes', [])}")
    
    for i, scopes in enumerate(scope_variations, 1):
        print(f"\n🧪 Testing scope variation {i}: {scopes}")
        
        try:
            # Create credentials with these scopes
            creds = Credentials.from_authorized_user_file(str(token_path), scopes)
            
            if not creds.valid:
                print(f"   ❌ Credentials not valid with these scopes")
                continue
            
            # Try to build service
            photos_service = build('photoslibrary', 'v1', credentials=creds, static_discovery=False)
            print(f"   ✅ Service built successfully")
            
            # Try API call
            response = photos_service.mediaItems().list(pageSize=1).execute()
            print(f"   ✅ API call successful! Found {len(response.get('mediaItems', []))} items")
            print(f"   🎉 This scope combination works!")
            return scopes
            
        except HttpError as e:
            print(f"   ❌ HTTP Error: {e}")
        except Exception as e:
            print(f"   ❌ Error: {e}")
    
    print("\n❌ No scope combination worked")
    return None

if __name__ == "__main__":
    test_different_scopes()
