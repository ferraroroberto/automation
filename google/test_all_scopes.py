#!/usr/bin/env python3
"""
Test all Google API scopes to find the exact combination needed
"""

import json
import os
import sys
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def delete_old_token():
    """Delete existing token file"""
    script_dir = Path(__file__).parent
    token_path = script_dir / "token.json"
    
    if token_path.exists():
        print("🗑️  Deleting old token file...")
        os.remove(token_path)
        print("✅ Old token deleted")

def authenticate_with_all_scopes():
    """Authenticate with ALL possible scopes"""
    print("\n🔐 Requesting ALL scopes from your OAuth consent screen...")
    
    # ALL scopes from your OAuth consent screen
    SCOPES = [
        # Photos Library API - Core scopes
        'https://www.googleapis.com/auth/photoslibrary',
        'https://www.googleapis.com/auth/photoslibrary.readonly',
        'https://www.googleapis.com/auth/photoslibrary.appendonly',
        'https://www.googleapis.com/auth/photoslibrary.sharing',
        'https://www.googleapis.com/auth/photoslibrary.edit.appcreateddata',
        'https://www.googleapis.com/auth/photoslibrary.readonly.appcreateddata',
        
        # Gmail API - All scopes you've added
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/gmail.compose',
        'https://www.googleapis.com/auth/gmail.modify',
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/gmail.metadata',
        'https://www.googleapis.com/auth/gmail.insert',
        'https://www.googleapis.com/auth/gmail.labels',
        'https://www.googleapis.com/auth/gmail.settings.basic',
        'https://www.googleapis.com/auth/gmail.settings.sharing',
        'https://www.googleapis.com/auth/gmail.addons.current.action.compose',
        'https://www.googleapis.com/auth/gmail.addons.current.message.action',
        'https://www.googleapis.com/auth/gmail.addons.current.message.metadata',
        'https://www.googleapis.com/auth/gmail.addons.current.message.readonly',
        'https://mail.google.com/'  # Full Gmail access
    ]
    
    script_dir = Path(__file__).parent
    creds_path = script_dir / "client_secret.json"
    token_path = script_dir / "token_all_scopes.json"  # Different file for testing
    
    if not creds_path.exists():
        print("❌ client_secret.json not found!")
        return None
    
    print(f"📋 Requesting {len(SCOPES)} scopes:")
    for scope in SCOPES:
        print(f"   • {scope}")
    
    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(creds_path), 
            SCOPES
        )
        
        creds = flow.run_local_server(
            port=0,
            authorization_prompt_message='Please authorize ALL scopes in your browser.',
            success_message='Authorization successful! You can close this window.',
            open_browser=True,
            access_type='offline',
            prompt='consent'
        )
        
        # Save the credentials
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
        
        print(f"\n✅ Token saved with {len(creds.scopes)} scopes granted")
        print("\n🔍 Actually granted scopes:")
        for scope in creds.scopes:
            print(f"   ✓ {scope}")
            
        return creds
        
    except Exception as e:
        print(f"❌ Authentication failed: {e}")
        return None

def test_photos_api(creds):
    """Test various Photos API operations"""
    print("\n📷 Testing Google Photos API operations...")
    results = {}
    
    try:
        service = build(
            'photoslibrary', 
            'v1', 
            credentials=creds,
            discoveryServiceUrl='https://photoslibrary.googleapis.com/$discovery/rest?version=v1',
            cache_discovery=False
        )
        
        # Test 1: List albums
        try:
            response = service.albums().list(pageSize=1).execute()
            print("✅ List albums: SUCCESS")
            results['list_albums'] = True
        except HttpError as e:
            print(f"❌ List albums: FAILED - {e.resp.status}")
            results['list_albums'] = False
        
        # Test 2: Create album
        try:
            album = service.albums().create(
                body={"album": {"title": "Test Album - Delete Me 2"}}
            ).execute()
            print(f"✅ Create album: SUCCESS (ID: {album['id']})")
            results['create_album'] = True
            
            # Test 3: Share album
            try:
                share = service.albums().share(
                    albumId=album['id'],
                    body={"sharedAlbumOptions": {"isCollaborative": False, "isCommentable": True}}
                ).execute()
                print("✅ Share album: SUCCESS")
                results['share_album'] = True
            except HttpError as e:
                print(f"❌ Share album: FAILED - {e.resp.status}")
                results['share_album'] = False
                
        except HttpError as e:
            print(f"❌ Create album: FAILED - {e.resp.status}")
            results['create_album'] = False
            results['share_album'] = False
        
        # Test 4: List media items
        try:
            response = service.mediaItems().list(pageSize=1).execute()
            print("✅ List media items: SUCCESS")
            results['list_media'] = True
        except HttpError as e:
            print(f"❌ List media items: FAILED - {e.resp.status}")
            results['list_media'] = False
        
        # Test 5: Search media items
        try:
            from datetime import datetime, timedelta
            yesterday = datetime.now() - timedelta(days=1)
            filters = {
                "dateFilter": {
                    "ranges": [{
                        "startDate": {
                            "year": yesterday.year,
                            "month": yesterday.month,
                            "day": yesterday.day
                        },
                        "endDate": {
                            "year": yesterday.year,
                            "month": yesterday.month,
                            "day": yesterday.day
                        }
                    }]
                }
            }
            response = service.mediaItems().search(body={"filters": filters, "pageSize": 1}).execute()
            print("✅ Search media items: SUCCESS")
            results['search_media'] = True
        except HttpError as e:
            print(f"❌ Search media items: FAILED - {e.resp.status}")
            results['search_media'] = False
            
    except Exception as e:
        print(f"❌ Failed to build Photos service: {e}")
        return results
    
    return results

def test_gmail_api(creds):
    """Test various Gmail API operations"""
    print("\n📧 Testing Gmail API operations...")
    results = {}
    
    try:
        service = build('gmail', 'v1', credentials=creds)
        
        # Test 1: Get profile
        try:
            profile = service.users().getProfile(userId='me').execute()
            print(f"✅ Get profile: SUCCESS (Email: {profile.get('emailAddress')})")
            results['get_profile'] = True
        except HttpError as e:
            print(f"❌ Get profile: FAILED - {e.resp.status}")
            results['get_profile'] = False
        
        # Test 2: List labels
        try:
            labels = service.users().labels().list(userId='me').execute()
            print(f"✅ List labels: SUCCESS ({len(labels.get('labels', []))} labels)")
            results['list_labels'] = True
        except HttpError as e:
            print(f"❌ List labels: FAILED - {e.resp.status}")
            results['list_labels'] = False
        
        # Test 3: Create draft
        try:
            import base64
            from email.mime.text import MIMEText
            
            message = MIMEText("Test message body")
            message['to'] = 'test@example.com'
            message['subject'] = 'Test Draft'
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
            
            draft = service.users().drafts().create(
                userId='me',
                body={'message': {'raw': raw}}
            ).execute()
            print(f"✅ Create draft: SUCCESS (ID: {draft['id']})")
            results['create_draft'] = True
            
            # Clean up - delete the draft
            service.users().drafts().delete(userId='me', id=draft['id']).execute()
            
        except HttpError as e:
            print(f"❌ Create draft: FAILED - {e.resp.status}")
            results['create_draft'] = False
        
        # Test 4: List messages (just metadata)
        try:
            messages = service.users().messages().list(userId='me', maxResults=1).execute()
            print("✅ List messages: SUCCESS")
            results['list_messages'] = True
        except HttpError as e:
            print(f"❌ List messages: FAILED - {e.resp.status}")
            results['list_messages'] = False
            
    except Exception as e:
        print(f"❌ Failed to build Gmail service: {e}")
        return results
    
    return results

def find_minimal_scopes(creds):
    """Determine minimal required scopes"""
    print("\n🔬 Analyzing which scopes are actually needed...")
    
    granted_scopes = set(creds.scopes)
    
    # Essential Photos scopes for your use case
    essential_photos = {
        'https://www.googleapis.com/auth/photoslibrary',  # Create albums
        'https://www.googleapis.com/auth/photoslibrary.sharing',  # Share albums
        'https://www.googleapis.com/auth/photoslibrary.readonly'  # Read photos
    }
    
    # Essential Gmail scopes for your use case
    essential_gmail = {
        'https://www.googleapis.com/auth/gmail.send',  # Send emails
    }
    
    # Check if we might need additional Gmail scopes
    gmail_alternatives = {
        'https://www.googleapis.com/auth/gmail.compose',  # Compose and send
        'https://mail.google.com/',  # Full access
        'https://www.googleapis.com/auth/gmail.modify'  # Modify messages
    }
    
    print("\n📋 Minimal scopes needed for your automation:")
    print("\nPhotos API:")
    for scope in essential_photos:
        status = "✅" if scope in granted_scopes else "❌"
        print(f"   {status} {scope}")
    
    print("\nGmail API (need at least one of these):")
    for scope in essential_gmail:
        status = "✅" if scope in granted_scopes else "❌"
        print(f"   {status} {scope}")
    
    print("\nAlternative Gmail scopes that might work:")
    for scope in gmail_alternatives:
        if scope in granted_scopes:
            print(f"   ✅ {scope} (GRANTED - this should work!)")

def main():
    """Main test function"""
    print("🔧 Comprehensive Google API Scope Tester")
    print("=" * 50)
    
    # Delete old token
    delete_old_token()
    
    # Authenticate with all scopes
    creds = authenticate_with_all_scopes()
    if not creds:
        print("\n❌ Authentication failed")
        sys.exit(1)
    
    # Test APIs
    photos_results = test_photos_api(creds)
    gmail_results = test_gmail_api(creds)
    
    # Analyze results
    print("\n" + "=" * 50)
    print("📊 Test Results Summary:")
    
    print("\n📷 Photos API:")
    photos_working = sum(photos_results.values())
    photos_total = len(photos_results)
    print(f"   {photos_working}/{photos_total} operations successful")
    for op, success in photos_results.items():
        status = "✅" if success else "❌"
        print(f"   {status} {op}")
    
    print("\n📧 Gmail API:")
    gmail_working = sum(gmail_results.values())
    gmail_total = len(gmail_results)
    print(f"   {gmail_working}/{gmail_total} operations successful")
    for op, success in gmail_results.items():
        status = "✅" if success else "❌"
        print(f"   {status} {op}")
    
    # Find minimal scopes
    find_minimal_scopes(creds)
    
    # Final recommendation
    print("\n" + "=" * 50)
    if photos_working >= 3 and gmail_working >= 1:
        print("🎉 SUCCESS! Your APIs are working!")
        print("\n📝 Next steps:")
        print("1. Copy token_all_scopes.json to token.json:")
        print("   cp token_all_scopes.json token.json")
        print("2. Run your weekly photo automation")
    else:
        print("⚠️ Some operations are still failing.")
        print("\n🔍 Troubleshooting:")
        print("1. Check that ALL scopes are added to your OAuth consent screen")
        print("2. If in 'Testing' mode, ensure your email is added as a test user")
        print("3. Try publishing your OAuth app (set to 'In Production')")
        print("4. Create a new OAuth 2.0 Client ID and try again")

if __name__ == "__main__":
    main()