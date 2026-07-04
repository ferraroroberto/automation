#!/usr/bin/env python3
"""
Weekly Photo Album Automation
Creates weekly Google Photos albums for children and shares them via email.
"""

# Standard library imports
import argparse
import logging
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

# Third-party library imports
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError, UnknownApiNameOrVersion
import pytz

import _auth

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class WeeklyPhotoAutomation:
    """Manages weekly photo album creation and sharing."""
    
    def __init__(self, config_path: str = "config.json"):
        """
        Initialize the automation with configuration.
        
        Args:
            config_path: Path to the configuration file
        """
        self.config = self.load_config(config_path)
        self.photos_service = None
        self.gmail_service = None
        self.credentials = None
        
    def load_config(self, config_path: str) -> Dict[str, Any]:
        """
        Load configuration from JSON file.

        Args:
            config_path: Path to configuration file

        Returns:
            Configuration dictionary
        """
        return _auth.load_config(config_path, Path(__file__).parent)

    def authenticate_google_services(self) -> None:
        """Authenticate with Google Photos and Gmail APIs."""
        logger.info("🔐 Authenticating with Google services")

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

        creds = _auth.authenticate(self.config, Path(__file__).parent, SCOPES)
        self.credentials = creds

        # Build Google Photos Library API service.
        # Only catch discovery-related failures here (missing/outdated static
        # discovery doc, or an ImportError from the discovery machinery) —
        # a real HttpError (e.g. auth/scope error) must propagate immediately
        # instead of being swallowed and masked by three retries that will
        # all fail the same way (audit issue #67).
        try:
            # First try with static discovery disabled
            self.photos_service = build('photoslibrary', 'v1', credentials=creds, static_discovery=False)
            logger.info("✅ Google Photos API service built successfully")
        except (ImportError, UnknownApiNameOrVersion) as e:
            logger.error(f"❌ Failed to build Photos API service with static_discovery=False: {e}")
            try:
                # Try with cache discovery disabled
                self.photos_service = build('photoslibrary', 'v1', credentials=creds, cache_discovery=False)
                logger.info("✅ Google Photos API service built with cache_discovery=False")
            except (ImportError, UnknownApiNameOrVersion) as e2:
                logger.error(f"❌ Failed to build Photos API service with cache_discovery=False: {e2}")
                try:
                    # Try with discovery service URL
                    self.photos_service = build('photoslibrary', 'v1', credentials=creds,
                                             discoveryServiceUrl='https://photoslibrary.googleapis.com/$discovery/rest?version=v1')
                    logger.info("✅ Google Photos API service built with custom discovery URL")
                except (ImportError, UnknownApiNameOrVersion) as e3:
                    logger.error(f"❌ All Photos API build attempts failed: {e3}")
                    raise Exception(f"Could not build Photos API service after multiple attempts. Last error: {e3}")
        
        # Build Gmail API service
        try:
            self.gmail_service = build('gmail', 'v1', credentials=creds)
            logger.info("✅ Gmail API service built successfully")
        except Exception as e:
            logger.error(f"❌ Failed to build Gmail API service: {e}")
            raise Exception(f"Could not build Gmail API service: {e}")
            
        logger.info("✅ Authentication successful")
        logger.info(f"🔑 Granted scopes: {creds.scopes}")

        
    def calculate_week_range(self) -> Tuple[datetime, datetime]:
        """
        Calculate the date range for the last complete week (Saturday to Friday).
        
        Returns:
            Tuple of (start_date, end_date) for the week
        """
        today = datetime.now()
        
        # Find the most recent Friday
        days_since_friday = (today.weekday() - 4) % 7
        if days_since_friday == 0 and today.hour < 23:  # If today is Friday but not complete
            days_since_friday = 7
            
        last_friday = today - timedelta(days=days_since_friday)
        last_friday = last_friday.replace(hour=23, minute=59, second=59)
        
        # Calculate the Saturday before (start of week)
        last_saturday = last_friday - timedelta(days=6)
        last_saturday = last_saturday.replace(hour=0, minute=0, second=0)
        
        logger.info(f"📅 Week range: {last_saturday.date()} to {last_friday.date()}")
        return last_saturday, last_friday
        
    def calculate_week_numbers(self, week_start: datetime) -> Dict[str, int]:
        """
        Calculate week numbers for each child based on their birth dates.
        
        Args:
            week_start: Start date of the current week
            
        Returns:
            Dictionary with child names and their week numbers
        """
        week_numbers = {}
        
        for child in self.config['children']:
            birth_date = datetime.strptime(child['birth_date'], '%Y-%m-%d')
            weeks_since_birth = (week_start - birth_date).days // 7 + 1
            week_numbers[child['name']] = weeks_since_birth
            logger.info(f"👶 {child['name']}: Week {weeks_since_birth}")
            
        return week_numbers
        
    def get_photos_for_week(self, start_date: datetime, end_date: datetime) -> List[Dict]:
        """
        Retrieve photos from Google Photos for the specified week.
        
        Args:
            start_date: Start of the week
            end_date: End of the week
            
        Returns:
            List of photo items
        """
        logger.info("🔍 Fetching photos from Google Photos")
        
        # Convert dates to Google Photos filter format
        filters = {
            "dateFilter": {
                "ranges": [{
                    "startDate": {
                        "year": start_date.year,
                        "month": start_date.month,
                        "day": start_date.day
                    },
                    "endDate": {
                        "year": end_date.year,
                        "month": end_date.month,
                        "day": end_date.day
                    }
                }]
            }
        }
        
        photos = []
        next_page_token = None
        
        try:
            while True:
                body = {"filters": filters, "pageSize": 100}
                if next_page_token:
                    body["pageToken"] = next_page_token
                    
                response = self.photos_service.mediaItems().search(body=body).execute()
                
                if 'mediaItems' in response:
                    photos.extend(response['mediaItems'])
                    
                next_page_token = response.get('nextPageToken')
                if not next_page_token:
                    break
                    
            logger.info(f"📷 Found {len(photos)} photos for the week")
            return photos
            
        except HttpError as error:
            logger.error(f"❌ Error fetching photos: {error}")
            raise
            
    def create_album(self, album_title: str, photos: List[Dict]) -> Optional[Dict]:
        """
        Create a Google Photos album with the specified photos.
        
        Args:
            album_title: Title for the album
            photos: List of photo items to add
            
        Returns:
            Album information or None if creation failed
        """
        if not photos:
            logger.warning(f"⚠️  No photos to add to album: {album_title}")
            return None
            
        logger.info(f"📁 Creating album: {album_title}")
        
        try:
            # Create the album
            album_body = {"album": {"title": album_title}}
            album_response = self.photos_service.albums().create(body=album_body).execute()
            album_id = album_response['id']
            
            # Add photos to the album (max 50 at a time)
            for i in range(0, len(photos), 50):
                batch = photos[i:i+50]
                media_items = [{"mediaItemId": photo['id']} for photo in batch]
                
                add_body = {
                    "mediaItemIds": [item["mediaItemId"] for item in media_items]
                }
                
                self.photos_service.albums().batchAddMediaItems(
                    albumId=album_id,
                    body=add_body
                ).execute()
                
            logger.info(f"✅ Album created with {len(photos)} photos")
            return album_response
            
        except HttpError as error:
            logger.error(f"❌ Error creating album: {error}")
            return None
            
    def create_shareable_link(self, album_id: str) -> Optional[str]:
        """
        Create a shareable link for the album.
        
        Args:
            album_id: ID of the album
            
        Returns:
            Shareable URL or None if creation failed
        """
        logger.info("🔗 Creating shareable link")
        
        try:
            share_response = self.photos_service.albums().share(
                albumId=album_id,
                body={"sharedAlbumOptions": {"isCollaborative": False, "isCommentable": True}}
            ).execute()
            
            share_url = share_response['shareInfo']['shareableUrl']
            logger.info("✅ Shareable link created")
            return share_url
            
        except HttpError as error:
            logger.error(f"❌ Error creating shareable link: {error}")
            return None
            
    def send_email_notification(self, album_links: Dict[str, str], week_dates: Tuple[datetime, datetime]) -> bool:
        """
        Send email notification with album links.
        
        Args:
            album_links: Dictionary of child names to album URLs
            week_dates: Tuple of (start_date, end_date) for the week
            
        Returns:
            True if email sent successfully, False otherwise
        """
        logger.info("📧 Sending email notifications")
        
        start_date, end_date = week_dates
        week_period = f"{start_date.strftime('%B %d')} - {end_date.strftime('%B %d, %Y')}"
        
        # Build email content
        subject = f"Weekly Photo Albums - {week_period}"
        
        body = f"""
        <html>
            <body>
                <h2>Weekly Photo Albums 📷</h2>
                <p>Here are this week's photo albums ({week_period}):</p>
                <ul>
        """
        
        for child_name, url in album_links.items():
            if url:
                body += f'<li><strong>{child_name}:</strong> <a href="{url}">View Album</a></li>'
            else:
                body += f'<li><strong>{child_name}:</strong> No photos this week</li>'
                
        body += """
                </ul>
                <p>Enjoy the memories!</p>
            </body>
        </html>
        """
        
        try:
            # Create the email message
            message = {
                'raw': self._create_message(
                    self.config['email']['recipients'],
                    subject,
                    body
                )
            }
            
            # Send the email
            self.gmail_service.users().messages().send(
                userId='me',
                body=message
            ).execute()
            
            logger.info("✅ Email notifications sent successfully")
            return True
            
        except HttpError as error:
            logger.error(f"❌ Error sending email: {error}")
            return False
            
    def _create_message(self, recipients: List[str], subject: str, body: str) -> str:
        """
        Create email message in base64 format.
        
        Args:
            recipients: List of recipient email addresses
            subject: Email subject
            body: HTML email body
            
        Returns:
            Base64 encoded message
        """
        import base64
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        
        message = MIMEMultipart()
        message['to'] = ', '.join(recipients)
        message['subject'] = subject
        
        msg = MIMEText(body, 'html')
        message.attach(msg)
        
        return base64.urlsafe_b64encode(message.as_bytes()).decode()
        
    def run(self, dry_run: bool = False) -> None:
        """
        Execute the weekly photo automation.
        
        Args:
            dry_run: If True, don't create albums or send emails
        """
        logger.info("🚀 Starting weekly photo automation")
        
        try:
            # Authenticate
            self.authenticate_google_services()
            
            # Calculate week range
            start_date, end_date = self.calculate_week_range()
            
            # Calculate week numbers for children
            week_numbers = self.calculate_week_numbers(start_date)
            
            # Get photos for the week
            photos = self.get_photos_for_week(start_date, end_date)
            
            if not photos:
                logger.warning("⚠️  No photos found for this week")
                return
                
            album_links = {}
            
            # Create albums for each child
            for child in self.config['children']:
                child_name = child['name']
                week_num = week_numbers[child_name]
                album_title = f"{child_name}'s {week_num} week"
                
                if dry_run:
                    logger.info(f"🔍 [DRY RUN] Would create album: {album_title}")
                    album_links[album_title] = "https://example.com/dry-run"
                else:
                    album = self.create_album(album_title, photos)
                    if album:
                        share_url = self.create_shareable_link(album['id'])
                        album_links[album_title] = share_url
                    else:
                        album_links[album_title] = None
                        
            # Send email notification
            if not dry_run and album_links:
                self.send_email_notification(album_links, (start_date, end_date))
                
            logger.info("✅ Weekly photo automation completed successfully")
            
        except Exception as e:
            logger.error(f"❌ Automation failed: {e}")
            raise


def main():
    """Main execution function."""
    parser = argparse.ArgumentParser(
        description="Automate weekly photo album creation and sharing"
    )
    parser.add_argument(
        '--config',
        type=str,
        default='config_weekly_photo.json',
        help='Path to configuration file (default: config_weekly_photo.json)'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Run without creating albums or sending emails'
    )
    parser.add_argument(
        '--debug',
        action='store_true',
        help='Enable debug logging'
    )
    
    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        
    try:
        automation = WeeklyPhotoAutomation(args.config)
        automation.run(dry_run=args.dry_run)
        logger.info("✅ Script completed successfully")
        
    except Exception as e:
        logger.error(f"❌ Script failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()