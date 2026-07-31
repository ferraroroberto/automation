#!/usr/bin/env python3
"""
Gmail and Google Drive Automation
Automatically creates and maintains a Google Sheets file with email data.
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
from googleapiclient.errors import HttpError
import pytz

import _auth

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class GmailDriveAutomation:
    """Manages Gmail monitoring and Google Drive spreadsheet automation."""
    
    def __init__(self, config_path: str = "config.json"):
        """
        Initialize the automation with configuration.
        
        Args:
            config_path: Path to the configuration file
        """
        self.config = self.load_config(config_path)
        self.drive_service = None
        self.gmail_service = None
        self.sheets_service = None
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
        """Authenticate with Google Drive, Gmail, and Sheets APIs."""
        logger.info("🔐 Authenticating with Google services")

        SCOPES = [
            # Drive API - Core scopes
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.file',

            # Gmail API - Core scopes
            'https://www.googleapis.com/auth/gmail.readonly',
            'https://www.googleapis.com/auth/gmail.metadata',

            # Sheets API - Core scopes
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/spreadsheets.readonly'
        ]

        creds = _auth.authenticate(self.config, Path(__file__).parent, SCOPES)
        self.credentials = creds

        # Build Google Drive API service
        try:
            self.drive_service = build('drive', 'v3', credentials=creds)
            logger.info("✅ Google Drive API service built successfully")
        except Exception as e:
            logger.error(f"❌ Failed to build Drive API service: {e}")
            raise Exception(f"Could not build Drive API service: {e}")
        
        # Build Gmail API service
        try:
            self.gmail_service = build('gmail', 'v1', credentials=creds)
            logger.info("✅ Gmail API service built successfully")
        except Exception as e:
            logger.error(f"❌ Failed to build Gmail API service: {e}")
            raise Exception(f"Could not build Gmail API service: {e}")
            
        # Build Google Sheets API service
        try:
            self.sheets_service = build('sheets', 'v4', credentials=creds)
            logger.info("✅ Google Sheets API service built successfully")
        except Exception as e:
            logger.error(f"❌ Failed to build Sheets API service: {e}")
            raise Exception(f"Could not build Sheets API service: {e}")
            
        logger.info("✅ Authentication successful")
        logger.info(f"🔑 Granted scopes: {creds.scopes}")

    def find_or_create_spreadsheet(self) -> str:
        """
        Find existing email tracking spreadsheet or create a new one.
        
        Returns:
            Spreadsheet ID
        """
        logger.info("📊 Looking for existing email tracking spreadsheet")
        
        # Search for existing spreadsheet by name
        query = f"name='{self.config['spreadsheet']['name']}' and mimeType='application/vnd.google-apps.spreadsheet'"
        
        try:
            results = self.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name, createdTime)'
            ).execute()
            
            files = results.get('files', [])
            
            if files:
                # Use the most recently created one
                latest_file = max(files, key=lambda x: x['createdTime'])
                logger.info(f"✅ Found existing spreadsheet: {latest_file['name']} (ID: {latest_file['id']})")
                return latest_file['id']
            else:
                logger.info("📝 No existing spreadsheet found, creating new one")
                return self.create_new_spreadsheet()
                
        except HttpError as error:
            logger.error(f"❌ Error searching for spreadsheet: {error}")
            raise

    def create_new_spreadsheet(self) -> str:
        """
        Create a new Google Sheets spreadsheet for email tracking.
        
        Returns:
            Spreadsheet ID
        """
        logger.info("📝 Creating new email tracking spreadsheet")
        
        try:
            # Create the spreadsheet
            spreadsheet_body = {
                'properties': {
                    'title': self.config['spreadsheet']['name'],
                    'timeZone': self.config['settings']['timezone']
                },
                'sheets': [
                    {
                        'properties': {
                            'title': 'Email Tracking',
                            'gridProperties': {
                                'rowCount': 1000,
                                'columnCount': 4
                            }
                        }
                    }
                ]
            }
            
            spreadsheet = self.sheets_service.spreadsheets().create(
                body=spreadsheet_body
            ).execute()
            
            spreadsheet_id = spreadsheet['spreadsheetId']
            logger.info(f"✅ Created new spreadsheet: {spreadsheet_id}")
            
            # Set up headers
            self.setup_spreadsheet_headers(spreadsheet_id)
            
            return spreadsheet_id
            
        except HttpError as error:
            logger.error(f"❌ Error creating spreadsheet: {error}")
            raise

    def setup_spreadsheet_headers(self, spreadsheet_id: str) -> None:
        """
        Set up the spreadsheet headers and formatting.
        
        Args:
            spreadsheet_id: ID of the spreadsheet to format
        """
        logger.info("📋 Setting up spreadsheet headers")
        
        try:
            # Set headers
            headers = [
                ['Subject', 'Date', 'From', 'Content Preview']
            ]
            
            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='A1:D1',
                valueInputOption='RAW',
                body={'values': headers}
            ).execute()
            
            # Format headers (make them bold)
            requests = [
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': 0,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                            'startColumnIndex': 0,
                            'endColumnIndex': 4
                        },
                        'cell': {
                            'userEnteredFormat': {
                                'textFormat': {'bold': True},
                                'backgroundColor': {
                                    'red': 0.9,
                                    'green': 0.9,
                                    'blue': 0.9
                                }
                            }
                        },
                        'fields': 'userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor'
                    }
                },
                {
                    'autoResizeDimensions': {
                        'dimensions': {
                            'sheetId': 0,
                            'dimension': 'COLUMNS',
                            'startIndex': 0,
                            'endIndex': 4
                        }
                    }
                }
            ]
            
            self.sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()
            
            logger.info("✅ Spreadsheet headers set up successfully")
            
        except HttpError as error:
            logger.error(f"❌ Error setting up headers: {error}")
            raise

    def search_emails(self, query: str = None) -> List[Dict[str, Any]]:
        """
        Search Gmail for emails matching criteria.
        
        Args:
            query: Gmail search query (optional)
            
        Returns:
            List of email data dictionaries
        """
        logger.info("🔍 Searching Gmail for emails")
        
        if not query:
            query = self.config['gmail']['search_query']
        
        try:
            # Search for emails
            results = self.gmail_service.users().messages().list(
                userId='me',
                q=query,
                maxResults=self.config['gmail']['max_results']
            ).execute()
            
            messages = results.get('messages', [])
            logger.info(f"📧 Found {len(messages)} emails matching criteria")
            
            if not messages:
                return []
            
            # Get detailed information for each email
            email_data = []
            for message in messages:
                try:
                    msg = self.gmail_service.users().messages().get(
                        userId='me',
                        id=message['id'],
                        format='metadata',
                        metadataHeaders=['Subject', 'From', 'Date']
                    ).execute()
                    
                    headers = msg['payload']['headers']
                    subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
                    sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown Sender')
                    date = next((h['value'] for h in headers if h['name'] == 'Date'), 'Unknown Date')
                    
                    # Parse date
                    try:
                        parsed_date = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
                        formatted_date = parsed_date.strftime('%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        formatted_date = date
                    
                    email_data.append({
                        'id': message['id'],
                        'subject': subject,
                        'from': sender,
                        'date': formatted_date,
                        'snippet': msg.get('snippet', 'No content')
                    })
                    
                except HttpError as e:
                    logger.warning(f"⚠️ Could not fetch email {message['id']}: {e}")
                    continue
            
            return email_data
            
        except HttpError as error:
            logger.error(f"❌ Error searching Gmail: {error}")
            raise

    def check_existing_records(self, spreadsheet_id: str, email_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Check which emails are not already in the spreadsheet.
        
        Args:
            spreadsheet_id: ID of the spreadsheet
            email_data: List of email data to check
            
        Returns:
            List of new emails not in spreadsheet
        """
        logger.info("🔍 Checking for existing records in spreadsheet")
        
        try:
            # Get existing data from spreadsheet
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range='A:D'
            ).execute()
            
            existing_values = result.get('values', [])
            
            # Skip header row and create set of existing subjects for quick lookup
            existing_subjects = set()
            if len(existing_values) > 1:
                for row in existing_values[1:]:  # Skip header
                    if row and len(row) > 0:
                        existing_subjects.add(row[0].strip())  # Subject is in column A
            
            # Filter out emails that already exist
            new_emails = []
            for email in email_data:
                if email['subject'].strip() not in existing_subjects:
                    new_emails.append(email)
                else:
                    logger.debug(f"📧 Email already exists: {email['subject'][:50]}...")
            
            logger.info(f"✅ Found {len(new_emails)} new emails to add")
            return new_emails
            
        except HttpError as error:
            logger.error(f"❌ Error checking existing records: {error}")
            raise

    def add_emails_to_spreadsheet(self, spreadsheet_id: str, emails: List[Dict[str, Any]]) -> None:
        """
        Add new emails to the spreadsheet.
        
        Args:
            spreadsheet_id: ID of the spreadsheet
            emails: List of email data to add
        """
        if not emails:
            logger.info("📧 No new emails to add")
            return
        
        logger.info(f"📝 Adding {len(emails)} new emails to spreadsheet")
        
        try:
            # Prepare data for spreadsheet
            rows = []
            for email in emails:
                rows.append([
                    email['subject'],
                    email['date'],
                    email['from'],
                    email['snippet'][:100] + '...' if len(email['snippet']) > 100 else email['snippet']
                ])
            
            # Get current row count to append at the end
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range='A:A'
            ).execute()
            
            current_rows = len(result.get('values', []))
            start_row = current_rows + 1
            
            # Add new rows
            range_name = f'A{start_row}:D{start_row + len(rows) - 1}'
            
            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': rows}
            ).execute()
            
            logger.info(f"✅ Added {len(emails)} emails to spreadsheet starting at row {start_row}")
            
        except HttpError as error:
            logger.error(f"❌ Error adding emails to spreadsheet: {error}")
            raise

    def run_automation(self) -> None:
        """Run the complete email tracking automation."""
        logger.info("🚀 Starting Gmail and Drive automation")
        
        try:
            # Authenticate with Google services
            self.authenticate_google_services()
            
            # Find or create spreadsheet
            spreadsheet_id = self.find_or_create_spreadsheet()
            
            # Search for emails
            email_data = self.search_emails()
            
            if not email_data:
                logger.info("📧 No emails found matching criteria")
                return
            
            # Check for existing records
            new_emails = self.check_existing_records(spreadsheet_id, email_data)
            
            # Add new emails to spreadsheet
            self.add_emails_to_spreadsheet(spreadsheet_id, new_emails)
            
            logger.info("✅ Automation completed successfully")
            
        except Exception as e:
            logger.error(f"❌ Automation failed: {e}")
            raise


def main(config_path: str = "config.json") -> None:
    """
    Main function to run the automation.

    Args:
        config_path: Path to the configuration file
    """
    try:
        automation = GmailDriveAutomation(config_path)
        automation.run_automation()
        logger.info("✅ Script completed successfully")
    except Exception as e:
        logger.error("❌ Script failed: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    try:
        parser = argparse.ArgumentParser(description='Gmail and Google Drive Automation')
        parser.add_argument('--config', default='config.json', help='Path to configuration file')
        parser.add_argument('--debug', action='store_true', help='Enable debug logging')

        args = parser.parse_args()

        if args.debug:
            logging.getLogger().setLevel(logging.DEBUG)
            logger.setLevel(logging.DEBUG)

        main(args.config)

    except Exception as e:
        logger.error("❌ Script failed: %s", e)
        sys.exit(1)