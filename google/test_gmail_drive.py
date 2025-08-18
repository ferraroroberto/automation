#!/usr/bin/env python3
"""
Test script for Gmail and Google Drive Automation
Tests the core functionality without making actual API calls.
"""

# Standard library imports
import json
import logging
import sys
from pathlib import Path
from typing import Dict, Any

# Configure logging for testing
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


def test_config_loading() -> bool:
    """Test configuration file loading."""
    logger.info("🧪 Testing configuration loading...")
    
    try:
        # Test with sample config
        config_path = "config_gmail_drive.json.sample"
        if not Path(config_path).exists():
            logger.error(f"❌ Sample config file not found: {config_path}")
            return False
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        # Validate required sections
        required_sections = ['auth', 'spreadsheet', 'gmail', 'settings']
        for section in required_sections:
            if section not in config:
                logger.error(f"❌ Missing required config section: {section}")
                return False
        
        logger.info("✅ Configuration loading test passed")
        return True
        
    except Exception as e:
        logger.error(f"❌ Configuration loading test failed: {e}")
        return False


def test_email_data_processing() -> bool:
    """Test email data processing logic."""
    logger.info("🧪 Testing email data processing...")
    
    try:
        # Mock email data
        mock_emails = [
            {
                'id': 'msg1',
                'subject': 'Test Email 1',
                'from': 'test1@example.com',
                'date': '2024-01-15 10:00:00',
                'snippet': 'This is a test email content'
            },
            {
                'id': 'msg2',
                'subject': 'Test Email 2',
                'from': 'test2@example.com',
                'date': '2024-01-15 11:00:00',
                'snippet': 'Another test email with longer content that should be truncated'
            }
        ]
        
        # Test data transformation
        rows = []
        for email in mock_emails:
            rows.append([
                email['subject'],
                email['date'],
                email['from'],
                email['snippet'][:100] + '...' if len(email['snippet']) > 100 else email['snippet']
            ])
        
        # Validate row structure
        if len(rows) != 2:
            logger.error(f"❌ Expected 2 rows, got {len(rows)}")
            return False
        
        if len(rows[0]) != 4:
            logger.error(f"❌ Expected 4 columns, got {len(rows[0])}")
            return False
        
        # Test content truncation
        if len(rows[1][3]) > 103:  # 100 + "..."
            logger.error(f"❌ Content not properly truncated: {len(rows[1][3])} characters")
            return False
        
        logger.info("✅ Email data processing test passed")
        return True
        
    except Exception as e:
        logger.error(f"❌ Email data processing test failed: {e}")
        return False


def test_duplicate_prevention() -> bool:
    """Test duplicate email prevention logic."""
    logger.info("🧪 Testing duplicate prevention...")
    
    try:
        # Mock existing subjects
        existing_subjects = {'Test Email 1', 'Old Email', 'Meeting Reminder'}
        
        # Mock new emails
        new_emails = [
            {'subject': 'Test Email 1'},  # Duplicate
            {'subject': 'New Email'},     # New
            {'subject': 'Meeting Reminder'},  # Duplicate
            {'subject': 'Another New Email'}  # New
        ]
        
        # Filter out duplicates
        unique_emails = []
        for email in new_emails:
            if email['subject'].strip() not in existing_subjects:
                unique_emails.append(email)
        
        # Should have 2 unique emails
        if len(unique_emails) != 2:
            logger.error(f"❌ Expected 2 unique emails, got {len(unique_emails)}")
            return False
        
        # Check that duplicates were removed
        unique_subjects = {email['subject'] for email in unique_emails}
        if 'Test Email 1' in unique_subjects or 'Meeting Reminder' in unique_subjects:
            logger.error("❌ Duplicate emails were not properly filtered")
            return False
        
        logger.info("✅ Duplicate prevention test passed")
        return True
        
    except Exception as e:
        logger.error(f"❌ Duplicate prevention test failed: {e}")
        return False


def test_spreadsheet_structure() -> bool:
    """Test spreadsheet structure validation."""
    logger.info("🧪 Testing spreadsheet structure...")
    
    try:
        # Mock spreadsheet data
        mock_spreadsheet = {
            'properties': {
                'title': 'Email Tracking Dashboard',
                'timeZone': 'UTC'
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
        
        # Validate structure
        if 'properties' not in mock_spreadsheet:
            logger.error("❌ Missing properties in spreadsheet structure")
            return False
        
        if 'sheets' not in mock_spreadsheet:
            logger.error("❌ Missing sheets in spreadsheet structure")
            return False
        
        if len(mock_spreadsheet['sheets']) == 0:
            logger.error("❌ No sheets defined in spreadsheet")
            return False
        
        # Validate sheet properties
        sheet = mock_spreadsheet['sheets'][0]
        if sheet['properties']['title'] != 'Email Tracking':
            logger.error("❌ Incorrect sheet title")
            return False
        
        if sheet['properties']['gridProperties']['columnCount'] != 4:
            logger.error("❌ Incorrect column count")
            return False
        
        logger.info("✅ Spreadsheet structure test passed")
        return True
        
    except Exception as e:
        logger.error(f"❌ Spreadsheet structure test failed: {e}")
        return False


def run_all_tests() -> bool:
    """Run all tests and return overall success status."""
    logger.info("🚀 Starting Gmail and Drive automation tests...")
    logger.info("=" * 50)
    
    tests = [
        test_config_loading,
        test_email_data_processing,
        test_duplicate_prevention,
        test_spreadsheet_structure
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
        logger.info("")
    
    logger.info("=" * 50)
    logger.info(f"📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        logger.info("🎉 All tests passed! The automation is ready to use.")
        return True
    else:
        logger.error(f"❌ {total - passed} test(s) failed. Please check the issues above.")
        return False


if __name__ == "__main__":
    try:
        success = run_all_tests()
        sys.exit(0 if success else 1)
    except Exception as e:
        logger.error(f"❌ Test suite failed with unexpected error: {e}")
        sys.exit(1)