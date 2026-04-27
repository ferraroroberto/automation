#!/usr/bin/env python3
"""
Setup helper for Weekly Photo Album Automation.
Guides users through initial configuration.
"""

import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

log = logging.getLogger(__name__)


def validate_email(email: str) -> bool:
    """
    Basic email validation.
    
    Args:
        email: Email address to validate
        
    Returns:
        True if email appears valid
    """
    return '@' in email and '.' in email.split('@')[1]


def validate_date(date_str: str) -> bool:
    """
    Validate date format.
    
    Args:
        date_str: Date string to validate
        
    Returns:
        True if date is valid YYYY-MM-DD format
    """
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
        return True
    except ValueError:
        return False


def get_children_info() -> List[Dict[str, str]]:
    """
    Interactively get children information from user.
    
    Returns:
        List of children dictionaries
    """
    children = []
    print("\n📝 Child Information")
    print("-" * 40)
    
    while True:
        print(f"\nChild #{len(children) + 1}")
        
        name = input("Enter child's name (or press Enter to finish): ").strip()
        if not name:
            if not children:
                print("❌ You must add at least one child!")
                continue
            break
            
        while True:
            birth_date = input(f"Enter {name}'s birth date (YYYY-MM-DD): ").strip()
            if validate_date(birth_date):
                break
            print("❌ Invalid date format. Please use YYYY-MM-DD")
            
        children.append({
            "name": name,
            "birth_date": birth_date
        })
        
        print(f"✅ Added {name} (born {birth_date})")
        
    return children


def get_email_recipients() -> List[str]:
    """
    Interactively get email recipients from user.
    
    Returns:
        List of email addresses
    """
    recipients = []
    print("\n📧 Email Recipients")
    print("-" * 40)
    print("Enter email addresses to receive weekly album links")
    
    while True:
        email = input(f"Email #{len(recipients) + 1} (or press Enter to finish): ").strip()
        if not email:
            if not recipients:
                print("❌ You must add at least one recipient!")
                continue
            break
            
        if validate_email(email):
            recipients.append(email)
            print(f"✅ Added {email}")
        else:
            print("❌ Invalid email format. Please try again.")
            
    return recipients


def check_credentials_file() -> bool:
    """
    Check if credentials.json exists.
    
    Returns:
        True if credentials file exists
    """
    creds_path = Path("credentials.json")
    if creds_path.exists():
        print("✅ Found credentials.json")
        return True
    else:
        print("\n⚠️  credentials.json not found!")
        print("\nTo get your credentials file:")
        print("1. Go to https://console.cloud.google.com")
        print("2. Create or select a project")
        print("3. Enable Google Photos Library API and Gmail API")
        print("4. Go to APIs & Services > Credentials")
        print("5. Create OAuth 2.0 Client ID (Desktop application)")
        print("6. Download the credentials and save as 'credentials.json'")
        print("\nPlace credentials.json in this directory and run setup again.")
        return False


def create_config_file(config: Dict[str, Any]) -> None:
    """
    Create the configuration file.
    
    Args:
        config: Configuration dictionary
    """
    config_path = Path("config.json")
    
    if config_path.exists():
        backup = input("\n⚠️  config.json already exists. Create backup? (y/n): ")
        if backup.lower() == 'y':
            backup_path = Path(f"config.backup.{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            config_path.rename(backup_path)
            print(f"✅ Backup created: {backup_path}")
            
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
        
    print(f"\n✅ Configuration saved to config.json")


def main():
    """Main setup function."""
    print("=" * 50)
    print("   Weekly Photo Album Automation Setup")
    print("=" * 50)
    
    # Check for credentials file
    if not check_credentials_file():
        print("\n❌ Setup cannot continue without credentials.json")
        input("\nPress Enter to exit...")
        sys.exit(1)
        
    print("\nLet's configure your automation settings...")
    
    # Get children information
    children = get_children_info()
    
    # Get email recipients
    recipients = get_email_recipients()
    
    # Get timezone
    print("\n🌍 Timezone Configuration")
    print("-" * 40)
    print("Common timezones:")
    print("  - US/Eastern, US/Central, US/Pacific")
    print("  - Europe/London, Europe/Paris, Europe/Madrid")
    print("  - Asia/Tokyo, Asia/Shanghai, Australia/Sydney")
    timezone = input("\nEnter your timezone (default: Europe/Madrid): ").strip()
    if not timezone:
        timezone = "Europe/Madrid"
        
    # Build configuration
    config = {
        "auth": {
            "credentials_file": "credentials.json",
            "token_file": "token.json"
        },
        "children": children,
        "email": {
            "recipients": recipients
        },
        "settings": {
            "timezone": timezone,
            "max_photos_per_album": 500,
            "retry_attempts": 3,
            "retry_delay_seconds": 5
        }
    }
    
    # Display summary
    print("\n" + "=" * 50)
    print("   Configuration Summary")
    print("=" * 50)
    print(f"\n👶 Children ({len(children)}):")
    for child in children:
        birth = datetime.strptime(child['birth_date'], '%Y-%m-%d')
        weeks_old = (datetime.now() - birth).days // 7
        print(f"  - {child['name']} (born {child['birth_date']}, currently week {weeks_old})")
        
    print(f"\n📧 Recipients ({len(recipients)}):")
    for email in recipients:
        print(f"  - {email}")
        
    print(f"\n🌍 Timezone: {timezone}")
    
    # Confirm and save
    confirm = input("\n💾 Save this configuration? (y/n): ")
    if confirm.lower() == 'y':
        create_config_file(config)
        
        print("\n" + "=" * 50)
        print("   Setup Complete!")
        print("=" * 50)
        print("\n✅ Your automation is configured and ready to use!")
        print("\nNext steps:")
        print("1. Run a test: python weekly_photo_automation.py --dry-run")
        print("2. Authenticate with Google (browser will open)")
        print("3. Run the automation: python weekly_photo_automation.py")
        print("\nFor automated scheduling, see README.md")
    else:
        print("\n❌ Setup cancelled. No changes were made.")
        
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    main()