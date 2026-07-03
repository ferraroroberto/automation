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
    log.info("📝 Child Information")
    log.info("-" * 40)

    while True:
        log.info("Child #%d", len(children) + 1)

        name = input("Enter child's name (or press Enter to finish): ").strip()
        if not name:
            if not children:
                log.warning("❌ You must add at least one child!")
                continue
            break

        while True:
            birth_date = input(f"Enter {name}'s birth date (YYYY-MM-DD): ").strip()
            if validate_date(birth_date):
                break
            log.warning("❌ Invalid date format. Please use YYYY-MM-DD")

        children.append({
            "name": name,
            "birth_date": birth_date
        })

        log.info("✅ Added %s (born %s)", name, birth_date)

    return children


def get_email_recipients() -> List[str]:
    """
    Interactively get email recipients from user.
    
    Returns:
        List of email addresses
    """
    recipients = []
    log.info("📧 Email Recipients")
    log.info("-" * 40)
    log.info("Enter email addresses to receive weekly album links")

    while True:
        email = input(f"Email #{len(recipients) + 1} (or press Enter to finish): ").strip()
        if not email:
            if not recipients:
                log.warning("❌ You must add at least one recipient!")
                continue
            break

        if validate_email(email):
            recipients.append(email)
            log.info("✅ Added %s", email)
        else:
            log.warning("❌ Invalid email format. Please try again.")

    return recipients


def check_credentials_file() -> bool:
    """
    Check if credentials.json exists.
    
    Returns:
        True if credentials file exists
    """
    creds_path = Path("credentials.json")
    if creds_path.exists():
        log.info("✅ Found credentials.json")
        return True
    else:
        log.warning("⚠️  credentials.json not found!")
        log.info("To get your credentials file:")
        log.info("1. Go to https://console.cloud.google.com")
        log.info("2. Create or select a project")
        log.info("3. Enable Google Photos Library API and Gmail API")
        log.info("4. Go to APIs & Services > Credentials")
        log.info("5. Create OAuth 2.0 Client ID (Desktop application)")
        log.info("6. Download the credentials and save as 'credentials.json'")
        log.info("Place credentials.json in this directory and run setup again.")
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
            log.info("✅ Backup created: %s", backup_path)

    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    log.info("✅ Configuration saved to config.json")


def main():
    """Main setup function."""
    log.info("=" * 50)
    log.info("   Weekly Photo Album Automation Setup")
    log.info("=" * 50)

    # Check for credentials file
    if not check_credentials_file():
        log.error("❌ Setup cannot continue without credentials.json")
        input("\nPress Enter to exit...")
        sys.exit(1)

    log.info("Let's configure your automation settings...")

    # Get children information
    children = get_children_info()

    # Get email recipients
    recipients = get_email_recipients()

    # Get timezone
    log.info("🌍 Timezone Configuration")
    log.info("-" * 40)
    log.info("Common timezones:")
    log.info("  - US/Eastern, US/Central, US/Pacific")
    log.info("  - Europe/London, Europe/Paris, Europe/Madrid")
    log.info("  - Asia/Tokyo, Asia/Shanghai, Australia/Sydney")
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
    log.info("=" * 50)
    log.info("   Configuration Summary")
    log.info("=" * 50)
    log.info("👶 Children (%d):", len(children))
    for child in children:
        birth = datetime.strptime(child['birth_date'], '%Y-%m-%d')
        weeks_old = (datetime.now() - birth).days // 7
        log.info("  - %s (born %s, currently week %d)", child['name'], child['birth_date'], weeks_old)

    log.info("📧 Recipients (%d):", len(recipients))
    for email in recipients:
        log.info("  - %s", email)

    log.info("🌍 Timezone: %s", timezone)

    # Confirm and save
    confirm = input("\n💾 Save this configuration? (y/n): ")
    if confirm.lower() == 'y':
        create_config_file(config)

        log.info("=" * 50)
        log.info("   Setup Complete!")
        log.info("=" * 50)
        log.info("✅ Your automation is configured and ready to use!")
        log.info("Next steps:")
        log.info("1. Run a test: python weekly_photo_automation.py --dry-run")
        log.info("2. Authenticate with Google (browser will open)")
        log.info("3. Run the automation: python weekly_photo_automation.py")
        log.info("For automated scheduling, see README.md")
    else:
        log.warning("❌ Setup cancelled. No changes were made.")

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    main()