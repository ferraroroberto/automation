import logging
import os

import requests
from dotenv import load_dotenv

log = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Get access token from environment variable
ACCESS_TOKEN = os.getenv("IMGUR_ACCESS_TOKEN")

def check_imgur_rate_limit():
    """
    Fetches and displays the current Imgur API rate limits for an authenticated user.
    """
    if not ACCESS_TOKEN:
        log.error("❌ IMGUR_ACCESS_TOKEN not found in environment variables")
        return

    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    response = requests.get("https://api.imgur.com/3/credits", headers=headers)

    if response.status_code == 200:
        rate_info = response.json()["data"]
        log.info("📊 Imgur API Rate Limits")
        log.info("💡 Remaining Requests: %s / %s (per hour)", rate_info['ClientRemaining'], rate_info['ClientLimit'])
        log.info("📸 Remaining Uploads: %s / %s (per day)", rate_info['UserRemaining'], rate_info['UserLimit'])
        log.info("🔄 Resets in: %s seconds (~%d minutes)", rate_info['UserReset'], rate_info['UserReset'] // 60)
    else:
        log.error("❌ Failed to fetch rate limits: %s - %s", response.status_code, response.text)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    check_imgur_rate_limit()
