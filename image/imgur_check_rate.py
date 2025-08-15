import requests
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get access token from environment variable
ACCESS_TOKEN = os.getenv("IMGUR_ACCESS_TOKEN")

def check_imgur_rate_limit():
    """
    Fetches and displays the current Imgur API rate limits for an authenticated user.
    """
    if not ACCESS_TOKEN:
        print("❌ IMGUR_ACCESS_TOKEN not found in environment variables")
        return
        
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    response = requests.get("https://api.imgur.com/3/credits", headers=headers)

    if response.status_code == 200:
        rate_info = response.json()["data"]
        print("\n📊 **Imgur API Rate Limits**")
        print(f"💡 Remaining Requests: {rate_info['ClientRemaining']} / {rate_info['ClientLimit']} (per hour)")
        print(f"📸 Remaining Uploads: {rate_info['UserRemaining']} / {rate_info['UserLimit']} (per day)")
        print(f"🔄 Resets in: {rate_info['UserReset']} seconds (~{rate_info['UserReset'] // 60} minutes)\n")
    else:
        print(f"❌ Failed to fetch rate limits: {response.status_code} - {response.text}")

if __name__ == "__main__":
    check_imgur_rate_limit()
