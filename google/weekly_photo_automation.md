# Weekly Photo Album Automation

Automatically creates weekly Google Photos albums for your children and shares them with family members via email.

## 🚀 Overview

This automation tool:
- Collects photos from the previous complete week (Saturday to Friday)
- Creates individual albums for each child with sequential week numbers since birth
- Generates shareable links for the albums
- Sends email notifications to specified recipients

The script intelligently handles week boundaries - if run on any day, it will always process the most recent complete Saturday-Friday week.

## ⚠️ Important: Google Photos Sharing API Limitation

### The Issue
The Google Photos API has a restriction on the album sharing functionality (`albums.share()` method). Even with the correct OAuth scopes (`photoslibrary.sharing`), the API returns a 403 Forbidden error when attempting to programmatically share albums. This is a known limitation where Google requires app verification for sharing capabilities, which is typically only granted to published/verified applications, not personal automation scripts.

### What Works ✅
- Creating albums
- Adding photos to albums
- Searching and listing photos
- Sending email notifications with album links
- All Gmail API functions

### What Doesn't Work ❌
- Programmatically sharing albums via API (`albums.share()` returns 403)
- Generating public shareable links automatically

### The Solution
The automation works with a semi-manual approach:
1. **Automated**: Creates weekly albums with photos
2. **Automated**: Sends emails to recipients with album links
3. **Manual**: Recipients request access via the link, and you approve it in Google Photos (takes 30 seconds)

Alternatively, after albums are created, you can manually share them in the Google Photos app with your family group.

### Technical Details
- **Error**: `403 Forbidden` on `albums.share()` API call
- **Cause**: Google's API restrictions on sharing for unverified apps
- **Scopes affected**: `https://www.googleapis.com/auth/photoslibrary.sharing`
- **Not fixable by**: Adding more scopes, changing authentication, or modifying code

This is a Google-imposed limitation, not a bug in the code. The automation still saves significant time by handling album creation, organization, and email notifications automatically.

## 📋 Prerequisites

Before using this automation, you need:
1. A Google Cloud Project with enabled APIs
2. OAuth 2.0 credentials
3. Python 3.7 or higher

### Setting up Google Cloud APIs

1. **Create a Google Cloud Project**:
   - Go to [Google Cloud Console](https://console.cloud.google.com)
   - Create a new project or select an existing one

2. **Enable Required APIs**:
   - Enable the **Google Photos Library API**
   - Enable the **Gmail API**

3. **Create OAuth 2.0 Credentials**:
   - Go to APIs & Services > Credentials
   - Click "Create Credentials" > "OAuth client ID"
   - Choose "Desktop app" as the application type
   - Download the credentials as `credentials.json`

4. **Configure OAuth Consent Screen**:
   - Go to APIs & Services > OAuth consent screen
   - Choose "External" user type (unless you have a Google Workspace organization)
   - Fill in the required app information:
     - **App name**: `automation` (or your preferred name)
     - **User support email**: Your email address
     - **Developer contact information**: Your email address
   
   **⚠️ IMPORTANT: Audience Settings for Testing**
   - In the "Test users" section, add your email address (`roberto.ferraro@gmail.com`)
   - This allows you to test the app while it's in development mode
   - **Without adding test users, you'll get "Access blocked" errors**
   
   **Scopes to Add**:
   - `https://www.googleapis.com/auth/photoslibrary` (Google Photos access)
   - `https://www.googleapis.com/auth/gmail.send` (Gmail send permission)
   - `https://www.googleapis.com/auth/userinfo.email` (Basic profile info)
   
   **Testing vs Production**:
   - **Testing Mode**: App can only be used by added test users
   - **Production Mode**: Requires Google verification (can take weeks)
   - For personal automation, testing mode is sufficient

## 🔄 OAuth Access Solutions Comparison

When you encounter the "Access blocked" error, you have several options to resolve it. Here's a detailed comparison:

### Option 1: Add Yourself as a Test User (Recommended for Development)
**What it is**: Add your email to the "Test users" section in OAuth consent screen.

**Pros**:
- ✅ **Immediate solution** - works within minutes
- ✅ **No verification process** - bypasses Google's review
- ✅ **Perfect for personal automation** - no external users needed
- ✅ **Free** - no additional costs
- ✅ **Easy to manage** - you control who has access

**Cons**:
- ❌ **Limited to test users only** - can't share with others
- ❌ **Manual management** - must add each user individually
- ❌ **Not suitable for public apps** - can't distribute widely

**Best for**: Personal automation, development/testing, small family use

---

### Option 2: Complete Google Verification Process
**What it is**: Submit your app for Google's official verification to remove all restrictions.

**Pros**:
- ✅ **No user limits** - anyone can use your app
- ✅ **Professional appearance** - no "unverified app" warnings
- ✅ **Scalable** - can distribute to unlimited users
- ✅ **Trust indicator** - users know app is Google-approved

**Cons**:
- ❌ **Time-consuming** - can take 6-8 weeks
- ❌ **Complex requirements** - privacy policy, terms of service, etc.
- ❌ **May require business verification** - if app handles sensitive data
- ❌ **Ongoing compliance** - must maintain verification standards
- ❌ **Potential rejection** - Google may deny verification

**Best for**: Commercial apps, public distribution, professional services

---

### Option 3: Use Service Account Authentication
**What it is**: Replace OAuth with service account credentials for server-to-server communication.

**Pros**:
- ✅ **No user consent needed** - works without user interaction
- ✅ **No verification required** - bypasses OAuth entirely
- ✅ **More secure** - no user tokens to manage
- ✅ **Suitable for automation** - designed for background processes
- ✅ **Unlimited access** - no user limits

**Cons**:
- ❌ **Limited to your own data** - can't access other users' accounts
- ❌ **More complex setup** - requires service account creation
- ❌ **Different API calls** - may need code modifications
- ❌ **Not suitable for user apps** - only for your own automation

**Best for**: Personal automation, server-side processes, your own data only

---

### Option 4: Continue in Testing Mode with Multiple Users
**What it is**: Keep app in testing mode but add multiple family members as test users.

**Pros**:
- ✅ **Quick setup** - add users in minutes
- ✅ **No verification needed** - bypasses Google review
- ✅ **Family-friendly** - can share with close family
- ✅ **Easy to manage** - add/remove users as needed

**Cons**:
- ❌ **Limited to 100 test users** - Google's maximum limit
- ❌ **Manual management** - must add each user individually
- ❌ **Not scalable** - impractical for large user bases
- ❌ **Still shows warnings** - users see "unverified app" message

**Best for**: Family automation, small team use, limited distribution

---

## 🎯 Recommendation for Your Use Case

**For your weekly photo automation tool**, I recommend **Option 1 (Test User)** because:

1. **It's personal automation** - you don't need external users
2. **Immediate solution** - fixes your current error right now
3. **No ongoing costs** - completely free to maintain
4. **Simple management** - just add your email and you're done
5. **Perfect fit** - designed exactly for this type of personal tool

**If you later want to share with family**, you can easily add them as test users (Option 4) without going through the full verification process.

## 🔧 Installation

### Step 1: Clone and Setup Environment

**Windows PowerShell:**
```powershell
# Clone repository (or copy files to a folder)
mkdir weekly-photo-automation; cd weekly-photo-automation

# Create virtual environment
python -m venv .venv

# Activate virtual environment
.\.venv\Scripts\Activate.ps1

# Install requirements
pip install -r requirements.txt
```

**Unix/Linux/MacOS:**
```bash
# Clone repository (or copy files to a folder)
mkdir weekly-photo-automation && cd weekly-photo-automation

# Create virtual environment
python3 -m venv .venv

# Activate virtual environment
source .venv/bin/activate

# Install requirements
pip install -r requirements.txt
```

### Step 2: Configuration

1. **Copy your `credentials.json`** file (from Google Cloud Console) to the project directory

2. **Edit `config.json`** with your information:
   ```json
   {
     "children": [
       {
         "name": "Child Name 1",
         "birth_date": "2023-01-15"  // Format: YYYY-MM-DD
       },
       {
         "name": "Child Name 2",
         "birth_date": "2024-03-20"
       }
     ],
     "email": {
       "recipients": [
         "family1@example.com",
         "family2@example.com",
         "family3@example.com"
       ]
     }
   }
   ```

3. **First-time authentication**:
   ```bash
   python weekly_photo_automation.py --dry-run
   ```
   This will open a browser for Google authentication. Grant permissions for:
   - View and manage your Google Photos library
   - Send email on your behalf

## 📊 Usage

### Basic Usage
```bash
# Run the automation
python weekly_photo_automation.py

# Dry run (test without creating albums or sending emails)
python weekly_photo_automation.py --dry-run

# Enable debug logging
python weekly_photo_automation.py --debug

# Use custom configuration file
python weekly_photo_automation.py --config custom_config.json
```

### Scheduling Automation

**Windows Task Scheduler:**
```xml
<!-- Save as weekly_photos.xml and import to Task Scheduler -->
<Task>
  <Triggers>
    <CalendarTrigger>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Saturday />
        </DaysOfWeek>
      </ScheduleByWeek>
      <StartBoundary>2024-01-01T10:00:00</StartBoundary>
    </CalendarTrigger>
  </Triggers>
  <Actions>
    <Exec>
      <Command>C:\path\to\.venv\Scripts\python.exe</Command>
      <Arguments>C:\path\to\weekly_photo_automation.py</Arguments>
    </Exec>
  </Actions>
</Task>
```

**Linux/MacOS (cron):**
```bash
# Edit crontab
crontab -e

# Add this line to run every Saturday at 10 AM
0 10 * * 6 /path/to/.venv/bin/python /path/to/weekly_photo_automation.py
```

## 🔧 Configuration Options

### config.json Structure

```json
{
  "auth": {
    "credentials_file": "credentials.json",  // OAuth credentials from Google
    "token_file": "token.json"              // Stored auth token (auto-generated)
  },
  "children": [
    {
      "name": "Child Name",                 // Used in album titles
      "birth_date": "YYYY-MM-DD"           // For calculating week numbers
    }
  ],
  "email": {
    "recipients": [                        // Email addresses to notify
      "email1@example.com",
      "email2@example.com"
    ]
  },
  "settings": {
    "timezone": "Your/Timezone",           // Your local timezone
    "max_photos_per_album": 500,          // Maximum photos per album
    "retry_attempts": 3,                  // API retry attempts
    "retry_delay_seconds": 5              // Delay between retries
  }
}
```

## 📊 Output

### Album Naming Convention
Albums are named: `[Child Name]'s [Week Number] week`

Examples:
- Child Name 1's 52 week
- Child Name 2's 26 week

### Email Format
Recipients receive an HTML email containing:
- Week date range (Saturday - Friday)
- Direct links to each child's album
- Indication if no photos were found for the week

### Logging Output
```
2024-01-06 10:00:00 - INFO - 🚀 Starting weekly photo automation
2024-01-06 10:00:01 - INFO - 🔐 Authenticating with Google services
2024-01-06 10:00:02 - INFO - ✅ Authentication successful
2024-01-06 10:00:02 - INFO - 📅 Week range: 2023-12-30 to 2024-01-05
2024-01-06 10:00:02 - INFO - 👶 Child Name 1: Week 52
2024-01-06 10:00:02 - INFO - 👶 Child Name 2: Week 26
2024-01-06 10:00:03 - INFO - 📷 Found 127 photos for the week
2024-01-06 10:00:05 - INFO - 📁 Creating album: Child Name 1's 52 week
2024-01-06 10:00:07 - INFO - ✅ Album created with 127 photos
2024-01-06 10:00:08 - INFO - 📧 Sending email notifications
2024-01-06 10:00:09 - INFO - ✅ Script completed successfully
```

## 🔒 Security Considerations

- **Never commit** `credentials.json` or `token.json` to version control
- **Add to .gitignore**:
  ```
  credentials.json
  token.json
  config.json
  *.log
  .env
  ```
- **Store sensitive files** outside the code directory in production
- **Use environment variables** for production deployments

## 🛠️ Troubleshooting

### Common Issues

**Authentication Errors**:
- Delete `token.json` and re-authenticate
- Ensure OAuth consent screen is configured correctly
- Verify API scopes include Photos and Gmail

**OAuth Consent Screen Issues**:
- **"Access blocked" error**: Add your email to "Test users" in OAuth consent screen
- **"App not verified" warning**: This is normal in testing mode - click "Advanced" > "Go to [app name]"
- **Missing scopes**: Ensure all required scopes are added in OAuth consent screen
- **Test user not working**: Verify email is exactly as shown in Google account

**No Photos Found:**
- Check date calculations with `--debug` flag
- Verify photos exist in the specified date range
- Ensure Google Photos sync is enabled on your devices

**Email Not Sending:**
- Verify Gmail API is enabled
- Check recipient email addresses are valid
- Ensure OAuth scope includes Gmail send permission

**Rate Limiting:**
- The script includes automatic retry logic
- Adjust `retry_delay_seconds` in config if needed
- Consider running during off-peak hours

### Debug Mode
Run with `--debug` flag for detailed information:
```bash
python weekly_photo_automation.py --debug
```

## 📝 Advanced Usage

### Custom Week Definitions
To modify the week definition (default: Saturday-Friday), edit the `calculate_week_range()` method in the script.

### Multiple Family Groups
Create separate configuration files for different recipient groups:
```bash
python weekly_photo_automation.py --config family_group1.json
python weekly_photo_automation.py --config family_group2.json
```

### Filtering Photos
To add photo filtering (e.g., by location or album), modify the `get_photos_for_week()` method to include additional filters in the API request.

## 📄 License

This project follows the coding standards defined in RULES.md and is designed for personal use with Google Photos and Gmail APIs.

## 🤝 Contributing

When contributing, please follow the coding standards in RULES.md:
- Use snake_case for functions and variables
- Include type hints for all functions
- Add appropriate logging with emoji indicators
- Update documentation for new features
- Include configuration examples

## 🤖 Source

Created with Claude Opus 4.1 
Private chat link > https://claude.ai/chat/5dcc874f-c0c8-4410-b2f0-3d5212696150