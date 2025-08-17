# Pull Request: Weekly Photo Album Automation

## Purpose
This PR introduces a new automation module that creates weekly Google Photos albums for children and automatically shares them with family members via email. The automation handles the complete workflow from photo collection to email notification, ensuring family members receive regular photo updates.

## Changes Made

### New Files Created
- `weekly_photo_automation.py` - Main automation script with Google Photos and Gmail integration
- `setup_automation.py` - Interactive setup helper for initial configuration
- `config.json` - Configuration template with children info and email recipients
- `requirements.txt` - Python dependencies for Google API integration
- `run_weekly_photos.bat` - Windows batch script for easy execution
- `README.md` - Comprehensive documentation following project standards
- `PULL_REQUEST.md` - This documentation file

### Key Features Implemented
1. **Intelligent Week Calculation**: Always processes the most recent complete Saturday-Friday week
2. **Sequential Week Numbering**: Tracks weeks since each child's birth date
3. **Google Photos Integration**: Creates albums and generates shareable links
4. **Email Notifications**: Sends HTML emails with album links to specified recipients
5. **Dry Run Mode**: Test functionality without creating albums or sending emails
6. **Comprehensive Logging**: Uses emoji indicators following project standards
7. **Error Handling**: Implements retry logic and graceful error recovery

## Testing Instructions

### Environment Setup - Windows PowerShell
```powershell
# Clone and enter directory
git clone <repo-url> -b weekly-photo-automation weekly-photo-test; cd weekly-photo-test

# Create virtual environment
python -m venv .venv

# Activate virtual environment
.\.venv\Scripts\Activate.ps1

# Install requirements
pip install -r requirements.txt
```

### Environment Setup - Unix/Linux/MacOS
```bash
# Clone and enter directory
git clone <repo-url> -b weekly-photo-automation weekly-photo-test && cd weekly-photo-test

# Create virtual environment
python3 -m venv .venv

# Activate virtual environment
source .venv/bin/activate

# Install requirements
pip install -r requirements.txt
```

### Configuration Steps
1. **Obtain Google Cloud Credentials**:
   - Visit https://console.cloud.google.com
   - Enable Google Photos Library API and Gmail API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download as `credentials.json`

2. **Run Setup Helper**:
   ```bash
   python setup_automation.py
   ```
   Follow prompts to configure children and email recipients

3. **Test Dry Run**:
   ```bash
   python weekly_photo_automation.py --dry-run --debug
   ```
   This will authenticate with Google (browser opens) without creating albums

4. **Run Full Test**:
   ```bash
   python weekly_photo_automation.py --debug
   ```

### Manual Testing Checklist
- [ ] Setup script creates valid config.json
- [ ] Google authentication flow completes successfully
- [ ] Week calculation correctly identifies Saturday-Friday range
- [ ] Week numbers calculate correctly from birth dates
- [ ] Photos are retrieved for the correct date range
- [ ] Albums are created with proper naming convention
- [ ] Shareable links are generated successfully
- [ ] Email is sent to all configured recipients
- [ ] Dry run mode prevents album creation and email sending
- [ ] Debug flag provides detailed logging output
- [ ] Error handling works for missing photos
- [ ] Batch script works on Windows

## Requirements

### Minimal Dependencies
```
google-auth==2.23.4
google-auth-oauthlib==1.1.0
google-auth-httplib2==0.1.1
google-api-python-client==2.108.0
pytz==2023.3.post1
```

### API Requirements
- Google Photos Library API (enabled in Google Cloud Console)
- Gmail API (enabled in Google Cloud Console)
- OAuth 2.0 Desktop application credentials

### Python Version
- Python 3.7 or higher (uses type hints and f-strings)

## Breaking Changes
None - this is a new standalone module that doesn't affect existing functionality.

## Configuration Example
```json
{
  "children": [
    {
      "name": "Valentina",
      "birth_date": "2023-01-15"
    },
    {
      "name": "Luca",
      "birth_date": "2024-03-20"
    }
  ],
  "email": {
    "recipients": [
      "grandmother@example.com",
      "parent@example.com"
    ]
  }
}
```

## Deployment Notes
1. **Credentials Security**: Never commit `credentials.json` or `token.json`
2. **Scheduling**: Can be automated via cron (Linux) or Task Scheduler (Windows)
3. **Rate Limits**: Google Photos API has quotas - script includes retry logic
4. **Token Refresh**: OAuth token auto-refreshes when expired

## Compliance with RULES.md
✅ Uses snake_case for functions and variables
✅ Implements type hints for all functions
✅ JSON-based configuration
✅ Comprehensive logging with emoji indicators
✅ Argparse for command line interface
✅ Proper error handling with retry logic
✅ Follows project structure guidelines
✅ Includes debug mode support
✅ Complete documentation with examples