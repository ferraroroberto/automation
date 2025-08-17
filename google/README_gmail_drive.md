# Gmail and Google Drive Automation

Automatically creates and maintains a Google Sheets spreadsheet to track emails from Gmail based on customizable search criteria.

## 🚀 Overview

This automation tool:
- Monitors Gmail for emails matching specific search criteria
- Creates a Google Sheets spreadsheet in Google Drive (if it doesn't exist)
- Automatically adds new email records to the spreadsheet
- Prevents duplicate entries by checking existing records
- Maintains a clean, formatted spreadsheet with email data

The system intelligently handles:
- **First run**: Creates a new spreadsheet with proper headers and formatting
- **Subsequent runs**: Finds existing spreadsheet and adds only new emails
- **Duplicate prevention**: Checks existing records to avoid adding the same email twice
- **Data organization**: Maintains columns for Subject, Date, From, and Content Preview

## 📊 Spreadsheet Structure

The automation creates a spreadsheet with the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| A | Subject | "Meeting Reminder - Tomorrow" |
| B | Date | "2024-01-15 14:30:00" |
| C | From | "john.doe@company.com" |
| D | Content Preview | "Hi team, just a reminder about..." |

## 🔧 Prerequisites

Before using this automation, you need:
1. A Google Cloud Project with enabled APIs
2. OAuth 2.0 credentials
3. Python 3.7 or higher

### Setting up Google Cloud APIs

1. **Create a Google Cloud Project**:
   - Go to [Google Cloud Console](https://console.cloud.google.com)
   - Create a new project or select an existing one

2. **Enable Required APIs**:
   - Enable the **Google Drive API**
   - Enable the **Gmail API**
   - Enable the **Google Sheets API**

3. **Create OAuth 2.0 Credentials**:
   - Go to APIs & Services > Credentials
   - Click "Create Credentials" > "OAuth client ID"
   - Choose "Desktop app" as the application type
   - Download the credentials as `credentials.json`

4. **Configure OAuth Consent Screen**:
   - Go to APIs & Services > OAuth consent screen
   - Choose "External" user type (unless you have a Google Workspace organization)
   - Fill in the required app information:
     - **App name**: `Gmail Drive Automation` (or your preferred name)
     - **User support email**: Your email address
     - **Developer contact information**: Your email address
   
   **⚠️ IMPORTANT: Audience Settings for Testing**
   - In the "Test users" section, add your email address
   - This allows you to test the app while it's in development mode
   - **Without adding test users, you'll get "Access blocked" errors**
   
   **Scopes to Add**:
   - `https://www.googleapis.com/auth/drive` (Google Drive access)
   - `https://www.googleapis.com/auth/gmail.readonly` (Gmail read permission)
   - `https://www.googleapis.com/auth/spreadsheets` (Google Sheets access)
   
   **Testing vs Production**:
   - **Testing Mode**: App can only be used by added test users
   - **Production Mode**: Requires Google verification (can take weeks)
   - For personal automation, testing mode is sufficient

## 🔄 How It Works

### 1. Authentication
- Uses OAuth 2.0 to authenticate with Google services
- Stores authentication tokens locally for future use
- Automatically refreshes expired tokens

### 2. Spreadsheet Management
- **First run**: Creates a new spreadsheet with headers and formatting
- **Subsequent runs**: Locates existing spreadsheet by name
- Automatically formats headers with bold text and background color
- Auto-resizes columns for optimal viewing

### 3. Email Processing
- Searches Gmail using configurable search criteria
- Extracts key information: Subject, From, Date, Content snippet
- Filters out emails already in the spreadsheet
- Adds new emails as rows to the spreadsheet

### 4. Duplicate Prevention
- Maintains a set of existing email subjects
- Compares new emails against existing records
- Only adds emails that aren't already tracked

## 📋 Configuration

### Configuration File Structure

Create a `config.json` file based on the `config_gmail_drive.json.sample`:

```json
{
    "auth": {
        "credentials_file": "credentials.json",
        "token_file": "token.json"
    },
    "spreadsheet": {
        "name": "Email Tracking Dashboard"
    },
    "gmail": {
        "search_query": "is:important",
        "max_results": 100
    },
    "settings": {
        "timezone": "UTC",
        "retry_attempts": 3,
        "retry_delay_seconds": 5
    }
}
```

### Configuration Options

#### Authentication (`auth`)
- `credentials_file`: Path to your Google OAuth credentials file
- `token_file`: Path where authentication tokens will be stored

#### Spreadsheet (`spreadsheet`)
- `name`: Name of the spreadsheet to create/find in Google Drive

#### Gmail (`gmail`)
- `search_query`: Gmail search query to filter emails (e.g., "is:important", "from:boss@company.com")
- `max_results`: Maximum number of emails to process per run

#### Settings (`settings`)
- `timezone`: Timezone for date formatting
- `retry_attempts`: Number of retry attempts for API calls
- `retry_delay_seconds`: Delay between retry attempts

### Gmail Search Query Examples

- `is:important` - Important emails
- `from:specific@email.com` - Emails from specific sender
- `subject:meeting` - Emails with "meeting" in subject
- `has:attachment` - Emails with attachments
- `after:2024/01/01` - Emails after specific date
- `label:work` - Emails with specific label

## 🚀 Installation

### Step 1: Clone and Setup Environment

**Windows PowerShell:**
```powershell
# Clone repository (or copy files to a folder)
mkdir gmail-drive-automation; cd gmail-drive-automation

# Create virtual environment
python -m venv .venv

# Activate virtual environment
.\.venv\Scripts\Activate.ps1
```

**Unix/Linux/macOS:**
```bash
# Clone repository (or copy files to a folder)
mkdir gmail-drive-automation; cd gmail-drive-automation

# Create virtual environment
python3 -m venv .venv

# Activate virtual environment
source .venv/bin/activate
```

### Step 2: Install Dependencies

```bash
pip install -r requirements_gmail_drive.txt
```

### Step 3: Setup Configuration

1. Copy `config_gmail_drive.json.sample` to `config.json`
2. Place your `credentials.json` file in the same directory
3. Customize the configuration as needed

### Step 4: Run the Automation

```bash
# Basic run
python gmail_drive_automation.py

# With custom config file
python gmail_drive_automation.py --config my_config.json

# With debug logging
python gmail_drive_automation.py --debug
```

## 📊 Output

### Spreadsheet Creation
- Creates a new Google Sheets file in your Google Drive
- Sets up formatted headers with bold text and background color
- Auto-resizes columns for optimal viewing

### Data Population
- Adds email data as new rows
- Formats dates consistently
- Truncates long content to fit in preview column
- Maintains chronological order

### Logging
- Provides detailed logging of all operations
- Shows progress and success/failure messages
- Includes debug information when `--debug` flag is used

## 🔍 Troubleshooting

### Common Issues

#### "Access blocked" Error
- **Cause**: Your email is not added as a test user
- **Solution**: Add your email to the OAuth consent screen test users

#### "File not found" Error
- **Cause**: Configuration file or credentials file not found
- **Solution**: Ensure all files are in the correct directory

#### "API not enabled" Error
- **Cause**: Required Google APIs are not enabled
- **Solution**: Enable Google Drive, Gmail, and Sheets APIs in Google Cloud Console

#### "Quota exceeded" Error
- **Cause**: API usage limits reached
- **Solution**: Wait for quota reset or request quota increase

### Debug Mode
Use the `--debug` flag for detailed logging:
```bash
python gmail_drive_automation.py --debug
```

This will show:
- Detailed API request/response information
- Step-by-step execution details
- Performance metrics
- Error traces

## 🔄 Automation Ideas

### Scheduled Execution
- **Windows Task Scheduler**: Run daily/weekly
- **Linux Cron**: Schedule regular execution
- **GitHub Actions**: Automated runs on schedule

### Custom Search Criteria
- **Work emails**: `from:*@company.com`
- **Important notifications**: `is:important OR subject:urgent`
- **Meeting reminders**: `subject:meeting OR subject:calendar`
- **Project updates**: `subject:project OR subject:update`

### Data Analysis
- **Email volume tracking**: Count emails by date
- **Sender analysis**: Group by sender domain
- **Content categorization**: Analyze email subjects and content
- **Response time tracking**: Monitor email patterns

## 📚 API References

- [Google Drive API v3](https://developers.google.com/drive/api/v3/reference)
- [Gmail API v1](https://developers.google.com/gmail/api/v1/reference)
- [Google Sheets API v4](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets)

## 🤝 Contributing

This automation follows the project's coding standards:
- **Type hints**: All functions use proper type annotations
- **Error handling**: Graceful error handling with retry logic
- **Logging**: Comprehensive logging with emoji indicators
- **Configuration**: JSON-based configuration management
- **Documentation**: Clear docstrings and inline comments

## 📄 License

This automation is part of the larger automation project and follows the same licensing terms.