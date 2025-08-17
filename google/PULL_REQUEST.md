# Pull Request: Gmail and Google Drive Automation

## 🎯 Purpose
This PR introduces a new Google Drive and Gmail integration automation that demonstrates how to:
- Monitor Gmail for emails matching specific criteria
- Automatically create and maintain a Google Sheets spreadsheet in Google Drive
- Prevent duplicate email entries by checking existing records
- Follow the project's established coding standards and structure

This automation serves as a practical demonstration of Google API integration, replacing the non-functional Google Photos API with working Google Drive and Gmail APIs.

## 🔄 Changes Made
- **New automation module**: `gmail_drive_automation.py` - Core automation logic
- **Configuration file**: `config_gmail_drive.json.sample` - Sample configuration
- **Requirements file**: `requirements_gmail_drive.txt` - Minimal dependencies
- **Documentation**: `README_gmail_drive.md` - Comprehensive usage guide
- **Windows batch file**: `run_gmail_drive.bat` - Easy execution on Windows
- **Unix/Linux script**: `run_gmail_drive.sh` - Easy execution on Unix/Linux
- **Test script**: `test_gmail_drive.py` - Functionality testing without API calls

### Key Features
- **Intelligent spreadsheet management**: Creates new spreadsheet on first run, finds existing on subsequent runs
- **Duplicate prevention**: Checks existing records to avoid adding the same email twice
- **Configurable email search**: Uses Gmail search queries to filter emails
- **Professional formatting**: Auto-formats headers with bold text and background colors
- **Comprehensive logging**: Detailed logging with emoji indicators following project standards

## 🧪 Testing Instructions

### **Prerequisites**
- Python 3.8+ installed
- Git installed
- Access to the repository
- Google Cloud Project with enabled APIs (Drive, Gmail, Sheets)

### **Environment Setup Commands**

**Windows PowerShell:**
```powershell
# 1. Clone repo and checkout specific branch into new folder
git clone -b feature/gmail-drive-automation https://github.com/ferraroroberto/automation.git test-gmail-drive

# 2. Navigate to the folder
cd test-gmail-drive

# 3. Navigate to google subfolder
cd google

# 4. Create virtual environment
python -m venv .venv

# 5. Activate virtual environment
.\.venv\Scripts\Activate.ps1

# 6. Install ad-hoc requirements
pip install -r requirements_gmail_drive.txt
```

**Unix/Linux/macOS:**
```bash
# 1. Clone repo and checkout specific branch into new folder
git clone -b feature/gmail-drive-automation https://github.com/ferraroroberto/automation.git test-gmail-drive

# 2. Navigate to the folder
cd test-gmail-drive

# 3. Navigate to google subfolder
cd google

# 4. Create virtual environment
python3 -m venv .venv

# 5. Activate virtual environment
source .venv/bin/activate

# 6. Install ad-hoc requirements
pip install -r requirements_gmail_drive.txt
```

### **Test Execution**

1. **Run unit tests** (no API calls required):
   ```bash
   python test_gmail_drive.py
   ```

2. **Setup Google credentials**:
   - Copy `config_gmail_drive.json.sample` to `config.json`
   - Place your `credentials.json` file in the same directory
   - Customize configuration as needed

3. **Run the automation**:
   ```bash
   # Basic run
   python gmail_drive_automation.py
   
   # With debug logging
   python gmail_drive_automation.py --debug
   ```

4. **Alternative execution**:
   - **Windows**: Double-click `run_gmail_drive.bat`
   - **Unix/Linux**: Run `./run_gmail_drive.sh`

### **Expected Results**
- Unit tests should pass (4/4 tests)
- First run creates a new Google Sheets file in your Drive
- Subsequent runs add new emails to existing spreadsheet
- No duplicate emails are added
- Spreadsheet maintains proper formatting and structure

## 📋 Requirements for Testing
- **Minimal dependencies**: Only Google API libraries and timezone support
- **Google Cloud setup**: Drive, Gmail, and Sheets APIs enabled
- **OAuth credentials**: Desktop app credentials with proper scopes
- **Test user access**: Your email added to OAuth consent screen test users

### **Dependencies**
```
google-auth==2.23.4
google-auth-oauthlib==1.1.0
google-auth-httplib2==0.1.1
google-api-python-client==2.108.0
pytz==2023.3.post1
```

## 🚨 Breaking Changes
- **None**: This is a new feature that doesn't affect existing functionality
- **Compatible**: Uses same authentication pattern as existing Google automation
- **Independent**: Can coexist with existing Google Photos automation

## 📚 Additional Notes

### **Google API Scopes Required**
- `https://www.googleapis.com/auth/drive` - Google Drive access
- `https://www.googleapis.com/auth/gmail.readonly` - Gmail read permission
- `https://www.googleapis.com/auth/spreadsheets` - Google Sheets access

### **Configuration Options**
- **Email search**: Customizable Gmail search queries (e.g., "is:important", "from:boss@company.com")
- **Spreadsheet name**: Configurable spreadsheet title
- **Max results**: Limit number of emails processed per run
- **Timezone**: Date formatting timezone setting

### **Use Cases**
- **Email tracking**: Monitor important emails automatically
- **Project management**: Track project-related communications
- **Customer service**: Monitor support email volumes
- **Compliance**: Maintain audit trails of business communications

### **Automation Ideas**
- **Scheduled execution**: Daily/weekly runs via cron or Task Scheduler
- **Custom filters**: Domain-specific email monitoring
- **Data analysis**: Email volume and pattern analysis
- **Integration**: Connect with other automation tools

This automation demonstrates the project's coding standards while providing a practical, working example of Google API integration that can be easily extended and customized.