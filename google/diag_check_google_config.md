# Google Cloud Configuration Checker

## Overview

The `check_google_config.py` script is a diagnostic tool that helps troubleshoot Google Cloud authentication and configuration issues. It performs comprehensive checks on your gcloud CLI setup, credentials, and project configuration to identify common problems.

## What the Script Does

### 🔍 Gcloud Configuration Check
- Verifies gcloud CLI is properly installed and accessible
- Shows current active gcloud project
- Lists active authentication accounts
- Handles Windows-specific gcloud detection

### 🔑 Credentials File Check
- Validates `client_secret.json` file exists and is readable
- Extracts project ID and client ID from credentials
- Ensures OAuth2 credentials are properly configured

### 🎫 Token File Check
- Verifies `token.json` file exists
- Lists granted OAuth scopes
- Checks token validity and expiration

### 🧪 Authentication Testing
- Tests authentication with Google Photos API
- Tests authentication with Gmail API
- Attempts token refresh if expired
- Provides detailed error messages for failed API calls

### 🔗 Project Alignment Check
- Compares gcloud project with credentials project
- Identifies project mismatches that cause authentication issues
- Provides warnings when projects don't align

## Prerequisites

### Required Files
- `client_secret.json` - OAuth2 credentials from Google Cloud Console
- `token.json` - Generated authentication tokens (created on first run)

### Required APIs
- Google Photos Library API
- Gmail API

## Installation and Setup

### 1. Install Google Cloud SDK on Windows

#### Option A: Official Installer (Recommended)
1. Download from [Google Cloud SDK](https://cloud.google.com/sdk/docs/install)
2. Run the installer with administrator privileges
3. Default installation path: `C:\Program Files (x86)\Google\Cloud SDK\`

#### Option B: Manual Installation
```powershell
# Download and extract to C:\Program Files\Google\Cloud SDK\
# Add to PATH: C:\Program Files\Google\Cloud SDK\google-cloud-sdk\bin
```

### 2. Verify Installation
```powershell
# Check if gcloud is accessible
gcloud --version

# Check installation path
Get-Command gcloud | Select-Object -ExpandProperty Source
```

### 3. Initial Setup
```powershell
# Login to Google account
gcloud auth login

# Set default project
gcloud config set project YOUR_PROJECT_ID

# Verify configuration
gcloud config list
```

## Usage

### Basic Run
```powershell
cd automation/google
python check_google_config.py
```

### What to Expect
The script will output a comprehensive diagnostic report:

```
🚀 Google Cloud Configuration Diagnostic
==================================================
🔗 Checking project alignment...
🔍 Checking gcloud configuration...
   📋 Current gcloud project: your-project-id
   👤 Active gcloud accounts:
      • user@example.com (ACTIVE)

🔑 Checking credentials file...
   ✅ Credentials file found
   📋 Project ID: your-project-id
   🆔 Client ID: 123456789-abcdef.apps.googleusercontent.com

🎫 Checking token file...
   ✅ Token file found
   🔐 Granted scopes (3):
      • https://www.googleapis.com/auth/photoslibrary.readonly
      • https://www.googleapis.com/auth/gmail.readonly

🧪 Testing authentication...
   ✅ Credentials loaded successfully
   🔑 Token scopes: ['https://www.googleapis.com/auth/photoslibrary.readonly', ...]
   ✅ Photos API: Authentication successful
   ✅ Gmail API: Authentication successful (Email: user@example.com)

💡 Solutions:
   1. If projects don't match, switch gcloud to the correct project:
      gcloud config set project automation-469306
   2. Re-authenticate with the correct project:
      gcloud auth login
   3. Delete token.json and re-run authentication to get fresh tokens
   4. Ensure the Google Cloud project has the required APIs enabled
   5. Check that your OAuth consent screen includes the required scopes

📊 Summary:
   Project alignment: ✅
   Authentication: ✅
```

## Common Issues and Solutions

### ❌ "gcloud CLI not found"
**Problem**: Script can't locate gcloud executable
**Solution**: 
- Ensure Google Cloud SDK is installed
- Add SDK bin directory to PATH
- Use full path to gcloud.cmd or gcloud.ps1

### ❌ "Project mismatch"
**Problem**: gcloud project differs from credentials project
**Solution**:
```powershell
# Switch to correct project
gcloud config set project automation-469306

# Re-authenticate
gcloud auth login
```

### ❌ "Credentials not valid"
**Problem**: OAuth tokens expired or invalid
**Solution**:
```powershell
# Delete old tokens
Remove-Item token.json

# Re-run authentication flow
python your_auth_script.py
```

### ❌ "API not enabled"
**Problem**: Required APIs not enabled in Google Cloud project
**Solution**:
1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Navigate to APIs & Services > Library
3. Enable:
   - Google Photos Library API
   - Gmail API

## Windows-Specific Notes

### PowerShell vs Command Prompt
- **PowerShell**: gcloud works as `gcloud` or `gcloud.ps1`
- **Command Prompt**: Use `gcloud.cmd` explicitly

### PATH Environment Variable
The script automatically detects common installation paths:
- `C:\Program Files (x86)\Google\Cloud SDK\google-cloud-sdk\bin\`
- `C:\Program Files\Google\Cloud SDK\google-cloud-sdk\bin\`

### Administrator Rights
Some operations require administrator privileges:
- Installing Google Cloud SDK
- Updating components: `gcloud components update`
- Modifying system PATH

## Troubleshooting Commands

### Check gcloud Status
```powershell
# Version and components
gcloud --version

# Current configuration
gcloud config list

# Active accounts
gcloud auth list

# Current project
gcloud config get-value project
```

### Reset Configuration
```powershell
# Reset to default configuration
gcloud config configurations activate default

# Create new configuration
gcloud config configurations create my-config

# Delete configuration
gcloud config configurations delete my-config
```

### Update Components
```powershell
# Update all components
gcloud components update

# List available components
gcloud components list

# Install specific component
gcloud components install COMPONENT_NAME
```

## File Structure

```
automation/google/
├── check_google_config.py      # Main diagnostic script
├── check_google_config.md      # This documentation
├── client_secret.json          # OAuth2 credentials (required)
├── token.json                  # Authentication tokens (generated)
└── README.md                   # General project documentation
```

## Security Notes

- **Never commit** `client_secret.json` or `token.json` to version control
- Keep credentials secure and restrict access
- Rotate credentials regularly
- Use service accounts for production applications

## Support

If you encounter issues:
1. Run the diagnostic script first
2. Check the error messages and suggested solutions
3. Verify your Google Cloud project configuration
4. Ensure all required APIs are enabled
5. Check OAuth consent screen configuration
