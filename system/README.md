# System Automation Module

Collection of system-level automation scripts for Windows environments.

## 🚀 Overview

This module contains various system automation tools including WiFi password extraction, system monitoring, and utility functions.

## 📋 Available Scripts

### WiFi Passwords (`wifi_passwords.py`)

Extracts saved WiFi network information and passwords from Windows systems.

**Features:**
- Exports all saved WLAN profiles with clear text keys
- Parses XML profile files for network details
- Outputs to both console and JSON file
- Configurable output format and logging

**Security Note:** This script extracts WiFi passwords in clear text. Use responsibly and ensure output files are properly secured.

## 🔧 Configuration

### WiFi Passwords Configuration

Configuration is stored in `wifi_passwords.json` in the same directory as the script:

```json
{
  "default_output": "wifi_passwords.json",
  "temp_dir_prefix": "wifi_export_",
  "logging": {
    "level": "INFO",
    "format": "%(asctime)s - %(levelname)s - %(message)s"
  },
  "output_format": {
    "include_metadata": true,
    "pretty_print": true,
    "ensure_ascii": false
  },
  "xml_parsing": {
    "namespace": "http://www.microsoft.com/networking/WLAN/profile/v1",
    "fallback_to_filename": true
  }
}
```

## 📊 Output Format

### WiFi Passwords Output

The script generates a JSON file with the following structure:

```json
{
  "metadata": {
    "total_networks": 5,
    "export_timestamp": "1234567890.0",
    "source": "Windows WLAN Profile Export",
    "config_source": "wifi_passwords.json"
  },
  "networks": [
    {
      "ssid": "MyWiFi",
      "authentication": "WPA2PSK",
      "encryption": "AES",
      "password": "mypassword123"
    }
  ]
}
```

## 🚀 Usage

### Basic Usage

```bash
# Export to default file (wifi_passwords.json)
python wifi_passwords.py

# Specify custom output file
python wifi_passwords.py --output my_networks.json

# Enable debug logging
python wifi_passwords.py --debug
```

### Command Line Options

- `--output, -o`: Specify output JSON file path
- `--debug`: Enable debug logging for troubleshooting

## 🔒 Security Considerations

- **Password Exposure**: This script extracts WiFi passwords in clear text
- **File Permissions**: Ensure output files have appropriate permissions
- **Temporary Files**: XML files are created temporarily and automatically cleaned up
- **Logging**: Debug logs may contain sensitive information

## 🛠️ Troubleshooting

### Common Issues

1. **Permission Denied**: Run as Administrator for WiFi profile access
2. **No Networks Found**: Ensure WiFi profiles exist and are accessible
3. **XML Parsing Errors**: Check if profiles are corrupted or incomplete

### Debug Mode

Use `--debug` flag for detailed logging:

```bash
python wifi_passwords.py --debug
```

## 📝 Dependencies

- Python 3.6+
- Standard library modules (no external dependencies)
- Windows environment with `netsh` command available

## 🔧 Development

### Code Standards

This module follows the project's RULES.md standards:
- Type hints for all functions
- Comprehensive error handling
- Structured logging with emojis
- JSON-based configuration
- Command-line argument parsing with argparse
- Graceful fallbacks for configuration errors
