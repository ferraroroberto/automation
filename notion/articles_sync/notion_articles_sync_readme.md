# Notion Articles Sync

This module syncs articles from a source Notion database to an archive database with advanced threading and rate limiting capabilities. It creates new entries, updates existing ones, and archives target entries when they no longer exist in the source.

## 🚀 Overview

The sync process:
1. **Watches** the source "articles" database for changes
2. **Creates** new entries in the archive database
3. **Updates** existing archive entries when source changes
4. **Archives** target entries when source records are deleted
5. **Tracks** bidirectional relationships between databases
6. **Processes** operations in parallel with configurable threading
7. **Limits** API requests using intelligent rate limiting
8. **Validates** configuration using JSON schema validation
9. **Secures** sensitive data using environment variables

## 📁 Files

- `notion_articles_sync.py` - Main Python module with threading support and comprehensive type hints
- `notion_articles_sync.json` - Configuration file with environment variable placeholders
- `notion_articles_sync_last_time.txt` - Tracks last successful sync time
- `notion_articles_sync_readme.md` - This comprehensive documentation
- `.env` - Environment variables file (located in project root)
- `notion_articles_sync.log` - Created when `--debug` is enabled

## 🔧 Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Environment Configuration
Create a `.env` file in the **project root directory** (not in the notion subfolder):

```bash
# Notion API Configuration
NOTION_API_TOKEN=your_actual_api_token_here
```

**Note**: Only the API token is stored in the `.env` file for security. All other configuration (database IDs, sync settings, threading options) is defined in the `notion_articles_sync.json` file with sensible defaults.

### 3. Notion Integration Setup
1. Go to https://www.notion.so/my-integrations
2. Create an integration and copy the Internal Integration Token
3. Share both databases with this integration
4. Update the `.env` file with your actual values

### 4. Configuration File
The `notion_articles_sync.json` file now uses environment variable placeholders. The system will automatically replace these with values from your `.env` file.

## 🚀 Usage

### Basic Usage
```bash
# Run continuous sync
python notion_articles_sync.py --config notion_articles_sync.json

# Run once and exit
python notion_articles_sync.py --config notion_articles_sync.json --once

# Force a full comparison (detect deletions) for this run
python notion_articles_sync.py --config notion_articles_sync.json --full-sync --once

# Reset sync time to force a full sync on next run
python notion_articles_sync.py --config notion_articles_sync.json --reset-sync-time

# Show current sync status
python notion_articles_sync.py --config notion_articles_sync.json --status

# Enable verbose logging and write a log file next to the config
python notion_articles_sync.py --config notion_articles_sync.json --debug
```

### Command Line Options

- `--config` - Path to configuration JSON file (default: file next to the script)
- `--once` - Run one sync cycle and exit
- `--full-sync` - Force a full comparison (also detects deletions)
- `--reset-sync-time` - Reset the last sync time to force a full sync on next run
- `--status` - Show current sync status and exit
- `--debug` - Enable debug logging and write `notion_articles_sync.log`

## ⚙️ Configuration

### Environment Variables
The system now uses environment variables only for sensitive data:

| Environment Variable | Description | Required |
|---------------------|-------------|----------|
| `NOTION_API_TOKEN` | Your Notion API token | Yes |

**Note**: All other configuration (database IDs, sync settings, threading options) is defined in the `notion_articles_sync.json` file with sensible defaults. This approach provides security for sensitive data while maintaining easy configuration management.

### Field Mapping
The module maps properties by name (not encoded property IDs). Defaults shipped in `notion_articles_sync.json`:

| Source Field | Target Field |
|--------------|--------------|
| `article` | `article` |
| `summary` | `summary` |
| `topic` | `topic` |
| `link` | `link` |
| `niche` | `niche` |
| `post` | `post` |
| `author export` | `author or source` |
| `news export` | `newsletter` |
| `rowid` | `source rowid` |
| `created` | `created` |

### Sync Rules
- **Exclude Archive**: Items with `exclude archive = true` in the source are skipped
- **Cascade Deletions**: When a source item disappears, the corresponding target page is archived
- **Bidirectional Tracking**: The source item receives `target rowid` with the created/updated target page ID

### Sync Time Tracking
The module properly tracks when the last sync actually occurred:

- **Persistent Storage**: Sync time is saved to `notion_articles_sync_last_time.txt`
- **Accurate Incremental Sync**: Only processes items that have actually changed since the last sync
- **Manual Control**: Use `--reset-sync-time` to force a full sync, or `--status` to check current sync state
- **Automatic Updates**: Sync time is automatically updated after each successful sync cycle

## 🔒 Security Features

### Environment Variable Management
- **Minimal Sensitive Data**: Only API token stored in `.env` file
- **Root Level Configuration**: `.env` file located in project root for shared access
- **Automatic Loading**: Uses `python-dotenv` for seamless environment variable loading
- **Simplified Processing**: Only processes the API token from environment variables

### Configuration Validation
- **JSON Schema Validation**: Uses `jsonschema` for configuration validation
- **Required Field Checking**: Ensures all necessary configuration is present
- **Data Type Validation**: Validates configuration value types and ranges
- **Early Error Detection**: Fails fast if configuration is invalid

## 🏗️ Architecture

### Code Quality Standards
- **Comprehensive Type Hints**: Full type annotation following RULES.md standards
- **Enhanced Documentation**: Detailed docstrings with parameter descriptions
- **Error Handling**: Graceful error handling with retry logic
- **Logging Standards**: Consistent emoji-based logging throughout

### Threading & Performance
- **Parallel Database Queries**: Multiple workers can fetch database pages simultaneously
- **Parallel Operations**: Create, update, and delete operations run in parallel
- **Configurable Workers**: Adjust `max_workers` based on your system capabilities
- **Batch Processing**: Operations are grouped into configurable batch sizes

### Rate Limiting
- **Token Bucket Algorithm**: Intelligent rate limiting that respects Notion API limits
- **Configurable Limits**: Set `requests_per_second` and `burst_size` in config
- **Thread-Safe**: Rate limiter works correctly with multiple threads
- **Automatic Backoff**: Built-in retry logic with exponential backoff

## 📊 Monitoring & Logging

### Log Levels
- **INFO**: Default level with emoji-based visual indicators
- **DEBUG**: Detailed information when `--debug` flag is used
- **ERROR**: Error conditions with context information

### Log Format
```
2024-01-15 10:30:45 - INFO - ✅ Notion sync initialized
2024-01-15 10:30:45 - INFO - 📊 Source: 67fbcee66711465c852ebf97303787a3
2024-01-15 10:30:45 - INFO - 🚀 Threading: 5 workers, 3.0 req/s
```

### Debug Mode Features
When `--debug` is enabled:
- **File Logging**: Creates `notion_articles_sync.log` file
- **API Details**: Logs all API requests and responses
- **Schema Information**: Shows database schema usage
- **Performance Metrics**: Tracks API call counts and timing

## 🚨 Troubleshooting

### Common Issues

#### 1. Environment Variable Errors
**Problem**: `❌ Environment variable not found: NOTION_API_TOKEN`
**Solution**: 
- Ensure `.env` file exists in project root directory
- Check that variable names match exactly (case-sensitive)
- Verify no extra spaces or quotes around values

#### 2. Configuration Validation Failures
**Problem**: `❌ Configuration validation failed`
**Solution**:
- Check that all required fields are present in `.env`
- Ensure database IDs are valid Notion UUIDs
- Verify numeric values are within valid ranges

#### 3. API Token Issues
**Problem**: `❌ Cannot access source database`
**Solution**:
- Verify `NOTION_API_TOKEN` in `.env` file
- Ensure integration has access to both databases
- Check that integration is properly shared

#### 4. Database Access Denied
**Problem**: `❌ Cannot access target database`
**Solution**:
- Double-check database IDs in `.env` file
- Ensure integration has read/write permissions
- Verify databases are shared with the integration

#### 5. Rate Limiting Issues
**Problem**: Slow performance or API errors
**Solution**:
- Reduce `THREADING_REQUESTS_PER_SECOND` in `.env`
- Increase `THREADING_BURST_SIZE` for better burst handling
- Reduce `THREADING_MAX_WORKERS` if experiencing connection issues

#### 6. Field Mapping Errors
**Problem**: Properties not syncing correctly
**Solution**:
- Ensure both databases define the mapped properties
- Check property types match between source and target
- Verify field names in configuration file

### Deep Diagnostics

#### Enable Debug Mode
```bash
python notion_articles_sync.py --config notion_articles_sync.json --debug
```

#### Check Log Files
- Review `notion_articles_sync.log` for detailed information
- Look for API request/response details
- Check for schema validation errors

#### Verify Configuration
```bash
python notion_articles_sync.py --config notion_articles_sync.json --status
```

#### Test Connections
The system automatically tests database connections on startup. Look for:
- `✅ Source database accessible`
- `✅ Target database accessible`

## 🔧 Performance Tuning

### Threading Configuration
- **Start with**: `THREADING_MAX_WORKERS=5`
- **Adjust based on**: System capabilities and API performance
- **Monitor**: API call success rates and response times

### Rate Limiting
- **Safe Default**: `THREADING_REQUESTS_PER_SECOND=3.0`
- **Burst Handling**: `THREADING_BURST_SIZE=10`
- **Adjust if**: Experiencing API rate limit errors

### Batch Sizes
- **Database Queries**: `SYNC_BATCH_SIZE=100` (API efficiency)
- **Operation Batching**: `THREADING_OPERATION_BATCH_SIZE=10` (memory usage)

### Polling Intervals
- **Responsive**: `SYNC_POLLING_INTERVAL=300` (5 minutes)
- **Conservative**: `SYNC_POLLING_INTERVAL=600` (10 minutes)
- **Balance**: Between responsiveness and API usage

## 🚀 Advanced Features

### Full Sync Mode
Force a complete database comparison:
```bash
python notion_articles_sync.py --full-sync --once
```

### Incremental Sync
Default mode that only processes changes since last sync:
```bash
python notion_articles_sync.py --once
```

### Sync Time Management
```bash
# Check current sync status
python notion_articles_sync.py --status

# Reset sync time for full sync
python notion_articles_sync.py --reset-sync-time

# Force full sync on next run
python notion_articles_sync.py --full-sync --once
```

## 📋 Support & Maintenance

### Regular Maintenance
1. **Monitor Logs**: Check for errors and performance issues
2. **Update Dependencies**: Keep `python-dotenv` and `jsonschema` updated
3. **Rotate Tokens**: Periodically update Notion API tokens
4. **Review Configuration**: Adjust settings based on usage patterns

### Getting Help
1. **Check Logs**: Enable debug mode and review log files
2. **Verify Configuration**: Ensure all environment variables are set
3. **Test Permissions**: Verify Notion integration access
4. **Review Documentation**: Check this README for solutions

### Reporting Issues
When reporting issues, include:
- Error messages from logs
- Configuration file contents (without sensitive data)
- Environment variable names (not values)
- Notion integration permissions
- System information and Python version

## 🔄 Migration from Previous Versions

### Breaking Changes
- **Environment Variables**: Only API token now uses `.env` file
- **Configuration Validation**: JSON schema validation is now mandatory
- **Type Hints**: All functions now have comprehensive type annotations

### Migration Steps
1. **Create `.env` file** in project root with your configuration
2. **Update configuration file** to use environment variable placeholders
3. **Install new dependencies**: `pip install python-dotenv jsonschema`
4. **Test configuration**: Run with `--status` flag first

### Backward Compatibility
- Existing functionality remains unchanged
- Configuration file structure is preserved
- All command line options work as before

---

**Note**: This module follows the automation project RULES.md standards for code quality, security, and maintainability. For questions or issues, refer to the troubleshooting section above or check the project documentation.
