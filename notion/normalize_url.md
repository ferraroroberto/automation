# Notion URL Normalizer

## 🚀 Overview

Automatically cleans URLs in a Notion database by removing unnecessary query parameters (like UTM tags and tracking codes) while preserving parameters for specific domains that require them (like YouTube or Vimeo). This tool ensures your link repository stays clean and functional without broken video embeds.

## 📋 Usage

### Basic Usage

```bash
# Normalize URLs from the last 14 days
python normalize_url.py --days 14 --config normalize_url.json

# Preview changes without updating (dry run) with validation check
python normalize_url.py --days 7 --dry-run --testing

# Test normalization on a specific string
python normalize_url.py --test "https://example.com?utm_source=test"
```

### Command Line Arguments

- `--days` (default: 14): Number of days to look back for created articles
- `--config` (default: normalize_url.json): Path to JSON configuration file
- `--debug`: Enable debug logging for detailed output
- `--test <string>`: Test mode - clean a specific URL string without processing the database
- `--dry-run`: Preview changes without updating Notion
- `--testing`: Enable validation mode to check if the cleaned URL is accessible (returns 200 OK)

## 🔧 Configuration

The tool uses a JSON configuration file (`normalize_url.json`) with the following structure:

```json
{
  "notion_api_key": "${NOTION_API_TOKEN}",
  "database_id": "your-database-id",
  "domains_preserving_params": [
    "youtube.com",
    "www.youtube.com",
    "youtu.be",
    "vimeo.com",
    "twitter.com",
    "x.com"
  ]
}
```

### Configuration Fields

- **notion_api_key**: Notion API token (can use `${NOTION_API_TOKEN}` env variable)
- **database_id**: Notion database ID to process
- **domains_preserving_params**: List of domains where query parameters should NOT be removed (e.g., video platforms where `?v=id` is required)

## 📊 Normalization Rules

The normalizer follows these rules:

### 1. Parameter Stripping
- For most domains, **all** query parameters (everything after `?`) are removed.
- Example: `https://example.com/post?utm_source=newsletter` → `https://example.com/post`

### 2. Domain Whitelisting
- Domains listed in `domains_preserving_params` are **skipped** entirely.
- Example: `https://www.youtube.com/watch?v=dQw4w9WgXcQ` → Unchanged

### 3. Validation (Optional)
- When `--testing` is used, the tool performs a `HEAD` request (falling back to `GET`) to verify the cleaned URL exists.
- Status codes 200-399 are considered valid (✅).
- 400+ codes are flagged as errors (❌).

## 🔍 How It Works

### Processing Flow

1. **Query Notion Database**: Retrieves articles created in the specified time window
2. **Extract URLs**: Extracts the "link" property (URL type) from each page
3. **Clean URL**:
   - Checks if domain is whitelisted
   - If not, rebuilds URL dropping `query` components
4. **Validate (Optional)**: Checks if the link is alive
5. **Update Notion**: Updates the link property with the cleaned URL (unless dry-run)

### Key Methods

- `_clean_url()`: Main cleaning logic using `urllib.parse`
- `_check_url_validity()`: Verification logic using `requests`
- `process_database()`: Orchestrates the batch processing

## 📝 Examples

### Basic Cleaning
```
Input:  "https://techcrunch.com/article?utm_campaign=daily&ref=newsletter"
Output: "https://techcrunch.com/article"
```

### Preserved Domains (YouTube)
```
Input:  "https://www.youtube.com/watch?v=abcdef"
Output: "https://www.youtube.com/watch?v=abcdef"
```

### Validation Output
```
📝 Cleaned: "http://bit.ly/xyz?s=1" → "http://bit.ly/xyz" [✅ Link: OK (200)]
```

## ⚠️ Important Notes

1. **API Credentials**: Requires Notion API token set in environment variable `NOTION_API_TOKEN`
2. **Database Property**: Expects a "link" property of type "url" in the Notion database
3. **Validation Speed**: Using `--testing` will slow down processing due to network requests
4. **False Negatives**: Some sites block automated validation requests (403 Forbidden), though a browser User-Agent is used to minimize this.

## 📚 Dependencies

- `requests`: HTTP library for Notion API calls and validation
- `python-dotenv`: Environment variable management
- `urllib.parse`: Standard library for URL manipulation

