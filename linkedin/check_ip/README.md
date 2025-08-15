# LinkedIn Image Search Application

A modular Python application that performs reverse image searches using Google Lens via SerpAPI to find where your images appear across the web, with a focus on LinkedIn and other social media platforms.

## 🚀 Features

- **Metadata Management**: Tracks images already processed and their Imgur URLs
- **Cumulative Database**: Preserves historical search results and user annotations
- **API History Database**: Stores complete details of all API calls including raw JSON responses
- **Smart Optimization**: Avoids re-uploading images that were previously processed
- **Social Media Detection**: Extracts and formats post dates from various platforms (LinkedIn, Twitter/X)
- **Time-based Processing**: Skips images processed within configurable time thresholds
- **Configurable Search Types**: Control whether to perform exact matches, similar matches, or both
- **Excel Integration**: Preserves custom columns and formatting in Excel files
- **Retry Logic**: Robust handling of API rate limits and network issues

## 📁 Project Structure

```
linkedin/
├── main.py                          # Main orchestration script
├── config.py                        # Configuration management
├── storage.py                       # Excel file operations
├── imgur_client.py                  # Imgur upload functionality
├── search_client.py                 # Google Lens search via SerpAPI
├── data_processor.py                # Data processing utilities
├── linkedin_search_image.json       # Configuration file
├── linkedin_search_image.py         # [LEGACY] Original monolithic script
└── README.md                        # This documentation
```

## 🔧 Module Overview

### `config.py`
- **Purpose**: Handles configuration loading from JSON files and command-line argument parsing
- **Key Classes**: `Config`
- **Key Functions**: `parse_arguments()`, `get_app_config()`

### `storage.py`
- **Purpose**: Manages all Excel file operations including loading, saving, and data management
- **Key Classes**: `StorageManager`
- **Features**: Column width preservation, custom column handling, backup creation

### `imgur_client.py`
- **Purpose**: Handles image upload operations to Imgur with retry logic
- **Key Classes**: `ImgurClient`
- **Features**: Exponential backoff, rate limit handling, credential validation

### `search_client.py`
- **Purpose**: Performs Google Lens searches via SerpAPI with API history tracking
- **Key Classes**: `SearchClient`
- **Features**: Search type configuration, API history logging, response processing

### `data_processor.py`
- **Purpose**: Data processing utilities for search results
- **Key Classes**: `DataProcessor`
- **Features**: Date extraction, source identification, duplicate detection

### `main.py`
- **Purpose**: Main orchestration script that coordinates all modules
- **Key Classes**: `LinkedInImageSearchApp`
- **Features**: Complete workflow management, error handling, progress tracking

## ⚙️ Configuration

### Configuration File: `linkedin_search_image.json`

```json
{
  "api_keys": {
    "imgur_client_id": "your_imgur_client_id",
    "imgur_access_token": "your_imgur_access_token",
    "serpapi_key": "your_serpapi_key"
  },
  "folders": {
    "images_folder": "path/to/your/images",
    "metadata_folder": "path/to/metadata/storage"
  },
  "file_names": {
    "metadata_file": "metadata.xlsx",
    "results_file": "search_results_database.xlsx",
    "api_history_file": "api_history_database.xlsx"
  },
  "processing_thresholds": {
    "linkedin_high_threshold": 30,
    "high_linkedin_days": 6,
    "standard_days": 180
  },
  "search_settings": {
    "do_exact_search": true,
    "do_similar_search": false
  },
  "allowed_extensions": [".jpg", ".jpeg", ".png", ".gif", ".bmp"],
  "api_endpoints": {
    "search_endpoint": "https://serpapi.com/search?engine=google_lens"
  }
}
```

### Processing Thresholds Explained

The application uses intelligent processing thresholds to optimize performance:

- **High LinkedIn Threshold**: Images with ≥30 LinkedIn mentions are processed every 6 days
- **Standard Threshold**: Other images are processed every 180 days
- **Rationale**: Popular images on LinkedIn may get new mentions frequently, while others remain static

## 🚀 Installation & Setup

### Prerequisites

```bash
pip install pandas numpy requests openpyxl
```

### Required API Keys

1. **Imgur API**: Get from [Imgur API](https://api.imgur.com/)
   - Create an application to get Client ID and Access Token
   
2. **SerpAPI**: Get from [SerpAPI](https://serpapi.com/)
   - Sign up for an account to get your API key

### Initial Setup

1. Clone or download the application files
2. Create your configuration file by copying and editing `linkedin_search_image.json`
3. Add your API keys to the configuration file
4. Set your images folder and metadata storage paths
5. Run the application!

## 📖 Usage Examples

### Basic Usage

```bash
# Use defaults from linkedin_search_image.json
python main.py
```

### Configuration File Override

```bash
# Use a different configuration file
python main.py --config "path/to/custom_config.json"
```

### Command Line Overrides

```bash
# Override specific settings
python main.py --images-folder "C:\\custom\\images" --standard-days 20
```

### Enable Similar Matches

```bash
# Enable similar matches search (in addition to exact matches)
python main.py --enable-similar-search
```

### Full Example

```bash
python main.py \
    --config "custom_config.json" \
    --images-folder "C:\\Users\\rober\\custom_images" \
    --linkedin-high-threshold 30 \
    --enable-similar-search
```

## 📊 Output Files

The application creates three Excel files in your metadata folder:

### 1. `metadata.xlsx`
Tracks each image and its processing metadata:
- `filename`: Image filename
- `last_processed_date`: When the image was last processed
- `imgur_url`: Uploaded Imgur URL
- `total_links`: Total number of links found
- `exact_match_count`: Number of exact matches
- `similar_match_count`: Number of similar matches
- `linkedin_count`: Number of LinkedIn mentions
- `instagram_count`: Number of Instagram mentions
- `twitter_count`: Number of Twitter/X mentions
- `facebook_count`: Number of Facebook mentions
- `pinterest_count`: Number of Pinterest mentions

### 2. `search_results_database.xlsx`
Contains all search results across all images:
- `local_image`: Original image filename
- `uploaded_url`: Imgur URL
- `found_link`: URL where the image was found
- `title`: Page title where the image appears
- `duplicate`: Duplicate status (0=unique, 1=secondary duplicate, 2=primary duplicate)
- `match_type`: "Exact Match" or "Similar Match"
- `source`: Social media platform (LinkedIn, Twitter/X, etc.)
- `post_date`: Extracted post date (for LinkedIn and Twitter/X)
- `search_date`: When this result was found
- `order`: Order of appearance in search results

### 3. `api_history_database.xlsx`
Detailed log of all API calls made:
- `search_type`: Type of search performed
- `local_image`: Image filename
- `imgur_url`: Imgur URL used
- `search_date`: When the search was performed
- `api_id`: SerpAPI request ID
- `raw_json`: Complete API response (for debugging)
- `playground_link`: Link to replay the search in SerpAPI playground
- Plus additional metadata fields

## 🔍 Advanced Features

### Custom Column Preservation
The application preserves any custom columns you add to the Excel files, making it safe to add your own annotations and data.

### Duplicate Detection
The application intelligently marks duplicate URLs:
- **0**: Unique URL (appears only once)
- **1**: Secondary duplicate (appears multiple times, this is not the primary entry)
- **2**: Primary duplicate (appears multiple times, this is the oldest/primary entry)

### Date Extraction
Automatically extracts post dates from:
- **LinkedIn**: Uses activity ID to calculate post timestamp
- **Twitter/X**: Uses tweet ID snowflake algorithm to calculate timestamp
- **Other platforms**: Returns None (date extraction not supported)

### Retry Logic
Robust handling of:
- **Imgur uploads**: Exponential backoff for rate limits and server errors
- **SerpAPI searches**: Proper error handling and logging
- **File operations**: Retry logic for locked Excel files

## 🐛 Troubleshooting

### Common Issues

1. **Excel file is locked**: Close Excel and the application will retry automatically
2. **Imgur rate limits**: The application will wait and retry with exponential backoff
3. **SerpAPI errors**: Check your API key and account limits
4. **Missing dependencies**: Install required packages with pip

### Logging

The application provides detailed logging with emojis for easy reading:
- 🚀 Process start/image upload
- ✅ Success operations
- ❌ Errors
- ⚠️ Warnings
- 📊 Statistics and data operations
- 🔍 Search operations
- 💾 File operations

### Debug Mode

For more detailed logging, modify the logging level in `main.py`:

```python
logging.basicConfig(level=logging.DEBUG)
```

## 🔄 Migration from Legacy Script

If you're upgrading from the original `linkedin_search_image.py`:

1. Your existing Excel files will work with the new modular version
2. The new version preserves all existing functionality
3. Configuration is now in JSON format instead of global variables
4. Command line arguments remain the same
5. All output formats are backward compatible

## 📈 Performance Considerations

- **Image Upload Optimization**: Images are only uploaded to Imgur once, then reused
- **Smart Processing**: Recent images are skipped based on configurable thresholds
- **Rate Limit Handling**: Built-in exponential backoff prevents API exhaustion
- **Memory Efficiency**: Processes images one at a time to avoid memory issues
- **Parallel Operations**: Could be enhanced with asyncio for better performance

## 🤝 Contributing

This modular architecture makes it easy to contribute:

1. **Add new social platforms**: Extend `DataProcessor.identify_social_media_source()`
2. **Add new date extractors**: Extend `DataProcessor.extract_post_date()`
3. **Add new storage backends**: Implement new storage classes
4. **Add new search engines**: Implement new search client classes
5. **Improve UI**: Add web interface or GUI components

## 📝 License

This project is provided as-is for educational and personal use.

## 👤 Authors

- Roberto (Original implementation)
- Claude (Refactoring and modularization)

---

**Note**: This application requires valid API keys for Imgur and SerpAPI. Please ensure you comply with their respective terms of service and rate limits. 