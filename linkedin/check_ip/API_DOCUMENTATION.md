# API Documentation

## Module: `config.py`

### Class: `Config`

#### Methods:
- `__init__(config_path=None)` - Initialize configuration manager
- `get_api_keys()` - Get API keys dictionary
- `get_folders()` - Get folder paths dictionary  
- `get_file_names()` - Get file names dictionary
- `get_processing_thresholds()` - Get processing thresholds dictionary
- `get_search_settings()` - Get search settings dictionary
- `get_allowed_extensions()` - Get allowed file extensions list
- `get_api_endpoints()` - Get API endpoints dictionary

#### Functions:
- `parse_arguments()` - Parse command line arguments
- `get_app_config(args)` - Build application configuration from arguments

## Module: `storage.py`

### Class: `StorageManager`

#### Methods:
- `__init__(metadata_folder)` - Initialize storage manager
- `setup_folders_and_files(images_folder, metadata_file, results_file, api_history_file)` - Setup folder structure
- `load_metadata(metadata_path)` - Load metadata Excel file
- `load_results_database(results_path)` - Load results database Excel file
- `load_api_history_database(history_path)` - Load API history Excel file
- `save_excel_with_retry(df, file_path, max_retries=5, retry_delay=5)` - Save Excel with retry logic

## Module: `imgur_client.py`

### Class: `ImgurClient`

#### Methods:
- `__init__(client_id, access_token)` - Initialize Imgur client
- `upload_image(image_path, retries=20)` - Upload image with retry logic
- `validate_credentials()` - Validate Imgur API credentials

## Module: `search_client.py`

### Class: `SearchClient`

#### Methods:
- `__init__(api_key, search_endpoint)` - Initialize search client
- `search_google_lens(image_url, local_image, api_history_df, search_type="all")` - Perform Google Lens search
- `validate_api_key()` - Validate SerpAPI key

## Module: `data_processor.py`

### Class: `DataProcessor`

#### Static Methods:
- `extract_post_date(url)` - Extract post date from LinkedIn/Twitter URLs
- `identify_social_media_source(url)` - Identify social media platform from URL
- `is_recently_processed(last_processed_date, linkedin_count, thresholds...)` - Check if image was recently processed
- `update_metadata_counts(metadata_df, results_df)` - Update link counts in metadata
- `mark_duplicate_urls(results_df)` - Mark duplicate URLs in results
- `collect_valid_image_paths(images_folder, allowed_extensions)` - Get valid image files
- `process_search_results(search_data, search_type, img_filename, img_url, current_time_str, existing_urls)` - Process search API results

## Module: `main.py`

### Class: `LinkedInImageSearchApp`

#### Methods:
- `__init__(config)` - Initialize application with configuration
- `run()` - Run the main image search process

#### Functions:
- `main()` - Main entry point

## Usage Patterns

### Basic Configuration Loading
```python
from config import Config, parse_arguments, get_app_config

args = parse_arguments()
config = get_app_config(args)
```

### Storage Operations
```python
from storage import StorageManager

storage = StorageManager(metadata_folder)
metadata_df = storage.load_metadata(metadata_path)
storage.save_excel_with_retry(df, file_path)
```

### Image Upload
```python
from imgur_client import ImgurClient

client = ImgurClient(client_id, access_token)
img_url = client.upload_image(image_path)
```

### Search Operations
```python
from search_client import SearchClient

search_client = SearchClient(api_key, endpoint)
data, history_df = search_client.search_google_lens(url, filename, history_df, "exact_matches")
```

### Data Processing
```python
from data_processor import DataProcessor

processor = DataProcessor()
post_date = processor.extract_post_date(url)
source = processor.identify_social_media_source(url)
results_df = processor.mark_duplicate_urls(results_df)
``` 