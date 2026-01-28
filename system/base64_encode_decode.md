# Base64 Encoder/Decoder Tool

A comprehensive GUI tool for Base64 encoding and decoding files, with advanced split and join capabilities for handling large files.

## Features

### Basic Operations
- **Encode File**: Convert binary files to Base64 encoded text files
- **Decode File**: Convert Base64 encoded text files back to original binary format
- **Encode Folder**: Batch encode all files in a folder (skips .txt files)
- **Decode Folder**: Batch decode all .txt files in a folder

### Advanced Operations
- **Encode & Split File**: Encode a file and split it into chunks of specified size (in MB)
- **Join & Decode File**: Automatically find and join split files, then decode back to original

## Usage

### GUI Interface
Run the script to launch the tkinter-based GUI:

```bash
python base64_encode_decode.py
```

The interface provides buttons for all operations with intuitive file selection dialogs.

### Split and Join Workflow

#### Splitting Files
1. Click "Encode & Split File"
2. Select the file you want to encode and split
3. Enter the desired chunk size in MB (e.g., 1 for 1MB chunks)
4. Choose a base filename for the output files
5. The tool will create files named like: `filename_001.txt`, `filename_002.txt`, etc.

#### Joining Files
1. Click "Join & Decode File"
2. Select ANY of the split files (e.g., `filename_005.txt`)
3. Choose where to save the decoded original file
4. The tool automatically finds all related split files and joins them in order

## Technical Details

### File Naming Convention
Split files use the pattern: `basename_XXX.txt` where XXX is a 3-digit sequential number starting from 001.

### Chunk Size Calculation
- Chunk size is specified in MB but converted to bytes internally
- Base64 encoding increases file size by ~33%, so actual chunk sizes may vary slightly
- The tool splits the Base64 encoded string, not the original binary data

### Automatic File Detection
When joining files, the tool:
1. Extracts the base name from the selected file
2. Searches the same directory for all files matching the pattern `basename_XXX.txt`
3. Sorts files numerically by their sequence number
4. Reads and concatenates them in order
5. Decodes the combined Base64 string back to the original file

## Requirements

- Python 3.6+
- tkinter (usually included with Python)
- pathlib
- base64 (standard library)

## Error Handling

The tool includes comprehensive error handling for:
- Invalid file selections
- File access permissions
- Base64 decoding errors
- Incorrect split file patterns
- Missing related files during join operations

## Logging

All operations are logged with detailed information including:
- File processing status
- Number of chunks created/found
- Success/failure messages
- Error details

## Examples

### Example 1: Basic Encoding
```python
from pathlib import Path
from base64_encode_decode import encode_path

# Encode a binary file
encode_path(Path("document.pdf"), Path("document.txt"))
```

### Example 2: Split Large File
```python
from pathlib import Path
from base64_encode_decode import encode_and_split_path

# Split a large file into 5MB chunks
num_chunks = encode_and_split_path(
    Path("large_video.mp4"),
    Path("video_split.txt"),  # This becomes video_split_001.txt, etc.
    5  # 5MB chunks
)
print(f"Created {num_chunks} chunks")
```

### Example 3: Join and Decode
```python
from pathlib import Path
from base64_encode_decode import join_and_decode_path

# Join split files and decode (select any split file)
join_and_decode_path(
    Path("video_split_005.txt"),  # Any split file works
    Path("restored_video.mp4")
)
```

## Notes

- Split files must be in the same directory for automatic detection
- The tool preserves file metadata during encoding/decoding
- Large files may take time to process depending on system resources
- Base64 encoding increases file size by approximately 33%