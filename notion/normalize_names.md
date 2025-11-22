# Notion Article Name Normalizer

## 🚀 Overview

Automatically normalizes article names in a Notion database to sentence case while preserving proper names, acronyms, and special tokens. This tool ensures consistent capitalization across your Notion articles while maintaining important formatting like company names, acronyms, and emphasis.

## 📋 Usage

### Basic Usage

```bash
# Normalize articles from the last 14 days
python normalize_names.py --days 14 --config normalize_names.json

# Preview changes without updating (dry run)
python normalize_names.py --days 7 --dry-run

# Test normalization on a specific string
python normalize_names.py --test "TEST ALL CAPS"
```

### Command Line Arguments

- `--days` (default: 14): Number of days to look back for created articles
- `--config` (default: normalize_names.json): Path to JSON configuration file
- `--debug`: Enable debug logging for detailed output
- `--test <string>`: Test mode - normalize a specific string without processing the database
- `--dry-run`: Preview changes without updating Notion

## 🔧 Configuration

The tool uses a JSON configuration file (`normalize_names.json`) with the following structure:

```json
{
  "notion_api_key": "${NOTION_API_TOKEN}",
  "database_id": "your-database-id",
  "use_spacy": false,
  "proper_name_whitelist": [
    "Apple", "Google", "Microsoft", "AI", "GPT-4"
  ],
  "special_cases": [
    "Moore's Law",
    "Parkinson's Law"
  ],
  "common_words": [
    "the", "and", "of", "for", "in", "on", "at"
  ],
  "common_words_with_punct": [
    "a", "an", "the", "and", "or", "but"
  ]
}
```

### Configuration Fields

- **notion_api_key**: Notion API token (can use `${NOTION_API_TOKEN}` env variable)
- **database_id**: Notion database ID to process
- **use_spacy**: Enable spaCy NLP for entity detection (optional, requires spaCy installation)
- **proper_name_whitelist**: List of proper names to preserve (companies, people, products)
- **special_cases**: Multi-word proper names that need special handling
- **common_words**: Common words that shouldn't be treated as proper names
- **common_words_with_punct**: Common words with punctuation for contraction handling

## 📊 Normalization Rules

The normalizer follows these rules in order:

### 1. Sentence Case
- First word is capitalized
- Subsequent words are lowercase unless they match preservation rules

### 2. ALL CAPS Preservation
- Words with 2+ uppercase letters are preserved as ALL CAPS
- Example: `"TEST ALL CAPS"` → `"TEST ALL CAPS"`

### 3. Proper Names
- Words in the `proper_name_whitelist` are preserved with correct capitalization
- Example: `"apple"` → `"Apple"` (if in whitelist)

### 4. Multi-Word Proper Names
- Multi-word proper names are matched and preserved
- Example: `"adam grant"` → `"Adam Grant"` (if "Adam Grant" in whitelist)

### 5. Sentence-Ending Punctuation
- Words ending with `.`, `?`, or `!` are normalized correctly
- The next word after sentence-ending punctuation is capitalized
- Example: `"Hello World. How Are You?"` → `"Hello world. How are you?"`

### 6. Acronyms
- Acronyms with periods (like `U.S.A.`) don't trigger sentence capitalization
- Example: `"U.S.A. Is Great"` → `"U.S.A. is great"` (not `"U.S.A. Is great"`)

### 7. Contractions
- Contractions are preserved with correct structure
- Example: `"Here's The Thing"` → `"Here's the thing"`

### 8. Pronoun "I"
- The pronoun "I" is always capitalized
- Example: `"i think"` → `"I think"`

### 9. Optional: spaCy NLP
- If enabled, uses spaCy to detect person names automatically
- Requires: `pip install spacy && python -m spacy download en_core_web_sm`

## 🔍 How It Works

### Processing Flow

1. **Query Notion Database**: Retrieves articles created in the specified time window
2. **Extract Article Names**: Extracts the "article" property (title type) from each page
3. **Normalize Each Name**: Applies normalization rules while preserving special cases
4. **Update Notion**: Updates the article property with the normalized name (unless dry-run)

### Normalization Algorithm

1. Split text into tokens (words)
2. Apply basic sentence case (capitalize first letter only)
3. For each token:
   - Extract sentence-ending punctuation (`.`, `?`, `!`)
   - Check if previous token ended with sentence punctuation → capitalize
   - Apply preservation rules (proper names, ALL CAPS, etc.)
   - Reattach sentence-ending punctuation
4. Join tokens back into normalized text

### Key Methods

- `_normalize_name()`: Main normalization method
- `_extract_sentence_punctuation()`: Separates sentence punctuation from acronyms
- `_should_capitalize_token()`: Determines if a token should be capitalized
- `_restore_token_capitalization()`: Applies preservation rules
- `_handle_punctuated_token()`: Handles contractions and other punctuation
- `_reconstruct_with_punctuation()`: Rebuilds tokens with normalized word parts

## 📝 Examples

### Basic Normalization
```
Input:  "HOW TO BUILD A STARTUP"
Output: "How to build a startup"
```

### Preserving Proper Names
```
Input:  "apple releases new iphone"
Output: "Apple releases new iPhone"  (if Apple and iPhone in whitelist)
```

### ALL CAPS Preservation
```
Input:  "TEST ALL CAPS. More Text Here"
Output: "TEST ALL CAPS. More text here"
```

### Sentence-Ending Punctuation
```
Input:  "Hello World. How Are You?"
Output: "Hello world. How are you?"
```

### Acronyms
```
Input:  "U.S.A. Is Great. Really It Is"
Output: "U.S.A. is great. Really it is"
```

### Contractions
```
Input:  "Here's The Thing. It's Important"
Output: "Here's the thing. It's important"
```

### Multi-Word Proper Names
```
Input:  "adam grant writes about leadership"
Output: "Adam Grant writes about leadership"  (if "Adam Grant" in whitelist)
```

## ⚠️ Important Notes

1. **API Credentials**: Requires Notion API token set in environment variable `NOTION_API_TOKEN` or in config file
2. **Database Property**: Expects an "article" property of type "title" in the Notion database
3. **Dry Run First**: Always use `--dry-run` first to preview changes
4. **Backup**: Consider backing up your Notion database before bulk updates
5. **Rate Limits**: Notion API has rate limits; the tool handles pagination automatically

## 🐛 Troubleshooting

### Common Issues

**Issue**: "Missing required configuration values"
- **Solution**: Ensure `NOTION_API_TOKEN` is set or `notion_api_key` is in config

**Issue**: "Page missing or invalid article property"
- **Solution**: Ensure your database has an "article" property of type "title"

**Issue**: "Failed to update page"
- **Solution**: Check API permissions and rate limits

**Issue**: Proper names not being preserved
- **Solution**: Add names to `proper_name_whitelist` in config file

## 📚 Dependencies

- `requests`: HTTP library for Notion API calls
- `python-dotenv`: Environment variable management
- `spacy` (optional): NLP library for entity detection

## 🔗 See Also

- [Notion API Documentation](https://developers.notion.com/)
- Configuration file: `normalize_names.json`

