# Journal Automation

Automatically processes Notion journal database entries and creates a consolidated weekly summary optimized for LLM analysis.

## 🚀 Overview

This automation tool:
- Connects directly to Notion API (no manual exports needed)
- Queries journal entries for a specific week (Monday to Sunday)
- Processes 12 text fields and 4 checkbox fields
- Creates a single consolidated text file with all journal data
- Includes frequency counters for duplicate entries
- Tracks daily practice checkboxes (e.g., "patience: 5 days out of 7")
- **Opens the output folder in Windows Explorer** when the run finishes so the file is visible immediately

The system intelligently handles:
- **Date range calculation**: Automatically gets previous week (Monday to Sunday)
- **Sunday handling**: If today is Sunday, uses today as end date
- **Frequency tracking**: Shows how many times an entry appears (e.g., "play together (3x)")
- **Checkbox counting**: Tracks daily practices across the week
- **Duplicate removal**: Consolidates repeated entries with counters

## 📋 Usage

### Basic Usage

```bash
# Run with default settings (previous week)
python journal_automation.py

# Or use the batch file
journal_automation.bat
```

### Date Adjustment

When prompted, enter a number to adjust the date:
- `0` = current week (default)
- `-7` = one week ago
- `7` = next week

### Example Output

```
📂 Loading configuration...
🔐 Loading environment variables...

🚀 Processing journal entries...
📅 Enter an integer number to apply a timedelta to the current date (default is 0): 
📅 Date range: 2026-01-13 to 2026-01-19
🔍 Querying Notion database...
📥 Retrieved 7 entries (total: 7)
📊 Retrieved 7 journal entries for date range
🔄 Processing field: GRATITUDE
🔄 Processing field: GOALS FOR TODAY
...
🔄 Processing checkboxes
📝 Writing consolidated output: 2026-01-13 to 2026-01-19-journal-consolidated.txt
✅ Saved: E://automation//notion-automation-files//journal-output//2026-01-13 to 2026-01-19-journal-consolidated.txt
📄 Total characters: 3500

✅ Journal processing completed!
📁 Output file: E://automation//notion-automation-files//journal-output//...
```

When the script finishes successfully, **the output folder opens in Windows Explorer** so you can see the new file right away.

## 🔧 Prerequisites

Before using this automation, you need:

1. **Python 3.7 or higher**
2. **Required packages**:
   ```bash
   pip install requests python-dotenv
   ```
3. **Notion API token** in `.env` file
4. **Notion integration** with access to your journal database

### Setting up Notion Integration

1. **Create a Notion Integration**:
   - Go to [Notion Integrations](https://www.notion.so/my-integrations)
   - Click "New integration"
   - Give it a name (e.g., "Journal Automation")
   - Select your workspace
   - Copy the "Internal Integration Token"

2. **Share Database with Integration**:
   - Open your journal database in Notion
   - Click "..." (three dots) → "Add connections"
   - Select your integration

3. **Add Token to .env**:
   - Create `.env` file in `E:\automation\automation\`
   - Add: `NOTION_API_TOKEN=your_token_here`

## 📊 Output Structure

The automation creates a consolidated text file with the following sections:

```
================================================================================
WEEKLY JOURNAL SUMMARY
Period: 2026-01-13 to 2026-01-19
================================================================================

--------------------------------------------------------------------------------
GRATITUDE
(Things I'm grateful for this week)
--------------------------------------------------------------------------------
time with Ana in the morning (3x)
play with Luca (2x)
exercise with Valentina

--------------------------------------------------------------------------------
GOALS FOR TODAY
(Daily goals and objectives)
--------------------------------------------------------------------------------
prepare posts and newsletter
play with Luca and Valentina

... [10 more field sections] ...

--------------------------------------------------------------------------------
DAILY PRACTICES
(Days practiced this week)
--------------------------------------------------------------------------------
self-centered: 7 days out of 7
patience: 5 days out of 7
last day to live: 6 days out of 7
share a moment: 4 days out of 7
```

## ⚙️ Configuration

### journal_automation.json

The configuration file defines:

```json
{
  "description": "Configuration for journal automation",
  "input": {
    "database_id": "884de9f611464ca69e6028a6cbbafc00",
    "database_url": "https://www.notion.so/...",
    "date_property": "Date"
  },
  "output": {
    "directory": "E://automation//notion-automation-files//journal-output//",
    "format": "single_file",
    "filename_pattern": "{start_date} to {end_date}-journal-consolidated.txt"
  },
  "fields": [
    {
      "name": "gratitude",
      "column": "gratitude",
      "split_comma": true,
      "section_title": "GRATITUDE",
      "description": "Things I'm grateful for this week"
    }
    // ... more fields
  ],
  "checkboxes": [
    {
      "column": "self-centered",
      "display_name": "self-centered"
    }
    // ... more checkboxes
  ]
}
```

### Field Configuration

Each field supports:
- `name`: Internal identifier
- `column`: Exact Notion property name (case-sensitive)
- `split_comma`: Whether to split entries on commas
- `section_title`: Display title in output
- `description`: Section description

### Checkbox Configuration

Each checkbox supports:
- `column`: Exact Notion checkbox property name
- `display_name`: Display name in output

## 📝 Notion Database Structure

### Required Properties

| Property Name | Type | Description |
|--------------|------|-------------|
| `Date` | Date | Entry date (used for filtering) |

### Text Properties (12 fields)

| Property Name | Type | Split Comma | Description |
|--------------|------|-------------|-------------|
| `gratitude` | Rich Text | Yes | Things you're grateful for |
| `goals for today` | Rich Text | No | Daily goals |
| `I will express` | Multi-select | Yes | Core values |
| `who I will support` | Rich Text | No | People to support |
| `help` | Rich Text | No | Who did I help |
| `initiate` | Rich Text | No | What did I initiate |
| `learning` | Rich Text | No | Personal learning |
| `personal` | Rich Text | No | Family activities |
| `struggle` | Rich Text | No | Challenges |
| `win` | Rich Text | No | Victories |
| `work` | Rich Text | No | Work activities |
| `emotion` | Multi-select | Yes | Emotional states |

### Checkbox Properties (4 fields)

| Property Name | Type | Description |
|--------------|------|-------------|
| `self-centered` | Checkbox | Self-centered awareness |
| `patience` | Checkbox | Patience practice |
| `last day to live` | Checkbox | Last day mindset |
| `share a moment` | Checkbox | Shared a moment |

## 🔍 How It Works

### Date Range Calculation

1. Gets current date (or adjusted date from user input)
2. Checks if today is Sunday:
   - **If Sunday**: Uses today as end date
   - **If not Sunday**: Uses previous Sunday as end date
3. Calculates start date: 6 days before end date (Monday)

### Data Processing

1. **Query Notion API**: Filters entries by date range
2. **Extract values**: Gets data from each property
3. **Count frequencies**: Tracks how many times each value appears
4. **Format output**: Adds frequency counters (e.g., "play together (3x)")
5. **Process checkboxes**: Counts checked days (e.g., "patience: 5 days out of 7")
6. **Generate file**: Creates consolidated text file

### Frequency Tracking

- Single occurrence: `play with Luca`
- Multiple occurrences: `play with Luca (3x)`
- Works for all text fields
- Preserves original order of first occurrence

### Checkbox Tracking

- Counts how many days each checkbox was checked
- Format: `checkbox_name: X days out of Y`
- Example: `patience: 5 days out of 7` (checked 5 out of 7 days)

## 🛠️ Troubleshooting

### "ModuleNotFoundError: No module named 'requests'"

**Solution**: Install required packages
```bash
pip install requests python-dotenv
```

### "NOTION_API_TOKEN not found in environment variables"

**Solution**: 
1. Create `.env` file in `E:\automation\automation\`
2. Add: `NOTION_API_TOKEN=your_token_here`
3. Get token from https://www.notion.so/my-integrations

### "No entries found for this date range"

**Possible causes**:
1. Date range doesn't match your journal entries
2. Date property name doesn't match (should be `Date`)
3. Notion integration doesn't have access to database

**Solutions**:
- Check Notion database for entries in calculated date range
- Verify date property name in database
- Share database with your Notion integration

### "Property 'XXX' not found"

**Solution**: Property name in config doesn't match Notion database
1. Open database in Notion
2. Check exact column name (case-sensitive)
3. Update `column` value in `journal_automation.json`

### Wrong Date Range

**Solution**: Adjust date when prompted
- Enter `-7` for one week ago
- Enter `0` for current week
- Enter `7` for next week

## 📁 Files

- `journal_automation.py` - Main script
- `journal_automation.json` - Configuration file
- `journal_automation.bat` - Windows launcher
- `journal_automation.md` - This documentation

## 🔄 Workflow

### Weekly Process

1. **Run script**: Double-click `journal_automation.bat` or run `python journal_automation.py`
2. **Adjust date**: Press Enter for default (previous week) or enter offset
3. **Wait**: Script queries Notion and processes data
4. **Get output**: The output folder opens in Explorer when done; the consolidated file is there
5. **Use with LLM**: Copy entire file to ChatGPT/Claude for weekly summary

### LLM Integration

The consolidated output is optimized for LLM analysis:
- Single file (easy to copy/paste)
- Structured sections (clear context)
- Frequency counters (shows patterns)
- Checkbox tracking (shows consistency)
- Clean formatting (readable by AI)

Example LLM prompt:
```
Based on the following weekly journal entries, provide a summary highlighting:
1. Key themes and patterns
2. Achievements and wins
3. Learning and growth areas
4. Challenges faced
5. Daily practice consistency
6. Recommendations for next week

[Paste journal content here]
```

## 💡 Tips

- **Run weekly**: Best on Monday mornings for previous week
- **Check output**: Review file before sending to LLM
- **Adjust fields**: Edit `journal_automation.json` to add/remove fields
- **Backup data**: Output files are timestamped and preserved
- **Track patterns**: Frequency counters show what matters most

## 🔐 Security

- **API token**: Stored in `.env` (gitignored, never committed)
- **Database ID**: Not sensitive (just a reference)
- **Output files**: Contain personal data (keep private)
- **Configuration**: Can be shared (no secrets)

## 📊 Benefits

- ✅ No manual exports from Notion
- ✅ Always up-to-date data
- ✅ Frequency tracking shows patterns
- ✅ Checkbox tracking shows consistency
- ✅ Single consolidated file
- ✅ LLM-ready format
- ✅ Automatic date calculation
- ✅ Output folder opens in Explorer when done
- ✅ Easy to customize

## 🔄 Version History

- **v2.1** (2026-03): Open output folder in Explorer after save; logging instead of print; code cleanup and docs
- **v2.0** (2026-01-18): Notion API integration, frequency counters, checkbox tracking
- **v1.0** (2023-10-08): Initial Excel-based version
