# Automation Project Rules & Standards

*This document defines coding standards, project structure, and organizational principles for maintaining consistency and quality across the codebase.*

## 📁 Project Structure

### Directory Organization
- **Separate concerns**: Each major functionality gets its own directory
- **Clear naming**: Use descriptive, lowercase names with underscores
- **Hierarchical structure**: Group related functionality together
- **Configuration separation**: Keep all config files in dedicated `config/` directories

### File Naming Conventions
- **Python files**: `snake_case.py` (e.g., `email_automation_classify.py`)
- **SQL files**: Descriptive names with underscores (e.g., `contactabilidad_join.sql`)
- **Configuration**: `.json` or `.py` extensions
- **Documentation**: `.md` extension for markdown files
- **Batch files**: `.bat` extension for Windows scripts

### Code Naming Conventions
- **Functions & Variables**: `snake_case` (e.g., `load_config`, `api_config`)
- **Constants**: `UPPER_SNAKE_CASE` (e.g., `MAX_RETRIES`)
- **Classes**: `PascalCase` (e.g., `DataProcessor`)

## 🐍 Python Standards

### Import Organization
```python
# 1. Standard library imports
# 2. Third-party library imports  
# 3. Local/relative imports
```

### Type Hints
- **Always use type hints** for function parameters and return values
- **Use Optional[]** for parameters that can be None
- **Use Union[]** for parameters that can be multiple types
- **Use Dict[str, Any]** for flexible dictionary parameters
- **Use List[]** for list parameters

### Function Structure
- Follow PEP 20 (Zen of Python)
- Use docstrings with clear parameter descriptions
- Implement graceful error handling with retry logic
- Log errors with context and attempt information

### Logging Standards
Configure logging at module level, no persistent file logging by default.

```python
logger.debug("📂 Loading configuration file")
logger.info("✅ Configuration loaded successfully")
logger.error(f"❌ Error: Configuration file not found at {config_path}")
logger.info(f"🔍 Fetching {platform_key} data")
logger.info(f"⏳ Retrying in {wait_time} seconds...")
```

## ⚙️ Configuration Management

### Configuration Rules
1. **JSON-based configuration** for flexibility and readability
2. **No hardcoding**: Use config files for files, paths, and settings
3. **Environment variables**: Use `.env` files for sensitive data
4. **Default values**: Provide defaults for all configuration options
5. **Validation**: Validate configuration values on load

### Configuration Structure
- **Simple modules**: JSON config with same name as Python module, in same folder
- **Complex modules**: JSON in dedicated `config/` folder within project
- **Sensitive data**: Store in `.env` files, never commit to version control

## 🗄️ Database & SQL Standards

### SQL File Organization
- **Process files**: Use descriptive names indicating business process
- **Temporary tables**: Prefix with `TMP_` (e.g., `TMP_ROB_001_MAPPING`)
- **Export tables**: Suffix with `_EXP` (e.g., `TABLE_NAME_EXP`)
- **Documentation**: Create corresponding `.md` files for complex processes

### SQL Coding Standards
```sql
-- Use descriptive table aliases: WITH t1_with_canonical AS (...)
-- Use COALESCE for handling NULL values in joins
-- use UPPERCASE for table and field names, lowercase for functions
```

### BigQuery Standards
- **Project references**: Use full project IDs in table references
- **Schema fields**: Define explicit schemas for data loading
- **Partitioning**: Use appropriate partitioning strategies for large tables
- **Clustering**: Implement clustering for frequently queried columns

## ☁️ Cloud Automation

### Google Cloud Platform
```python
def connex_bigquery(project_id, location):
    try:
        client = bigquery.Client(project=project_id, location=location)
        print(f"Connected to BigQuery project: {project_id} in region {location}")
        return client
    except Exception as error:
        print("Error connecting to BigQuery:", error)
        raise

# Configuration constants
PROJECT_ID = "your-project-id"
BIGQUERY_REGION = "europe-southwest1"
EXPORT_FOLDER = "/path/to/exports"
```

## 📊 Data Processing

### Pandas & Excel
- Use explicit data types and handle missing values
- Preserve formatting and column widths in Excel exports
- Implement data validation and error handling

## 🔧 Utilities & Helpers

### File Operations
```python
def ensure_folder_exists(folder_path):
def get_file_extension(file_path):
def is_valid_file(file_path, allowed_extensions):
```

### Command Line Interface
- **Always use argparse**: Besides JSON config, use argparse for everything else
- **Default values**: Set defaults so `__main__` function always works
- **Debug flag**: Always support `--debug` command line flag

## 📝 Documentation

### Code Comments
- **Function headers**: Use docstrings with clear parameter descriptions
- **Complex logic**: Explain business logic and algorithms
- **TODO comments**: Mark areas that need future attention
- **Inline comments**: Explain non-obvious code sections

### README Structure
```markdown
# Module Name

Brief description of the module's purpose.

## 🚀 Overview
Detailed explanation of functionality.

## 📋 Usage
Code examples and usage instructions.

## 🔧 Configuration
Configuration options and examples.

## 📊 Output
Description of output format and structure.
```

### Documentation Rules
1. **Emoji headers**: Use emojis for visual section identification
2. **Code examples**: Provide practical usage examples
3. **Configuration docs**: Document all configuration options
4. **Output format**: Describe data structures and formats
5. **Troubleshooting**: Include common issues and solutions

## 🚀 Deployment & Execution

### Script Execution
```python
if __name__ == "__main__":
    try:
        config = load_config()
        main(config)
        print("✅ Script completed successfully")
    except Exception as e:
        print(f"❌ Script failed: {e}")
        sys.exit(1)
```

### Error Handling & Debug
- **Graceful degradation**: Handle errors without crashing
- **User feedback**: Provide clear error messages and recovery suggestions
- **Logging**: Log all errors for debugging
- **Exit codes**: Use appropriate exit codes for automation
- **Debug mode**: Provide detailed information, error traces, and performance metrics

## 🔒 Security & Best Practices

### Sensitive Information
- **Never commit**: API keys, passwords, or sensitive data
- **Configuration files**: Use JSON or environment variables
- **Git ignore**: Exclude sensitive files from version control
- **Parameter files**: Store outside of code directories

### Data Validation
```python
def validate_input_parameters(params, required_keys):
    """Validate that all required parameters are present."""
    missing_keys = [key for key in required_keys if key not in params]
    if missing_keys:
        raise ValueError(f"Missing required parameters: {missing_keys}")
    return True
```

### Dependency Management
1. **Version pinning**: Pin exact versions for reproducibility
2. **Purpose comments**: Document why each dependency is needed
3. **Minimal dependencies**: Only include necessary packages
4. **Security updates**: Regularly update dependencies for security patches
5. **Virtual environments**: Use virtual environments for isolation

## 📋 Planning Before Code Changes

### Mandatory Planning Phase

**Before making any code changes, you MUST outline a detailed plan that includes:**

1. **Change Summary**: Clear description of what needs to be accomplished
2. **Files to Modify**: Specific list of files that will be changed, including:
   - Full file paths
   - Type of modification (create, modify, delete)
   - Brief description of what will change in each file
3. **Dependencies**: Any new packages, imports, or external resources needed
4. **Testing Strategy**: How the changes will be validated
5. **Risk Assessment**: Potential impacts on existing functionality

### Scope Limitation Rule

**ONLY implement what is explicitly requested:**

- **No unnecessary improvements**: Do not add features, optimizations, or enhancements unless specifically asked
- **No excessive debugging**: Do not add debug logs, error handling, or validation beyond what's requested
- **Ask before expanding**: If you believe additional changes would be beneficial, present them as suggestions and wait for approval
- **Stick to the brief**: Implement exactly what was requested, nothing more, nothing less

### Plan Validation Process

**The user must validate the plan before any code changes proceed:**

- Present the complete plan in a structured format
- Wait for explicit user approval before implementing changes
- If the plan is rejected or needs modification, update and re-present
- Only proceed with implementation after plan approval

### Plan Documentation Format

```markdown
## 📋 Implementation Plan

### 🎯 Objective
[Clear description of what needs to be accomplished]

### 📁 Files to Modify
- `path/to/file1.py` - [Type: create/modify/delete] - [Description of change]
- `path/to/file2.py` - [Type: create/modify/delete] - [Description of change]

### 🔗 Dependencies
- [List any new packages or imports needed]

### 🧪 Testing Strategy
- [How changes will be validated]

### ⚠️ Risk Assessment
- [Potential impacts on existing functionality]
```

## 🔧 Code Maintenance & Pull Request Process

### Pull Request Documentation Requirements

When creating a pull request, **always** include a `PULL_REQUEST.md` file in the root of your changes. Use the `PULL_REQUEST_TEMPLATE.md` in the project root as a reference.

### Required Documentation Sections

Your PULL_REQUEST.md must include:
- **Purpose**: Clear explanation of what the PR accomplishes
- **Changes Made**: List of specific changes and files modified
- **Testing Instructions**: Complete environment setup commands for both Windows PowerShell and Unix/Linux systems
- **Requirements**: Minimal dependencies needed for testing
- **Breaking Changes**: Any changes affecting existing functionality

### Environment Setup Requirements

The PULL_REQUEST.md must include commands for:
- Clone repo and checkout specific branch into new folder (single command)
- Create virtual environment (.venv)
- Activate virtual environment
- Install ad-hoc requirements