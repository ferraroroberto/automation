# Pull Request: Fix Google Photos scope handling

## 🎯 Purpose
Ensure the weekly photo automation reauthenticates when stored credentials lack required Google API scopes, preventing "insufficient authentication scopes" errors.

## 🔄 Changes Made
- Check for missing OAuth scopes in `weekly_photo_automation.py` and trigger reauthorization when needed.

## 🧪 Testing Instructions

### **Prerequisites**
- Python 3.8+
- Git

### **Environment Setup Commands**

**Windows PowerShell:**
```powershell
# 1. Clone repo and checkout specific branch into new folder
git clone -b <branch-name> <repo-url> <folder-name>

# 2. Navigate to the folder
cd <folder-name>

# 3. Create virtual environment
python -m venv .venv

# 4. Activate virtual environment
.\.venv\Scripts\Activate.ps1

# 5. Install ad-hoc requirements
pip install -r requirements.txt
```

**Unix/Linux/macOS:**
```bash
# 1. Clone repo and checkout specific branch into new folder
git clone -b <branch-name> <repo-url> <folder-name>

# 2. Navigate to the folder
cd <folder-name>

# 3. Create virtual environment
python3 -m venv .venv

# 4. Activate virtual environment
source .venv/bin/activate

# 5. Install ad-hoc requirements
pip install -r requirements.txt
```

### **Test Execution**
1. Run unit tests:
   ```bash
   pytest -q
   ```

## 📋 Requirements for Testing
Minimal dependencies are listed in `requirements.txt`.

## 🚨 Breaking Changes
None.

## 📚 Additional Notes
N/A
