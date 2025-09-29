# Installation Guide

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the V2 implementation:**
   ```bash
   python3 main2.py
   ```

3. **Compare V1 vs V2:**
   ```bash
   python3 compare.py
   ```

## Dependencies

### Required Packages
- **pandas** (>=1.5.0): Data manipulation and analysis
- **openpyxl** (>=3.0.0): Primary Excel file processing engine
- **xlrd** (>=2.0.0): Fallback Excel file processing engine
- **typing-extensions** (>=4.0.0): Enhanced type hints support

### Optional Packages
- **xlsxwriter**: For writing Excel files (future feature)
- **pytest**: For running tests
- **black**: For code formatting
- **flake8**: For code linting
- **mypy**: For type checking

## Python Version
- **Python 3.7+** required
- **Python 3.9+** recommended for best performance

## Installation Methods

### Using pip
```bash
pip install -r requirements.txt
```

### Using conda
```bash
conda install pandas openpyxl xlrd typing-extensions
```

### Using virtual environment (recommended)
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Verification

After installation, verify everything works:
```bash
python3 -c "from access_profile_automation_v2 import OpenAccessProfilesXLSXV2; print('✅ V2 implementation ready!')"
```
