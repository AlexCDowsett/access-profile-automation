# Access Profile Automation V1

A Python script for parsing Excel access profile data and converting it to a structured dictionary format.

## Overview

This implementation processes Excel files containing access profile information, extracting categories, headings, and data pairs to create a nested dictionary structure for further processing.

## Features

- **Excel Processing**: Reads `.xlsx` files using openpyxl
- **Data Structure**: Creates nested dictionaries with categories, headings, and values
- **Progress Tracking**: Built-in progress bar for processing status
- **JSON Export**: Converts data to JSON format for external use

## Installation

### Prerequisites
- Python 3.6 or higher
- openpyxl library

### Install Dependencies
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage
```bash
python3 main.py
```

### Programmatic Usage
```python
from access_profile_automation import OpenAccessProfilesXLSX

# Initialize parser
parser = OpenAccessProfilesXLSX()

# Access the data
profiles = parser.access_profile_dict
filters = parser.filter

# Convert to JSON
json_data = parser.to_json()
```

## File Structure

- **`access_profile_automation.py`**: Main implementation class
- **`main.py`**: Example usage script
- **`AccessProfilesTEAM.xlsx`**: Sample Excel data file
- **`requirements.txt`**: Python dependencies

## Data Structure

The parser creates a nested dictionary structure:
```
{
  "Profile Name": {
    "Category": {
      "Heading": ["operator", "value"]
    }
  }
}
```

## Excel File Format

The Excel file should contain:
1. **Row 1**: Categories (Conductor, storm Contact, UC, etc.)
2. **Row 2**: Headings (Actions, Announcements, Menus, etc.)
3. **Row 3**: Column headers
4. **Row 4+**: Data rows with profile names, filters, and operator/value pairs

## Example Output

```json
{
  "TEAM - AZDE_BER_KS-SachB_PS_1": {
    "UC": {
      "User Groups": ["equals", "AZDE_BER_KS-SachB_PS_1"]
    }
  }
}
```

## Limitations

- Requires specific Excel file format
- Limited error handling
- Basic data validation
- Single-threaded processing

## Enhanced Version

For a more robust implementation with comprehensive error handling, logging, and improved structure, see the `backup/` folder which contains the V2 implementation.

## Requirements

- **openpyxl** (>=3.0.0): Excel file processing

## License

This project is for internal use and processing of access profile data.
