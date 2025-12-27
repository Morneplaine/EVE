# EVE Manufacturing Database - Setup Guide

## Virtual Environment Setup

A virtual environment has been created for this project. Follow these steps to activate it and install dependencies:

### Activate the Virtual Environment

**PowerShell:**
```powershell
.\venv\Scripts\Activate.ps1
```

**Command Prompt:**
```cmd
venv\Scripts\activate.bat
```

### Install Dependencies

Once the virtual environment is activated, install the required packages:

```bash
pip install -r requirements.txt
```

### Deactivate the Virtual Environment

When you're done working, you can deactivate the virtual environment:

```bash
deactivate
```

## Running the Script

After activating the virtual environment and installing dependencies, you can run the script:

```bash
python eve_manufacturing_database.py
```

## Dependencies

- pandas: Data manipulation and analysis
- openpyxl: Excel file reading/writing
- requests: HTTP library for API calls
- xlsxwriter: Excel file writing with formatting

