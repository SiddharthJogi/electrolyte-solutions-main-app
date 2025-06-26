# Electrolyte Report Converter

**This is a simple internal tool for a single user. The codebase is intentionally kept simple and easy to understand.**

A modern desktop application for converting CRM CSV reports from electrolyte manufacturing into formatted Excel (XLSX) outputs.

## Features

- Simple, user-friendly interface (no tooltips, splash screens, or multi-user features)
- Drag-and-drop support for CSV files
- Batch processing capability
- Progress tracking
- Conversion history logging
- Professional Excel output formatting

## Installation

1. Ensure you have Python 3.9+ installed
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the app:
   ```
   python app.py
   ```

## Usage

- Click "Browse CSV Files" to select one or more CRM CSV files
- Click "Convert to XLSX (Batch)" to generate Excel reports
- View conversion history in the "History" tab
- Use the "Dark Mode" switch for a dark theme

## Expected CSV Format

The input CSV should be exported from your CRM and include columns like:
- Case Number
- SLA
- Customer Name
- Customer Phone
- Street, City, State/Province, Zip/Postal Code
- Customer Complaint
- Product Description
- LineItem Status
- Technician Name
- Technician Remarks
- Created Date

## What to Include/Exclude in the Repo

**Include:**
- `app.py`, `gui.py`, `converter.py`, `database.py`, `requirements.txt`, `README.md`, `config.ini`, `assets/`

**Do NOT include:**
- `__pycache__/`, `*.pyc`, `build/`, `dist/`, `*.db`, `*.csv`, `*.xlsx`, `*.zip`, binaries in `electrolyte/Feed_Electrolyte.../Debug/`

## Contributing

This project is for internal use by a single user. If you need to make changes, keep the code as simple and readable as possible.

## Assets

The logo is in the `assets/` folder.

## Production/Distribution

If you want to create a standalone `.exe` for easy use, see below.

---

# Creating a Standalone .exe

You can use [PyInstaller](https://pyinstaller.org/) to create a Windows executable:

```
pip install pyinstaller
pyinstaller --onefile --windowed app.py
```

The `.exe` will appear in the `dist/` folder. Copy the `assets/` folder alongside the `.exe` for the logo to work.

---

**Note:** This app is for internal use only. Do not upload data files, binaries, or build artifacts to GitHub. 