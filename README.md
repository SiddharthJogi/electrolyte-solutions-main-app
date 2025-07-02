# Electrolyte CRM Dashboard

## Overview
This application is a comprehensive dashboard for managing and processing company data, including advanced file conversion and reporting tools for Atomberg and Orient. Symphony and Usha dashboards are coming soon.

## Features
- Multi-company dashboard (Usha, Symphony, Orient, Atomberg)
- Atomberg: Three advanced file processing tools:
  - **General File Conversion**: Converts CSV to styled Excel with SLA and pivot table
  - **Feed Remark**: Adds technician remarks to processed files
  - **VOC-VOT Remark**: Adds customer feedback/remarks to processed files
- Orient: ZIP-to-Excel processing with VLOOKUP and pivot table support
- Daily task management (Atomberg only)
- Performance dashboard, feedback call, salary counter (Atomberg only)
- Symphony/Usha: Dashboards coming soon

## Setup Instructions

### 1. Install Python dependencies
```sh
pip install -r requirements.txt
```

### 2. Run the Application
```sh
python gui.py
```

### 3. Windows .exe Fallback
- For Atomberg (General File Conversion) and Orient, `.exe` files are provided for Windows users without Python. Double-click the `.exe` to run the tool.
- Feed Remark and VOC-VOT Remark require Python.

## How to Use
1. Launch the app and log in.
2. Select a company:
   - **Atomberg**: Choose a processing type and process your CSV file.
   - **Orient**: Select a ZIP file containing a CSV and follow the prompts.
   - **Symphony/Usha**: Dashboards will show "Coming soon".
3. Processed files are saved in the `output/` folder.

## Platform Notes
- **Mac**: All features work natively using Python scripts (except Excel COM automation for VLOOKUP, which is Windows-only).
- **Windows**: `.exe` fallback available for Atomberg (General) and Orient. For advanced features, install Python and dependencies.

## Requirements
- Python 3.8+
- See `requirements.txt` for dependencies
- For Windows: `.exe` fallback for Atomberg (General) and Orient
- For VLOOKUP automation: Microsoft Excel (Windows only)

## Troubleshooting
- If you encounter errors, check the console output for details.
- For VLOOKUP features, ensure your lookup Excel files are formatted as required.
- If using `.exe` and nothing happens, try running from the command prompt to see error messages.

## .gitignore Example
```
__pycache__/
.venv/
*.pyc
output/
*.log
*.exe
.DS_Store
.idea/
```

## Credits
- Atomberg File Conversion Logic: Shrey, Ronit, Aaryan, Shakti (SARS)
- Feed Remark & VOC-VOT Remark: Saniya, Prem, Atharva, Manaswi (SPAM)
- Orient: (Your Orient script authors)

---
For any issues, please contact the project maintainer. 