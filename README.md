# Electrolyte Report Converter

A modern desktop application for converting CRM CSV reports from electrolyte manufacturing into formatted Excel (XLSX) outputs.

## Features

- Modern, user-friendly interface with dark/light theme
- Drag-and-drop support for CSV files
- Batch processing capability
- Progress tracking
- Conversion history logging
- Professional Excel output formatting

## Installation

1. Ensure you have Python 3.9+ installed
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python app.py
   ```

2. Use the application:
   - Drag and drop CSV files onto the application window, or use the "Browse CSV Files" button
   - Click "Convert to XLSX (Batch)" to process the files
   - Select an output directory for the converted files
   - Monitor conversion progress in the status bar
   - View conversion history in the "History" tab

## Data Format

The application expects CRM CSV files with columns such as:
- Case Number
- SLA
- Customer Name
- Customer Phone
- Street, City, State/Province, Zip/Postal Code
- Customer Complaint
- Product Description
- LineItem Status (e.g., New, Completed, etc.)
- Technician Name
- Technician Remarks
- Created Date

The output Excel file will be formatted with:
- Proper column headers
- Center alignment
- Text wrapping
- Auto-adjusted column widths
- Professional styling

## What NOT to Upload to GitHub

**Do NOT upload these files/folders:**
- `__pycache__/`, `*.pyc` files
- `build/`, `dist/`, and all their contents
- `conversion_logs.db` or any `.db`/`.sqlite` files
- Any `.csv`, `.xlsx`, or `.zip` files in the `electrolyte/` folder (these are data/output, not code)
- All binaries in `electrolyte/Feed_Electrolyte-*/Feed_Electrolyte/Debug/` (e.g., `.exe`, `.dll`, `.pdb`, `.xml`, `.config`)
- Large files like `context.txt`, `context.exe`, or any build artifacts

**Recommended:** Add a `.gitignore` file to enforce this.

## What to Upload
- All `.py` source code files
- `requirements.txt`
- `README.md`
- `Electrolyte_Report_Converter_Requirements.md`
- The `assets/` folder (for your logo and static images)

## Contributing

Pull requests are welcome! Please:
- Exclude all data, binaries, and build artifacts from your commits
- Follow PEP8 style for Python code
- Update the README if you add new features

## License

MIT License 