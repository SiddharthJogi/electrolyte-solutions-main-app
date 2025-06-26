
# 📄 Project Requirements Document  
**Project Title:** Electrolyte Report Converter  
**Prepared For:** Development Assistants  
**Objective:** Build a modern desktop application that converts raw electrolyte `.csv` report files into well-formatted `.xlsx` output files.  
**Target Users:** Interns and engineers managing production/report data.

---

## 🧠 Overview  
The application should accept CSV report files, extract only necessary fields, format them into an Excel sheet exactly like the demo file, and log all activity into a local database. The tool must offer a user-friendly GUI and allow file selections dynamically.

---

## ✅ Core Functional Requirements

### 1. 📥 Input Handling
- Users should be able to **browse and select** `.csv` files.
- Application must **remember the last used input path** and suggest it as default.
- CSV parsing should be done using a robust library like **pandas**.
- Allow support for **missing or extra columns** gracefully — only process those that match required keys.

### 2. 📤 Output Generation
- Output must be saved as an `.xlsx` file with the following:
  - Only key fields (e.g. `Timestamp`, `Voltage`, `Current`, `Temperature`)
  - Renamed columns to user-friendly headers (e.g., `Timestamp → Date`)
  - A single sheet titled **“Report”**
  - All cells **center-aligned** and **wrap text** enabled
- User should select the save location, with default path from last usage.

### 3. 🖥️ Graphical User Interface (GUI)
- Use **PyQt6** or similar modern GUI framework
- GUI should have:
  - File input field with browse button
  - Output path selector (save as dialog)
  - Convert button
  - Status label or message boxes
- GUI must be clean and responsive (fixed size acceptable for now)

---

## 🗄️ Data Logging Requirements

### 4. 🗃️ Local Database Logging
- Use **SQLite** to store conversion history
- Required table: `logs`
  - Fields:
    - ID (auto-increment)
    - Filename
    - Converted_At (timestamp)
    - Status (`Success` or error message)
    - Row Count
- Log every attempt, even if conversion fails

---

## ⚙️ Configuration and State Persistence

### 5. ⚙️ Config File
- Use a `.ini` file to store:
  - Last used CSV path
  - Last used output path
- Read these on app launch and write after every conversion

---

## 🧪 Error Handling
- Show user-friendly error messages if:
  - File not found
  - Output location invalid
  - Required columns are missing
- Errors should also be logged in the database with the status

---

## 🪄 Nice-to-Have Features (Optional Phase 2)
- Drag-and-drop support for CSV files
- Multi-file batch conversion
- Live preview of data before export
- Theming: dark/light mode toggle
- In-app history viewer from the database
- Export logs as Excel file
- Auto-format Excel columns (width adjustment, filters)

---

## 💾 Executable Packaging
- The final product must be delivered as a **Windows `.exe`** using **PyInstaller**
- The app should:
  - Not require Python installation
  - Be fully standalone
  - Include a custom icon (optional)

---

## 🧱 Suggested Tools / Libraries
| Purpose            | Tool / Library       |
|--------------------|----------------------|
| GUI                | PyQt6                |
| CSV/Excel handling | pandas, openpyxl     |
| Database           | sqlite3              |
| Config             | configparser         |
| Packaging          | PyInstaller          |

---

## 📂 Folder Structure Example
```
ElectrolyteConverter/
├── app.py
├── converter.py         # CSV → XLSX logic
├── database.py          # SQLite setup and logging
├── gui.py               # UI components
├── config.ini
├── conversion_logs.db
├── resources/
│   └── icon.ico
└── dist/                # Final .exe after packaging
```

---

## 📌 Deliverables
- Working `.exe` file
- All Python source files
- `README.md` with instructions
- `conversion_logs.db` with sample logs
- `config.ini` auto-generated
