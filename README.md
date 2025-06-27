# Electrolyte CRM - Multi-Company Dashboard

A modern, cross-platform desktop tool for managing technician performance, file-based reporting, and daily operations for multiple companies (Usha, Symphony, Orient, Atomberg).

## Features
- Clean, intuitive UI with company selector and dashboards
- Role-based authentication (main_admin, admin, user)
- Company-specific dashboards: Daily Task, Performance Dashboard, Feedback Call, Salary Counter, File Processing
- File processing integration (including Atomberg logic)
- Shared SQLite database for all app data
- Responsive, accessible design

## Setup
1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   cd electrolyte-solutions-main-app-main
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   python app.py
   ```

## Usage
- Log in with your assigned username and password.
- Select your company and access the relevant dashboard sections.
- Only admins/main_admins can access all features; users have limited access.

## Security
**Do not commit real user credentials or passwords to public repositories.**

## License
MIT License 