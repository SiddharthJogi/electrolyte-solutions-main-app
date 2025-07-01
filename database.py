import sqlite3
from datetime import datetime
import hashlib
import os

DB_PATH = "electrolyte_crm.db"

USERS = [
    {"username": "main_admin", "password_hash": hashlib.sha256("M@inAdm1n!23".encode()).hexdigest(), "role": "main_admin"},
    {"username": "admin1", "password_hash": hashlib.sha256("Adm1n#2024!".encode()).hexdigest(), "role": "admin"},
    {"username": "admin2", "password_hash": hashlib.sha256("Adm2n$2024!".encode()).hexdigest(), "role": "admin"},
    {"username": "user1", "password_hash": hashlib.sha256("Us3r!1@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user2", "password_hash": hashlib.sha256("Us3r!2@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user3", "password_hash": hashlib.sha256("Us3r!3@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user4", "password_hash": hashlib.sha256("Us3r!4@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user5", "password_hash": hashlib.sha256("Us3r!5@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user6", "password_hash": hashlib.sha256("Us3r!6@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user7", "password_hash": hashlib.sha256("Us3r!7@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user8", "password_hash": hashlib.sha256("Us3r!8@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user9", "password_hash": hashlib.sha256("Us3r!9@2024".encode()).hexdigest(), "role": "user"},
    {"username": "user10", "password_hash": hashlib.sha256("Us3r!10@2024".encode()).hexdigest(), "role": "user"},
]

def setup_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Users table for authentication
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('main_admin', 'admin', 'user')),
        company TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        is_active BOOLEAN DEFAULT 1
    );''')
    
    # Companies table
    cursor.execute('''CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        logo_path TEXT,
        color_theme TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );''')
    
    # Daily tasks table
    cursor.execute('''CREATE TABLE IF NOT EXISTS daily_tasks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company TEXT NOT NULL,
        task_title TEXT NOT NULL,
        task_description TEXT,
        assigned_to TEXT,
        assigned_by TEXT,
        status TEXT DEFAULT 'pending' CHECK(status IN ('pending', 'in_progress', 'completed', 'cancelled')),
        priority TEXT DEFAULT 'medium' CHECK(priority IN ('low', 'medium', 'high', 'urgent')),
        due_date TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    );''')
    
    # Performance logs table
    cursor.execute('''CREATE TABLE IF NOT EXISTS performance_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company TEXT NOT NULL,
        technician_name TEXT NOT NULL,
        activity_type TEXT NOT NULL,
        activity_details TEXT,
        performance_score REAL,
        date TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );''')
    
    # Feedback calls table
    cursor.execute('''CREATE TABLE IF NOT EXISTS feedback_calls (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company TEXT NOT NULL,
        customer_name TEXT,
        phone_number TEXT,
        call_date TEXT NOT NULL,
        feedback_type TEXT CHECK(feedback_type IN ('positive', 'negative', 'neutral')),
        feedback_details TEXT,
        technician_name TEXT,
        resolution_status TEXT DEFAULT 'open' CHECK(resolution_status IN ('open', 'resolved', 'escalated')),
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );''')
    
    # Salary data table
    cursor.execute('''CREATE TABLE IF NOT EXISTS salary_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company TEXT NOT NULL,
        technician_name TEXT NOT NULL,
        month TEXT NOT NULL,
        year INTEGER NOT NULL,
        base_salary REAL,
        performance_bonus REAL,
        total_salary REAL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );''')
    
    # File processing logs table
    cursor.execute('''CREATE TABLE IF NOT EXISTS file_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company TEXT NOT NULL,
        filename TEXT NOT NULL,
        file_type TEXT NOT NULL,
        processed_at TEXT DEFAULT CURRENT_TIMESTAMP,
        status TEXT CHECK(status IN ('success', 'error', 'processing')),
        output_path TEXT,
        error_message TEXT,
        processed_by TEXT
    );''')
    
    # Insert default main admin user
    cursor.execute('''INSERT OR IGNORE INTO users (username, password_hash, role, company) 
                      VALUES (?, ?, ?, ?)''', 
                   ('admin', hash_password('admin123'), 'main_admin', 'all'))
    
    # Insert default companies
    companies = [
        ('Usha', 'assets/usha logo.png', '#FF6B35'),
        ('Symphony', 'assets/symphony logo.jpg', '#4ECDC4'),
        ('Orient', 'assets/orient logo.png', '#45B7D1'),
        ('Atomberg', 'assets/business-atomb-list-logo.png', '#FFD93D')
    ]
    
    for company in companies:
        cursor.execute('''INSERT OR IGNORE INTO companies (name, logo_path, color_theme) 
                          VALUES (?, ?, ?)''', company)
    
    conn.commit()
    conn.close()

def hash_password(password):
    """Hash a password using SHA-256"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    for user in USERS:
        if user["username"] == username and user["password_hash"] == password_hash:
            return {"username": user["username"], "role": user["role"]}
    return None

def create_user(username, password, role, company, created_by):
    """Create a new user"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute('''INSERT INTO users (username, password_hash, role, company) 
                          VALUES (?, ?, ?, ?)''', 
                       (username, hash_password(password), role, company))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    finally:
        conn.close()
    
    return success

def get_all_users():
    """Get all users for admin management"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''SELECT id, username, role, company, created_at, is_active 
                      FROM users 
                      ORDER BY created_at DESC''')
    
    users = cursor.fetchall()
    conn.close()
    
    return [{
        'id': user[0],
        'username': user[1],
        'role': user[2],
        'company': user[3],
        'created_at': user[4],
        'is_active': user[5]
    } for user in users]

def get_companies():
    """Get all companies"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('SELECT name, logo_path, color_theme FROM companies ORDER BY name')
    companies = cursor.fetchall()
    conn.close()
    
    return [{
        'name': company[0],
        'logo_path': company[1],
        'color_theme': company[2]
    } for company in companies]

# Daily Tasks functions
def add_daily_task(company, task_title, task_description, assigned_to, assigned_by, priority, due_date):
    """Add a new daily task"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''INSERT INTO daily_tasks 
                      (company, task_title, task_description, assigned_to, assigned_by, priority, due_date)
                      VALUES (?, ?, ?, ?, ?, ?, ?)''',
                   (company, task_title, task_description, assigned_to, assigned_by, priority, due_date))
    
    conn.commit()
    conn.close()

def get_daily_tasks(company, status=None):
    """Get daily tasks for a company"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    if status:
        cursor.execute('''SELECT * FROM daily_tasks 
                          WHERE company = ? AND status = ?
                          ORDER BY created_at DESC''', (company, status))
    else:
        cursor.execute('''SELECT * FROM daily_tasks 
                          WHERE company = ?
                          ORDER BY created_at DESC''', (company,))
    
    tasks = cursor.fetchall()
    conn.close()
    
    return tasks

def update_task_status(task_id, status):
    """Update task status"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''UPDATE daily_tasks 
                      SET status = ?, updated_at = CURRENT_TIMESTAMP
                      WHERE id = ?''', (status, task_id))
    
    conn.commit()
    conn.close()

# Performance functions
def add_performance_log(company, technician_name, activity_type, activity_details, performance_score, date):
    """Add performance log"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''INSERT INTO performance_logs 
                      (company, technician_name, activity_type, activity_details, performance_score, date)
                      VALUES (?, ?, ?, ?, ?, ?)''',
                   (company, technician_name, activity_type, activity_details, performance_score, date))
    
    conn.commit()
    conn.close()

def get_performance_summary(company, start_date=None, end_date=None):
    """Get performance summary for a company"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    if start_date and end_date:
        cursor.execute('''SELECT technician_name, 
                          COUNT(*) as total_activities,
                          AVG(performance_score) as avg_score
                          FROM performance_logs 
                          WHERE company = ? AND date BETWEEN ? AND ?
                          GROUP BY technician_name''', (company, start_date, end_date))
    else:
        cursor.execute('''SELECT technician_name, 
                          COUNT(*) as total_activities,
                          AVG(performance_score) as avg_score
                          FROM performance_logs 
                          WHERE company = ?
                          GROUP BY technician_name''', (company,))
    
    summary = cursor.fetchall()
    conn.close()
    
    return summary

# Feedback functions
def add_feedback_call(company, customer_name, phone_number, call_date, feedback_type, feedback_details, technician_name):
    """Add feedback call record"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''INSERT INTO feedback_calls 
                      (company, customer_name, phone_number, call_date, feedback_type, feedback_details, technician_name)
                      VALUES (?, ?, ?, ?, ?, ?, ?)''',
                   (company, customer_name, phone_number, call_date, feedback_type, feedback_details, technician_name))
    
    conn.commit()
    conn.close()

def get_feedback_calls(company, status=None):
    """Get feedback calls for a company"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    if status:
        cursor.execute('''SELECT * FROM feedback_calls 
                          WHERE company = ? AND resolution_status = ?
                          ORDER BY call_date DESC''', (company, status))
    else:
        cursor.execute('''SELECT * FROM feedback_calls 
                          WHERE company = ?
                          ORDER BY call_date DESC''', (company,))
    
    calls = cursor.fetchall()
    conn.close()
    
    return calls

# Salary functions
def add_salary_data(company, technician_name, month, year, base_salary, performance_bonus):
    """Add salary data"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    total_salary = base_salary + performance_bonus
    
    cursor.execute('''INSERT INTO salary_data 
                      (company, technician_name, month, year, base_salary, performance_bonus, total_salary)
                      VALUES (?, ?, ?, ?, ?, ?, ?)''',
                   (company, technician_name, month, year, base_salary, performance_bonus, total_salary))
    
    conn.commit()
    conn.close()

def get_salary_data(company, month=None, year=None):
    """Get salary data for a company"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    if month and year:
        cursor.execute('''SELECT * FROM salary_data 
                          WHERE company = ? AND month = ? AND year = ?
                          ORDER BY technician_name''', (company, month, year))
    else:
        cursor.execute('''SELECT * FROM salary_data 
                          WHERE company = ?
                          ORDER BY year DESC, month DESC, technician_name''', (company,))
    
    data = cursor.fetchall()
    conn.close()
    
    return data

# File processing functions
def log_file_processing(company, filename, file_type, status, output_path=None, error_message=None, processed_by=None):
    """Log file processing activity"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''INSERT INTO file_logs 
                      (company, filename, file_type, status, output_path, error_message, processed_by)
                      VALUES (?, ?, ?, ?, ?, ?, ?)''',
                   (company, filename, file_type, status, output_path, error_message, processed_by))
    
    conn.commit()
    conn.close() 

def get_file_logs(company):
    """Get file processing logs for a company"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''SELECT * FROM file_logs 
                      WHERE company = ?
                      ORDER BY processed_at DESC''', (company,))
    
    logs = cursor.fetchall()
    conn.close()
    
    return logs

# Legacy function for backward compatibility
def log_conversion(filename, status, rows):
    """Legacy function for backward compatibility"""
    log_file_processing('legacy', filename, 'csv', status, error_message=f"Rows: {rows}") 