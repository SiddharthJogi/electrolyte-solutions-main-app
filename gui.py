import sys
import os
import configparser
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QStatusBar, QFrame, QSizePolicy, QSpacerItem, QMessageBox, QListWidget, QListWidgetItem, QSplashScreen, QCheckBox, QProgressBar, QDialog, QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit, QTextEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox, QGroupBox, QFormLayout, QStackedWidget, QGridLayout, QScrollArea, QGraphicsDropShadowEffect
)
from PyQt6.QtGui import QPixmap, QIcon, QPainter, QColor, QBrush, QAction, QFont, QPalette
from PyQt6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, QThread, pyqtSignal, QDate
from database import log_conversion, setup_database, verify_user, get_companies, get_daily_tasks, add_daily_task, get_performance_summary, add_performance_log, get_feedback_calls, add_feedback_call, get_salary_data, add_salary_data, log_file_processing, get_file_logs, USERS
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import threading
from datetime import datetime
import pandas as pd

CONFIG_PATH = "config.ini"
LOGO_PATH = "assets/electrolye logo.png"

# Company-specific color schemes
COMPANY_COLORS = {
    'Usha': {
        'primary': '#FF6B35',
        'secondary': '#FF8A65',
        'accent': '#FF5722',
        'background': '#FFF3E0',
        'text': '#3E2723'
    },
    'Symphony': {
        'primary': '#4ECDC4',
        'secondary': '#81C784',
        'accent': '#26A69A',
        'background': '#E8F5E8',
        'text': '#1B5E20'
    },
    'Orient': {
        'primary': '#45B7D1',
        'secondary': '#64B5F6',
        'accent': '#1976D2',
        'background': '#E3F2FD',
        'text': '#0D47A1'
    },
    'Atomberg': {
        'primary': '#FFD93D',
        'secondary': '#FFE082',
        'accent': '#FFC107',
        'background': '#FFFDE7',
        'text': '#F57F17'
    }
}

class SplashScreen(QSplashScreen):
    def __init__(self):
        pixmap = QPixmap(LOGO_PATH).scaledToHeight(180, Qt.TransformationMode.SmoothTransformation)
        bg = QPixmap(pixmap.width() + 40, pixmap.height() + 60)
        bg.fill(QColor('#FFFFFF'))
        painter = QPainter(bg)
        painter.drawPixmap(20, 20, pixmap)
        painter.end()
        super().__init__(bg)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setStyleSheet("background: #FFFFFF; border-radius: 20px;")
        self.showMessage("Loading Electrolyte CRM Tool...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, QColor('#2c3e50'))

class LoginPage(QWidget):
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        self.user_data = None
        self.init_ui()

    def init_ui(self):
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFFFFF, stop:1 #F5F7FA);
            }
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(40)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Logo
        logo_label = QLabel()
        logo_pixmap = QPixmap(LOGO_PATH)
        logo_label.setPixmap(logo_pixmap.scaledToHeight(120, Qt.TransformationMode.SmoothTransformation))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label)
        # Title
        title_label = QLabel("Welcome to Electrolyte CRM")
        title_label.setFont(QFont("Arial", 30, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin: 20px 0;")
        layout.addWidget(title_label)
        # Subtitle
        subtitle_label = QLabel("Please login to continue")
        subtitle_label.setFont(QFont("Arial", 15))
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle_label.setStyleSheet("color: #7f8c8d; margin-bottom: 30px;")
        layout.addWidget(subtitle_label)
        # Centered input area
        input_area = QWidget()
        input_layout = QVBoxLayout(input_area)
        input_layout.setSpacing(18)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_area.setMaximumWidth(420)
        input_area.setMinimumWidth(320)
        # Username
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("Enter your username")
        self.username_edit.setMinimumHeight(44)
        self.username_edit.setStyleSheet("""
            QLineEdit {
                padding: 12px 18px;
                border: 2px solid #1976D2;
                border-radius: 10px;
                font-size: 16px;
                background: white;
                color: #2c3e50;
                font-weight: 500;
            }
            QLineEdit:focus {
                border-color: #0057B8;
                background: #FAFAFA;
            }
        """)
        input_layout.addWidget(self.username_edit)
        # Password
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Enter your password")
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_edit.setMinimumHeight(44)
        self.password_edit.setStyleSheet("""
            QLineEdit {
                padding: 12px 18px;
                border: 2px solid #1976D2;
                border-radius: 10px;
                font-size: 16px;
                background: white;
                color: #2c3e50;
                font-weight: 500;
            }
            QLineEdit:focus {
                border-color: #0057B8;
                background: #FAFAFA;
            }
        """)
        input_layout.addWidget(self.password_edit)
        # Login button
        self.login_button = QPushButton("Login")
        self.login_button.setMinimumHeight(44)
        self.login_button.setStyleSheet("""
            QPushButton {
                background: #0057B8;
                color: #fff;
                border: none;
                padding: 12px 0px;
                border-radius: 10px;
                font-size: 18px;
                font-weight: bold;
                margin-top: 10px;
            }
            QPushButton:hover {
                background: #003974;
            }
            QPushButton:pressed {
                background: #002244;
            }
        """)
        input_layout.addWidget(self.login_button)
        self.login_button.clicked.connect(self.login)
        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #e74c3c; font-size: 14px; margin-top: 15px; font-weight: bold;")
        input_layout.addWidget(self.status_label)
        layout.addWidget(input_area, alignment=Qt.AlignmentFlag.AlignCenter)
        self.username_edit.setFocus()
        self.username_edit.returnPressed.connect(self.login)
        self.password_edit.returnPressed.connect(self.login)

    def login(self):
        username = self.username_edit.text().strip()
        password = self.password_edit.text().strip()
        if not username or not password:
            self.status_label.setText("Please enter both username and password")
            return
        self.user_data = verify_user(username, password)
        if self.user_data:
            self.main_app.user_data = self.user_data
            self.main_app.show_company_selector()
        else:
            self.status_label.setText("Invalid username or password")
            self.password_edit.clear()
            self.password_edit.setFocus()

class CompanySelector(QWidget):
    def __init__(self, user_data, main_app, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.main_app = main_app  # Store reference to main app
        self.selected_company = None
        self.init_ui()
        
    def init_ui(self):
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFFFFF, stop:1 #F5F7FA);
            }
        """)
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        # Header
        header_layout = QHBoxLayout()
        header_layout.setSpacing(24)
        header_layout.setContentsMargins(32, 32, 32, 0)
        logo_label = QLabel()
        logo_pixmap = QPixmap(LOGO_PATH)
        logo_label.setPixmap(logo_pixmap.scaledToHeight(72, Qt.TransformationMode.SmoothTransformation))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        header_layout.addWidget(logo_label)
        title_label = QLabel("Select Your Company")
        title_label.setFont(QFont("Arial", 32, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #22314a; margin-left: 24px; background: transparent;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        user_label = QLabel(f"👤 {self.user_data['username']} ({self.user_data['role']})")
        user_label.setStyleSheet("color: #7f8c8d; font-size: 16px; font-weight: bold;")
        header_layout.addWidget(user_label)
        logout_button = QPushButton("Logout")
        logout_button.setStyleSheet("""
            QPushButton {
                background: #e74c3c;
                color: white;
                border: none;
                padding: 12px 28px;
                border-radius: 12px;
                font-size: 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #c0392b;
            }
        """)
        logout_button.setMinimumWidth(100)
        logout_button.clicked.connect(self.logout)
        header_layout.addWidget(logout_button, alignment=Qt.AlignmentFlag.AlignRight)
        main_layout.addLayout(header_layout)
        main_layout.addSpacing(32)
        # Company grid in a scroll area for responsiveness
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        grid_container = QWidget()
        grid_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        grid_layout = QGridLayout(grid_container)
        grid_layout.setSpacing(18)
        grid_layout.setContentsMargins(32, 0, 32, 32)
        companies = get_companies()
        if not companies:
            error_label = QLabel("No companies found in database. Please check the database setup.")
            error_label.setStyleSheet("color: #e74c3c; font-size: 16px; font-weight: bold; text-align: center;")
            main_layout.addWidget(error_label)
            return
        n = len(companies)
        cols = 2
        rows = (n + cols - 1) // cols
        for i, company in enumerate(companies):
            row = i // cols
            col = i % cols
            button = self.create_company_button(company)
            button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
            grid_layout.addWidget(button, row, col)
        for r in range(rows):
            grid_layout.setRowStretch(r, 1)
        for c in range(cols):
            grid_layout.setColumnStretch(c, 1)
        grid_container.setLayout(grid_layout)
        scroll.setWidget(grid_container)
        main_layout.addWidget(scroll, stretch=1)

    def create_company_button(self, company):
        button = QPushButton()
        button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        colors = COMPANY_COLORS.get(company['name'], COMPANY_COLORS['Usha'])
        button.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 white, stop:1 {colors['background']});
                border: 3px solid {colors['primary']};
                border-radius: 16px;
                padding: 0px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {colors['primary']}, stop:1 {colors['secondary']});
                border-color: {colors['accent']};
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {colors['secondary']}, stop:1 {colors['primary']});
            }}
        """)
        button_layout = QVBoxLayout(button)
        button_layout.setSpacing(16)
        button_layout.setContentsMargins(0, 24, 0, 24)
        # Logo
        logo_label = QLabel()
        logo_path = company['logo_path']
        if company['name'] == 'Atomberg':
            logo_path = 'assets/business-atomb-list-logo.png'
        logo_pixmap = QPixmap(logo_path)
        logo_label.setPixmap(logo_pixmap.scaledToHeight(80, Qt.TransformationMode.SmoothTransformation))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        button_layout.addWidget(logo_label, alignment=Qt.AlignmentFlag.AlignCenter)
        # Company name
        name_label = QLabel(company['name'])
        name_label.setFont(QFont("Arial", 22, QFont.Weight.Bold))
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_label.setStyleSheet(f"color: {colors['primary']};")
        button_layout.addWidget(name_label, alignment=Qt.AlignmentFlag.AlignCenter)
        button.clicked.connect(lambda checked, name=company['name']: self.select_company(name))
        return button
        
    def select_company(self, company_name):
        self.selected_company = company_name
        if hasattr(self.main_app, 'on_company_selected'):
            self.main_app.on_company_selected(company_name)
            
    def logout(self):
        # Use the main app reference directly
        if hasattr(self.main_app, 'logout'):
            self.main_app.logout()

class CompanyDashboard(QWidget):
    def __init__(self, company_name, user_data, main_app, parent=None):
        super().__init__(parent)
        self.company_name = company_name
        self.user_data = user_data
        self.main_app = main_app  # Store reference to main app
        self.colors = COMPANY_COLORS.get(company_name, COMPANY_COLORS['Usha'])
        self.init_ui()
        
    def init_ui(self):
        self.setStyleSheet(f"""
            QWidget {{
                background: {self.colors['background']};
                color: {self.colors['text']};
            }}
        """)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        center_widget = QWidget()
        center_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        center_layout = QVBoxLayout(center_widget)
        center_layout.setContentsMargins(10, 10, 10, 10)
        center_layout.setSpacing(6)
        # Header
        header_layout = QHBoxLayout()
        companies = get_companies()
        company_info = next((c for c in companies if c['name'] == self.company_name), None)
        if company_info:
            logo_label = QLabel()
            logo_path = company_info['logo_path']
            if self.company_name == 'Atomberg':
                logo_path = 'assets/business-atomb-list-logo.png'
            logo_pixmap = QPixmap(logo_path)
            logo_label.setPixmap(logo_pixmap.scaledToHeight(32, Qt.TransformationMode.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            header_layout.addWidget(logo_label)
            company_label = QLabel(self.company_name)
            company_label.setFont(QFont("Arial", 18, QFont.Weight.Bold))
            company_label.setStyleSheet(f"color: {self.colors['primary']}; margin-left: 8px;")
            header_layout.addWidget(company_label)
        header_layout.addStretch()
        user_label = QLabel(f"👤 {self.user_data['username']} ({self.user_data['role']})")
        user_label.setStyleSheet("color: #7f8c8d; font-size: 13px; font-weight: bold;")
        header_layout.addWidget(user_label)
        logout_button = QPushButton("Logout")
        logout_button.setStyleSheet("""
            QPushButton {
                background: #e74c3c;
                color: white;
                border: none;
                padding: 8px 18px;
                border-radius: 8px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #c0392b;
            }
        """)
        logout_button.clicked.connect(self.logout)
        header_layout.addWidget(logout_button)
        center_layout.addLayout(header_layout)
        # Dashboard grid in a scroll area for responsiveness
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        grid_container = QWidget()
        grid_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        dashboard_layout = QGridLayout(grid_container)
        dashboard_layout.setSpacing(10)
        dashboard_layout.setContentsMargins(0, 0, 0, 0)
        # 3x2 grid, File Processing first
        sections = []
        if self.company_name == 'Symphony' or self.company_name == 'Usha':
            coming_soon_label = QLabel("Dashboard for this company is coming soon.")
            coming_soon_label.setFont(QFont("Arial", 18, QFont.Weight.Bold))
            coming_soon_label.setStyleSheet("color: #888; margin: 40px;")
            coming_soon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            center_layout.addWidget(coming_soon_label)
        else:
            # Only show three cards: Performance Dashboard, Daily Task, Feedback Calling
            if self.user_data['role'] in ['main_admin', 'admin']:
                sections.append(self.create_dashboard_card("Performance Dashboard", "View technician performance metrics", "performance"))
                sections.append(self.create_dashboard_card("Daily Task", "Feed Remark and VOC-VOT Remark processing", "daily_task"))
                sections.append(self.create_dashboard_card("Feedback Call", "Manage customer feedback calls", "feedback"))
            cols = 3
            for idx, card in enumerate(sections):
                card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
                row = idx // cols
                col = idx % cols
                dashboard_layout.addWidget(card, row, col)
            for r in range((len(sections) + cols - 1) // cols):
                dashboard_layout.setRowStretch(r, 1)
            for c in range(cols):
                dashboard_layout.setColumnStretch(c, 1)
            grid_container.setLayout(dashboard_layout)
            scroll.setWidget(grid_container)
            center_layout.addWidget(scroll, stretch=1)
        back_button = QPushButton('← Back to Company Selector')
        back_button.setStyleSheet('''
            QPushButton {
                background: #222;
                color: #fff;
                border: none;
                padding: 10px 24px;
                border-radius: 8px;
                font-size: 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #FFD93D;
                color: #222;
            }
        ''')
        back_button.setFixedWidth(260)
        back_button.clicked.connect(self.back_to_company_selection)
        center_layout.addWidget(back_button, alignment=Qt.AlignmentFlag.AlignHCenter)
        main_layout.addWidget(center_widget)
        
    def create_dashboard_card(self, title, description, section_type):
        card = QFrame()
        card.setFrameStyle(QFrame.Shape.StyledPanel)
        card.setStyleSheet(f"""
            QFrame {{
                background: white;
                border: 2px solid {self.colors['primary']};
                border-radius: 14px;
                padding: 8px;
                margin: 4px;
            }}
            QFrame:hover {{
                border-color: {self.colors['accent']};
                background: {self.colors['background']};
            }}
        """)
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(card)
        layout.setSpacing(6)
        layout.setContentsMargins(8, 8, 8, 8)
        # Title
        title_label = QLabel(title)
        title_label.setFont(QFont("Arial", 15, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {self.colors['primary']};")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        # Description
        desc_label = QLabel(description)
        desc_label.setFont(QFont("Arial", 11))
        desc_label.setStyleSheet("color: #444;")
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(desc_label)
        # Action button
        action_button = QPushButton("Open")
        action_button.setStyleSheet(f"""
            QPushButton {{
                background: #222;
                color: #fff;
                border: none;
                padding: 8px 18px;
                border-radius: 8px;
                font-size: 13px;
                font-weight: bold;
                margin-top: 6px;
            }}
            QPushButton:hover {{
                background: {self.colors['primary']};
                color: #fff;
            }}
        """)
        action_button.clicked.connect(lambda: self.open_section(section_type))
        action_button.setToolTip(f"Open the {title} section")
        layout.addWidget(action_button)
        card.setToolTip(description)
        return card
    
    def open_section(self, section_type):
        if section_type == "daily_task":
            self.open_daily_tasks()
        elif section_type == "performance":
            self.open_performance_dashboard()
        elif section_type == "feedback":
            self.open_feedback_calls()
        else:
            QMessageBox.information(self, "Coming Soon", f"The {section_type.replace('_', ' ').title()} section will be implemented soon!")
    
    def open_daily_tasks(self):
        dialog = DailyTasksDialog(self.company_name, self.user_data, self)
        dialog.exec()
    
    def open_performance_dashboard(self):
        dialog = PerformanceDialog(self.company_name, self)
        dialog.exec()
    
    def open_feedback_calls(self):
        dialog = FeedbackDialog(self.company_name, self)
        dialog.exec()
    
    def logout(self):
        # Use the main app reference directly
        if hasattr(self.main_app, 'logout'):
            self.main_app.logout()
    
    def back_to_company_selection(self):
        # Use the main app reference directly
        if hasattr(self.main_app, 'show_company_selector'):
            self.main_app.show_company_selector()

class ElectrolyteCRMApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.user_data = None
        self.current_company = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Electrolyte CRM - Multi-Company Dashboard")
        self.setMinimumSize(900, 600)
        self.setWindowIcon(QIcon(LOGO_PATH))
        self.setStyleSheet("""
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFFFFF, stop:1 #F5F7FA);
            }
        """)
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        self.show_login()

    def show_login(self):
        while self.stacked_widget.count() > 0:
            widget = self.stacked_widget.widget(0)
            self.stacked_widget.removeWidget(widget)
            widget.deleteLater()
        login_page = LoginPage(self)
        self.stacked_widget.addWidget(login_page)
        self.stacked_widget.setCurrentWidget(login_page)

    def show_company_selector(self):
        while self.stacked_widget.count() > 0:
            widget = self.stacked_widget.widget(0)
            self.stacked_widget.removeWidget(widget)
            widget.deleteLater()
        company_selector = CompanySelector(self.user_data, self, self)
        self.stacked_widget.addWidget(company_selector)
        self.stacked_widget.setCurrentWidget(company_selector)

    def on_company_selected(self, company_name):
        self.current_company = company_name
        while self.stacked_widget.count() > 0:
            widget = self.stacked_widget.widget(0)
            self.stacked_widget.removeWidget(widget)
            widget.deleteLater()
        dashboard = CompanyDashboard(company_name, self.user_data, self, self)
        self.stacked_widget.addWidget(dashboard)
        self.stacked_widget.setCurrentWidget(dashboard)

    def logout(self):
        self.user_data = None
        self.current_company = None
        self.show_login()

# Legacy classes for backward compatibility
class CSVFileHandler(FileSystemEventHandler):
    def __init__(self, converter_app):
        self.converter_app = converter_app
        self.output_folder = ""
        
    def set_output_folder(self, folder):
        self.output_folder = folder
        
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith('.csv'):
            time.sleep(2)
            if os.path.exists(event.src_path):
                self.converter_app.auto_convert_file(event.src_path)

class FileWatcherThread(QThread):
    status_updated = pyqtSignal(str)
    
    def __init__(self, input_folder, output_folder, converter_app):
        super().__init__()
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.converter_app = converter_app
        self.observer = None
        self.running = False
        
    def run(self):
        try:
            self.observer = Observer()
            event_handler = CSVFileHandler(self.converter_app)
            event_handler.set_output_folder(self.output_folder)
            self.observer.schedule(event_handler, self.input_folder, recursive=False)
            self.observer.start()
            self.running = True
            self.status_updated.emit("Monitoring started")
            
            while self.running:
                time.sleep(1)
                
        except Exception as e:
            self.status_updated.emit(f"Error: {str(e)}")
            
    def stop(self):
        self.running = False
        if self.observer:
            self.observer.stop()
            self.observer.join()
        self.status_updated.emit("Monitoring stopped")

class AnimatedTabWidget(QTabWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.currentChanged.connect(self.animate_tab)
        self.anim = QPropertyAnimation(self, b"windowOpacity")
        self.anim.setDuration(300)
        self.anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        
    def animate_tab(self, idx):
        self.setWindowOpacity(0.7)
        self.anim.stop()
        self.anim.setStartValue(0.7)
        self.anim.setEndValue(1.0)
        self.anim.start()

# Legacy ConverterApp class for backward compatibility
class ConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_config()
        self.file_watcher_thread = None
        self.auto_convert_enabled = False
        self.init_ui()

    def init_config(self):
        self.config = configparser.ConfigParser()
        self.config.read(CONFIG_PATH)
        if 'paths' not in self.config:
            self.config['paths'] = {}
        self.last_csv_path = self.config['paths'].get('last_csv_path', "")
        self.last_output_path = self.config['paths'].get('last_output_path', "")
        if 'auto_convert' not in self.config:
            self.config['auto_convert'] = {}
        self.auto_input_folder = self.config['auto_convert'].get('input_folder', "")
        self.auto_output_folder = self.config['auto_convert'].get('output_folder', "")

    def save_config(self):
        self.config['paths']['last_csv_path'] = self.last_csv_path
        self.config['paths']['last_output_path'] = self.last_output_path
        self.config['auto_convert']['input_folder'] = self.auto_input_folder
        self.config['auto_convert']['output_folder'] = self.auto_output_folder
        with open(CONFIG_PATH, 'w') as f:
            self.config.write(f)

    def init_ui(self):
        # Legacy UI implementation - kept for backward compatibility
        self.setWindowTitle("Electrolyte CRM Report Tool")
        self.setMinimumSize(800, 600)
        self.setWindowIcon(QIcon(LOGO_PATH))
        
        # Create the new CRM app instead
        self.crm_app = ElectrolyteCRMApp()
        self.setCentralWidget(self.crm_app.stacked_widget)
        
    def browse_files(self):
        pass  # Legacy method
        
    def remove_selected_files(self):
        pass  # Legacy method
        
    def convert_files(self):
        pass  # Legacy method
        
    def auto_convert_file(self, csv_path):
        pass  # Legacy method

# Dashboard Section Dialogs
class FileProcessingDialog(QDialog):
    def __init__(self, company_name, user_data, parent=None):
        super().__init__(parent)
        self.company_name = company_name
        self.user_data = user_data
        self.colors = COMPANY_COLORS.get(company_name, COMPANY_COLORS['Usha'])
        self.init_ui()
        self.processing_type = 'General'
        if self.company_name == 'Atomberg':
            self.processing_type = 'General'
        # For Orient, no processing type needed

    def init_ui(self):
        self.setWindowTitle(f"{self.company_name} - File Processing")
        self.setFixedSize(600, 550)
        self.setStyleSheet(f"""
            QDialog {{
                background: {self.colors['background']};
                border-radius: 15px;
            }}
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        title_label = QLabel("📁 File Processing")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {self.colors['primary']};")
        layout.addWidget(title_label)
        # Atomberg: Add processing type selector
        if self.company_name == 'Atomberg':
            from PyQt6.QtWidgets import QComboBox
            self.type_selector = QComboBox()
            self.type_selector.addItems(["General File Conversion", "Feed Remark", "VOC-VOT Remark"])
            self.type_selector.setCurrentIndex(0)
            self.type_selector.setStyleSheet(f"color: {self.colors['primary']}; font-weight: bold; font-size: 15px;")
            self.type_selector.currentIndexChanged.connect(self.on_type_changed)
            layout.addWidget(QLabel("Select Processing Type:"))
            layout.addWidget(self.type_selector)
        elif self.company_name == 'Orient':
            info_label = QLabel("This tool processes ZIP files containing a CSV and generates a formatted Excel report. Optionally, you can add VLOOKUPs for REMARKS and SO_NUMBER.")
            info_label.setWordWrap(True)
            info_label.setStyleSheet("color: #444; font-size: 14px;")
            layout.addWidget(info_label)
        # File selection
        file_group = QGroupBox("Select File to Process")
        file_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                color: {self.colors['primary']};
                border: 2px solid {self.colors['primary']};
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
        """)
        file_layout = QVBoxLayout(file_group)
        self.file_path_label = QLabel("No file selected")
        self.file_path_label.setStyleSheet("""
            QLabel {
                padding: 10px;
                background: white;
                border: 2px solid #E0E0E0;
                border-radius: 8px;
                color: #666;
            }
        """)
        file_layout.addWidget(self.file_path_label)
        browse_button = QPushButton("Browse File")
        browse_button.setStyleSheet(f"""
            QPushButton {{
                background: {self.colors['primary']};
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 8px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background: {self.colors['accent']};
            }}
        """)
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)
        layout.addWidget(file_group)
        self.process_button = QPushButton("Process File")
        self.process_button.setStyleSheet(f"""
            QPushButton {{
                background: #27ae60;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 10px;
                font-size: 16px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background: #229954;
            }}
            QPushButton:disabled {{
                background: #bdc3c7;
            }}
        """)
        self.process_button.clicked.connect(self.process_file)
        self.process_button.setEnabled(False)
        layout.addWidget(self.process_button)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 2px solid {self.colors['primary']};
                border-radius: 8px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background: {self.colors['primary']};
                border-radius: 6px;
            }}
        """)
        layout.addWidget(self.progress_bar)
        self.status_label = QLabel("Ready to process files")
        self.status_label.setStyleSheet("color: #27ae60; font-weight: bold;")
        layout.addWidget(self.status_label)
        history_group = QGroupBox("Processing History")
        history_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                color: {self.colors['primary']};
                border: 2px solid {self.colors['primary']};
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
        """)
        history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget()
        self.history_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background: white;
            }
        """)
        history_layout.addWidget(self.history_list)
        layout.addWidget(history_group)
        self.load_history()
        
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "", "CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path_label.setText(os.path.basename(file_path))
            self.selected_file_path = file_path
            self.process_button.setEnabled(True)
            
    def on_type_changed(self, idx):
        if idx == 0:
            self.processing_type = 'General'
        elif idx == 1:
            self.processing_type = 'Feed_Remark'
        elif idx == 2:
            self.processing_type = 'VOC-VOT_Remark'
    
    def process_file(self):
        if not hasattr(self, 'selected_file_path'):
            return
        self.process_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("Processing file...")
        self.status_label.setStyleSheet("color: #f39c12; font-weight: bold;")
        processing_type = getattr(self, 'processing_type', 'General')
        self.processor_thread = FileProcessorThread(
            self.selected_file_path, self.company_name, self.user_data['username'], processing_type=processing_type
        )
        self.processor_thread.finished.connect(self.on_processing_finished)
        self.processor_thread.start()
        
    def on_processing_finished(self, success, message):
        self.progress_bar.setVisible(False)
        self.process_button.setEnabled(True)

        # Check for fallback/exe usage in logs (simple approach: check for a marker file or log)
        fallback_marker = 'output/atomberg_fallback_used.txt'
        exe_marker = 'output/atomberg_exe_used.txt'
        feedback_msg = None
        if os.path.exists(fallback_marker):
            feedback_msg = "Note: Atomberg fallback processing was used. Some advanced features may not be available."
            os.remove(fallback_marker)
        elif os.path.exists(exe_marker):
            feedback_msg = "Note: Atomberg main.exe was used for processing."
            os.remove(exe_marker)

        if success:
            self.status_label.setText("File processed successfully!")
            self.status_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            msg = "File processed successfully!"
            if feedback_msg:
                msg += f"\n\n{feedback_msg}"
            QMessageBox.information(self, "Success", msg)
        else:
            self.status_label.setText("Processing failed")
            self.status_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
            msg = f"Processing failed: {message}"
            if feedback_msg:
                msg += f"\n\n{feedback_msg}"
            QMessageBox.warning(self, "Error", msg)

        self.load_history()
        
    def load_history(self):
        self.history_list.clear()
        logs = get_file_logs(self.company_name)
        
        for log in logs[:10]:  # Show last 10 logs
            item_text = f"{log[4]} - {log[2]} ({log[3]})"
            if log[6]:  # error message
                item_text += f" - Error: {log[6]}"
            self.history_list.addItem(item_text)

class FileProcessorThread(QThread):
    finished = pyqtSignal(bool, str)
    
    def __init__(self, file_path, company_name, processed_by, processing_type='General'):
        super().__init__()
        self.file_path = file_path
        self.company_name = company_name
        self.processed_by = processed_by
        self.processing_type = processing_type
        
    def run(self):
        try:
            log_file_processing(
                self.company_name,
                os.path.basename(self.file_path),
                "csv",
                "processing",
                processed_by=self.processed_by
            )
            if self.company_name == "Atomberg":
                success = self.process_atomberg_file()
            elif self.company_name == "Orient":
                success = self.process_orient_file()
            else:
                success = self.process_default_file()
            if success:
                log_file_processing(
                    self.company_name,
                    os.path.basename(self.file_path),
                    "csv",
                    "success",
                    output_path="output/",
                    processed_by=self.processed_by
                )
                self.finished.emit(True, "File processed successfully")
            else:
                log_file_processing(
                    self.company_name,
                    os.path.basename(self.file_path),
                    "csv",
                    "error",
                    error_message="Processing failed",
                    processed_by=self.processed_by
                )
                self.finished.emit(False, "Processing failed")
        except Exception as e:
            log_file_processing(
                self.company_name,
                os.path.basename(self.file_path),
                "csv",
                "error",
                error_message=str(e),
                processed_by=self.processed_by
            )
            self.finished.emit(False, str(e))
            
    def process_atomberg_file(self):
        import platform
        import traceback
        import sys
        import subprocess
        import importlib.util
        try:
            print(f"[DEBUG] Atomberg processing type: {self.processing_type}")
            sys.path.append('atomberg')
            script_map = {
                'General': ('file conversion logic', 'main', 'process_file_simple'),
                'Feed_Remark': ('Feed_Remark', 'main', 'process_file'),
                'VOC-VOT_Remark': ('VOC-VOT_Remark', 'main', 'process_file'),
            }
            folder, module_name, func_name = script_map.get(self.processing_type, script_map['General'])
            module_path = os.path.join('atomberg', folder, f'{module_name}.py')
            spec = importlib.util.spec_from_file_location(module_name, module_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            func = getattr(module, func_name)
            os.makedirs('output', exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"output/Atomberg_{self.processing_type}_Output_{timestamp}.xlsx"
            # For Feed_Remark and VOC-VOT_Remark, use their process_file signature
            if self.processing_type == 'General':
                func(self.file_path, output_filename)
            else:
                func(self.file_path, output_filename)
            print(f"[DEBUG] Atomberg {self.processing_type} processing complete. Output: {output_filename}")
            return True
        except Exception as e:
            print(f"[ERROR] Exception in Atomberg {self.processing_type} import/process: {e}")
            print(traceback.format_exc())
            return False
    
    def process_orient_file(self):
        import platform
        import traceback
        import sys
        import subprocess
        import importlib.util
        try:
            print(f"[DEBUG] Starting Orient file processing for: {self.file_path}")
            sys.path.append('orient')
            module_path = os.path.join('orient', 'orient.py')
            try:
                spec = importlib.util.spec_from_file_location('orient', module_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                func = getattr(module, 'main')
                # Simulate command-line input for orient.py (it uses tkinter dialogs for file selection)
                # Instead, just run the script as a subprocess for now
                # If on Windows and .exe exists, use orient.exe as fallback
                if platform.system() == "Windows" and os.path.exists(os.path.join('orient', 'orient.exe')):
                    exe_path = os.path.join('orient', 'orient.exe')
                    result = subprocess.run([exe_path], capture_output=True, text=True)
                    print(f"[DEBUG] orient.exe stdout: {result.stdout}")
                    print(f"[DEBUG] orient.exe stderr: {result.stderr}")
                    return result.returncode == 0
                else:
                    # Run the script as a subprocess (so tkinter dialogs work)
                    result = subprocess.run([sys.executable, module_path], capture_output=True, text=True)
                    print(f"[DEBUG] orient.py stdout: {result.stdout}")
                    print(f"[DEBUG] orient.py stderr: {result.stderr}")
                    return result.returncode == 0
            except Exception as e:
                print(f"[ERROR] Exception in Orient import/process: {e}")
                print(traceback.format_exc())
                return False
        except Exception as e:
            print(f"[ERROR] Error processing Orient file: {e}")
            print(traceback.format_exc())
            return False
    
    def process_default_file(self):
        try:
            # Default processing logic (placeholder)
            # This can be replaced with specific logic for other companies
            import pandas as pd
            
            # Read CSV
            df = pd.read_csv(self.file_path)
            
            # Create output directory
            os.makedirs('output', exist_ok=True)
            
            # Generate output filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"output/{self.company_name}_Output_{timestamp}.xlsx"
            
            # Save as Excel
            df.to_excel(output_filename, index=False)
            
            return True
        except Exception as e:
            print(f"Error processing default file: {e}")
            return False

class DailyTaskProcessorThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, task_type, file_path, company_name, processed_by, vlookup_enabled=False, lookup_file_path=None):
        super().__init__()
        self.task_type = task_type
        self.file_path = file_path
        self.company_name = company_name
        self.processed_by = processed_by
        self.vlookup_enabled = vlookup_enabled
        self.lookup_file_path = lookup_file_path

    def run(self):
        import importlib.util
        import os
        import sys
        from database import log_file_processing
        import traceback
        try:
            log_file_processing(
                self.company_name,
                os.path.basename(self.file_path),
                "csv",
                "processing",
                processed_by=self.processed_by
            )
            success = False
            if self.company_name == "Atomberg":
                sys.path.append('atomberg')
                script_map = {
                    'Feed_Remark': ('Feed_Remark', 'main', 'process_file'),
                    'VOC-VOT_Remark': ('VOC-VOT_Remark', 'main', 'process_file'),
                }
                folder, module_name, func_name = script_map.get(self.task_type, (None, None, None))
                if folder:
                    module_path = os.path.join('atomberg', folder, f'{module_name}.py')
                    spec = importlib.util.spec_from_file_location(module_name, module_path)
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    func = getattr(module, func_name)
                    output_filename = self.file_path.replace('.csv', '_output.xlsx')
                    if self.vlookup_enabled and self.lookup_file_path:
                        # Call process_file and then apply vlookup
                        func(self.file_path, output_filename)
                        # Use the public vlookup function if available
                        if hasattr(module, 'apply_vlookup_with_excel_com'):
                            module.apply_vlookup_with_excel_com(output_filename, self.lookup_file_path)
                        success = True
                    else:
                        func(self.file_path, output_filename)
                        success = True
            elif self.company_name == "Orient":
                # Similar logic for Orient if needed
                success = True  # Placeholder
            if success:
                log_file_processing(
                    self.company_name,
                    os.path.basename(self.file_path),
                    "csv",
                    "success",
                    output_path="output/",
                    processed_by=self.processed_by
                )
                self.finished.emit(True, "File processed successfully")
            else:
                log_file_processing(
                    self.company_name,
                    os.path.basename(self.file_path),
                    "csv",
                    "error",
                    error_message="Processing failed",
                    processed_by=self.processed_by
                )
                self.finished.emit(False, "Processing failed")
        except Exception as e:
            log_file_processing(
                self.company_name,
                os.path.basename(self.file_path),
                "csv",
                "error",
                error_message=str(e),
                processed_by=self.processed_by
            )
            self.finished.emit(False, str(e))

class DailyTasksDialog(QDialog):
    def __init__(self, company_name, user_data, parent=None):
        super().__init__(parent)
        self.company_name = company_name
        self.user_data = user_data
        self.colors = COMPANY_COLORS.get(company_name, COMPANY_COLORS['Usha'])
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(f"{self.company_name} - Daily Tasks")
        self.setFixedSize(800, 600)
        self.setStyleSheet(f"""
            QDialog {{
                background: {self.colors['background']};
                border-radius: 15px;
            }}
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        title_label = QLabel("📋 Daily Tasks Management")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {self.colors['primary']};")
        layout.addWidget(title_label)
        # --- Feed Remark Section ---
        feed_group = QGroupBox("Feed Remark Processing")
        feed_layout = QHBoxLayout(feed_group)
        self.feed_vlookup_checkbox = QCheckBox("Enable VLOOKUP")
        self.feed_vlookup_checkbox.setChecked(False)
        feed_layout.addWidget(self.feed_vlookup_checkbox)
        feed_btn = QPushButton("Process Feed Remark CSV")
        feed_btn.setStyleSheet(f"background: {self.colors['primary']}; color: white; font-weight: bold; padding: 10px 20px; border-radius: 8px;")
        feed_btn.clicked.connect(self.process_feed_remark)
        feed_layout.addWidget(feed_btn)
        layout.addWidget(feed_group)
        # --- VOC-VOT Remark Section ---
        voc_group = QGroupBox("VOC-VOT Remark Processing")
        voc_layout = QHBoxLayout(voc_group)
        self.voc_vlookup_checkbox = QCheckBox("Enable VLOOKUP")
        self.voc_vlookup_checkbox.setChecked(False)
        voc_layout.addWidget(self.voc_vlookup_checkbox)
        voc_btn = QPushButton("Process VOC-VOT Remark CSV")
        voc_btn.setStyleSheet(f"background: {self.colors['primary']}; color: white; font-weight: bold; padding: 10px 20px; border-radius: 8px;")
        voc_btn.clicked.connect(self.process_voc_vot_remark)
        voc_layout.addWidget(voc_btn)
        layout.addWidget(voc_group)
        # --- Orient Section (if company is Orient) ---
        if self.company_name == 'Orient':
            orient_group = QGroupBox("Orient ZIP Processing")
            orient_layout = QHBoxLayout(orient_group)
            orient_label = QLabel("Only ZIP files containing a single CSV are accepted.")
            orient_layout.addWidget(orient_label)
            orient_btn = QPushButton("Process Orient ZIP File")
            orient_btn.setStyleSheet(f"background: {self.colors['primary']}; color: white; font-weight: bold; padding: 10px 20px; border-radius: 8px;")
            orient_btn.clicked.connect(self.process_orient_zip)
            orient_layout.addWidget(orient_btn)
            layout.addWidget(orient_group)
        # --- History Section ---
        history_group = QGroupBox("Daily Task Processing History")
        history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget()
        self.history_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background: white;
            }
        """)
        history_layout.addWidget(self.history_list)
        layout.addWidget(history_group)
        self.load_history()
        # Spacer
        layout.addStretch(1)
        self.setLayout(layout)

    def process_feed_remark(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Feed Remark CSV File", "", "CSV Files (*.csv)")
        if not file_path:
            return
        vlookup_enabled = self.feed_vlookup_checkbox.isChecked()
        lookup_file_path = None
        if vlookup_enabled:
            lookup_file_path, _ = QFileDialog.getOpenFileName(self, "Select Lookup Excel File", "", "Excel Files (*.xlsx)")
            if not lookup_file_path:
                QMessageBox.warning(self, "VLOOKUP Cancelled", "No lookup file selected. VLOOKUP will be skipped.")
                vlookup_enabled = False
        self.run_processing_thread('Feed_Remark', file_path, vlookup_enabled, lookup_file_path)

    def process_voc_vot_remark(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select VOC-VOT Remark CSV File", "", "CSV Files (*.csv)")
        if not file_path:
            return
        vlookup_enabled = self.voc_vlookup_checkbox.isChecked()
        lookup_file_path = None
        if vlookup_enabled:
            lookup_file_path, _ = QFileDialog.getOpenFileName(self, "Select Lookup Excel File", "", "Excel Files (*.xlsx)")
            if not lookup_file_path:
                QMessageBox.warning(self, "VLOOKUP Cancelled", "No lookup file selected. VLOOKUP will be skipped.")
                vlookup_enabled = False
        self.run_processing_thread('VOC-VOT_Remark', file_path, vlookup_enabled, lookup_file_path)

    def process_orient_zip(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Orient ZIP File", "", "ZIP Files (*.zip)")
        if not file_path:
            return
        self.run_processing_thread('Orient', file_path, False, None)

    def run_processing_thread(self, task_type, file_path, vlookup_enabled, lookup_file_path):
        self.processor_thread = DailyTaskProcessorThread(
            task_type, file_path, self.company_name, self.user_data['username'], vlookup_enabled, lookup_file_path
        )
        self.processor_thread.finished.connect(self.on_processing_finished)
        self.processor_thread.start()

    def on_processing_finished(self, success, message):
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.warning(self, "Error", message)
        self.load_history()

    def load_history(self):
        from database import get_file_logs
        self.history_list.clear()
        logs = get_file_logs(self.company_name)
        # Only show logs for daily task conversions (Feed_Remark, VOC-VOT_Remark, Orient)
        for log in logs:
            if log[2].endswith('.csv') and (log[1] == self.company_name):
                # Optionally, filter further by filename or status if needed
                self.history_list.addItem(f"{log[4]} - {log[2]} ({log[3]})" + (f" - Error: {log[6]}" if log[6] else ""))

class PerformanceDialog(QDialog):
    def __init__(self, company_name, parent=None):
        super().__init__(parent)
        self.company_name = company_name
        self.colors = COMPANY_COLORS.get(company_name, COMPANY_COLORS['Usha'])
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle(f"{self.company_name} - Performance Dashboard")
        self.setFixedSize(700, 500)
        self.setStyleSheet(f"""
            QDialog {{
                background: {self.colors['background']};
                border-radius: 15px;
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Title
        title_label = QLabel("📊 Performance Dashboard")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {self.colors['primary']};")
        layout.addWidget(title_label)
        
        # Performance summary
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(3)
        self.summary_table.setHorizontalHeaderLabels([
            "Technician", "Total Activities", "Average Score"
        ])
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.summary_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background: white;
            }
            QHeaderView::section {
                background: #f8f9fa;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.summary_table)
        
        self.load_performance_data()
        
    def load_performance_data(self):
        summary = get_performance_summary(self.company_name)
        self.summary_table.setRowCount(len(summary))
        
        for i, record in enumerate(summary):
            self.summary_table.setItem(i, 0, QTableWidgetItem(record[0]))
            self.summary_table.setItem(i, 1, QTableWidgetItem(str(record[1])))
            self.summary_table.setItem(i, 2, QTableWidgetItem(f"{record[2]:.2f}" if record[2] else "N/A"))

class FeedbackDialog(QDialog):
    def __init__(self, company_name, parent=None):
        super().__init__(parent)
        self.company_name = company_name
        self.colors = COMPANY_COLORS.get(company_name, COMPANY_COLORS['Usha'])
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle(f"{self.company_name} - Feedback Call Management")
        self.setFixedSize(900, 600)
        self.setStyleSheet(f"""
            QDialog {{
                background: {self.colors['background']};
                border-radius: 15px;
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Title
        title_label = QLabel("📞 Feedback Call Management")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {self.colors['primary']};")
        layout.addWidget(title_label)
        
        # Feedback table
        self.feedback_table = QTableWidget()
        self.feedback_table.setColumnCount(7)
        self.feedback_table.setHorizontalHeaderLabels([
            "Customer", "Phone", "Date", "Type", "Details", "Technician", "Status"
        ])
        self.feedback_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.feedback_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background: white;
            }
            QHeaderView::section {
                background: #f8f9fa;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.feedback_table)
        
        self.load_feedback_data()
        
    def load_feedback_data(self):
        feedback = get_feedback_calls(self.company_name)
        self.feedback_table.setRowCount(len(feedback))
        
        for i, record in enumerate(feedback):
            self.feedback_table.setItem(i, 0, QTableWidgetItem(record[2] or ""))
            self.feedback_table.setItem(i, 1, QTableWidgetItem(record[3] or ""))
            self.feedback_table.setItem(i, 2, QTableWidgetItem(record[4]))
            self.feedback_table.setItem(i, 3, QTableWidgetItem(record[5] or ""))
            self.feedback_table.setItem(i, 4, QTableWidgetItem(record[6] or ""))
            self.feedback_table.setItem(i, 5, QTableWidgetItem(record[7] or ""))
            self.feedback_table.setItem(i, 6, QTableWidgetItem(record[8])) 