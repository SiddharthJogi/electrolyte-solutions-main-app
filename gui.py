import sys
import os
import configparser
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QStatusBar, QFrame, QSizePolicy, QSpacerItem, QMessageBox, QListWidget, QListWidgetItem, QSplashScreen, QCheckBox, QProgressBar, QDialog, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtGui import QPixmap, QIcon, QPainter, QColor, QBrush, QAction
from PyQt6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from converter import convert_csv_to_xlsx
from database import log_conversion

CONFIG_PATH = "config.ini"
LOGO_PATH = "assets/electrolye logo.png"
COMPANY_COLORS = {
    'yellow': '#FFDE00',
    'blue': '#6C9DFE',
    'dark': '#030408',
    'white': '#FFFFFF'
}

class SplashScreen(QSplashScreen):
    def __init__(self):
        pixmap = QPixmap(LOGO_PATH).scaledToHeight(180, Qt.TransformationMode.SmoothTransformation)
        # Create a white background pixmap
        bg = QPixmap(pixmap.width(), pixmap.height())
        bg.fill(QColor(COMPANY_COLORS['white']))
        painter = QPainter(bg)
        painter.drawPixmap(0, 0, pixmap)
        painter.end()
        super().__init__(bg)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setStyleSheet(f"background: {COMPANY_COLORS['white']};")
        self.showMessage("Loading Electrolyte CRM Tool...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, Qt.GlobalColor.black)

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

class ConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_config()
        self.init_ui()

    def init_config(self):
        self.config = configparser.ConfigParser()
        self.config.read(CONFIG_PATH)
        if 'paths' not in self.config:
            self.config['paths'] = {}
        self.last_csv_path = self.config['paths'].get('last_csv_path', "")
        self.last_output_path = self.config['paths'].get('last_output_path', "")

    def save_config(self):
        self.config['paths']['last_csv_path'] = self.last_csv_path
        self.config['paths']['last_output_path'] = self.last_output_path
        with open(CONFIG_PATH, 'w') as f:
            self.config.write(f)

    def init_ui(self):
        self.setWindowTitle("Electrolyte CRM Report Tool")
        self.setMinimumSize(900, 650)
        self.setWindowIcon(QIcon(LOGO_PATH))
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFFFFF, stop:1 #F5F7FA); color: {COMPANY_COLORS['dark']}; }}
            QPushButton {{ background: {COMPANY_COLORS['blue']}; color: {COMPANY_COLORS['white']}; border-radius: 16px; padding: 14px 32px; font-size: 18px; font-weight: 600; margin: 0 8px; border: 1px solid #b0b0b0; }}
            QPushButton:hover {{ background: {COMPANY_COLORS['yellow']}; color: {COMPANY_COLORS['dark']}; }}
            QLabel#logoLabel {{ background: transparent; margin-bottom: 12px; }}
            QTabWidget::pane {{ border: 2px solid {COMPANY_COLORS['blue']}; border-radius: 16px; margin-top: 8px; }}
            QTabBar::tab:selected {{ background: {COMPANY_COLORS['yellow']}; color: {COMPANY_COLORS['dark']}; border-radius: 12px 12px 0 0; font-weight: bold; }}
            QTabBar::tab:!selected {{ background: {COMPANY_COLORS['blue']}; color: {COMPANY_COLORS['white']}; border-radius: 12px 12px 0 0; }}
            QTabBar::tab:hover {{ background: #e6e6e6; color: {COMPANY_COLORS['blue']}; }}
            QStatusBar {{ background: #f0f0f0; border-top: 1px solid #e0e0e0; font-size: 15px; padding: 6px 16px; border-radius: 0 0 12px 12px; }}
            QListWidget {{ border-radius: 12px; border: 1px solid #e0e0e0; margin: 8px 0; font-size: 15px; }}
        """)
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(32, 24, 32, 24)
        main_layout.setSpacing(16)
        logo_label = QLabel()
        logo_label.setObjectName("logoLabel")
        logo_pixmap = QPixmap(LOGO_PATH)
        logo_label.setPixmap(logo_pixmap.scaledToHeight(70, Qt.TransformationMode.SmoothTransformation))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(logo_label)
        # Add dark mode toggle as a visible switch
        self.dark_mode_switch = QCheckBox('Dark Mode')
        self.dark_mode_switch.setChecked(False)
        self.dark_mode_switch.setStyleSheet('QCheckBox { font-size: 16px; padding: 6px 16px; }')
        self.dark_mode_switch.stateChanged.connect(self.toggle_dark_mode)
        # Place at top-right
        top_bar = QHBoxLayout()
        top_bar.addStretch(1)
        top_bar.addWidget(self.dark_mode_switch)
        main_layout.insertLayout(0, top_bar)
        tabs = AnimatedTabWidget()
        tabs.setTabPosition(QTabWidget.TabPosition.North)
        tabs.setStyleSheet('QTabBar::tab { font-size: 18px; padding: 12px 32px; min-height: 40px; }')
        # Convert Tab
        convert_tab = QWidget()
        convert_layout = QVBoxLayout(convert_tab)
        convert_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        convert_layout.setContentsMargins(24, 18, 24, 18)
        convert_layout.setSpacing(18)
        self.label = QLabel("Select one or more CSV Report Files")
        self.label.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {COMPANY_COLORS['blue']}; margin-bottom: 18px;")
        convert_layout.addWidget(self.label)
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(18)
        self.select_button = QPushButton("Browse CSV Files")
        self.select_button.clicked.connect(self.browse_files)
        btn_layout.addWidget(self.select_button)
        self.remove_button = QPushButton("Remove Selected")
        self.remove_button.clicked.connect(self.remove_selected_files)
        btn_layout.addWidget(self.remove_button)
        self.convert_button = QPushButton("Convert to XLSX (Batch)")
        self.convert_button.clicked.connect(self.convert_files)
        btn_layout.addWidget(self.convert_button)
        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        convert_layout.addWidget(self.progress_bar)
        convert_layout.addLayout(btn_layout)
        self.file_list = QListWidget()
        convert_layout.addWidget(self.file_list)
        convert_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))
        self.status_bar = QStatusBar()
        main_layout.addWidget(self.status_bar)
        tabs.addTab(convert_tab, QIcon(LOGO_PATH), "Convert")
        # History Tab (placeholder)
        history_tab = QWidget()
        history_layout = QVBoxLayout(history_tab)
        history_label = QLabel("Conversion history and logs will appear here.")
        history_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        history_label.setStyleSheet("font-size: 18px; color: #888;")
        history_layout.addWidget(history_label)
        tabs.addTab(history_tab, "History")
        main_layout.addWidget(tabs)
        self.csv_paths = []
        self.output_path = ""
        # Drag-and-drop support
        self.file_list.setAcceptDrops(True)
        self.file_list.dragEnterEvent = self.dragEnterEvent
        self.file_list.dropEvent = self.dropEvent
        tabs.currentChanged.connect(self.on_tab_changed)

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Open CSV Files", self.last_csv_path, "CSV files (*.csv)")
        if files:
            for f in files:
                if f not in self.csv_paths:
                    self.csv_paths.append(f)
                    self.add_file_with_animation(f)
            self.last_csv_path = os.path.dirname(files[0])
            self.label.setText(f"Selected {len(self.csv_paths)} file(s)")
            self.status_bar.showMessage("")

    def add_file_with_animation(self, file):
        item = QListWidgetItem(file)
        self.file_list.addItem(item)
        self.file_list.setCurrentItem(item)

    def remove_selected_files(self):
        """
        Remove selected files from the list and keep self.csv_paths in sync.
        """
        selected = self.file_list.selectedItems()
        for item in selected:
            idx = self.file_list.row(item)
            self.file_list.takeItem(idx)
            del self.csv_paths[idx]
        self.label.setText(f"Selected {len(self.csv_paths)} file(s)")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith('.csv') and file_path not in self.csv_paths:
                self.csv_paths.append(file_path)
                self.add_file_with_animation(file_path)
        self.label.setText(f"Selected {len(self.csv_paths)} file(s)")

    def toggle_dark_mode(self):
        self.dark_mode = self.dark_mode_switch.isChecked()
        if self.dark_mode:
            self.setStyleSheet("""
                QMainWindow, QWidget { background: #23272e; color: #f0f0f0; }
                QPushButton { background: #444c5e; color: #f0f0f0; border-radius: 16px; padding: 14px 32px; font-size: 18px; font-weight: 600; margin: 0 8px; border: 1px solid #222; }
                QPushButton:hover { background: #FFDE00; color: #23272e; }
                QLabel#logoLabel { background: transparent; margin-bottom: 12px; }
                QTabWidget::pane { border: 2px solid #6C9DFE; border-radius: 16px; margin-top: 8px; }
                QTabBar::tab:selected { background: #FFDE00; color: #23272e; border-radius: 12px 12px 0 0; font-weight: bold; }
                QTabBar::tab:!selected { background: #6C9DFE; color: #f0f0f0; border-radius: 12px 12px 0 0; }
                QTabBar::tab:hover { background: #e6e6e6; color: #6C9DFE; }
                QStatusBar { background: #23272e; border-top: 1px solid #444; font-size: 15px; padding: 6px 16px; border-radius: 0 0 12px 12px; }
                QListWidget { border-radius: 12px; border: 1px solid #444; margin: 8px 0; font-size: 15px; }
                QCheckBox { color: #f0f0f0; }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow, QWidget { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFFFFF, stop:1 #F5F7FA); color: #030408; }
                QPushButton { background: #6C9DFE; color: #FFFFFF; border-radius: 16px; padding: 14px 32px; font-size: 18px; font-weight: 600; margin: 0 8px; border: 1px solid #b0b0b0; }
                QPushButton:hover { background: #FFDE00; color: #030408; }
                QLabel#logoLabel { background: transparent; margin-bottom: 12px; }
                QTabWidget::pane { border: 2px solid #6C9DFE; border-radius: 16px; margin-top: 8px; }
                QTabBar::tab:selected { background: #FFDE00; color: #030408; border-radius: 12px 12px 0 0; font-weight: bold; }
                QTabBar::tab:!selected { background: #6C9DFE; color: #FFFFFF; border-radius: 12px 12px 0 0; }
                QTabBar::tab:hover { background: #e6e6e6; color: #6C9DFE; }
                QStatusBar { background: #f0f0f0; border-top: 1px solid #e0e0e0; font-size: 15px; padding: 6px 16px; border-radius: 0 0 12px 12px; }
                QListWidget { border-radius: 12px; border: 1px solid #e0e0e0; margin: 8px 0; font-size: 15px; }
                QCheckBox { color: #030408; }
            """)

    def convert_files(self):
        if not self.csv_paths:
            self.status_bar.showMessage("No files selected.")
            QMessageBox.warning(self, "No Files Selected", "Please select one or more CSV files to convert.")
            return
        output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory", self.last_output_path)
        if not output_dir:
            return
        self.output_path = output_dir
        self.status_bar.showMessage("Converting batch...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(self.csv_paths))
        self.progress_bar.setValue(0)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        summary = []
        for i, csv_path in enumerate(self.csv_paths):
            base = os.path.splitext(os.path.basename(csv_path))[0]
            out_file = os.path.join(output_dir, base + ".xlsx")
            try:
                success, info = convert_csv_to_xlsx(csv_path, out_file)
                if not success:
                    log_conversion(os.path.basename(csv_path), f"Error: {info}", 0)
                    summary.append((base, f"Error: {info}"))
                elif info == 0:
                    msg = "No data matched the filter. Output file is empty."
                    log_conversion(os.path.basename(csv_path), f"Warning: {msg}", 0)
                    summary.append((base, msg))
                else:
                    log_conversion(os.path.basename(csv_path), "Success", info)
                    summary.append((base, f"Complete ({info} rows)", ))
            except Exception as e:
                log_conversion(os.path.basename(csv_path), f"Exception: {e}", 0)
                summary.append((base, f"Exception: {e}"))
            self.progress_bar.setValue(i+1)
            QApplication.processEvents()
        QApplication.restoreOverrideCursor()
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage("Batch conversion complete.")
        self.save_config()
        self.show_summary_dialog(summary)
        self.update_history_tab()
        # Reset UI after dialog
        self.label.setText("Select one or more CSV Report Files")
        self.file_list.clear()
        self.csv_paths = []
        self.select_button.setEnabled(True)
        self.remove_button.setEnabled(True)
        self.convert_button.setEnabled(True)

    def show_summary_dialog(self, summary):
        dialog = QDialog(self)
        dialog.setWindowTitle("Conversion Summary")
        layout = QVBoxLayout(dialog)
        table = QTableWidget(len(summary), 2)
        table.setHorizontalHeaderLabels(["File", "Status"])
        for row, (file, status) in enumerate(summary):
            table.setItem(row, 0, QTableWidgetItem(file))
            table.setItem(row, 1, QTableWidgetItem(status))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(table)
        btn = QPushButton("OK")
        btn.clicked.connect(dialog.accept)
        layout.addWidget(btn)
        dialog.exec()

    def update_history_tab(self):
        # Show actual conversion logs from the database
        import sqlite3
        from datetime import datetime
        DB_PATH = "conversion_logs.db"
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT filename, converted_at, status, rows FROM logs ORDER BY converted_at DESC LIMIT 100")
        rows = c.fetchall()
        conn.close()
        table = QTableWidget(len(rows), 4)
        table.setHorizontalHeaderLabels(["File", "Converted At", "Status", "Rows"])
        for row, (filename, converted_at, status, nrows) in enumerate(rows):
            table.setItem(row, 0, QTableWidgetItem(filename))
            table.setItem(row, 1, QTableWidgetItem(converted_at))
            table.setItem(row, 2, QTableWidgetItem(status))
            table.setItem(row, 3, QTableWidgetItem(str(nrows)))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # Replace the placeholder in the history tab
        history_tab = self.centralWidget().findChild(QTabWidget).widget(1)
        layout = history_tab.layout()
        for i in reversed(range(layout.count())):
            widget = layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        layout.addWidget(table)

    def on_tab_changed(self, idx):
        tabs = self.centralWidget().findChild(QTabWidget)
        if idx == 1:  # History tab
            self.update_history_tab() 