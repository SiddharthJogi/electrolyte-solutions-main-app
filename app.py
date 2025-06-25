from database import setup_database
from gui import ConverterApp, SplashScreen
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTimer
import sys

if __name__ == "__main__":
    setup_database()
    app = QApplication(sys.argv)
    splash = SplashScreen()
    splash.show()
    def start_main():
        window = ConverterApp()
        window.show()
        splash.finish(window)
        # Keep a reference to the window so it doesn't get garbage collected
        app.window = window
    QTimer.singleShot(2000, start_main)
    sys.exit(app.exec()) 