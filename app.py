from database import setup_database
from gui import ElectrolyteCRMApp, SplashScreen
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTimer
import sys

if __name__ == "__main__":
    setup_database()
    app = QApplication(sys.argv)
    
    # Show splash screen
    splash = SplashScreen()
    splash.show()
    
    def start_main():
        # Create the new CRM app
        window = ElectrolyteCRMApp()
        window.show()
        splash.finish(window)
        # Keep a reference to the window so it doesn't get garbage collected
        app.window = window
    
    # Start the main app after 2 seconds
    QTimer.singleShot(2000, start_main)
    
    sys.exit(app.exec()) 