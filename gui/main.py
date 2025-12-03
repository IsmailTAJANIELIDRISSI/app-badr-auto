#!/usr/bin/env python3
"""
BADR Automation - Main Entry Point
Launches the desktop application
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox
import logging

# IMPORTANT: Change working directory to the script's location
# This fixes the issue when launching from Start Menu (which uses C:\Windows\system32)
if getattr(sys, 'frozen', False):
    # Running as .exe - change to exe directory
    application_path = os.path.dirname(sys.executable)
else:
    # Running as .py - change to script directory
    application_path = os.path.dirname(os.path.abspath(__file__))
    application_path = os.path.dirname(application_path)  # Go up to project root

os.chdir(application_path)

# Add parent directory to path to import existing scripts
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gui.app import BADRApp

def setup_logging():
    """Configure logging for the application"""
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # Log file in the application directory (not system32)
    log_file = os.path.join(application_path, 'badr_gui.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

def main():
    """Main entry point"""
    try:
        # Setup logging
        setup_logging()
        logger = logging.getLogger(__name__)
        logger.info("Starting BADR Automation Application...")
        
        # Create root window
        root = tk.Tk()
        
        # Create application
        app = BADRApp(root)
        
        # Start event loop
        logger.info("Application initialized successfully")
        root.mainloop()
        
    except Exception as e:
        error_msg = f"Failed to start application: {e}"
        logging.error(error_msg, exc_info=True)
        messagebox.showerror("Erreur de Démarrage", error_msg)
        sys.exit(1)

if __name__ == "__main__":
    main()
