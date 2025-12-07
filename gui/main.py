#!/usr/bin/env python3
"""
BADR Automation - Main Entry Point
Launches the desktop application
"""

import sys
import os
import subprocess
import tkinter as tk
from tkinter import messagebox
import logging
# salam
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

# ============================================================================
# AUTO-UPDATE FROM GITHUB (SILENT)
# ============================================================================
# This ensures employees automatically get:
# - License renewals (updated LTA_sys_ts and LTA_validity)
# - Script improvements and bug fixes
# - New features
# All without manual intervention
try:
    # CREATE_NO_WINDOW prevents terminal windows from appearing on Windows
    creation_flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    
    # Check if git is available
    _git_check = subprocess.run(
        ["git", "--version"],
        capture_output=True,
        text=True,
        timeout=5,
        cwd=application_path,
        creationflags=creation_flags
    )
    
    # Check if we're in a git repository
    _git_status_check = subprocess.run(
        ["git", "rev-parse", "--git-dir"],
        capture_output=True,
        text=True,
        timeout=5,
        cwd=application_path,
        creationflags=creation_flags
    )
    
    if _git_status_check.returncode == 0:
        # Pull updates silently with --autostash
        # This will:
        # 1. Stash any local changes (like LTA folders added by employees)
        # 2. Pull updates from GitHub (license, scripts, GUI)
        # 3. Reapply stashed changes
        # Local files/folders are preserved!
        subprocess.run(
            ["git", "pull", "--autostash", "origin", "main"],
            capture_output=True,
            text=True,
            timeout=30,
            cwd=application_path,
            creationflags=creation_flags
        )
except:
    # Silent fail - continue with current version
    # This allows the app to work even without git installed
    pass

# Add parent directory to path to import existing scripts
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gui.app import BADRApp
from gui.utils.license_validator import validate_and_continue

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
        root.withdraw()  # Hide until validation passes
        
        # Validate license before starting
        if not validate_and_continue(root, show_warnings=True):
            logger.error("License validation failed - application will exit")
            # Don't destroy root - it's already handled by the validator
            sys.exit(1)
        
        # Show the window after validation passes
        root.deiconify()
        
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
