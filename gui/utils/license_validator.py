#!/usr/bin/env python3
"""
License Validation Module
Checks system validity for BADR Automation application
"""

import sys
import os
import json
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

def _load_license_config():
    """Load license configuration from external JSON file"""
    try:
        # Always read from the git repository root config folder
        # This allows git pull to update the license without any manual copying
        if getattr(sys, 'frozen', False):
            # When frozen (.exe), find the git repo root (go up from dist/)
            exe_dir = os.path.dirname(sys.executable)
            # Check if we're in a 'dist' folder
            if os.path.basename(exe_dir) == 'dist':
                app_path = os.path.dirname(exe_dir)  # Go up to repo root
            else:
                app_path = exe_dir  # Already at root
        else:
            app_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        config_path = os.path.join(app_path, 'config', 'license.json')
        
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Parse the expiration date (format: YYYY-MM-DD)
                expiry_str = config.get('LTA_sys_validity', '2025-01-09')
                expiry_date = datetime.strptime(expiry_str, '%Y-%m-%d')
                expiry_ts = int(expiry_date.timestamp())
                
                # Return a fake start timestamp and calculate validity days
                # This maintains compatibility with existing code
                current_ts = int(datetime.now().timestamp())
                days_remaining = (expiry_ts - current_ts) // (24 * 3600)
                
                # Return values that make the license valid until expiry_date
                return expiry_ts - (days_remaining * 24 * 3600), days_remaining
        else:
            # Fallback to default values if file not found (expired)
            return 1763251200, 17
    except Exception as e:
        # Fallback to default values on any error
        return 1763251200, 17

# Load license configuration from external file
# This allows git pull to update the license without rebuilding the .exe
LTA_sys_ts, validity_days = _load_license_config()
LTA_validity = validity_days * 24 * 3600  # Convert days to seconds

def check_license_validity():
    """
    Check if the application license is still valid
    
    Returns:
        tuple: (is_valid: bool, days_remaining: int)
    """
    current_ts = int(datetime.now().timestamp())
    expiry_ts = LTA_sys_ts + LTA_validity
    
    if current_ts > expiry_ts:
        return False, 0
    
    # Calculate days remaining
    seconds_remaining = expiry_ts - current_ts
    days_remaining = seconds_remaining // (24 * 3600)
    
    return True, days_remaining

def show_license_expired_dialog(root=None):
    """
    Show professional license expiration dialog
    
    Args:
        root: Tkinter root window (optional)
    """
    message = """
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║       ⚙️  CALIBRAGE DU SYSTÈME OCR REQUIS                ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝

⚠️  Le moteur de reconnaissance optique nécessite une 
    recalibration pour garantir des résultats optimaux.

🔧 Maintenance préventive détectée :
   • Mise à jour des algorithmes de détection
   • Synchronisation des bases de données OCR
   • Optimisation des modèles d'apprentissage

📞 Support OCR :
   Contactez l'équipe technique pour la recalibration.

"""
    
    messagebox.showwarning(
        "BADR OCR Engine - Calibrage Requis",
        message,
        parent=root
    )

def show_license_warning(days_remaining, root=None):
    """
    Show warning when license is about to expire
    
    Args:
        days_remaining: Number of days until expiration
        root: Tkinter root window (optional)
    """
    message = f"""
⚠️  MAINTENANCE OCR PROGRAMMÉE

Le système OCR nécessitera un recalibrage dans {days_remaining} jour(s).

💡 Préparation recommandée :
   • Vérifier les mises à jour OCR (git pull)
   • Sauvegarder les configurations actuelles
   • Planifier la fenêtre de maintenance

✓ Le système continuera de fonctionner normalement.
"""
    
    messagebox.showinfo(
        "BADR OCR Engine - Maintenance Planifiée",
        message,
        parent=root
    )

def validate_and_continue(root=None, show_warnings=True):
    """
    Validate license and show appropriate messages
    
    Args:
        root: Tkinter root window (optional)
        show_warnings: Whether to show warnings for upcoming expiration
        
    Returns:
        bool: True if valid, False if expired
    """
    is_valid, days_remaining = check_license_validity()
    
    if not is_valid:
        show_license_expired_dialog(root)
        return False
    
    # Show warning if less than 5 days remaining
    if show_warnings and days_remaining <= 5:
        show_license_warning(days_remaining, root)
    
    return True
