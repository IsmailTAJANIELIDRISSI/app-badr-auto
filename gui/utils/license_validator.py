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
        # Get application path (handles both .exe and .py)
        if getattr(sys, 'frozen', False):
            app_path = os.path.dirname(sys.executable)
        else:
            app_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        config_path = os.path.join(app_path, 'config', 'license.json')
        
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get('LTA_sys_ts', 1763251200), config.get('LTA_validity_days', 17)
        else:
            # Fallback to default values if file not found
            return 1763251200, 17
    except:
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
║          ❌ LICENCE EXPIRÉE - ACCÈS REFUSÉ                ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝

⚠️  La licence du système a expiré.

Le service BADR Automation nécessite un renouvellement pour 
continuer à fonctionner.

💡 Actions requises :
   • Vérifier le statut de l'abonnement
   • Renouveler les services cloud
   • Mettre à jour l'application (git pull)

📞 Support Technique :
   Contactez votre administrateur système pour assistance.

"""
    
    messagebox.showerror(
        "BADR Automation - Licence Expirée",
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
⚠️  ATTENTION : Licence expire bientôt

Votre licence BADR Automation expire dans {days_remaining} jour(s).

💡 Action recommandée :
   • Préparer le renouvellement de la licence
   • Vérifier les mises à jour disponibles (git pull)
   • Contacter le support si nécessaire

L'application continuera de fonctionner jusqu'à expiration.
"""
    
    messagebox.showwarning(
        "BADR Automation - Avertissement Licence",
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
