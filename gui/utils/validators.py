#!/usr/bin/env python3
"""
Input Validators - Validation functions for user inputs
"""

import re
import logging

logger = logging.getLogger(__name__)

def normalize_ds_series(text):
    """
    Normalize DS series format by removing extra whitespace and newlines
    Formats to: "XXXX Y" (4 digits + space + letter)
    
    Args:
        text: Input text to normalize
        
    Returns:
        Normalized text in format "XXXX Y"
    """
    if not text:
        return ""
    
    # Remove all whitespace and newlines
    cleaned = ''.join(text.split())
    
    # Extract digits and letter
    # Pattern: extract 4 digits followed by a letter
    match = re.match(r'^(\d{4})([A-Z]?)$', cleaned.upper())
    
    if match:
        digits = match.group(1)
        letter = match.group(2) if match.group(2) else ''
        
        if letter:
            return f"{digits} {letter}"
        else:
            return digits
    
    # If pattern doesn't match exactly, try to extract what we can
    # Find 4 consecutive digits
    digits_match = re.search(r'\d{4}', cleaned)
    # Find a letter
    letter_match = re.search(r'[A-Z]', cleaned.upper())
    
    if digits_match and letter_match:
        return f"{digits_match.group()} {letter_match.group()}"
    elif digits_match:
        return digits_match.group()
    
    # Return cleaned text if nothing matches
    return cleaned.upper()

def validate_ds_series(text):
    """
    Validate DS series format: "XXXX Y" (digits + space + letter)
    
    Args:
        text: Input text to validate
        
    Returns:
        Tuple (is_valid, error_message)
    """
    if not text or not text.strip():
        return True, None  # Empty is valid (optional field)
    
    # Normalize first
    text = normalize_ds_series(text)
    
    # Pattern: 4 digits + space + single uppercase letter
    pattern = r'^\d{4}\s+[A-Z]$'
    
    if re.match(pattern, text):
        return True, None
    else:
        return False, "Format invalide. Attendu: '9913 G' (4 chiffres + espace + lettre)"

def validate_location(text):
    """
    Validate location name
    
    Args:
        text: Input text to validate
        
    Returns:
        Tuple (is_valid, error_message)
    """
    if not text or not text.strip():
        return True, None  # Empty is valid (optional field)
    
    text = text.strip()
    
    if len(text) < 3:
        return False, "Le nom du lieu doit contenir au moins 3 caractères"
    
    return True, None

def validate_signed_series(text):
    """
    Validate signed series format (same as DS series)
    
    Args:
        text: Input text to validate
        
    Returns:
        Tuple (is_valid, error_message)
    """
    if not text or not text.strip():
        return False, "La série signée est requise"
    
    return validate_ds_series(text)

def validate_folder_path(path):
    """
    Validate folder path exists
    
    Args:
        path: Folder path to validate
        
    Returns:
        Tuple (is_valid, error_message)
    """
    import os
    
    if not path or not path.strip():
        return False, "Veuillez sélectionner un dossier"
    
    if not os.path.exists(path):
        return False, "Le dossier n'existe pas"
    
    if not os.path.isdir(path):
        return False, "Le chemin n'est pas un dossier"
    
    return True, None

def validate_credentials(username, password):
    """
    Validate BADR credentials
    
    Args:
        username: Username
        password: Password
        
    Returns:
        Tuple (is_valid, error_message)
    """
    if not username or not username.strip():
        return False, "Le nom d'utilisateur est requis"
    
    if not password or not password.strip():
        return False, "Le mot de passe est requis"
    
    return True, None
