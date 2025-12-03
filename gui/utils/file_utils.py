#!/usr/bin/env python3
"""
File Utilities - Helper functions for file I/O operations
Handles reading/writing LTA files and shipper files
"""

import os
import glob
import logging

logger = logging.getLogger(__name__)

def clean_lta_reference(reference):
    """
    Clean LTA reference by removing /1 suffix
    
    Args:
        reference: LTA reference string (e.g., "607-50843822/1")
        
    Returns:
        Cleaned reference (e.g., "607-50843822")
    """
    if not reference:
        return reference
    
    # Remove /1 suffix if present
    if reference.endswith('/1'):
        return reference[:-2]
    
    return reference

def detect_ltas(folder_path):
    """
    Detect all LTA folders in the given path
    
    Args:
        folder_path: Path to search for LTA folders
        
    Returns:
        List of dicts with LTA information
    """
    ltas = []
    
    try:
        # Look for folders with "LTA" in the name
        all_items = os.listdir(folder_path)
        lta_folders = [item for item in all_items 
                      if os.path.isdir(os.path.join(folder_path, item)) 
                      and 'lta' in item.lower()]
        
        for folder_name in lta_folders:
            folder_full_path = os.path.join(folder_path, folder_name)
            
            # Look for shipper file
            shipper_pattern = f"{folder_name.replace(' ', '_')}_*.txt"
            shipper_files = glob.glob(os.path.join(folder_path, shipper_pattern))
            
            # Look for LTA file
            lta_file_pattern = f"{folder_name}.txt"
            lta_file_path = os.path.join(folder_path, lta_file_pattern)
            
            lta_info = {
                'name': folder_name,
                'folder_path': folder_full_path,
                'shipper_file': shipper_files[0] if shipper_files else None,
                'lta_file': lta_file_path if os.path.exists(lta_file_path) else None,
                'has_ds': False,  # Will be updated after reading shipper file
                'ds_series': None,  # Original DS from shipper file line 2
                'validated_ds': None,  # Validated DS from shipper file line 4
                'signed_ds': None,  # Signed DS from LTA file line 8
                'location': None,
                'lta_reference': None  # LTA reference from LTA file line 4
            }
            
            # Read shipper file if exists
            if lta_info['shipper_file']:
                shipper_data = read_shipper_file(lta_info['shipper_file'])
                if shipper_data:
                    lta_info['has_ds'] = bool(shipper_data.get('ds_series'))
                    lta_info['ds_series'] = shipper_data.get('ds_series')
                    lta_info['validated_ds'] = shipper_data.get('ds_reference')  # Line 4
                    lta_info['location'] = shipper_data.get('location')
            
            # Read LTA file if exists to get signed series
            if lta_info['lta_file']:
                lta_data = read_lta_file(lta_info['lta_file'])
                if lta_data:
                    lta_info['signed_ds'] = lta_data.get('signed_series')  # Line 8
                    raw_reference = lta_data.get('lta_reference')  # Line 4
                    lta_info['lta_reference'] = clean_lta_reference(raw_reference)
            
            ltas.append(lta_info)
        
        logger.info(f"Detected {len(ltas)} LTA folders")
        return sorted(ltas, key=lambda x: x['name'])
        
    except Exception as e:
        logger.error(f"Error detecting LTAs: {e}", exc_info=True)
        return []

def read_shipper_file(file_path):
    """
    Read shipper file (*_LTA_shipper_name.txt)
    
    Format:
        Line 1: Shipper company name
        Line 2: DS Serie (e.g., "9913 G") - optional
        Line 3: Location (e.g., "ABU DHABI INT") - optional
        Line 4: DS reference (added after Phase 1)
        Line 5: LTA reference validated (added after Phase 1)
    
    Returns:
        Dict with shipper data or None
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = [line.rstrip('\n\r') for line in f.readlines()]
        
        data = {
            'shipper_name': lines[0].strip() if len(lines) > 0 else None,
            'ds_series': lines[1].strip() if len(lines) > 1 else None,
            'location': lines[2].strip() if len(lines) > 2 else None,
            'ds_reference': lines[3].strip() if len(lines) > 3 else None,
            'lta_reference': lines[4].strip() if len(lines) > 4 else None
        }
        
        # Filter out empty strings
        data = {k: v if v else None for k, v in data.items()}
        
        return data
        
    except Exception as e:
        logger.error(f"Error reading shipper file {file_path}: {e}")
        return None

def write_shipper_file(file_path, ds_series=None, location=None):
    """
    Write DS series and location to shipper file (Lines 2 and 3)
    Preserves Line 1 (shipper name)
    
    Args:
        file_path: Path to shipper file
        ds_series: DS series to write (e.g., "9913 G")
        location: Location to write (e.g., "ABU DHABI INT")
    """
    try:
        # Read existing content
        existing_data = read_shipper_file(file_path)
        if not existing_data:
            logger.error(f"Cannot write to non-existent file: {file_path}")
            return False
        
        # Prepare lines
        lines = [
            existing_data.get('shipper_name', ''),
            ds_series if ds_series else '',
            location if location else ''
        ]
        
        # Preserve lines 4 and 5 if they exist
        if existing_data.get('ds_reference'):
            lines.append(existing_data['ds_reference'])
        if existing_data.get('lta_reference'):
            if len(lines) == 3:  # Add empty line 4 if needed
                lines.append('')
            lines.append(existing_data['lta_reference'])
        
        # Write to file
        with open(file_path, 'w', encoding='utf-8') as f:
            for line in lines:
                f.write(f"{line}\n")
        
        logger.info(f"Updated shipper file: {file_path}")
        return True
        
    except Exception as e:
        logger.error(f"Error writing shipper file {file_path}: {e}")
        return False

def read_lta_file(file_path):
    """
    Read LTA file (*eme LTA.txt)
    
    Returns:
        Dict with LTA data or None
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = [line.rstrip('\n\r') for line in f.readlines()]
        
        data = {
            'lta_name': lines[1].strip() if len(lines) > 1 else None,
            'mawb': lines[2].strip() if len(lines) > 2 else None,
            'lta_reference': lines[3].strip() if len(lines) > 3 else None,
            'shipper_name': lines[5].strip() if len(lines) > 5 else None,
            'signed_series': lines[7].strip() if len(lines) > 7 else None,  # Line 8 (index 7)
        }
        
        return data
        
    except Exception as e:
        logger.error(f"Error reading LTA file {file_path}: {e}")
        return None

def write_lta_signed_series(file_path, signed_series):
    """
    Write signed series to LTA file Line 8
    
    Args:
        file_path: Path to LTA file
        signed_series: Signed series (e.g., "9914 H")
    """
    try:
        # Read existing content
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = [line.rstrip('\n\r') for line in f.readlines()]
        
        # Ensure we have at least 8 lines
        while len(lines) < 8:
            lines.append('')
        
        # Update line 8 (index 7)
        lines[7] = signed_series
        
        # Write back
        with open(file_path, 'w', encoding='utf-8') as f:
            for line in lines:
                f.write(f"{line}\n")
        
        logger.info(f"Updated LTA file with signed series: {file_path}")
        return True
        
    except Exception as e:
        logger.error(f"Error writing LTA file {file_path}: {e}")
        return False

def get_dum_count(lta_folder_path):
    """
    Count number of DUMs in an LTA folder by counting Sheet files
    
    Args:
        lta_folder_path: Path to LTA folder
        
    Returns:
        Number of DUMs (Sheet files)
    """
    try:
        sheet_files = glob.glob(os.path.join(lta_folder_path, "Sheet *.xlsx"))
        return len(sheet_files)
    except Exception as e:
        logger.error(f"Error counting DUMs: {e}")
        return 0
