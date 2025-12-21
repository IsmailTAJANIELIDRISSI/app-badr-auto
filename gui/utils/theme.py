#!/usr/bin/env python3
"""
Simple MedAfrica Logistics Theme
Just header branding - everything else uses default tkinter theme
"""

import tkinter as tk
import os

# MedAfrica Logistics Orange Color
ORANGE_PRIMARY = "#FF6B35"

def get_logo_path():
    """Get path to MedAfrica logo"""
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    logo_path = os.path.join(script_dir, 'logo.png')
    if os.path.exists(logo_path):
        return logo_path
    return None

def set_window_icon(root):
    """Set window icon from logo file"""
    logo_path = get_logo_path()
    if logo_path:
        try:
            # On Windows, try to convert PNG to ICO and set as icon
            if os.name == 'nt':
                try:
                    from PIL import Image
                    import tempfile
                    img = Image.open(logo_path)
                    # Convert to ICO format
                    ico_path = os.path.join(tempfile.gettempdir(), 'medafrica_icon.ico')
                    # Save in multiple sizes for better quality
                    img.save(ico_path, format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
                    root.iconbitmap(ico_path)
                except Exception:
                    # Fallback: try PNG directly (works on some systems)
                    try:
                        root.iconbitmap(logo_path)
                    except Exception:
                        pass
            else:
                # On non-Windows, try to use PNG as icon
                try:
                    root.iconbitmap(logo_path)
                except Exception:
                    pass
        except Exception:
            pass

def create_header(parent):
    """Create simple orange header with company name and logo"""
    header_frame = tk.Frame(parent, bg=ORANGE_PRIMARY, height=70)
    header_frame.pack(fill=tk.X, padx=0, pady=0)
    header_frame.pack_propagate(False)
    
    # Logo (if available)
    logo_path = get_logo_path()
    if logo_path:
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo_path)
            # Resize to fit header (max height 50px)
            img.thumbnail((180, 50), Image.Resampling.LANCZOS)
            logo_img = ImageTk.PhotoImage(img)
            logo_label = tk.Label(header_frame, image=logo_img, bg=ORANGE_PRIMARY)
            logo_label.image = logo_img  # Keep a reference
            logo_label.pack(side=tk.LEFT, padx=15, pady=10)
        except Exception:
            # If PIL not available or image fails, continue without logo
            pass
    
    # Company name and title
    title_frame = tk.Frame(header_frame, bg=ORANGE_PRIMARY)
    title_frame.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.Y)
    
    # Application title - using bright white with better contrast
    title_label = tk.Label(
        title_frame,
        text="BADR Automation - Déclarations LTA",
        bg=ORANGE_PRIMARY,
        fg="#FFFFFF",  # Pure white for maximum contrast
        font=('Arial', 14, 'bold')
    )
    title_label.pack(side=tk.TOP, anchor=tk.W, pady=(0, 2))
    
    # Company name
    company_label = tk.Label(
        title_frame,
        text="MedAfrica Logistics",
        bg=ORANGE_PRIMARY,
        fg="#FFFFFF",  # Pure white for maximum contrast
        font=('Arial', 10, 'italic')
    )
    company_label.pack(side=tk.TOP, anchor=tk.W)
    
    return header_frame

def create_footer(parent):
    """Create footer with copyright"""
    footer_frame = tk.Frame(parent, bg="#E0E0E0", height=25)
    footer_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=0, pady=0)
    footer_frame.pack_propagate(False)
    
    copyright_label = tk.Label(
        footer_frame,
        text="© 2024 MedAfrica Logistics - Tous droits réservés",
        bg="#E0E0E0",
        fg="#666666",
        font=('Arial', 8)
    )
    copyright_label.pack(side=tk.RIGHT, padx=15, pady=5)
    
    return footer_frame

