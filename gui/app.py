#!/usr/bin/env python3
"""
BADR Automation - Main Application Window
Manages tabs and global application state
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
import os

from gui.screens.preparation import PreparationScreen
from gui.screens.phase1_ed import Phase1EDScreen
from gui.screens.phase2_dum import Phase2DUMScreen
from gui.screens.logs import LogsScreen

logger = logging.getLogger(__name__)

class BADRApp:
    """Main application class"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("BADR Automation - Déclarations LTA")
        self.root.geometry("1100x750")
        
        # Application state
        self.current_folder = None
        self.lta_list = []
        self.phase1_completed = False
        self.phase2_completed = False
        
        # Setup UI
        self._setup_ui()
        
        # Configure window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        logger.info("Application window initialized")
    
    def _setup_ui(self):
        """Setup the main UI with modern enhanced tabs"""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="0")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=0)
        
        # ========== Enhanced Tab Styling ==========
        style = ttk.Style()
        
        # Use a modern theme as base
        try:
            style.theme_use('clam')  # Modern, flat appearance
        except:
            style.theme_use('default')
        
        # Configure notebook (tab container)
        style.configure('TNotebook',
                       background='#ffffff',
                       borderwidth=0,
                       tabmargins=[5, 5, 0, 0])
        
        # Configure individual tabs - MUCH larger and more visible
        style.configure('TNotebook.Tab',
                       padding=[25, 15],  # Large padding: 25px horizontal, 15px vertical
                       font=('Segoe UI', 11, 'bold'),
                       background='#e9ecef',
                       foreground='#495057',
                       borderwidth=0,
                       focuscolor='none')
        
        # Tab states with better colors
        style.map('TNotebook.Tab',
                 background=[('selected', '#0066cc'),    # Blue when selected
                           ('active', '#dee2e6'),        # Light gray on hover
                           ('!selected', '#e9ecef')],    # Default gray
                 foreground=[('selected', '#ffffff'),    # White text when selected
                           ('active', '#212529'),        # Dark text on hover
                           ('!selected', '#495057')],    # Gray text default
                 expand=[('selected', [2, 2, 2, 0])],   # Expand selected tab
                 borderwidth=[('selected', 0)])
        
        # ========== Create Notebook ==========
        notebook_container = ttk.Frame(main_frame)
        notebook_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        notebook_container.columnconfigure(0, weight=1)
        notebook_container.rowconfigure(0, weight=1)
        
        self.notebook = ttk.Notebook(notebook_container)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=(5, 0))
        
        # Create tabs with icons
        self.prep_screen = PreparationScreen(self.notebook, self)
        self.phase1_screen = Phase1EDScreen(self.notebook, self)
        self.phase2_screen = Phase2DUMScreen(self.notebook, self)
        self.logs_screen = LogsScreen(self.notebook, self)
        
        # Add tabs to notebook with icons and better labels
        self.notebook.add(self.prep_screen.frame, text="  📋  Étape 1: Préparation  ")
        self.notebook.add(self.phase1_screen.frame, text="  📝  Étape 2: Phase ED  ")
        self.notebook.add(self.phase2_screen.frame, text="  📦  Étape 3: Phase DUM  ")
        self.notebook.add(self.logs_screen.frame, text="  📊  Logs & Console  ")
        
        # Bind tab change event for auto-refresh
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        # ========== Status Bar ==========
        status_frame = tk.Frame(main_frame, bg='#343a40', height=32)
        status_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        status_frame.grid_propagate(False)
        
        self.status_var = tk.StringVar(value="✓ Prêt")
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg='#343a40',
            fg='#ffffff',
            anchor=tk.W,
            padx=20,
            font=('Segoe UI', 9)
        )
        status_label.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Add version info on right side of status bar
        version_label = tk.Label(
            status_frame,
            text="v1.0",
            bg='#343a40',
            fg='#6c757d',
            anchor=tk.E,
            padx=20,
            font=('Segoe UI', 8)
        )
        version_label.pack(fill=tk.BOTH, side=tk.RIGHT)
    
    def set_status(self, message):
        """Update status bar message"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def _on_tab_changed(self, event):
        """Handle tab change event - auto-refresh LTAs"""
        current_tab = self.notebook.index(self.notebook.select())
        
        # Tab 1: Phase 1 ED - auto-refresh LTA list
        if current_tab == 1:
            if self.current_folder:
                self.phase1_screen.refresh_lta_list()
        
        # Tab 2: Phase 2 DUM - auto-refresh LTA list
        elif current_tab == 2:
            if self.current_folder:
                self.phase2_screen.refresh_lta_list()
    
    def enable_phase1_tab(self):
        """Enable Phase 1 tab after preparation"""
        # Visual feedback - already enabled
        logger.info("Phase 1 tab enabled")
    
    def enable_phase2_tab(self):
        """Enable Phase 2 tab after Phase 1"""
        # Visual feedback - already enabled
        logger.info("Phase 2 tab enabled")
    
    def log_message(self, message, level="INFO"):
        """Send message to logs screen"""
        self.logs_screen.add_log(message, level)
    
    def on_closing(self):
        """Handle window close event"""
        if messagebox.askokcancel("Quitter", "Voulez-vous vraiment quitter l'application?"):
            logger.info("Application closing")
            self.root.destroy()