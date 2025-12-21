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
from gui.utils.theme import create_header, create_footer, set_window_icon

logger = logging.getLogger(__name__)

class BADRApp:
    """Main application class"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("MedAfrica Logistics - BADR Automation")
        self.root.geometry("1000x750")
        
        # Set window icon (logo)
        set_window_icon(root)
        
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
        """Setup the main UI with tabs"""
        # Create orange header with company branding
        header = create_header(self.root)
        
        # Create main container (between header and footer)
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Create main frame with padding
        main_frame = ttk.Frame(main_container, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create tabs
        self.prep_screen = PreparationScreen(self.notebook, self)
        self.phase1_screen = Phase1EDScreen(self.notebook, self)
        self.phase2_screen = Phase2DUMScreen(self.notebook, self)
        self.logs_screen = LogsScreen(self.notebook, self)
        
        # Add tabs to notebook
        self.notebook.add(self.prep_screen.frame, text="1. Préparation")
        self.notebook.add(self.phase1_screen.frame, text="2. Phase ED")
        self.notebook.add(self.phase2_screen.frame, text="3. Phase Déd.")
        self.notebook.add(self.logs_screen.frame, text="Logs")
        
        # Bind tab change event for auto-refresh
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        # All tabs enabled from start (users can run any phase independently)
        # Phase 1 and Phase 2 will show warnings if prerequisites not met
        
        # Status bar
        self.status_var = tk.StringVar(value="Prêt")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Create footer with copyright
        footer = create_footer(self.root)
    
    def set_status(self, message):
        """Update status bar message"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def _on_tab_changed(self, event):
        """Handle tab change event - auto-refresh LTAs"""
        current_tab = self.notebook.index(self.notebook.select())
        tab_names = ["Préparation", "Phase ED", "Phase Déd.", "Logs"]
        self.log_message(f"Changement vers onglet {current_tab}: {tab_names[current_tab] if current_tab < len(tab_names) else 'Inconnu'}", "INFO")
        
        # Tab 1: Phase 1 ED - auto-refresh LTA list
        if current_tab == 1:
            if self.current_folder:
                self.log_message("Tab Phase 1: Auto-refresh des LTAs", "INFO")
                self.phase1_screen.refresh_lta_list()
            else:
                self.log_message("Tab Phase 1: current_folder non défini, pas d'auto-refresh", "WARNING")
        
        # Tab 2: Phase 2 DUM - auto-refresh LTA list
        elif current_tab == 2:
            if self.current_folder:
                self.log_message(f"Tab Phase 2: Auto-refresh des LTAs (dossier: {self.current_folder})", "INFO")
                self.phase2_screen.refresh_lta_list()
            else:
                self.log_message("Tab Phase 2: current_folder non défini, pas d'auto-refresh", "WARNING")
    
    def enable_phase1_tab(self):
        """Enable Phase 1 tab after preparation"""
        self.notebook.tab(1, state="normal")
        logger.info("Phase 1 tab enabled")
    
    def enable_phase2_tab(self):
        """Enable Phase 2 tab after Phase 1"""
        self.notebook.tab(2, state="normal")
        logger.info("Phase 2 tab enabled")
    
    def log_message(self, message, level="INFO"):
        """Send message to logs screen"""
        self.logs_screen.add_log(message, level)
    
    def on_closing(self):
        """Handle window close event"""
        if messagebox.askokcancel("Quitter", "Voulez-vous vraiment quitter l'application?"):
            logger.info("Application closing")
            self.root.destroy()
