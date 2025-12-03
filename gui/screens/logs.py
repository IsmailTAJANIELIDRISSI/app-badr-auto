#!/usr/bin/env python3
"""
Logs Screen - Screen 4
Displays real-time logs and results from scripts
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
from datetime import datetime
import os

logger = logging.getLogger(__name__)

class LogsScreen:
    """Screen 4: Logs and Results with separate sections for each script"""
    
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.frame = ttk.Frame(parent)
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the UI components"""
        # Title
        title = ttk.Label(self.frame, text="📊 LOGS & RÉSULTATS", 
                         font=('Arial', 14, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10, sticky=tk.W)
        
        # Info label
        info_label = ttk.Label(self.frame, 
                              text="💡 Utilisez Ctrl+F dans chaque section pour rechercher",
                              font=('Arial', 9, 'italic'),
                              foreground="blue")
        info_label.grid(row=1, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        # Buttons frame
        btn_frame = ttk.Frame(self.frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        ttk.Button(btn_frame, text="📥 Ouvrir Dossier", 
                  command=self.open_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📄 Exporter Tous Logs", 
                  command=self.export_all_logs).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ Effacer Tous Logs", 
                  command=self.clear_all_logs).pack(side=tk.LEFT, padx=5)
        
        # Create notebook for different log sections
        self.logs_notebook = ttk.Notebook(self.frame)
        self.logs_notebook.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Tab 1: General App Logs
        self.app_log_text = self._create_log_tab("App Général")
        
        # Tab 2: Fuzzy Match Script Logs
        self.fuzzy_log_text = self._create_log_tab("Script Préparation")
        
        # Tab 3: BADR Login Script Logs
        self.badr_log_text = self._create_log_tab("Script BADR")
        
        # Configure grid weights
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(3, weight=1)
        
        # Add initial message
        self.add_log("Application démarrée", "INFO", "app")
    
    def _create_log_tab(self, tab_name):
        """Create a log tab with text widget and scrollbar"""
        # Create frame for this tab
        tab_frame = ttk.Frame(self.logs_notebook)
        self.logs_notebook.add(tab_frame, text=tab_name)
        
        # Create text widget with scrollbar
        scrollbar = ttk.Scrollbar(tab_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        log_text = tk.Text(tab_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                          font=('Consolas', 9))
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=log_text.yview)
        
        # Configure tags for different log levels
        log_text.tag_config("INFO", foreground="black")
        log_text.tag_config("WARNING", foreground="orange")
        log_text.tag_config("ERROR", foreground="red")
        log_text.tag_config("SUCCESS", foreground="green")
        log_text.tag_config("DEBUG", foreground="gray")
        
        # Enable Ctrl+F for search
        log_text.bind('<Control-f>', lambda e: self._show_search_dialog(log_text))
        
        return log_text
    
    def _show_search_dialog(self, text_widget):
        """Show search dialog for finding text"""
        search_window = tk.Toplevel(self.frame)
        search_window.title("Rechercher")
        search_window.geometry("400x100")
        
        ttk.Label(search_window, text="Rechercher:").pack(pady=5)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_window, textvariable=search_var, width=40)
        search_entry.pack(pady=5)
        search_entry.focus()
        
        def find_text():
            # Remove previous highlights
            text_widget.tag_remove("search", "1.0", tk.END)
            
            search_term = search_var.get()
            if not search_term:
                return
            
            # Search and highlight
            start_pos = "1.0"
            count = 0
            while True:
                start_pos = text_widget.search(search_term, start_pos, tk.END, nocase=True)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(search_term)}c"
                text_widget.tag_add("search", start_pos, end_pos)
                count += 1
                start_pos = end_pos
            
            # Configure search highlight
            text_widget.tag_config("search", background="yellow", foreground="black")
            
            # Scroll to first match
            if count > 0:
                text_widget.see("search.first")
                messagebox.showinfo("Recherche", f"{count} résultat(s) trouvé(s)")
            else:
                messagebox.showinfo("Recherche", "Aucun résultat trouvé")
        
        btn_frame = ttk.Frame(search_window)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Rechercher", command=find_text).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Fermer", command=search_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key
        search_entry.bind('<Return>', lambda e: find_text())
    
    def add_log(self, message, level="INFO", source="app"):
        """
        Add a log message to the appropriate text widget
        
        Args:
            message: Log message
            level: Log level (INFO, WARNING, ERROR, SUCCESS, DEBUG)
            source: Source of log (app, fuzzy, badr)
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # Determine which text widget to use
        if source == "fuzzy":
            text_widget = self.fuzzy_log_text
        elif source == "badr":
            text_widget = self.badr_log_text
        else:  # app or default
            text_widget = self.app_log_text
        
        text_widget.insert(tk.END, log_entry, level)
        text_widget.see(tk.END)  # Auto-scroll to bottom
        
        # Also log to file
        if level == "ERROR":
            logger.error(f"[{source}] {message}")
        elif level == "WARNING":
            logger.warning(f"[{source}] {message}")
        else:
            logger.info(f"[{source}] {message}")
    
    def add_script_output(self, output, source="badr"):
        """
        Add raw script output to the appropriate log section
        
        Args:
            output: Raw output from script
            source: Source script (fuzzy, badr)
        """
        # Determine which text widget to use
        if source == "fuzzy":
            text_widget = self.fuzzy_log_text
        elif source == "badr":
            text_widget = self.badr_log_text
        else:
            text_widget = self.app_log_text
        
        # Add output without timestamp (script already has its own format)
        text_widget.insert(tk.END, output)
        text_widget.see(tk.END)
    
    def open_folder(self):
        """Open the working folder in file explorer"""
        folder = self.app.current_folder
        if folder and os.path.exists(folder):
            os.startfile(folder)
            self.add_log(f"Ouverture du dossier: {folder}", "INFO", "app")
        else:
            messagebox.showwarning("Attention", "Aucun dossier sélectionné")
    
    def export_all_logs(self):
        """Export all logs to separate files"""
        folder = filedialog.askdirectory(title="Sélectionner le dossier d'export")
        
        if folder:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            try:
                # Export each log section
                logs_to_export = [
                    ("app_logs", self.app_log_text),
                    ("fuzzy_match_logs", self.fuzzy_log_text),
                    ("badr_script_logs", self.badr_log_text)
                ]
                
                exported_files = []
                for name, text_widget in logs_to_export:
                    filename = os.path.join(folder, f"{name}_{timestamp}.txt")
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(text_widget.get("1.0", tk.END))
                    exported_files.append(filename)
                
                self.add_log(f"Logs exportés vers: {folder}", "SUCCESS", "app")
                messagebox.showinfo("Succès", f"{len(exported_files)} fichiers de logs exportés!")
            except Exception as e:
                self.add_log(f"Erreur export logs: {e}", "ERROR", "app")
                messagebox.showerror("Erreur", f"Impossible d'exporter les logs: {e}")
    
    def clear_all_logs(self):
        """Clear all logs from all text widgets"""
        if messagebox.askyesno("Confirmation", "Effacer tous les logs de toutes les sections?"):
            self.app_log_text.delete("1.0", tk.END)
            self.fuzzy_log_text.delete("1.0", tk.END)
            self.badr_log_text.delete("1.0", tk.END)
            self.add_log("Tous les logs effacés", "INFO", "app")
