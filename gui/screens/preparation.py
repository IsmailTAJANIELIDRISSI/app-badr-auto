#!/usr/bin/env python3
"""
Preparation Screen - Screen 1
Handles folder selection, script execution, and DS entry
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
from gui.utils.script_manager import ScriptManager
from gui.utils.file_utils import detect_ltas, write_shipper_file
from gui.utils.validators import validate_ds_series, validate_location, validate_folder_path, normalize_ds_series

logger = logging.getLogger(__name__)


class PreparationScreen:
    """Screen 1: Preparation and DS Entry"""
    
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.frame = ttk.Frame(parent)
        self.script_manager = ScriptManager(app)
        self.lta_inputs = []
        self.table_frame_content = None
        self.all_ltas = []
        self.lta_checkboxes = []
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the UI components"""
        # Configure main frame grid
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(4, weight=1)  # Table row expands
        
        # Title
        title = ttk.Label(
            self.frame,
            text="ÉTAPE 1: PRÉPARATION DES FICHIERS",
            font=('Arial', 14, 'bold')
        )
        title.grid(row=0, column=0, columnspan=3, pady=10, padx=10, sticky=tk.W)
        
        # Folder selection
        folder_frame = ttk.LabelFrame(self.frame, text="Dossier Source", padding="10")
        folder_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=10)
        folder_frame.columnconfigure(0, weight=1)
        
        self.folder_var = tk.StringVar()
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var)
        folder_entry.grid(row=0, column=0, padx=5, sticky=(tk.W, tk.E))
        
        browse_btn = ttk.Button(folder_frame, text="Parcourir...", command=self.browse_folder)
        browse_btn.grid(row=0, column=1, padx=5)
        
        detect_btn = ttk.Button(folder_frame, text="🔍 Détecter LTAs", command=self.detect_and_show_ltas)
        detect_btn.grid(row=0, column=2, padx=5)
        
        # Selection mode
        selection_frame = ttk.LabelFrame(self.frame, text="Sélection des LTAs", padding="10")
        selection_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=10)
        
        self.selection_mode = tk.StringVar(value="all")
        ttk.Radiobutton(
            selection_frame,
            text="Tous les LTAs",
            variable=self.selection_mode,
            value="all",
            command=self._on_selection_mode_change
        ).grid(row=0, column=0, padx=10, pady=5)
        
        ttk.Radiobutton(
            selection_frame,
            text="Sélection manuelle",
            variable=self.selection_mode,
            value="manual",
            command=self._on_selection_mode_change
        ).grid(row=0, column=1, padx=10, pady=5)
        
        # Run preparation button (before table)
        self.run_btn = ttk.Button(
            self.frame,
            text="▶️ Exécuter Script de Préparation",
            command=self.run_preparation
        )
        self.run_btn.grid(row=3, column=0, columnspan=3, pady=10, padx=10)
        
        # LTA Table (expandable)
        self.table_frame = ttk.LabelFrame(
            self.frame,
            text="Informations DS (depuis email)",
            padding="10"
        )
        self.table_frame.grid(
            row=4,
            column=0,
            columnspan=3,
            sticky=(tk.W, tk.E, tk.N, tk.S),
            pady=10,
            padx=10
        )
        self.table_frame.columnconfigure(0, weight=1)
        self.table_frame.rowconfigure(0, weight=1)
        
        # Placeholder
        self.table_label = ttk.Label(
            self.table_frame,
            text="Cliquez 'Détecter LTAs' pour afficher le formulaire"
        )
        self.table_label.grid(row=0, column=0, pady=50)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            self.frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=10)
        
        # Status label
        self.status_var = tk.StringVar(value="En attente...")
        status_label = ttk.Label(self.frame, textvariable=self.status_var)
        status_label.grid(row=6, column=0, columnspan=3, pady=5, padx=10)
        
        # Save button
        self.save_btn = ttk.Button(
            self.frame,
            text="💾 Sauvegarder et Continuer",
            command=self.save_and_continue,
            state="disabled"
        )
        self.save_btn.grid(row=7, column=0, columnspan=3, pady=10, padx=10)
        
        # Info label
        info_label = ttk.Label(
            self.frame,
            text="ℹ️ Laissez vide si DS non disponible (Phase 2 only)",
            font=('Arial', 9, 'italic')
        )
        info_label.grid(row=8, column=0, columnspan=3, pady=5, padx=10)
    
    def browse_folder(self):
        """Open folder browser dialog and auto-detect LTAs"""
        folder = filedialog.askdirectory(title="Sélectionner le dossier LTA")
        if folder:
            self.folder_var.set(folder)
            self.app.current_folder = folder
            logger.info(f"Selected folder: {folder}")
            self.app.log_message(f"Dossier sélectionné: {folder}")
            self.detect_and_show_ltas()
    
    def detect_and_show_ltas(self):
        """Detect LTAs in the folder and show form"""
        folder = self.folder_var.get()
        
        is_valid, error_msg = validate_folder_path(folder)
        if not is_valid:
            messagebox.showwarning("Attention", error_msg)
            return
        
        self.app.current_folder = folder
        
        import os
        lta_folders = []
        for item in os.listdir(folder):
            item_path = os.path.join(folder, item)
            if os.path.isdir(item_path) and ("LTA" in item or "lta" in item):
                lta_folders.append(item)
        
        if not lta_folders:
            messagebox.showinfo("Information", "Aucun dossier LTA détecté")
            return
        
        self.all_ltas = sorted(lta_folders)
        self.app.log_message(f"{len(self.all_ltas)} dossiers LTA détectés", "INFO")
        self.populate_lta_table()
    
    def _on_selection_mode_change(self):
        """Handle selection mode change"""
        mode = self.selection_mode.get()
        
        if hasattr(self, 'checkbox_widgets') and self.checkbox_widgets:
            if mode == "all":
                for cb_var, cb_widget in zip(self.lta_checkboxes, self.checkbox_widgets):
                    cb_var.set(True)
                    cb_widget.config(state="disabled")
            else:
                for cb_widget in self.checkbox_widgets:
                    cb_widget.config(state="normal")
    
    def run_preparation(self):
        """Run the preparation script on selected LTAs"""
        folder = self.folder_var.get()
        
        is_valid, error_msg = validate_folder_path(folder)
        if not is_valid:
            messagebox.showwarning("Attention", error_msg)
            return
        
        mode = self.selection_mode.get()
        if mode == "manual" and self.lta_checkboxes:
            selected_indices = [i for i, var in enumerate(self.lta_checkboxes) if var.get()]
            if not selected_indices:
                messagebox.showwarning("Attention", "Aucun LTA sélectionné")
                return
            selected_ltas = [self.all_ltas[i] for i in selected_indices]
        else:
            selected_ltas = None
        
        self.run_btn.config(state="disabled")
        self.status_var.set("⏳ Démarrage...")
        self.app.set_status("Exécution du script de préparation...")
        logger.info("Starting preparation script...")
        self.app.log_message("Démarrage du script de préparation...", "INFO")
        
        def on_progress(percent, message):
            self.progress_var.set(percent)
            self.status_var.set(message)
            self.app.set_status(message)
            self.app.log_message(message, "INFO")
        
        def on_complete(success=True, error=None):
            self.run_btn.config(state="normal")
            if success:
                self.status_var.set("✅ Traitement terminé")
                self.app.log_message("Script de préparation terminé avec succès", "SUCCESS")
                self.populate_lta_table()
                messagebox.showinfo(
                    "Succès",
                    "Script de préparation terminé!\nVous pouvez vérifier/modifier les informations DS."
                )
            else:
                self.status_var.set(f"❌ Erreur: {error[:50]}")
                self.app.log_message(f"Erreur: {error}", "ERROR")
                messagebox.showerror("Erreur", f"Le script a échoué:\n{error}")
        
        self.script_manager.run_preparation(folder, on_progress, on_complete, selected_ltas)
    
    def populate_lta_table(self):
        """Populate the LTA table with improved scrolling and responsive design"""
        folder = self.folder_var.get()
        ltas = detect_ltas(folder)
        
        if not ltas:
            messagebox.showwarning("Attention", "Aucun LTA détecté dans ce dossier")
            return
        
        # Clear existing
        if self.table_label:
            self.table_label.destroy()
        if self.table_frame_content:
            self.table_frame_content.destroy()
        
        # Create main container with proper scrolling
        main_container = ttk.Frame(self.table_frame)
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        self.table_frame_content = ttk.Frame(canvas)
        
        # Configure canvas scrolling
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Grid layout
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Create window in canvas
        canvas_window = canvas.create_window((0, 0), window=self.table_frame_content, anchor="nw")
        
        # Configure content columns to be responsive
        self.table_frame_content.columnconfigure(0, weight=0, minsize=40)   # Checkbox
        self.table_frame_content.columnconfigure(1, weight=1, minsize=200)  # LTA name + reference
        self.table_frame_content.columnconfigure(2, weight=1, minsize=120)  # DS Series
        self.table_frame_content.columnconfigure(3, weight=2, minsize=180)  # Location
        self.table_frame_content.columnconfigure(4, weight=0, minsize=40)   # Status
        
        # Headers
        headers = [
            ("☑", 0, 40),
            ("LTA", 1, 200),
            ("DS Série", 2, 120),
            ("Lieu Chargement", 3, 180),
            ("✓", 4, 40)
        ]
        
        for text, col, width in headers:
            header = ttk.Label(
                self.table_frame_content,
                text=text,
                font=('Arial', 9, 'bold'),
                width=width//8
            )
            header.grid(row=0, column=col, padx=5, pady=5, sticky=tk.W)
        
        # Separator
        ttk.Separator(self.table_frame_content, orient='horizontal').grid(
            row=1, column=0, columnspan=5, sticky=(tk.W, tk.E), pady=5
        )
        
        # Populate rows
        self.lta_inputs = []
        checkbox_vars = []
        checkbox_widgets = []
        
        for idx, lta in enumerate(ltas):
            row = idx + 2
            
            # Checkbox
            cb_var = tk.BooleanVar(value=True)
            checkbox_vars.append(cb_var)
            cb = ttk.Checkbutton(
                self.table_frame_content,
                variable=cb_var,
                state="disabled" if self.selection_mode.get() == "all" else "normal"
            )
            cb.grid(row=row, column=0, padx=5, pady=3)
            checkbox_widgets.append(cb)
            
            # LTA name
            lta_display = f"{lta['name']} - {lta.get('lta_reference', 'N/A')}"
            lta_label = ttk.Label(
                self.table_frame_content,
                text=lta_display,
                width=25
            )
            lta_label.grid(row=row, column=1, padx=5, pady=3, sticky=tk.W)
            
            # DS Series input
            ds_var = tk.StringVar(value=lta.get('ds_series', ''))
            ds_entry = ttk.Entry(self.table_frame_content, textvariable=ds_var, width=15)
            ds_entry.grid(row=row, column=2, padx=5, pady=3, sticky=(tk.W, tk.E))
            
            # Location input
            loc_var = tk.StringVar(value=lta.get('location', ''))
            loc_entry = ttk.Entry(self.table_frame_content, textvariable=loc_var, width=25)
            loc_entry.grid(row=row, column=3, padx=5, pady=3, sticky=(tk.W, tk.E))
            
            # Status indicator
            status_var = tk.StringVar(value="✓" if lta.get('has_ds') else "✗")
            status_label = ttk.Label(
                self.table_frame_content,
                textvariable=status_var,
                foreground="green" if lta.get('has_ds') else "gray",
                font=('Arial', 10, 'bold')
            )
            status_label.grid(row=row, column=4, padx=5, pady=3)
            
            self.lta_inputs.append({
                'lta': lta,
                'ds_var': ds_var,
                'loc_var': loc_var,
                'status_var': status_var,
                'status_label': status_label,
                'checkbox_var': cb_var,
                'checkbox_widget': cb
            })
        
        self.lta_checkboxes = checkbox_vars
        self.checkbox_widgets = checkbox_widgets
        
        # Update scroll region when content changes
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Make canvas window width match canvas width
            canvas_width = canvas.winfo_width()
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        self.table_frame_content.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)
        
        # Fix mousewheel scrolling (cross-platform)
        def on_mousewheel(event):
            # Windows and MacOS have different delta values
            if event.num == 5 or event.delta < 0:
                canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta > 0:
                canvas.yview_scroll(-1, "units")
        
        # Bind mousewheel to canvas and all children
        def bind_mousewheel(widget):
            widget.bind("<MouseWheel>", on_mousewheel)  # Windows/MacOS
            widget.bind("<Button-4>", on_mousewheel)    # Linux scroll up
            widget.bind("<Button-5>", on_mousewheel)    # Linux scroll down
            for child in widget.winfo_children():
                bind_mousewheel(child)
        
        bind_mousewheel(self.table_frame_content)
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Button-4>", on_mousewheel)
        canvas.bind("<Button-5>", on_mousewheel)
        
        # Enable save button
        self.save_btn.config(state="normal")
        self.app.log_message(f"{len(ltas)} LTA(s) détecté(s) - Formulaire prêt", "SUCCESS")
    
    def save_and_continue(self):
        """Save DS information and enable Phase 1"""
        logger.info("Saving DS information...")
        self.app.log_message("Sauvegarde des informations DS...", "INFO")
        
        saved_count = 0
        error_count = 0
        
        for input_data in self.lta_inputs:
            lta = input_data['lta']
            ds_series = input_data['ds_var'].get().strip()
            location = input_data['loc_var'].get().strip()
            
            if ds_series:
                ds_series = normalize_ds_series(ds_series)
                input_data['ds_var'].set(ds_series)
                
                is_valid, error_msg = validate_ds_series(ds_series)
                if not is_valid:
                    messagebox.showerror("Erreur de Validation", f"{lta['name']}: {error_msg}")
                    error_count += 1
                    continue
            
            if location:
                is_valid, error_msg = validate_location(location)
                if not is_valid:
                    messagebox.showerror("Erreur de Validation", f"{lta['name']}: {error_msg}")
                    error_count += 1
                    continue
            
            if lta['shipper_file']:
                success = write_shipper_file(lta['shipper_file'], ds_series, location)
                if success:
                    saved_count += 1
                    has_ds = bool(ds_series)
                    input_data['status_var'].set("✓" if has_ds else "✗")
                    input_data['status_label'].config(foreground="green" if has_ds else "gray")
                else:
                    error_count += 1
        
        if error_count == 0:
            self.app.log_message(f"Informations sauvegardées pour {saved_count} LTA(s)", "SUCCESS")
            self.app.enable_phase1_tab()
            messagebox.showinfo(
                "Succès",
                f"Informations sauvegardées pour {saved_count} LTA(s)!\n"
                "Vous pouvez passer à la Phase 1."
            )
        else:
            self.app.log_message(f"Sauvegarde terminée avec {error_count} erreur(s)", "WARNING")
            messagebox.showwarning("Attention", f"Sauvegarde terminée avec {error_count} erreur(s)")