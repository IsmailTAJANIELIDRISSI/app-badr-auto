#!/usr/bin/env python3
"""
Phase 1 ED Screen - Screen 2
Handles Etat de Dépotage creation and signed series entry
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging

from gui.utils.script_manager import ScriptManager
from gui.utils.file_utils import detect_ltas, write_lta_signed_series
from gui.utils.validators import validate_credentials, validate_signed_series, normalize_ds_series

logger = logging.getLogger(__name__)

class Phase1EDScreen:
    """Screen 2: Phase 1 - Etat de Dépotage"""
    
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.frame = ttk.Frame(parent)
        self.script_manager = ScriptManager(app)
        self.ltas_with_ds = []
        self.lta_inputs = []
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the UI components"""
        # Configure main frame grid
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(4, weight=1)  # Table expands
        
        # Title
        title = ttk.Label(
            self.frame,
            text="PHASE 1: CRÉATION ETAT DE DÉPOTAGE",
            font=('Arial', 14, 'bold')
        )
        title.grid(row=0, column=0, pady=10, padx=10, sticky=tk.W)
        
        # Info label
        info_label = ttk.Label(
            self.frame,
            text="Cette phase créera les Etats de Dépotage pour les LTAs avec DS série",
            font=('Arial', 9, 'italic')
        )
        info_label.grid(row=1, column=0, pady=5, padx=10, sticky=tk.W)
        
        # Top button row
        top_btn_frame = ttk.Frame(self.frame)
        top_btn_frame.grid(row=2, column=0, pady=5, padx=10, sticky=tk.W)
        
        refresh_btn = ttk.Button(
            top_btn_frame,
            text="🔄 Actualiser",
            command=self.refresh_lta_list
        )
        refresh_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.start_btn = ttk.Button(
            top_btn_frame,
            text="▶️ Démarrer Phase 1",
            command=self.start_phase1
        )
        self.start_btn.pack(side=tk.LEFT)
        
        # Selection mode
        selection_frame = ttk.LabelFrame(self.frame, text="Mode de sélection", padding="10")
        selection_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5, padx=10)
        
        self.selection_mode = tk.StringVar(value="all")
        
        ttk.Radiobutton(
            selection_frame,
            text="Tous les LTAs avec DS",
            variable=self.selection_mode,
            value="all",
            command=self._on_selection_mode_change
        ).grid(row=0, column=0, padx=10, pady=2)
        
        ttk.Radiobutton(
            selection_frame,
            text="Sélection manuelle",
            variable=self.selection_mode,
            value="manual",
            command=self._on_selection_mode_change
        ).grid(row=0, column=1, padx=10, pady=2)
        
        # Main table frame (expandable)
        self.table_frame = ttk.LabelFrame(
            self.frame,
            text="LTAs et Séries Signées",
            padding="10"
        )
        self.table_frame.grid(
            row=4,
            column=0,
            sticky=(tk.W, tk.E, tk.N, tk.S),
            pady=10,
            padx=10
        )
        self.table_frame.columnconfigure(0, weight=1)
        self.table_frame.rowconfigure(0, weight=1)
        
        # Placeholder
        self.table_label = ttk.Label(
            self.table_frame,
            text="Cliquez 'Actualiser' pour charger les LTAs"
        )
        self.table_label.grid(row=0, column=0, pady=50)
        
        # Progress
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            self.frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=5, padx=10)
        
        # Status
        self.status_var = tk.StringVar(value="En attente...")
        ttk.Label(self.frame, textvariable=self.status_var).grid(row=6, column=0, pady=5, padx=10)
        
        # Bottom action buttons
        bottom_btn_frame = ttk.Frame(self.frame)
        bottom_btn_frame.grid(row=7, column=0, pady=10, padx=10)
        
        self.save_btn = ttk.Button(
            bottom_btn_frame,
            text="💾 Enregistrer Séries",
            command=self.save_signed_series,
            state="disabled"
        )
        self.save_btn.pack(side=tk.LEFT, padx=5, ipadx=15, ipady=3)
        
        phase2_btn = ttk.Button(
            bottom_btn_frame,
            text="▶️ Continuer vers Phase 2",
            command=self._go_to_phase2
        )
        phase2_btn.pack(side=tk.LEFT, padx=5, ipadx=15, ipady=3)
        
        # Info label at bottom
        info_label = ttk.Label(
            self.frame,
            text="ℹ️ Entrez les séries signées après avoir créé les Etats de Dépotage",
            font=('Arial', 9, 'italic')
        )
        info_label.grid(row=8, column=0, pady=5, padx=10)
        
        # Initialize credential variables
        self.username_var = tk.StringVar(value="BK707345")
        self.password_var = tk.StringVar()
    
    def refresh_lta_list(self):
        """Refresh the list of LTAs with DS series and show unified table"""
        if not self.app.current_folder:
            messagebox.showwarning(
                "Attention",
                "Aucun dossier sélectionné.\nVeuillez d'abord exécuter la préparation."
            )
            return
        
        # Detect LTAs
        all_ltas = detect_ltas(self.app.current_folder)
        
        # Filter only those with DS series
        self.ltas_with_ds = [lta for lta in all_ltas if lta.get('has_ds')]
        
        if not self.ltas_with_ds:
            messagebox.showinfo(
                "Information",
                "Aucun LTA avec DS série détecté.\n"
                "Veuillez d'abord compléter la préparation (Étape 1)."
            )
            self.start_btn.config(state="disabled")
            return
        
        self.app.log_message(
            f"Phase 1: {len(self.ltas_with_ds)} LTA(s) avec DS série détecté(s)",
            "INFO"
        )
        
        # Show unified table
        self.populate_lta_table()
        self.start_btn.config(state="normal")
        self.save_btn.config(state="normal")
    
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
    
    def populate_lta_table(self):
        """Populate unified table with LTAs and all their information in one view"""
        # Clear existing
        if self.table_label:
            self.table_label.destroy()
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        # Create main container
        main_container = ttk.Frame(self.table_frame)
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        table_content = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        canvas_window = canvas.create_window((0, 0), window=table_content, anchor="nw")
        
        # Configure content columns to be responsive
        table_content.columnconfigure(0, weight=0, minsize=40)   # Checkbox
        table_content.columnconfigure(1, weight=1, minsize=200)  # LTA name + reference
        table_content.columnconfigure(2, weight=1, minsize=120)  # DS Série
        table_content.columnconfigure(3, weight=1, minsize=120)  # DS Validé
        table_content.columnconfigure(4, weight=0, minsize=80)   # Copy button
        table_content.columnconfigure(5, weight=1, minsize=120)  # Série Signée
        table_content.columnconfigure(6, weight=0, minsize=40)   # Status
        
        # Headers
        headers = [
            ("☑", 0),
            ("LTA", 1),
            ("DS Série", 2),
            ("DS Validé", 3),
            ("Action", 4),
            ("Série Signée", 5),
            ("✓", 6)
        ]
        
        for text, col in headers:
            header = ttk.Label(
                table_content,
                text=text,
                font=('Arial', 9, 'bold')
            )
            header.grid(row=0, column=col, padx=5, pady=5, sticky=tk.W)
        
        # Separator
        ttk.Separator(table_content, orient='horizontal').grid(
            row=1, column=0, columnspan=7, sticky=(tk.W, tk.E), pady=5
        )
        
        # Populate rows
        self.lta_inputs = []
        self.lta_checkboxes = []
        self.checkbox_widgets = []
        
        current_row = 2
        for idx, lta in enumerate(self.ltas_with_ds):
            # Checkbox
            cb_var = tk.BooleanVar(value=True)
            self.lta_checkboxes.append(cb_var)
            cb = ttk.Checkbutton(
                table_content,
                variable=cb_var,
                state="disabled" if self.selection_mode.get() == "all" else "normal"
            )
            cb.grid(row=current_row, column=0, padx=5, pady=3, sticky=tk.N)
            self.checkbox_widgets.append(cb)
            
            # Check if partial LTA
            if lta.get('is_partial') and lta.get('partial_config'):
                # PARTIAL LTA - Create expandable rows for each partial
                partials = lta['partial_config'].get('partials', [])
                
                # Main LTA header row
                lta_display = f"{lta['name']} - {lta.get('lta_reference', 'N/A')} 📦 ({len(partials)} partiels)"
                lta_label = ttk.Label(
                    table_content, 
                    text=lta_display,
                    font=('Arial', 9, 'bold')
                )
                lta_label.grid(row=current_row, column=1, padx=5, pady=3, sticky=tk.W, rowspan=len(partials))
                current_row += 1
                
                # Create a row for each partial
                for partial in partials:
                    partial_num = partial['partial_number']
                    
                    # Indent for sub-item
                    ttk.Label(
                        table_content,
                        text=f"  └ Partiel {partial_num}",
                        foreground="gray"
                    ).grid(row=current_row, column=1, padx=20, pady=2, sticky=tk.W)
                    
                    # DS Série for this partial
                    ds_serie = f"{partial['ds_serie']}/{partial['ds_cle']}"
                    ttk.Label(
                        table_content,
                        text=ds_serie,
                        foreground="purple"
                    ).grid(row=current_row, column=2, padx=5, pady=2, sticky=tk.W)
                    
                    # DS Validé placeholder (will be filled after Phase 1)
                    ttk.Label(
                        table_content,
                        text="En attente",
                        foreground="gray",
                        font=('Arial', 9, 'italic')
                    ).grid(row=current_row, column=3, padx=5, pady=2, sticky=tk.W)
                    
                    # Signed series input for this partial
                    signed_var = tk.StringVar(value="")
                    signed_entry = ttk.Entry(table_content, textvariable=signed_var, width=15)
                    signed_entry.grid(row=current_row, column=5, padx=5, pady=2, sticky=(tk.W, tk.E))
                    
                    # Store input for later saving
                    self.lta_inputs.append({
                        'lta': lta,
                        'partial_number': partial_num,
                        'signed_var': signed_var,
                        'is_partial': True
                    })
                    
                    # Status icon
                    ttk.Label(
                        table_content,
                        text="⏸️",
                        font=('Arial', 11),
                        foreground="gray"
                    ).grid(row=current_row, column=6, padx=5, pady=2)
                    
                    current_row += 1
                
            else:
                # REGULAR LTA - Single row
                lta_display = f"{lta['name']} - {lta.get('lta_reference', 'N/A')}"
                lta_label = ttk.Label(table_content, text=lta_display)
                lta_label.grid(row=current_row, column=1, padx=5, pady=3, sticky=tk.W)
                
                # DS Série
                ds_serie = lta.get('ds_series', 'N/A')
                ds_label = ttk.Label(
                    table_content,
                    text=ds_serie,
                    foreground="blue" if ds_serie != 'N/A' else "gray"
                )
                ds_label.grid(row=current_row, column=2, padx=5, pady=3, sticky=tk.W)
                
                # DS Validé
                validated_ds = lta.get('validated_ds', '')
                if validated_ds:
                    validated_label = ttk.Label(
                        table_content,
                        text=validated_ds,
                        font=('Arial', 9, 'bold'),
                        foreground="darkgreen"
                    )
                    validated_label.grid(row=current_row, column=3, padx=5, pady=3, sticky=tk.W)
                    
                    # Copy button
                    def make_copy_command(text, lta_name):
                        def copy_to_clipboard():
                            self.frame.clipboard_clear()
                            self.frame.clipboard_append(text)
                            self.app.log_message(f"DS copié: {text} ({lta_name})", "SUCCESS")
                        return copy_to_clipboard
                    
                    copy_btn = ttk.Button(
                        table_content,
                        text="📋 Copier",
                        width=10,
                        command=make_copy_command(validated_ds, lta['name'])
                    )
                    copy_btn.grid(row=current_row, column=4, padx=5, pady=3)
                else:
                    ttk.Label(
                        table_content,
                        text="En attente",
                        foreground="gray",
                        font=('Arial', 9, 'italic')
                    ).grid(row=current_row, column=3, padx=5, pady=3, sticky=tk.W)
                
                # Signed series input
                existing_signed = lta.get('signed_ds', '')
                signed_var = tk.StringVar(value=existing_signed)
                signed_entry = ttk.Entry(table_content, textvariable=signed_var, width=15)
                signed_entry.grid(row=current_row, column=5, padx=5, pady=3, sticky=(tk.W, tk.E))
                
                # Store input
                self.lta_inputs.append({
                    'lta': lta,
                    'signed_var': signed_var,
                    'is_partial': False
                })
                
                # Status icon
                status_var = tk.StringVar(value="✅" if existing_signed else "⏸️")
                status_label = ttk.Label(
                    table_content,
                    textvariable=status_var,
                    font=('Arial', 11),
                    foreground="green" if existing_signed else "gray"
                )
                status_label.grid(row=current_row, column=6, padx=5, pady=3)
                
                current_row += 1
        
        # Configure scrolling
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = canvas.winfo_width()
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        table_content.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)
        
        # Fix mousewheel scrolling
        def on_mousewheel(event):
            if event.num == 5 or event.delta < 0:
                canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta > 0:
                canvas.yview_scroll(-1, "units")
        
        def bind_mousewheel(widget):
            widget.bind("<MouseWheel>", on_mousewheel)
            widget.bind("<Button-4>", on_mousewheel)
            widget.bind("<Button-5>", on_mousewheel)
            for child in widget.winfo_children():
                bind_mousewheel(child)
        
        bind_mousewheel(table_content)
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Button-4>", on_mousewheel)
        canvas.bind("<Button-5>", on_mousewheel)
        
        self.app.log_message("Table Phase 1 affichée", "INFO")
    
    def start_phase1(self):
        """Start Phase 1 automation"""
        username = self.username_var.get()
        password = self.password_var.get()
        
        if not self.ltas_with_ds:
            messagebox.showwarning(
                "Attention",
                "Aucun LTA à traiter.\nCliquez sur 'Actualiser' pour charger les LTAs."
            )
            return
        
        # Determine selection
        mode = self.selection_mode.get()
        
        if mode == "all":
            lta_selection = "all"
            selected_lta_names = None  # None means all
            selected_count = len(self.ltas_with_ds)
        else:
            selected_indices = [i for i, var in enumerate(self.lta_checkboxes) if var.get()]
            
            if not selected_indices:
                messagebox.showwarning(
                    "Attention",
                    "Aucun LTA sélectionné.\nVeuillez sélectionner au moins un LTA."
                )
                return
            
            # Get folder names of selected LTAs
            selected_lta_names = [self.ltas_with_ds[i]['name'] for i in selected_indices]
            lta_selection = selected_indices
            selected_count = len(selected_indices)
        
        # Confirm start
        if not messagebox.askyesno(
            "Confirmation",
            f"Démarrer Phase 1 pour {selected_count} LTA(s)?"
        ):
            return
        
        self.start_btn.config(state="disabled")
        self.status_var.set("⏳ Démarrage Phase 1...")
        self.app.set_status("Exécution Phase 1...")
        logger.info("Starting Phase 1...")
        self.app.log_message(f"Démarrage de la Phase 1 pour {selected_count} LTA(s)...", "INFO")
        
        def on_progress(percent, message):
            self.progress_var.set(percent)
            self.status_var.set(message)
            self.app.set_status(message)
            self.app.log_message(message, "INFO")
        
        def on_complete(success=True, error=None):
            self.start_btn.config(state="normal")
            
            if success:
                self.status_var.set("✅ Phase 1 terminée")
                messagebox.showinfo(
                    "Succès",
                    "Phase 1 terminée!\nLes DS validés apparaissent dans la colonne 'DS Validé'.\n"
                    "Entrez maintenant les séries signées dans la colonne 'Série Signée'.",
                    parent=self.frame
                )
                # Refresh table to show validated DS
                self.refresh_lta_list()
            else:
                self.status_var.set(f"❌ Erreur: {error[:50] if error else 'Unknown'}")
                self.app.log_message(f"Erreur Phase 1: {error}", "ERROR")
                messagebox.showerror("Erreur", f"Phase 1 a échoué:\n{error}")
        
        credentials = {'username': username, 'password': password}
        self.script_manager.run_phase1(
            self.app.current_folder,
            credentials,
            lta_selection,
            on_progress,
            on_complete,
            selected_lta_names=selected_lta_names  # Pass folder names for filtering
        )
    
    def save_signed_series(self):
        """Save signed series to LTA files"""
        from gui.utils.file_utils import update_partial_signed_series
        
        logger.info("Saving signed series...")
        self.app.log_message("Sauvegarde des séries signées...", "INFO")
        
        saved_count = 0
        error_count = 0
        
        for input_data in self.lta_inputs:
            lta = input_data['lta']
            signed_series = input_data['signed_var'].get().strip()
            is_partial = input_data.get('is_partial', False)
            
            if not signed_series:
                continue
            
            # Normalize format
            signed_series = normalize_ds_series(signed_series)
            input_data['signed_var'].set(signed_series)
            
            # Validate
            is_valid, error_msg = validate_signed_series(signed_series)
            if not is_valid:
                messagebox.showerror(
                    "Erreur de Validation",
                    f"{lta['name']}: {error_msg}"
                )
                error_count += 1
                continue
            
            # Save based on LTA type
            if is_partial:
                # Save to partial config JSON
                partial_number = input_data['partial_number']
                success = update_partial_signed_series(
                    self.app.current_folder,
                    lta['name'],
                    partial_number,
                    signed_series
                )
                if success:
                    saved_count += 1
                    self.app.log_message(
                        f"✓ Série signée sauvegardée: {lta['name']} Partiel {partial_number} → {signed_series}",
                        "SUCCESS"
                    )
                else:
                    error_count += 1
            else:
                # Save to regular LTA file (line 8)
                if lta['lta_file']:
                    success = write_lta_signed_series(lta['lta_file'], signed_series)
                    if success:
                        saved_count += 1
                        self.app.log_message(
                            f"✓ Série signée sauvegardée: {lta['name']} → {signed_series}",
                            "SUCCESS"
                        )
                    else:
                        error_count += 1
        
        # Results
        if error_count == 0:
            self.app.log_message(
                f"Séries signées sauvegardées pour {saved_count} LTA(s)",
                "SUCCESS"
            )
            messagebox.showinfo(
                "Succès",
                f"Séries signées sauvegardées pour {saved_count} LTA(s)!\n"
                "Vous pouvez maintenant continuer vers la Phase 2."
            )
        else:
            self.app.log_message(
                f"Sauvegarde terminée avec {error_count} erreur(s)",
                "WARNING"
            )
            messagebox.showwarning(
                "Attention",
                f"Sauvegarde terminée avec {error_count} erreur(s)"
            )
    
    def _go_to_phase2(self):
        """Enable Phase 2 tab and switch to it"""
        self.app.enable_phase2_tab()
        self.app.notebook.select(2)
        self.app.log_message("Passage à Phase 2", "INFO")