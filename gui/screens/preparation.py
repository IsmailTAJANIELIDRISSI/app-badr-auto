#!/usr/bin/env python3
"""
Preparation Screen - Screen 1
Handles folder selection, script execution, and DS entry
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import os
import subprocess
import platform
from gui.utils.script_manager import ScriptManager
from gui.utils.file_utils import (detect_ltas, write_shipper_file, 
                                  get_lta_shipper_name, update_lta_shipper_name,
                                  get_lta_blocage_info, update_lta_blocage,
                                  find_lta_pdf)
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
        
        # Use the full detect_ltas function to get complete LTA information
        ltas = detect_ltas(folder)
        
        if not ltas:
            messagebox.showinfo("Information", "Aucun dossier LTA détecté")
            return
        
        self.all_ltas = [lta['name'] for lta in ltas]
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
        """Populate the LTA table with enhanced UI including shipper, location select, and blocage"""
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
        
        # Predefined locations
        self.locations = [
            "ISTAMBOUL ATATUR",
            "JEDDAH K/ABDUL A",
            "BAHREIN MOHARRAQ",
            "DOHA INT",
            "ABOU DHABI INT",
            "SHANGHAI PU DONG"
        ]
        
        # Create main container with scrolling
        main_container = ttk.Frame(self.table_frame)
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Canvas with scrollbar
        canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        self.table_frame_content = ttk.Frame(canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        canvas_window = canvas.create_window((0, 0), window=self.table_frame_content, anchor="nw")
        
        # Configure content columns
        self.table_frame_content.columnconfigure(0, weight=1)
        
        # Populate rows with card-style layout
        self.lta_inputs = []
        checkbox_vars = []
        checkbox_widgets = []
        
        for idx, lta in enumerate(ltas):
            # Card frame for each LTA
            card = ttk.LabelFrame(
                self.table_frame_content,
                text=f"  {lta['name']}  ",
                padding="15",
                relief="raised"
            )
            card.grid(row=idx, column=0, sticky=(tk.W, tk.E), padx=10, pady=8)
            card.columnconfigure(1, weight=1)
            
            row_num = 0
            
            # === ROW 1: Checkbox + LTA Reference ===
            cb_var = tk.BooleanVar(value=True)
            checkbox_vars.append(cb_var)
            cb = ttk.Checkbutton(
                card,
                text="Traiter ce LTA",
                variable=cb_var,
                state="disabled" if self.selection_mode.get() == "all" else "normal"
            )
            cb.grid(row=row_num, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
            checkbox_widgets.append(cb)
            
            # LTA Reference
            if lta.get('lta_reference'):
                ref_label = ttk.Label(
                    card,
                    text=f"Référence: {lta['lta_reference']}",
                    font=('Arial', 9, 'italic'),
                    foreground='gray'
                )
                ref_label.grid(row=row_num, column=2, sticky=tk.E, pady=(0, 10))
            
            row_num += 1
            
            # === ROW 2: Shipper Name with PDF Button ===
            ttk.Label(card, text="Expéditeur (Shipper):", font=('Arial', 9, 'bold')).grid(
                row=row_num, column=0, sticky=tk.W, pady=5
            )
            
            # Extract shipper from line 6 - use parent folder, not subfolder
            shipper_name = get_lta_shipper_name(folder, lta['name'])
            
            # Create frame for shipper entry + PDF button
            shipper_frame = ttk.Frame(card)
            shipper_frame.grid(row=row_num, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
            
            shipper_var = tk.StringVar(value=shipper_name)
            shipper_entry = ttk.Entry(shipper_frame, textvariable=shipper_var, width=45, font=('Arial', 9))
            shipper_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # PDF button to open MAWB PDF
            def open_lta_pdf(lta_folder=folder, lta_name=lta['name']):
                """Open the LTA MAWB PDF file"""
                pdf_path = find_lta_pdf(lta_folder, lta_name)
                if pdf_path:
                    try:
                        # Open PDF with default application
                        if platform.system() == 'Windows':
                            os.startfile(pdf_path)
                        elif platform.system() == 'Darwin':  # macOS
                            subprocess.run(['open', pdf_path])
                        else:  # Linux
                            subprocess.run(['xdg-open', pdf_path])
                        logger.info(f"Opened PDF: {pdf_path}")
                    except Exception as e:
                        messagebox.showerror("Erreur", f"Impossible d'ouvrir le PDF:\n{e}")
                        logger.error(f"Failed to open PDF {pdf_path}: {e}")
                else:
                    messagebox.showwarning("PDF Introuvable", 
                                        f"Le fichier PDF pour '{lta_name}' n'a pas été trouvé.\n\n"
                                        f"Vérifiez que le fichier existe avec le format:\n"
                                        f"{lta_name} - [MAWB].pdf")
            
            # Find PDF to determine if button should be enabled
            pdf_exists = find_lta_pdf(folder, lta['name']) is not None
            
            # Create professional PDF button
            if pdf_exists:
                pdf_button = tk.Button(
                    shipper_frame,
                    text="📄 Voir PDF",
                    font=('Segoe UI', 9),
                    fg='#ffffff',
                    bg='#007bff',
                    activebackground='#0056b3',
                    activeforeground='#ffffff',
                    relief='flat',
                    borderwidth=0,
                    cursor='hand2',
                    command=open_lta_pdf,
                    padx=12,
                    pady=6
                )
                
                # Hover effects for enabled button
                def on_enter(e, btn=pdf_button):
                    btn.config(bg='#0056b3')
                
                def on_leave(e, btn=pdf_button):
                    btn.config(bg='#007bff')
                
                pdf_button.bind('<Enter>', on_enter)
                pdf_button.bind('<Leave>', on_leave)
                
                # Tooltip for PDF button
                pdf_filename = os.path.basename(find_lta_pdf(folder, lta['name']))
                self._create_tooltip(pdf_button, f"Ouvrir {pdf_filename}")
            else:
                pdf_button = tk.Button(
                    shipper_frame,
                    text="📄 PDF",
                    font=('Segoe UI', 9, 'bold'),
                    fg='white',
                    bg='#e9ecef',
                    relief='flat',
                    borderwidth=0,
                    cursor='arrow',
                    state='disabled',
                    padx=12,
                    pady=6
                )
                self._create_tooltip(pdf_button, "PDF non disponible")
            pdf_button.pack(side=tk.LEFT, padx=(8, 0))
            
            row_num += 1
            
            # === ROW 3: DS Series ===
            ttk.Label(card, text="DS Série:", font=('Arial', 9, 'bold')).grid(
                row=row_num, column=0, sticky=tk.W, pady=5
            )
            
            ds_var = tk.StringVar(value=lta.get('ds_series', ''))
            ds_entry = ttk.Entry(card, textvariable=ds_var, width=20, font=('Arial', 9))
            ds_entry.grid(row=row_num, column=1, sticky=tk.W, pady=5, padx=(10, 0))
            
            # DS Validé status
            ds_status_var = tk.StringVar(value="✓" if lta.get('has_ds') else "✗")
            ds_status_label = ttk.Label(
                card,
                textvariable=ds_status_var,
                foreground="green" if lta.get('has_ds') else "gray",
                font=('Arial', 11, 'bold')
            )
            ds_status_label.grid(row=row_num, column=2, sticky=tk.E, pady=5)
            
            row_num += 1
            
            # === ROW 4: Location with Select/Custom ===
            ttk.Label(card, text="Lieu de Chargement:", font=('Arial', 9, 'bold')).grid(
                row=row_num, column=0, sticky=tk.W, pady=5
            )
            
            # Radio buttons for location mode
            loc_frame = ttk.Frame(card)
            loc_frame.grid(row=row_num, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
            
            loc_mode_var = tk.StringVar(value="select")
            loc_var = tk.StringVar(value=lta.get('location', ''))
            
            # Combobox for predefined locations
            loc_combo = ttk.Combobox(
                loc_frame,
                textvariable=loc_var,
                values=self.locations,
                width=25,
                font=('Arial', 9),
                state="readonly"
            )
            loc_combo.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            
            # Set to current value if exists
            if loc_var.get() and loc_var.get() in self.locations:
                loc_combo.set(loc_var.get())
            elif loc_var.get():
                loc_mode_var.set("custom")
            
            # Custom location entry (initially hidden)
            loc_custom_entry = ttk.Entry(loc_frame, textvariable=loc_var, width=25, font=('Arial', 9))
            
            # Radio buttons to switch mode
            def make_toggle_location(mode_var, combo, custom):
                def toggle():
                    if mode_var.get() == "select":
                        combo.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
                        custom.grid_remove()
                    else:
                        combo.grid_remove()
                        custom.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
                return toggle
            
            toggle_func = make_toggle_location(loc_mode_var, loc_combo, loc_custom_entry)
            
            ttk.Radiobutton(
                loc_frame,
                text="Liste",
                variable=loc_mode_var,
                value="select",
                command=toggle_func
            ).grid(row=0, column=1, padx=5)
            
            ttk.Radiobutton(
                loc_frame,
                text="Autre",
                variable=loc_mode_var,
                value="custom",
                command=toggle_func
            ).grid(row=0, column=2, padx=5)
            
            # Initialize visibility
            toggle_func()
            
            row_num += 1
            
            # === ROW 5: Blocage Section ===
            # Separator
            ttk.Separator(card, orient='horizontal').grid(
                row=row_num, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 10)
            )
            row_num += 1
            
            # Get blocage info - use parent folder
            blocage_info = get_lta_blocage_info(folder, lta['name'])
            
            blocage_var = tk.BooleanVar(value=blocage_info['is_blocage'])
            blocage_cb = ttk.Checkbutton(
                card,
                text="🔒 LTA Blocage",
                variable=blocage_var
            )
            blocage_cb.grid(row=row_num, column=0, columnspan=3, sticky=tk.W, pady=5)
            
            row_num += 1
            
            # Blocage weights frame (shown only if blocage checked)
            blocage_frame = ttk.Frame(card)
            blocage_frame.grid(row=row_num, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
            blocage_frame.columnconfigure(1, weight=1)
            blocage_frame.columnconfigure(3, weight=1)
            
            ttk.Label(blocage_frame, text="Poids Total:", font=('Arial', 9)).grid(
                row=0, column=0, sticky=tk.W, padx=(20, 5)
            )
            total_weight_var = tk.StringVar(value=blocage_info.get('total_weight', ''))
            total_weight_entry = ttk.Entry(
                blocage_frame,
                textvariable=total_weight_var,
                width=15,
                font=('Arial', 9)
            )
            total_weight_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
            
            ttk.Label(blocage_frame, text="Poids Bloqué:", font=('Arial', 9)).grid(
                row=0, column=2, sticky=tk.W, padx=(20, 5)
            )
            blocked_weight_var = tk.StringVar(value=blocage_info.get('blocked_weight', ''))
            blocked_weight_entry = ttk.Entry(
                blocage_frame,
                textvariable=blocked_weight_var,
                width=15,
                font=('Arial', 9)
            )
            blocked_weight_entry.grid(row=0, column=3, sticky=tk.W, padx=5)
            
            # Toggle blocage frame visibility - use factory function
            def make_toggle_blocage(bvar, bframe):
                def toggle():
                    if bvar.get():
                        bframe.grid()
                    else:
                        bframe.grid_remove()
                return toggle
            
            toggle_blocage_func = make_toggle_blocage(blocage_var, blocage_frame)
            blocage_cb.config(command=toggle_blocage_func)
            toggle_blocage_func()
            
            row_num += 1
            
            # === ROW 6: Partial LTA Section ===
            # Separator
            ttk.Separator(card, orient='horizontal').grid(
                row=row_num, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 10)
            )
            row_num += 1
            
            # Import partial functions
            from gui.utils.file_utils import get_lta_partial_info
            
            # Get partial info
            partial_info = get_lta_partial_info(folder, lta['name'])
            
            partial_var = tk.BooleanVar(value=partial_info is not None)
            partial_cb = ttk.Checkbutton(
                card,
                text="📦 LTA Partiel (Plusieurs vols)",
                variable=partial_var
            )
            partial_cb.grid(row=row_num, column=0, columnspan=2, sticky=tk.W, pady=5)
            
            # Configure button
            configure_btn = ttk.Button(
                card,
                text="⚙️ Configurer",
                command=lambda f=folder, l=lta['name'], p=partial_var: self._configure_partial(f, l, p)
            )
            configure_btn.grid(row=row_num, column=2, sticky=tk.E, pady=5)
            
            row_num += 1
            
            # Partial info display (shown only if configured)
            partial_info_frame = ttk.Frame(card)
            partial_info_frame.grid(row=row_num, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
            
            if partial_info:
                num_partials = len(partial_info.get('partials', []))
                info_text = f"✓ Configuré: {num_partials} partiel(s)"
                ttk.Label(partial_info_frame, text=info_text, font=('Arial', 9, 'italic'), foreground='green').grid(
                    row=0, column=0, sticky=tk.W, padx=(20, 0)
                )
            
            # Toggle configure button based on checkbox
            def make_toggle_partial(pvar, pbtn, pframe):
                def toggle():
                    if pvar.get():
                        pbtn.config(state="normal")
                        pframe.grid()
                    else:
                        pbtn.config(state="disabled")
                        pframe.grid_remove()
                return toggle
            
            toggle_partial_func = make_toggle_partial(partial_var, configure_btn, partial_info_frame)
            partial_cb.config(command=toggle_partial_func)
            toggle_partial_func()
            
            # Store all variables
            self.lta_inputs.append({
                'lta': lta,
                'checkbox_var': cb_var,
                'checkbox_widget': cb,
                'shipper_var': shipper_var,
                'ds_var': ds_var,
                'loc_var': loc_var,
                'loc_mode_var': loc_mode_var,
                'blocage_var': blocage_var,
                'total_weight_var': total_weight_var,
                'blocked_weight_var': blocked_weight_var,
                'partial_var': partial_var,
                'status_var': ds_status_var,
                'status_label': ds_status_label
            })
        
        self.lta_checkboxes = checkbox_vars
        self.checkbox_widgets = checkbox_widgets
        
        # Configure scroll region
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = canvas.winfo_width()
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        self.table_frame_content.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)
        
        # Mousewheel scrolling
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
        
        bind_mousewheel(self.table_frame_content)
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Button-4>", on_mousewheel)
        canvas.bind("<Button-5>", on_mousewheel)
        
        # Enable save button
        self.save_btn.config(state="normal")
        self.app.log_message(f"{len(ltas)} LTA(s) détecté(s) - Formulaire prêt", "SUCCESS")
    
    def save_and_continue(self):
        """Save DS information including shipper, location, and blocage data"""
        logger.info("Saving DS information...")
        self.app.log_message("Sauvegarde des informations DS...", "INFO")
        
        saved_count = 0
        error_count = 0
        skipped_count = 0
        folder = self.folder_var.get()
        
        for input_data in self.lta_inputs:
            lta = input_data['lta']
            ds_series = input_data['ds_var'].get().strip()
            location = input_data['loc_var'].get().strip()
            shipper_name = input_data['shipper_var'].get().strip()
            is_blocage = input_data['blocage_var'].get()
            total_weight = input_data['total_weight_var'].get().strip()
            blocked_weight = input_data['blocked_weight_var'].get().strip()
            
            logger.info(f"Processing LTA: {lta['name']}, DS: {ds_series}, Location: {location}, Shipper: {shipper_name}, Blocage: {is_blocage}")
            
            # Validate DS if provided
            if ds_series:
                ds_series = normalize_ds_series(ds_series)
                input_data['ds_var'].set(ds_series)
                
                is_valid, error_msg = validate_ds_series(ds_series)
                if not is_valid:
                    messagebox.showerror("Erreur de Validation", f"{lta['name']}: {error_msg}")
                    error_count += 1
                    continue
            
            # Validate location if provided
            if location:
                is_valid, error_msg = validate_location(location)
                if not is_valid:
                    messagebox.showerror("Erreur de Validation", f"{lta['name']}: {error_msg}")
                    error_count += 1
                    continue
            
            # Validate weights if blocage
            if is_blocage:
                if not total_weight or not blocked_weight:
                    messagebox.showerror(
                        "Erreur de Validation",
                        f"{lta['name']}: Les poids total et bloqué sont requis pour un LTA blocage"
                    )
                    error_count += 1
                    continue
                
                try:
                    float(total_weight)
                    float(blocked_weight)
                except ValueError:
                    messagebox.showerror(
                        "Erreur de Validation",
                        f"{lta['name']}: Les poids doivent être des nombres valides"
                    )
                    error_count += 1
                    continue
            
            try:
                # Update shipper name (line 6 of LTA file + line 1 of shipper file)
                # Use parent folder (folder) not subfolder
                if shipper_name:
                    update_lta_shipper_name(folder, lta['name'], shipper_name)
                
                # Update blocage information
                update_lta_blocage(
                    folder,
                    lta['name'],
                    is_blocage,
                    total_weight if is_blocage else '',
                    blocked_weight if is_blocage else ''
                )
                
                # Update DS and location in shipper file
                if lta.get('shipper_file'):
                    success = write_shipper_file(lta['shipper_file'], ds_series, location)
                    if success:
                        saved_count += 1
                        has_ds = bool(ds_series)
                        input_data['status_var'].set("✓" if has_ds else "✗")
                        input_data['status_label'].config(foreground="green" if has_ds else "gray")
                        logger.info(f"Successfully saved: {lta['name']}")
                    else:
                        error_count += 1
                        logger.error(f"Failed to save: {lta['name']}")
                else:
                    # No shipper file but still update LTA file
                    saved_count += 1
                    logger.info(f"Updated LTA file (no shipper): {lta['name']}")
                
            except Exception as e:
                logger.error(f"Error saving {lta['name']}: {e}")
                error_count += 1
        
        logger.info(f"Save summary: saved={saved_count}, errors={error_count}, skipped={skipped_count}")
        
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
    
    def _configure_partial(self, lta_folder_path, folder_name, partial_var):
        """Open partial LTA configuration dialog"""
        if not partial_var.get():
            messagebox.showinfo("Info", "Veuillez cocher la case 'LTA Partiel' d'abord")
            return
        
        # Import the partial configuration dialog
        from gui.screens.partial_config_dialog import PartialConfigDialog
        
        dialog = PartialConfigDialog(self.parent, lta_folder_path, folder_name)
        self.parent.wait_window(dialog.dialog)
        
        # Refresh the display to show updated info
        if dialog.config_saved:
            self.populate_lta_table()
            messagebox.showinfo("Succès", "Configuration partielle sauvegardée!")
    
    def _create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = tk.Label(
                tooltip,
                text=text,
                background="#ffffe0",
                relief=tk.SOLID,
                borderwidth=1,
                font=('Arial', 8)
            )
            label.pack()
            
            widget.tooltip = tooltip
        
        def hide_tooltip(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind('<Enter>', show_tooltip)
        widget.bind('<Leave>', hide_tooltip)