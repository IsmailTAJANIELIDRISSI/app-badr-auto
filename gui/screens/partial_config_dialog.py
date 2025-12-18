#!/usr/bin/env python3
"""
Partial LTA Configuration Dialog
Allows users to configure partial LTA processing
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import glob
import logging
from openpyxl import load_workbook
from gui.utils.file_utils import get_lta_partial_info, save_lta_partial_config

logger = logging.getLogger(__name__)


class PartialConfigDialog:
    """Dialog for configuring partial LTA processing"""
    
    def __init__(self, parent, lta_folder_path, folder_name):
        self.parent = parent
        self.lta_folder_path = lta_folder_path
        self.folder_name = folder_name
        self.config_saved = False
        
        # Load existing config if available
        self.existing_config = get_lta_partial_info(lta_folder_path, folder_name)
        
        # Load LTA data from generated_excel
        self.lta_data = self._load_lta_data()
        
        if not self.lta_data:
            messagebox.showerror("Erreur", "Impossible de charger les données LTA.\nVeuillez exécuter le script de préparation d'abord.")
            return
        
        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Configuration Partielle - {folder_name}")
        self.dialog.geometry("800x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._setup_ui()
    
    def _load_lta_data(self):
        """Load LTA data from generated_excel file"""
        try:
            lta_subfolder = os.path.join(self.lta_folder_path, self.folder_name)
            excel_files = glob.glob(os.path.join(lta_subfolder, "generated_excel*.xlsx"))
            
            if not excel_files:
                logger.error(f"No generated_excel file found in {lta_subfolder}")
                messagebox.showwarning(
                    "Fichier introuvable",
                    f"Le fichier 'generated_excel' n'a pas été trouvé dans:\n{lta_subfolder}\n\n"
                    "Veuillez exécuter la détection LTA d'abord."
                )
                return None
            
            logger.info(f"Loading LTA data from: {excel_files[0]}")
            wb = load_workbook(excel_files[0], data_only=True)
            
            # Check if Summary sheet exists
            if 'Summary' not in wb.sheetnames:
                logger.error(f"Summary sheet not found. Available sheets: {wb.sheetnames}")
                wb.close()
                messagebox.showerror(
                    "Erreur",
                    f"La feuille 'Summary' n'existe pas dans le fichier Excel.\n\n"
                    f"Feuilles disponibles: {', '.join(wb.sheetnames)}"
                )
                return None
            
            ws = wb['Summary']
            
            # Get total weight and positions from Summary sheet
            # Data is in column A (labels) and column B (values)
            total_weight = None
            total_positions = None
            
            # Search for "P,BRUT" and "P" labels in column A (rows 1-10)
            for row in range(1, 15):
                cell_a = ws[f'A{row}'].value
                if cell_a:
                    cell_a_str = str(cell_a).strip().upper()
                    if 'P,BRUT' in cell_a_str or 'P.BRUT' in cell_a_str:
                        val = ws[f'B{row}'].value
                        if val and isinstance(val, (int, float)):
                            total_weight = val
                            logger.info(f"Found total weight at B{row}: {total_weight}")
                    elif cell_a_str == 'P' and not total_positions:  # P for positions (before P,BRUT in file)
                        val = ws[f'B{row}'].value
                        if val and isinstance(val, (int, float)):
                            total_positions = val
                            logger.info(f"Found total positions at B{row}: {total_positions}")
            
            logger.info(f"Total weight: {total_weight}, Total positions: {total_positions}")
            
            # Count DUMs by checking C11, C18, C25... (DUM labels in column C)
            # DUM data structure: 
            # Row N: "DUM X" in column C
            # Row N+1: P (positions) - label in A, value in B
            # Row N+2: V (value) - label in A, value in B  
            # Row N+3: P,NET - label in A, value in B
            # Row N+4: P,BRUT (weight) - label in A, value in B
            dums = []
            for dum_idx in range(1, 10):
                row_num = 11 + (dum_idx - 1) * 7
                cell_value = ws[f'C{row_num}'].value
                
                if cell_value and 'DUM' in str(cell_value).upper():
                    # Get DUM positions and weight from column A (labels) and B (values)
                    # P is at row_num + 1, P,BRUT is at row_num + 4
                    dum_positions_row = row_num + 1  # P is 1 row below DUM label
                    dum_weight_row = row_num + 4     # P,BRUT is 4 rows below DUM label
                    
                    dum_positions = ws[f'B{dum_positions_row}'].value or 0
                    dum_weight = ws[f'B{dum_weight_row}'].value or 0
                    
                    logger.info(f"DUM {dum_idx} (row {row_num}): weight={dum_weight}, positions={dum_positions}")
                    
                    dums.append({
                        'number': dum_idx,
                        'weight': float(dum_weight) if dum_weight else 0,
                        'positions': int(dum_positions) if dum_positions else 0
                    })
                else:
                    break
            
            wb.close()
            
            logger.info(f"Loaded {len(dums)} DUMs")
            
            return {
                'total_weight': float(total_weight) if total_weight else 0,
                'total_positions': int(total_positions) if total_positions else 0,
                'dums': dums
            }
            
        except Exception as e:
            logger.error(f"Error loading LTA data: {e}", exc_info=True)
            messagebox.showerror(
                "Erreur",
                f"Erreur lors du chargement des données LTA:\n{str(e)}\n\n"
                f"Vérifiez le fichier Excel 'generated_excel' dans:\n{lta_subfolder}"
            )
            return None
    
    def _setup_ui(self):
        """Setup the dialog UI"""
        # Main container with scrollbar
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title = ttk.Label(
            main_frame,
            text=f"📦 Configuration LTA Partiel: {self.folder_name}",
            font=('Arial', 12, 'bold')
        )
        title.pack(pady=(0, 10))
        
        # LTA Totals
        totals_frame = ttk.LabelFrame(main_frame, text="Totaux LTA", padding="10")
        totals_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(totals_frame, text=f"Poids Total: {self.lta_data['total_weight']} kg").pack(anchor=tk.W)
        ttk.Label(totals_frame, text=f"Positions Totales: {self.lta_data['total_positions']}").pack(anchor=tk.W)
        ttk.Label(totals_frame, text=f"Nombre de DUMs: {len(self.lta_data['dums'])}").pack(anchor=tk.W)
        
        # Number of partials
        partials_frame = ttk.LabelFrame(main_frame, text="Nombre de Partiels", padding="10")
        partials_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(partials_frame, text="Nombre de vols partiels:").grid(row=0, column=0, padx=5)
        self.num_partials_var = tk.IntVar(value=2)
        num_partials_spinbox = ttk.Spinbox(
            partials_frame,
            from_=2,
            to=5,
            textvariable=self.num_partials_var,
            width=10
        )
        num_partials_spinbox.grid(row=0, column=1, padx=5)
        
        generate_btn = ttk.Button(
            partials_frame,
            text="Générer Formulaire",
            command=self._generate_partial_forms
        )
        generate_btn.grid(row=0, column=2, padx=10)
        
        # Exception case warning frame (initially hidden)
        self.exception_frame = ttk.LabelFrame(main_frame, text="⚠️ CAS D'EXCEPTION DÉTECTÉ", padding="10")
        self.exception_frame.pack(fill=tk.X, pady=5)
        self.exception_frame.pack_forget()  # Hide initially
        
        exception_info = ttk.Label(
            self.exception_frame,
            text="Un partiel a un poids inférieur au plus petit DUM.\n"
                 "Veuillez renseigner les informations de référence ci-dessous:",
            foreground="red",
            font=('Arial', 9, 'bold')
        )
        exception_info.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(self.exception_frame, text="Référence créée à l'aéroport:", font=('Arial', 9, 'bold')).grid(
            row=1, column=0, sticky=tk.W, padx=5, pady=2
        )
        self.airport_reference_var = tk.StringVar(value="")
        airport_ref_entry = ttk.Entry(self.exception_frame, textvariable=self.airport_reference_var, width=25)
        airport_ref_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(self.exception_frame, text="(ex: 157-41680645)", font=('Arial', 8, 'italic')).grid(
            row=1, column=2, sticky=tk.W, padx=5, pady=2
        )
        
        ttk.Label(self.exception_frame, text="Positions du plus petit partiel:", font=('Arial', 9, 'bold')).grid(
            row=2, column=0, sticky=tk.W, padx=5, pady=2
        )
        self.smallest_partial_positions_var = tk.StringVar(value="")
        positions_entry = ttk.Entry(self.exception_frame, textvariable=self.smallest_partial_positions_var, width=10)
        positions_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(self.exception_frame, text="(nombre de positions)", font=('Arial', 8, 'italic')).grid(
            row=2, column=2, sticky=tk.W, padx=5, pady=2
        )
        
        # Partials container (scrollable)
        self.partials_container = ttk.Frame(main_frame)
        self.partials_container.pack(fill=tk.BOTH, expand=True, pady=10)
        
        canvas = tk.Canvas(self.partials_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.partials_container, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas = canvas
        
        # Buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            buttons_frame,
            text="💾 Sauvegarder",
            command=self._save_config
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            buttons_frame,
            text="❌ Annuler",
            command=self.dialog.destroy
        ).pack(side=tk.LEFT, padx=5)
        
        # Load existing config if available
        if self.existing_config:
            self.num_partials_var.set(len(self.existing_config['partials']))
            # Load exception case data if exists
            if self.existing_config.get('partial_type') == 'exception':
                self.airport_reference_var.set(self.existing_config.get('smallest_partial_airport_reference', ''))
                self.smallest_partial_positions_var.set(str(self.existing_config.get('smallest_partial_positions', '')))
            self._generate_partial_forms(load_existing=True)
    
    def _generate_partial_forms(self, load_existing=False):
        """Generate forms for each partial"""
        # Clear existing forms
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.partial_forms = []
        num_partials = self.num_partials_var.get()
        
        for i in range(num_partials):
            partial_num = i + 1
            
            # Load existing data if available
            existing_data = None
            if load_existing and self.existing_config:
                for p in self.existing_config['partials']:
                    if p['partial_number'] == partial_num:
                        existing_data = p
                        break
            
            frame = self._create_partial_form(partial_num, existing_data)
            frame.pack(fill=tk.X, pady=5, padx=10)
            
        # Bind mousewheel
        def on_mousewheel(event):
            if event.num == 5 or event.delta < 0:
                self.canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
        
        self.canvas.bind("<MouseWheel>", on_mousewheel)
        self.canvas.bind("<Button-4>", on_mousewheel)
        self.canvas.bind("<Button-5>", on_mousewheel)
    
    def _create_partial_form(self, partial_num, existing_data=None):
        """Create form for a single partial"""
        frame = ttk.LabelFrame(
            self.scrollable_frame,
            text=f"Partiel {partial_num}",
            padding="10"
        )
        
        # Weight
        ttk.Label(frame, text="Poids (kg):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        weight_var = tk.StringVar(value=existing_data['weight'] if existing_data else "")
        weight_entry = ttk.Entry(frame, textvariable=weight_var, width=15)
        weight_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Calculated positions (read-only, will be auto-calculated)
        ttk.Label(frame, text="Positions (auto):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        positions_var = tk.StringVar(value="")
        positions_label = ttk.Label(frame, textvariable=positions_var, foreground="blue")
        positions_label.grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)
        
        # DS Serie
        ttk.Label(frame, text="DS Série:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ds_serie_var = tk.StringVar(value=existing_data['ds_serie'] if existing_data else "")
        ds_serie_entry = ttk.Entry(frame, textvariable=ds_serie_var, width=15)
        ds_serie_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # DS Cle
        ttk.Label(frame, text="DS Clé:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ds_cle_var = tk.StringVar(value=existing_data['ds_cle'] if existing_data else "")
        ds_cle_entry = ttk.Entry(frame, textvariable=ds_cle_var, width=5)
        ds_cle_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Loading Location
        ttk.Label(frame, text="Lieu de Chargement:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        location_var = tk.StringVar(value=existing_data['loading_location'] if existing_data else "")
        location_entry = ttk.Entry(frame, textvariable=location_var, width=30)
        location_entry.grid(row=3, column=1, columnspan=3, sticky=tk.W, padx=5, pady=2)
        
        # DUM Distribution Preview (read-only text widget)
        ttk.Label(frame, text="Distribution DUMs (auto):").grid(row=4, column=0, sticky=tk.NW, padx=5, pady=2)
        
        dums_text = tk.Text(frame, height=6, width=50, state='disabled', wrap=tk.WORD)
        dums_text.grid(row=4, column=1, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        # Scrollbar for DUM distribution
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=dums_text.yview)
        scrollbar.grid(row=4, column=4, sticky=(tk.N, tk.S))
        dums_text.configure(yscrollcommand=scrollbar.set)
        
        self.partial_forms.append({
            'partial_number': partial_num,
            'weight_var': weight_var,
            'positions_var': positions_var,
            'ds_serie_var': ds_serie_var,
            'ds_cle_var': ds_cle_var,
            'location_var': location_var,
            'dums_text': dums_text
        })
        
        # Trace weight changes to auto-calculate and update display
        weight_var.trace('w', lambda *args: self._update_distribution_preview())
        
        return frame
    
    def _update_distribution_preview(self):
        """Update the DUM distribution preview for all partials"""
        try:
            # Validate LTA data
            if not self.lta_data.get('dums') or self.lta_data.get('total_weight', 0) <= 0:
                # Show error message in preview
                for form_data in self.partial_forms:
                    form_data['positions_var'].set("0")
                    dums_text = form_data['dums_text']
                    dums_text.configure(state='normal')
                    dums_text.delete('1.0', tk.END)
                    dums_text.insert(tk.END, "⚠️ Données LTA invalides\n(Poids = 0 ou aucun DUM)")
                    dums_text.configure(state='disabled')
                return
            
            # Collect partial weights
            partial_weights = []
            for form_data in self.partial_forms:
                try:
                    weight = float(form_data['weight_var'].get().strip())
                    partial_weights.append(weight)
                except ValueError:
                    partial_weights.append(0)
            
            # Detect exception case: check if any partial weight < smallest DUM weight
            smallest_dum_weight = min(dum['weight'] for dum in self.lta_data['dums'])
            is_exception_case = any(w > 0 and w < smallest_dum_weight for w in partial_weights)
            
            if is_exception_case:
                # Show exception frame if hidden
                if not self.exception_frame.winfo_manager():
                    self.exception_frame.pack(fill=tk.X, pady=5, before=self.partials_container)
            else:
                # Hide exception frame
                if self.exception_frame.winfo_manager():
                    self.exception_frame.pack_forget()
            
            # Calculate distribution
            distribution = self._calculate_dum_distribution(partial_weights)
            
            # Update each partial's display
            for idx, form_data in enumerate(self.partial_forms):
                if idx < len(distribution):
                    partial_dist = distribution[idx]
                    
                    # Update positions
                    form_data['positions_var'].set(str(partial_dist['positions']))
                    
                    # Update DUM list
                    dums_text = form_data['dums_text']
                    dums_text.configure(state='normal')
                    dums_text.delete('1.0', tk.END)
                    
                    if not partial_dist['dums']:
                        dums_text.insert(tk.END, "Aucun DUM assigné")
                    else:
                        for dum_info in partial_dist['dums']:
                            dum_num = dum_info['dum_number']
                            dum_weight = dum_info['weight']
                            dum_positions = dum_info['positions']
                            is_split = dum_info['is_split']
                            split_id = dum_info.get('split_id', '')
                            
                            if is_split:
                                dums_text.insert(tk.END, f"DUM {dum_num} {split_id}: {dum_weight:.1f}kg, {dum_positions}p ⚠️ PARTIEL\n")
                            else:
                                dums_text.insert(tk.END, f"DUM {dum_num}: {dum_weight:.1f}kg, {dum_positions}p\n")
                    
                    dums_text.configure(state='disabled')
        except Exception as e:
            # Silently handle preview errors to avoid disrupting user input
            logger.error(f"Error updating distribution preview: {e}", exc_info=True)
    
    def _calculate_dum_distribution(self, partial_weights):
        """
        Automatically distribute DUMs across partials based on weights.
        Sequential distribution: Fill partials in order until weight is reached.
        Last DUM may be split if needed.
        """
        distribution = []
        
        total_lta_weight = self.lta_data['total_weight']
        total_lta_positions = self.lta_data['total_positions']
        dums = self.lta_data['dums']
        
        # Validate LTA data
        if not dums or total_lta_weight <= 0 or total_lta_positions <= 0:
            # Return empty distribution if LTA data is invalid
            for _ in partial_weights:
                distribution.append({'weight': 0, 'positions': 0, 'dums': []})
            return distribution
        
        current_dum_idx = 0
        remaining_dum_weight = dums[0]['weight'] if dums else 0
        remaining_dum_positions = dums[0]['positions'] if dums else 0
        is_continuing_split = False  # Track if we're continuing a split DUM
        
        for partial_idx, partial_weight in enumerate(partial_weights):
            if partial_weight <= 0:
                distribution.append({'weight': 0, 'positions': 0, 'dums': []})
                continue
            
            # Calculate positions for this partial (safe division)
            if total_lta_weight > 0:
                partial_positions = round((partial_weight * total_lta_positions) / total_lta_weight)
            else:
                partial_positions = 0
            
            partial_dums = []
            weight_accumulated = 0
            positions_accumulated = 0
            
            # Fill DUMs until we reach the target weight
            while weight_accumulated < partial_weight and current_dum_idx < len(dums):
                weight_needed = partial_weight - weight_accumulated
                
                if remaining_dum_weight <= weight_needed:
                    # Take entire remaining DUM (or remaining part of split DUM)
                    partial_dums.append({
                        'dum_number': dums[current_dum_idx]['number'],
                        'weight': remaining_dum_weight,
                        'positions': remaining_dum_positions,
                        'is_split': is_continuing_split,
                        'split_id': f"{dums[current_dum_idx]['number']}/{partial_idx + 1}" if is_continuing_split else ''
                    })
                    weight_accumulated += remaining_dum_weight
                    positions_accumulated += remaining_dum_positions
                    
                    # Move to next DUM
                    current_dum_idx += 1
                    is_continuing_split = False
                    if current_dum_idx < len(dums):
                        remaining_dum_weight = dums[current_dum_idx]['weight']
                        remaining_dum_positions = dums[current_dum_idx]['positions']
                else:
                    # Split the DUM - this is the last DUM for this partial
                    # Calculate positions to reach the target partial_positions
                    positions_needed = partial_positions - positions_accumulated
                    
                    partial_dums.append({
                        'dum_number': dums[current_dum_idx]['number'],
                        'weight': weight_needed,
                        'positions': positions_needed,
                        'is_split': True,
                        'split_id': f"{dums[current_dum_idx]['number']}/{partial_idx + 1}"
                    })
                    weight_accumulated += weight_needed
                    positions_accumulated += positions_needed
                    
                    # Update remaining DUM
                    remaining_dum_weight -= weight_needed
                    remaining_dum_positions -= positions_needed
                    is_continuing_split = True  # Mark that next partial continues this DUM
                    break
            
            distribution.append({
                'weight': weight_accumulated,
                'positions': partial_positions,  # Use calculated target positions, not accumulated
                'dums': partial_dums
            })
        
        return distribution
    
    def _save_config(self):
        """Validate and save configuration"""
        try:
            # Validate LTA data first
            if not self.lta_data.get('dums') or self.lta_data.get('total_weight', 0) <= 0:
                messagebox.showerror(
                    "Erreur",
                    "Données LTA invalides.\n\n"
                    "Le LTA doit avoir:\n"
                    "- Un poids total > 0\n"
                    "- Au moins un DUM\n\n"
                    "Vérifiez le fichier Excel du LTA."
                )
                return
            
            # Collect data from forms
            partials = []
            total_weight_check = 0
            
            # First collect partial weights to calculate distribution
            partial_weights = []
            for form_data in self.partial_forms:
                try:
                    weight = float(form_data['weight_var'].get().strip())
                    partial_weights.append(weight)
                except ValueError:
                    messagebox.showerror("Erreur", f"Poids invalide pour Partiel {form_data['partial_number']}")
                    return
            
            # Calculate DUM distribution automatically
            distribution = self._calculate_dum_distribution(partial_weights)
            
            # Build partials configuration using calculated distribution
            for idx, form_data in enumerate(self.partial_forms):
                partial_num = form_data['partial_number']
                
                # Validate required fields
                weight = form_data['weight_var'].get().strip()
                ds_serie = form_data['ds_serie_var'].get().strip()
                ds_cle = form_data['ds_cle_var'].get().strip()
                location = form_data['location_var'].get().strip()
                
                if not all([weight, ds_serie, ds_cle, location]):
                    messagebox.showerror(
                        "Validation",
                        f"Partiel {partial_num}: Tous les champs sont requis"
                    )
                    return
                
                # Validate weight
                try:
                    weight_float = float(weight)
                    total_weight_check += weight_float
                except ValueError:
                    messagebox.showerror(
                        "Validation",
                        f"Partiel {partial_num}: Poids invalide"
                    )
                    return
                
                # Get DUMs from calculated distribution
                partial_dist = distribution[idx]
                selected_dums = []
                
                for dum_info in partial_dist['dums']:
                    selected_dums.append({
                        'dum_number': dum_info['dum_number'],
                        'weight': dum_info['weight'],
                        'positions': dum_info['positions'],
                        'is_split': dum_info['is_split'],
                        'split_id': dum_info.get('split_id', '')
                    })
                
                # Validate distribution has DUMs
                if not selected_dums:
                    messagebox.showerror(
                        "Validation",
                        f"Partiel {partial_num}: Aucun DUM assigné par distribution automatique"
                    )
                    return
                
                partials.append({
                    'partial_number': partial_num,
                    'weight': weight_float,
                    'positions': partial_dist['positions'],
                    'ds_serie': ds_serie,
                    'ds_cle': ds_cle,
                    'loading_location': location,
                    'dums': selected_dums
                })
            
            # Validate weight tolerance (allow 1% difference)
            weight_diff = abs(total_weight_check - self.lta_data['total_weight'])
            weight_tolerance = self.lta_data['total_weight'] * 0.01
            
            if weight_diff > weight_tolerance:
                response = messagebox.askyesno(
                    "Attention",
                    f"La somme des poids partiels ({total_weight_check} kg) ne correspond pas exactement au poids total ({self.lta_data['total_weight']} kg).\n\n"
                    f"Différence: {weight_diff:.2f} kg\n\n"
                    "Continuer quand même?"
                )
                if not response:
                    return
            
            # Detect split DUMs from distribution
            split_dums = {}
            for partial in partials:
                for dum in partial['dums']:
                    if dum['is_split']:
                        dum_num = str(dum['dum_number'])
                        if dum_num not in split_dums:
                            split_dums[dum_num] = {
                                'total_weight': 0,
                                'splits': []
                            }
                        
                        split_dums[dum_num]['total_weight'] += dum['weight']
                        split_dums[dum_num]['splits'].append({
                            'partial': partial['partial_number'],
                            'split_id': dum['split_id'],
                            'weight': dum['weight'],
                            'positions': dum['positions']
                        })
            
            # Detect exception case
            smallest_dum_weight = min(dum['weight'] for dum in self.lta_data['dums'])
            smallest_partial_weight = min(partial_weights)
            is_exception_case = smallest_partial_weight < smallest_dum_weight
            
            # For exception case, validate additional fields
            smallest_partial_number = None
            smallest_partial_positions = None
            airport_reference = None
            
            if is_exception_case:
                # Find which partial is the smallest
                for idx, weight in enumerate(partial_weights):
                    if weight == smallest_partial_weight:
                        smallest_partial_number = idx + 1
                        break
                
                # Validate exception case fields
                airport_reference = self.airport_reference_var.get().strip()
                smallest_partial_positions_str = self.smallest_partial_positions_var.get().strip()
                
                if not airport_reference:
                    messagebox.showerror(
                        "Validation",
                        "Cas d'exception détecté: Veuillez renseigner la référence créée à l'aéroport"
                    )
                    return
                
                if not smallest_partial_positions_str:
                    messagebox.showerror(
                        "Validation",
                        "Cas d'exception détecté: Veuillez renseigner les positions du plus petit partiel"
                    )
                    return
                
                try:
                    smallest_partial_positions = int(smallest_partial_positions_str)
                    if smallest_partial_positions <= 0:
                        raise ValueError("Positions must be positive")
                except ValueError:
                    messagebox.showerror(
                        "Validation",
                        "Positions du plus petit partiel: valeur invalide (doit être un nombre > 0)"
                    )
                    return
            
            # Build config
            config = {
                'lta_reference': self._get_lta_reference(),
                'lta_total_weight': self.lta_data['total_weight'],
                'lta_total_positions': self.lta_data['total_positions'],
                'partial_type': 'exception' if is_exception_case else 'normal',
                'partials': partials,
                'split_dums': split_dums
            }
            
            # Add exception case fields if applicable
            if is_exception_case:
                config['smallest_partial_number'] = smallest_partial_number
                config['smallest_partial_positions'] = smallest_partial_positions
                config['smallest_partial_airport_reference'] = airport_reference
            
            # Save config
            success = save_lta_partial_config(
                self.lta_folder_path,
                self.folder_name,
                config
            )
            
            if success:
                self.config_saved = True
                messagebox.showinfo("Succès", "Configuration sauvegardée!")
                self.dialog.destroy()
            else:
                messagebox.showerror("Erreur", "Impossible de sauvegarder la configuration")
                
        except Exception as e:
            logger.error(f"Error saving partial config: {e}", exc_info=True)
            messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde:\n{e}")
    
    def _get_lta_reference(self):
        """Get LTA reference from LTA file"""
        try:
            lta_file_patterns = [
                f"{self.folder_name}.txt",
                f"{self.folder_name.replace(' ', '')}.txt",
                f"{self.folder_name.lower().replace(' ', '')}.txt"
            ]
            
            for pattern in lta_file_patterns:
                lta_file = os.path.join(self.lta_folder_path, pattern)
                if os.path.exists(lta_file):
                    with open(lta_file, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                    if len(lines) >= 4:
                        reference = lines[3].strip()  # Line 4 (index 3)
                        # Remove /1 suffix if present
                        if reference.endswith('/1'):
                            reference = reference[:-2]
                        return reference
            
            return "UNKNOWN"
            
        except Exception as e:
            logger.error(f"Error getting LTA reference: {e}")
            return "UNKNOWN"
