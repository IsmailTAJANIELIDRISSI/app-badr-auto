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
    
    def __init__(self, parent, lta_folder_path, folder_name, lta_file_path=None):
        self.parent = parent
        self.lta_folder_path = lta_folder_path
        self.folder_name = folder_name
        self.lta_file_path = lta_file_path
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
        self.dialog.geometry("850x750")  # Increased height to accommodate exception section + buttons
        self.dialog.minsize(800, 650)  # Set minimum size to ensure buttons are always visible
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
        
        # Add LTA Reference Field - Grid Layout
        ttk.Label(totals_frame, text="Référence LTA (MAWB):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        
        # Get initial reference
        initial_ref = self.existing_config.get('lta_reference') if self.existing_config else self._get_lta_reference()
        self.lta_reference_var = tk.StringVar(value=initial_ref)
        
        ref_entry = ttk.Entry(totals_frame, textvariable=self.lta_reference_var, width=20)
        ref_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Add totals using grid
        ttk.Label(totals_frame, text=f"Poids Total: {self.lta_data['total_weight']} kg").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(totals_frame, text=f"Positions Totales: {self.lta_data['total_positions']}").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(totals_frame, text=f"Nombre de DUMs: {len(self.lta_data['dums'])}").grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        
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
        
        # Exception confirmation button
        self.exception_confirmed = False
        self.exception_status_label = ttk.Label(
            self.exception_frame,
            text="",
            font=('Arial', 9),
            foreground="green"
        )
        self.exception_status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, padx=5, pady=(5, 0))
        
        confirm_exception_btn = ttk.Button(
            self.exception_frame,
            text="✅ Confirmer Informations Exception",
            command=self._confirm_exception_fields
        )
        confirm_exception_btn.grid(row=4, column=0, columnspan=2, padx=5, pady=10, sticky=tk.W)
        
        # Buttons frame - CRITICAL: Pack BEFORE content frames to keep at bottom
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
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
        
        # Partials container (scrollable) - packs AFTER buttons to fill remaining space
        self.partials_container = ttk.Frame(main_frame)
        self.partials_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
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
        
        # Load existing config if available
        if self.existing_config:
            self.num_partials_var.set(len(self.existing_config['partials']))
            # Load exception case data if exists
            if self.existing_config.get('partial_type') == 'exception':
                airport_ref = self.existing_config.get('smallest_partial_airport_reference', '')
                positions = str(self.existing_config.get('smallest_partial_positions', ''))
                self.airport_reference_var.set(airport_ref)
                self.smallest_partial_positions_var.set(positions)
                # If both fields are filled, mark as confirmed
                if airport_ref and positions:
                    self.exception_confirmed = True
                    self.exception_status_label.config(
                        text="Informations exception récupérées de la config existante",
                        foreground="green"
                    )
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
    
    def _confirm_exception_fields(self):
        """Validate and confirm exception case fields"""
        try:
            airport_reference = self.airport_reference_var.get().strip()
            smallest_partial_positions_str = self.smallest_partial_positions_var.get().strip()
            
            if not airport_reference:
                messagebox.showerror(
                    "Validation",
                    "Veuillez renseigner la référence créée à l'aéroport"
                )
                return
            
            if not smallest_partial_positions_str:
                messagebox.showerror(
                    "Validation",
                    "Veuillez renseigner les positions du plus petit partiel"
                )
                return
            
            try:
                positions = int(smallest_partial_positions_str)
                if positions <= 0:
                    raise ValueError("Positions must be positive")
            except ValueError:
                messagebox.showerror(
                    "Validation",
                    "Positions invalides: doit être un nombre > 0"
                )
                return
            
            # Mark as confirmed
            self.exception_confirmed = True
            self.exception_status_label.config(
                text="Informations exception confirmées",
                foreground="green",
                wraplength=600
            )
            
            # Update preview to reflect user-provided positions
            self._update_distribution_preview()
            
            # Update and scroll to show all content
            self.dialog.update_idletasks()
            self.dialog.after(100, lambda: self._scroll_to_bottom())
            
        except Exception as e:
            logger.error(f"Error confirming exception fields: {e}", exc_info=True)
            messagebox.showerror("Erreur", f"Erreur lors de la validation: {e}")
    
    def _scroll_to_bottom(self):
        """Scroll canvas to bottom to show all content"""
        try:
            if hasattr(self, 'canvas'):
                self.canvas.update_idletasks()
                # Scroll to bottom of canvas
                self.canvas.yview_moveto(1.0)
                # Ensure buttons frame is visible (it's outside canvas, so we just ensure dialog size is correct)
                self.dialog.update_idletasks()
        except Exception as e:
            logger.debug(f"Could not scroll to bottom: {e}")
    
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
                
                # CRITICAL FIX: Only reset confirmation if weights actually changed
                # Check if weights are different from last time
                current_weights_tuple = tuple(partial_weights)
                if not hasattr(self, '_last_exception_weights'):
                    self._last_exception_weights = current_weights_tuple
                
                if current_weights_tuple != self._last_exception_weights:
                    # Weights changed - reset confirmation
                    if self.exception_confirmed:
                        self.exception_confirmed = False
                        self.exception_status_label.config(text="⚠️ Poids modifié - veuillez reconfirmer", foreground="red")
                    self._last_exception_weights = current_weights_tuple
            else:
                # Hide exception frame
                if self.exception_frame.winfo_manager():
                    self.exception_frame.pack_forget()
                # Reset confirmation flag when exception is no longer detected
                self.exception_confirmed = False
                if hasattr(self, '_last_exception_weights'):
                    del self._last_exception_weights
            
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
                    
                    # Check if this is a smallest/largest partial in exception case
                    is_smallest = partial_dist.get('is_smallest_partial', False)
                    is_largest = partial_dist.get('is_largest_partial', False)
                    
                    if is_smallest:
                        dums_text.insert(tk.END, "⚠️ EXCEPTION: Déjà dégagé à l'aéroport (pas d'état de dépotage)\n\n", 'red')
                    elif is_largest:
                        dums_text.insert(tk.END, "✓ État de dépotage sera créé pour ce partiel\n\n", 'green')
                    
                    if not partial_dist['dums']:
                        dums_text.insert(tk.END, "Aucun DUM assigné")
                    else:
                        for dum_info in partial_dist['dums']:
                            dum_num = dum_info['dum_number']
                            dum_weight = dum_info['weight']
                            dum_positions = dum_info['positions']
                            is_split = dum_info['is_split']
                            split_id = dum_info.get('split_id', '')
                            is_exception_portion = dum_info.get('is_exception_portion', False)
                            adjusted_for_exception = dum_info.get('adjusted_for_exception', False)
                            
                            if is_exception_portion:
                                dums_text.insert(tk.END, f"DUM {dum_num} (portion aéroport): {dum_weight:.1f}kg, {dum_positions}p 🔴 EXCEPTION\n")
                            elif adjusted_for_exception:
                                dums_text.insert(tk.END, f"DUM {dum_num} (ajusté): {dum_weight:.1f}kg, {dum_positions}p ⚙️ AJUSTÉ\n")
                            elif is_split:
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
        Detects and handles exception cases differently from normal cases.
        """
        total_lta_weight = self.lta_data['total_weight']
        total_lta_positions = self.lta_data['total_positions']
        dums = self.lta_data['dums']
        
        # Validate LTA data
        if not dums or total_lta_weight <= 0 or total_lta_positions <= 0:
            # Return empty distribution if LTA data is invalid
            distribution = []
            for _ in partial_weights:
                distribution.append({'weight': 0, 'positions': 0, 'dums': []})
            return distribution
        
        # Detect exception case: any partial weight < smallest DUM weight
        smallest_dum_weight = min(dum['weight'] for dum in dums)
        is_exception_case = any(w > 0 and w < smallest_dum_weight for w in partial_weights)
        
        if is_exception_case:
            # Use exception-specific distribution
            return self._calculate_exception_distribution(partial_weights)
        else:
            # Use normal sequential distribution
            return self._calculate_normal_distribution(partial_weights)
    
    def _calculate_exception_distribution(self, partial_weights):
        """
        Exception case distribution: when smallest partial < smallest DUM.
        Logic: Subtract smallest partial from DUM 1, then distribute remaining DUMs to largest partial.
        """
        total_lta_weight = self.lta_data['total_weight']
        total_lta_positions = self.lta_data['total_positions']
        dums = self.lta_data['dums'].copy()  # Make a copy to avoid modifying original
        
        distribution = []
        
        # Find smallest and largest partial indices
        smallest_weight = min(w for w in partial_weights if w > 0)
        smallest_idx = partial_weights.index(smallest_weight)
        
        # Largest is the other one (works for 2 partials)
        largest_idx = 1 - smallest_idx if len(partial_weights) == 2 else max(
            range(len(partial_weights)),
            key=lambda i: partial_weights[i]
        )
        
        smallest_weight_val = partial_weights[smallest_idx]
        largest_weight_val = partial_weights[largest_idx]
        
        # DEBUG: Log weight values
        logger.info(f"Exception distribution - Smallest partial weight from form: {smallest_weight_val}kg")
        logger.info(f"Exception distribution - Largest partial weight from form: {largest_weight_val}kg")
        
        # Get user-provided positions for smallest partial from exception fields
        # This is the CRITICAL FIX: use user-entered value, not auto-calculated!
        try:
            smallest_positions_str = self.smallest_partial_positions_var.get().strip()
            if smallest_positions_str:
                smallest_positions = int(smallest_positions_str)
                logger.info(f"Exception distribution - Using USER-PROVIDED positions: {smallest_positions}p")
            else:
                # Fallback to proportional calculation if not entered yet
                smallest_positions = round((smallest_weight_val * total_lta_positions) / total_lta_weight)
                logger.warning(f"Exception distribution - Fallback to auto-calculated positions: {smallest_positions}p")
        except (ValueError, AttributeError) as e:
            # Fallback to proportional calculation if invalid or not available
            smallest_positions = round((smallest_weight_val * total_lta_positions) / total_lta_weight)
            logger.error(f"Exception distribution - Error reading positions, using auto-calc: {smallest_positions}p. Error: {e}")
        
        # DEBUG: Log DUM 1 original values
        logger.info(f"DUM 1 original: weight={dums[0]['weight']}kg, positions={dums[0]['positions']}p")
        
        # Calculate adjusted DUM 1 values
        adjusted_dum1_weight = dums[0]['weight'] - smallest_weight_val
        adjusted_dum1_positions = dums[0]['positions'] - smallest_positions
        
        # DEBUG: Log adjusted values
        logger.info(f"DUM 1 adjusted (calculation): {dums[0]['weight']} - {smallest_weight_val} = {adjusted_dum1_weight}kg")
        logger.info(f"DUM 1 adjusted (calculation): {dums[0]['positions']} - {smallest_positions} = {adjusted_dum1_positions}p")
        
        # Create distribution for each partial
        for idx, partial_weight in enumerate(partial_weights):
            if idx == smallest_idx:
                # Smallest partial - gets portion of DUM 1 only
                distribution.append({
                    'weight': smallest_weight_val,
                    'positions': smallest_positions,  # USER-PROVIDED VALUE!
                    'dums': [{
                        'dum_number': 1,
                        'weight': smallest_weight_val,
                        'positions': smallest_positions,  # USER-PROVIDED VALUE!
                        'is_split': False,
                        'split_id': '',
                        'is_exception_portion': True  # Mark as exception portion
                    }],
                    'is_smallest_partial': True  # Flag for no état de dépotage
                })
            elif idx == largest_idx:
                # Largest partial - gets DUM 1 (adjusted) + all other DUMs
                # Positions = TOTAL - USER-PROVIDED smallest positions
                largest_positions = total_lta_positions - smallest_positions
                
                # Subtract smallest partial from DUM 1
                adjusted_dum1_weight = dums[0]['weight'] - smallest_weight_val
                adjusted_dum1_positions = dums[0]['positions'] - smallest_positions  # Subtract USER-PROVIDED!
                
                partial_dums = []
                
                # Add adjusted DUM 1
                if adjusted_dum1_weight > 0:
                    partial_dums.append({
                        'dum_number': 1,
                        'weight': adjusted_dum1_weight,
                        'positions': adjusted_dum1_positions,  # ADJUSTED by user value!
                        'is_split': False,
                        'split_id': '',
                        'adjusted_for_exception': True 

 # Mark as adjusted
                    })
                
                # Add all remaining DUMs (2, 3, 4, etc.)
                for dum_idx in range(1, len(dums)):
                    partial_dums.append({
                        'dum_number': dums[dum_idx]['number'],
                        'weight': dums[dum_idx]['weight'],
                        'positions': dums[dum_idx]['positions'],
                        'is_split': False,
                        'split_id': ''
                    })
                
                distribution.append({
                    'weight': largest_weight_val,
                    'positions': largest_positions,  # TOTAL - user positions!
                    'dums': partial_dums,
                    'is_largest_partial': True  # Flag for état de dépotage creation
                })
            else:
                # Other partials (shouldn't happen in exception case with 2 partials)
                distribution.append({'weight': 0, 'positions': 0, 'dums': []})
        
        
        return distribution
    
    def _calculate_normal_distribution(self, partial_weights):
        """
        Normal case: Sequential distribution where partials are filled in order.
        Last DUM may be split if needed.
        """
        distribution = []
        
        total_lta_weight = self.lta_data['total_weight']
        total_lta_positions = self.lta_data['total_positions']
        dums = self.lta_data['dums']
        
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
                    dum_config = {
                        'dum_number': dum_info['dum_number'],
                        'weight': dum_info['weight'],
                        'positions': dum_info['positions'],
                        'is_split': dum_info['is_split'],
                        'split_id': dum_info.get('split_id', '')
                    }
                    
                    # Add exception-specific flags
                    if dum_info.get('is_exception_portion', False):
                        dum_config['is_exception_portion'] = True
                    if dum_info.get('adjusted_for_exception', False):
                        dum_config['adjusted_for_exception'] = True
                    
                    selected_dums.append(dum_config)
                
                # Validate distribution has DUMs
                if not selected_dums:
                    messagebox.showerror(
                        "Validation",
                        f"Partiel {partial_num}: Aucun DUM assigné par distribution automatique"
                    )
                    return
                
                # Build partial configuration
                partial_config = {
                    'partial_number': partial_num,
                    'weight': weight_float,
                    'positions': partial_dist['positions'],
                    'ds_serie': ds_serie,
                    'ds_cle': ds_cle,
                    'loading_location': location,
                    'dums': selected_dums
                }
                
                # Add exception-specific flags
                # For smallest partial: no état de dépotage needed (already cleared at airport)
                if partial_dist.get('is_smallest_partial', False):
                    partial_config['create_etat_depotage'] = False
                    partial_config['is_smallest_partial'] = True
                # For largest partial in exception case: état de dépotage with adjusted DUMs
                elif partial_dist.get('is_largest_partial', False):
                    partial_config['create_etat_depotage'] = True
                    partial_config['is_largest_partial'] = True
                else:
                    # Normal case: all partials get état de dépotage
                    partial_config['create_etat_depotage'] = True
                
                partials.append(partial_config)
            
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
                # Check if exception fields have been confirmed
                if not self.exception_confirmed:
                    messagebox.showwarning(
                        "Validation Requise",
                        "Cas d'exception détecté: Veuillez d'abord cliquer sur '✅ Confirmer Informations Exception'\n"
                        "et remplir tous les champs dans la section d'exception avant de sauvegarder."
                    )
                    # Scroll to exception frame
                    if hasattr(self, 'canvas'):
                        self.canvas.update_idletasks()
                        # Try to scroll to show exception frame
                        self.canvas.yview_moveto(0.1)
                    return
                
                # Find which partial is the smallest
                for idx, weight in enumerate(partial_weights):
                    if weight == smallest_partial_weight:
                        smallest_partial_number = idx + 1
                        break
                
                # Validate exception case fields (should already be validated by confirmation)
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
            
            # Validate LTA reference
            lta_reference = self.lta_reference_var.get().strip()
            if not lta_reference or lta_reference == "UNKNOWN":
                messagebox.showerror(
                    "Validation",
                    "Veuillez renseigner la Référence LTA (MAWB) avant de sauvegarder."
                )
                return

            # Build config
            # CRITICAL: For exception cases, lta_reference MUST be the airport reference
            if is_exception_case:
                config_lta_reference = airport_reference
                # Store original MAWB separately just in case
                main_lta_reference = lta_reference
            else:
                config_lta_reference = lta_reference
                main_lta_reference = lta_reference

            config = {
                'lta_reference': config_lta_reference,
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
                config['main_lta_reference'] = main_lta_reference  # Save MAWB here
            
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
            # Use passed LTA file path if available
            if self.lta_file_path and os.path.exists(self.lta_file_path):
                 with open(self.lta_file_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                 if len(lines) >= 4:
                    reference = lines[3].strip()  # Line 4 (index 3)
                    if reference.endswith('/1'):
                        reference = reference[:-2]
                    logger.info(f"Found LTA reference (from detection): {reference}")
                    return reference
            
            lta_file_patterns = [
                f"{self.folder_name}.txt",
                f"{self.folder_name.replace(' ', '')}.txt",
                f"{self.folder_name.lower().replace(' ', '')}.txt"
            ]
            
            for pattern in lta_file_patterns:
                lta_file = os.path.join(self.lta_folder_path, pattern)
                logger.info(f"Looking for LTA file: {lta_file}")
                
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
