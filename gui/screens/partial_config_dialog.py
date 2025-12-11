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
                return None
            
            wb = load_workbook(excel_files[0], data_only=True)
            ws = wb['Summary']
            
            # Get total weight and positions from Summary sheet
            # Weight is in column D, row 5 (P,BRUT)
            # Positions is in column D, row 6 (P)
            total_weight = ws['D5'].value
            total_positions = ws['D6'].value
            
            # Count DUMs by checking C11, C18, C25, C32, C39...
            dums = []
            for dum_idx in range(1, 10):
                row_num = 11 + (dum_idx - 1) * 7
                cell_value = ws[f'C{row_num}'].value
                
                if cell_value and 'DUM' in str(cell_value).upper():
                    # Get DUM weight (P,BRUT) and positions (P) from the same sheet
                    dum_weight_row = row_num + 2  # P,BRUT is 2 rows below
                    dum_positions_row = row_num + 3  # P is 3 rows below
                    
                    dum_weight = ws[f'D{dum_weight_row}'].value or 0
                    dum_positions = ws[f'D{dum_positions_row}'].value or 0
                    
                    dums.append({
                        'number': dum_idx,
                        'weight': float(dum_weight) if dum_weight else 0,
                        'positions': int(dum_positions) if dum_positions else 0
                    })
                else:
                    break
            
            wb.close()
            
            return {
                'total_weight': float(total_weight) if total_weight else 0,
                'total_positions': int(total_positions) if total_positions else 0,
                'dums': dums
            }
            
        except Exception as e:
            logger.error(f"Error loading LTA data: {e}", exc_info=True)
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
        location_entry.grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)
        
        # DUM Selection
        ttk.Label(frame, text="DUMs:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        
        dums_frame = ttk.Frame(frame)
        dums_frame.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        dum_vars = {}
        for dum in self.lta_data['dums']:
            dum_num = dum['number']
            is_selected = False
            
            if existing_data:
                for d in existing_data.get('dums', []):
                    if d['dum_number'] == dum_num:
                        is_selected = True
                        break
            
            var = tk.BooleanVar(value=is_selected)
            cb = ttk.Checkbutton(
                dums_frame,
                text=f"DUM {dum_num} ({dum['weight']}kg, {dum['positions']}p)",
                variable=var
            )
            cb.pack(anchor=tk.W)
            dum_vars[dum_num] = var
        
        self.partial_forms.append({
            'partial_number': partial_num,
            'weight_var': weight_var,
            'ds_serie_var': ds_serie_var,
            'ds_cle_var': ds_cle_var,
            'location_var': location_var,
            'dum_vars': dum_vars
        })
        
        return frame
    
    def _save_config(self):
        """Validate and save configuration"""
        try:
            # Collect data from forms
            partials = []
            total_weight_check = 0
            all_dums_assigned = set()
            
            for form_data in self.partial_forms:
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
                
                # Get selected DUMs
                selected_dums = []
                for dum_num, var in form_data['dum_vars'].items():
                    if var.get():
                        # Find DUM data
                        dum_data = next((d for d in self.lta_data['dums'] if d['number'] == dum_num), None)
                        if dum_data:
                            selected_dums.append({
                                'dum_number': dum_num,
                                'weight': dum_data['weight'],
                                'positions': dum_data['positions'],
                                'is_split': dum_num in all_dums_assigned  # Will be split if already assigned
                            })
                            all_dums_assigned.add(dum_num)
                
                if not selected_dums:
                    messagebox.showerror(
                        "Validation",
                        f"Partiel {partial_num}: Au moins un DUM doit être sélectionné"
                    )
                    return
                
                # Calculate positions proportionally
                positions = int(round((weight_float / self.lta_data['total_weight']) * self.lta_data['total_positions']))
                
                partials.append({
                    'partial_number': partial_num,
                    'weight': weight_float,
                    'positions': positions,
                    'ds_serie': ds_serie,
                    'ds_cle': ds_cle,
                    'loading_location': location,
                    'dums': selected_dums
                })
            
            # Validate all DUMs assigned
            if len(all_dums_assigned) != len(self.lta_data['dums']):
                missing = set(d['number'] for d in self.lta_data['dums']) - all_dums_assigned
                messagebox.showerror(
                    "Validation",
                    f"DUMs non assignés: {', '.join(map(str, missing))}"
                )
                return
            
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
            
            # Detect split DUMs
            split_dums = {}
            for dum_num in all_dums_assigned:
                appearances = []
                for partial in partials:
                    for dum in partial['dums']:
                        if dum['dum_number'] == dum_num:
                            appearances.append({
                                'partial': partial['partial_number'],
                                'split_id': f"{dum_num}/{partial['partial_number']}",
                                'weight': dum['weight'],
                                'positions': dum['positions']
                            })
                
                if len(appearances) > 1:
                    split_dums[str(dum_num)] = {
                        'total_weight': sum(a['weight'] for a in appearances),
                        'splits': appearances
                    }
            
            # Build config
            config = {
                'lta_reference': self._get_lta_reference(),
                'lta_total_weight': self.lta_data['total_weight'],
                'lta_total_positions': self.lta_data['total_positions'],
                'partials': partials,
                'split_dums': split_dums
            }
            
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
