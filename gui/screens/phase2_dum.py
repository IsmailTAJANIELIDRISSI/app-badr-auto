#!/usr/bin/env python3
"""
Phase 2 DUM Screen - Screen 3
Handles DUM declarations
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging

from gui.utils.script_manager import ScriptManager
from gui.utils.file_utils import detect_ltas, get_dum_count

logger = logging.getLogger(__name__)

class Phase2DUMScreen:
    """Screen 3: Phase 2 - DUM Declarations"""
    
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.frame = ttk.Frame(parent)
        self.script_manager = ScriptManager(app)
        self.all_ltas = []
        self.lta_checkboxes = []
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the UI components"""
        # Title
        title = ttk.Label(self.frame, text="PHASE 2: DÉCLARATIONS DÉDOUANEMENT", 
                         font=('Arial', 14, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10, sticky=tk.W)
        
        # Info label
        info_label = ttk.Label(self.frame, 
                              text="Cette phase créera les déclarations DUM pour chaque LTA sélectionné",
                              font=('Arial', 9, 'italic'))
        info_label.grid(row=1, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        # Refresh button
        refresh_btn = ttk.Button(self.frame, text="🔄 Actualiser la liste des LTAs",
                                command=self.refresh_lta_list)
        refresh_btn.grid(row=2, column=0, columnspan=2, pady=5)
        
        # Summary frame
        self.summary_frame = ttk.Frame(self.frame)
        self.summary_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        self.summary_label = ttk.Label(self.summary_frame, text="Chargez les LTAs pour commencer")
        self.summary_label.pack()
        
        # Selection options
        sel_frame = ttk.LabelFrame(self.frame, text="Sélection LTAs", padding="10")
        sel_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.selection_var = tk.StringVar(value="all")
        ttk.Radiobutton(sel_frame, text="Tous les LTAs", variable=self.selection_var, 
                       value="all", command=self.update_selection).grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Radiobutton(sel_frame, text="Seulement LTAs avec ED signé", 
                       variable=self.selection_var, value="signed", 
                       command=self.update_selection).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Radiobutton(sel_frame, text="Sélection manuelle", 
                       variable=self.selection_var, value="manual",
                       command=self.update_selection).grid(row=0, column=2, sticky=tk.W, padx=5)
        
        # LTA list with checkboxes (for manual selection)
        self.lta_list_frame = ttk.LabelFrame(self.frame, text="LTAs Disponibles", padding="10")
        self.lta_list_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Scrollable frame for checkboxes
        canvas = tk.Canvas(self.lta_list_frame, height=150)
        scrollbar = ttk.Scrollbar(self.lta_list_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Placeholder
        ttk.Label(self.scrollable_frame, text="Cliquez sur 'Actualiser' pour charger les LTAs").pack(pady=20)
        
        # Start button
        self.start_btn = ttk.Button(self.frame, text="▶️ Démarrer Phase 2",
                                    command=self.start_phase2, state="disabled")
        self.start_btn.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(self.frame, text="Progression", padding="10")
        progress_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # LTA progress
        ttk.Label(progress_frame, text="LTA:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.lta_progress_var = tk.DoubleVar()
        self.lta_progress = ttk.Progressbar(progress_frame, variable=self.lta_progress_var, 
                                           maximum=100, mode='determinate')
        self.lta_progress.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # DUM progress
        ttk.Label(progress_frame, text="DUM:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.dum_progress_var = tk.DoubleVar()
        self.dum_progress = ttk.Progressbar(progress_frame, variable=self.dum_progress_var, 
                                           maximum=100, mode='determinate')
        self.dum_progress.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        
        progress_frame.columnconfigure(1, weight=1)
        
        # Status
        self.status_var = tk.StringVar(value="En attente...")
        ttk.Label(self.frame, textvariable=self.status_var).grid(row=8, column=0, columnspan=2, pady=5)
        
        # Recent DUMs frame
        recent_frame = ttk.LabelFrame(self.frame, text="DUMs Récents", padding="10")
        recent_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Scrollable text for recent DUMs
        dum_scroll = ttk.Scrollbar(recent_frame)
        dum_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.dum_text = tk.Text(recent_frame, height=6, yscrollcommand=dum_scroll.set,
                               font=('Consolas', 9))
        self.dum_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dum_scroll.config(command=self.dum_text.yview)
        
        # Configure tags
        self.dum_text.tag_config("success", foreground="green")
        self.dum_text.tag_config("error", foreground="red")
        
        # Configure grid
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(5, weight=1)
        self.frame.rowconfigure(9, weight=1)
    
    def refresh_lta_list(self):
        """Refresh the list of LTAs"""
        if not self.app.current_folder:
            messagebox.showwarning("Attention", "Aucun dossier sélectionné.\nVeuillez d'abord exécuter la préparation.")
            return
        
        # Detect all LTAs
        self.all_ltas = detect_ltas(self.app.current_folder)
        
        if not self.all_ltas:
            messagebox.showwarning("Attention", "Aucun LTA détecté dans ce dossier")
            return
        
        # Count LTAs with signed ED
        ltas_with_ed = [lta for lta in self.all_ltas if lta.get('lta_file')]
        signed_count = sum(1 for lta in ltas_with_ed if self._has_signed_series(lta))
        
        # Update summary
        self.summary_label.config(
            text=f"📊 Total: {len(self.all_ltas)} LTAs | Avec ED signé: {signed_count} | Sans ED: {len(self.all_ltas) - signed_count}"
        )
        
        # Populate checkbox list
        self.populate_lta_checkboxes()
        
        # Enable start button
        self.start_btn.config(state="normal")
        
        # Apply current selection filter
        self.update_selection()
        
        self.app.log_message(f"Phase 2: {len(self.all_ltas)} LTA(s) détecté(s)", "INFO")
    
    def _has_signed_series(self, lta):
        """Check if LTA has signed series in Line 8"""
        if not lta.get('lta_file'):
            return False
        
        try:
            with open(lta['lta_file'], 'r', encoding='utf-8') as f:
                lines = f.readlines()
            if len(lines) > 7:
                signed_series = lines[7].strip()
                return bool(signed_series)
        except:
            pass
        return False
    
    def populate_lta_checkboxes(self):
        """Populate the LTA checkbox list"""
        # Clear existing checkboxes
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.lta_checkboxes = []
        
        for lta in self.all_ltas:
            var = tk.BooleanVar(value=True)
            
            # Get DUM count
            dum_count = get_dum_count(lta['folder_path']) if lta.get('folder_path') else 0
            
            # Check if has signed ED (from LTA file line 8)
            has_ed = bool(lta.get('signed_ds'))
            
            # Display signed DS if available, otherwise show "Pas d'ED"
            if has_ed:
                ed_status = f"ED: {lta.get('signed_ds')}"
            else:
                ed_status = "Pas d'ED"
            
            # Create checkbox with LTA name, reference, and status
            lta_ref = lta.get('lta_reference', 'N/A')
            cb_text = f"{lta['name']} - {lta_ref} - {dum_count} DUMs - {ed_status}"
            cb = ttk.Checkbutton(self.scrollable_frame, text=cb_text, variable=var)
            cb.pack(anchor=tk.W, padx=5, pady=2)
            
            self.lta_checkboxes.append({
                'lta': lta,
                'var': var,
                'has_ed': has_ed,
                'dum_count': dum_count
            })
    
    def update_selection(self):
        """Update checkbox states based on selection mode"""
        selection_mode = self.selection_var.get()
        
        for cb_data in self.lta_checkboxes:
            if selection_mode == "all":
                cb_data['var'].set(True)
            elif selection_mode == "signed":
                cb_data['var'].set(cb_data['has_ed'])
            # For "manual", leave as is (user controls)
    
    def start_phase2(self):
        """Start Phase 2 automation"""
        # Get selected LTAs
        selected_ltas = [cb['lta'] for cb in self.lta_checkboxes if cb['var'].get()]
        
        if not selected_ltas:
            messagebox.showwarning("Attention", "Aucun LTA sélectionné")
            return
        
        # Confirm start
        total_dums = sum(cb['dum_count'] for cb in self.lta_checkboxes if cb['var'].get())
        if not messagebox.askyesno("Confirmation", 
                                   f"Démarrer Phase 2 pour {len(selected_ltas)} LTA(s) ({total_dums} DUMs)?"):
            return
        
        # Disable button during execution
        self.start_btn.config(state="disabled")
        self.status_var.set("⏳ Démarrage Phase 2...")
        self.app.set_status("Exécution Phase 2...")
        logger.info(f"Starting Phase 2 for {len(selected_ltas)} LTAs...")
        self.app.log_message(f"Démarrage de la Phase 2 pour {len(selected_ltas)} LTA(s)...", "INFO")
        
        # Clear DUM text
        self.dum_text.delete("1.0", tk.END)
        
        # Progress callback
        def on_progress(percent, message):
            # Parse message for LTA/DUM progress
            if "LTA:" in message:
                self.lta_progress_var.set(percent)
            elif "DUM:" in message:
                self.dum_progress_var.set(percent)
            
            self.status_var.set(message)
            self.app.set_status(message)
            self.app.log_message(message, "INFO")
            
            # If message contains DUM reference, add to recent list
            if "DUM" in message and ":" in message:
                self.add_dum_reference(message)
        
        # Completion callback
        def on_complete(success=True, error=None):
            self.start_btn.config(state="normal")
            
            if success:
                self.status_var.set("✅ Phase 2 terminée")
                self.app.log_message("Phase 2 terminée avec succès", "SUCCESS")
                messagebox.showinfo("Succès", 
                                  f"Phase 2 terminée!\n{len(selected_ltas)} LTA(s) traité(s)")
            else:
                self.status_var.set(f"❌ Erreur: {error[:50] if error else 'Unknown'}")
                self.app.log_message(f"Erreur Phase 2: {error}", "ERROR")
                messagebox.showerror("Erreur", f"Phase 2 a échoué:\n{error}")
        
        # Execute Phase 2
        # Convert selected LTAs to indices
        lta_indices = []
        for lta in selected_ltas:
            # Find index of this LTA in all_ltas
            for idx, all_lta in enumerate(self.all_ltas):
                if all_lta['name'] == lta['name']:
                    lta_indices.append(idx)
                    break
        
        credentials = {'username': 'BK707345', 'password': ''}  # Will be loaded from .env by script
        self.script_manager.run_phase2(self.app.current_folder, lta_indices, credentials, on_progress, on_complete)
    
    def add_dum_reference(self, message):
        """Add DUM reference to the recent list"""
        # Determine tag based on message content
        tag = "success" if "✅" in message or "succès" in message.lower() else "error" if "❌" in message or "erreur" in message.lower() else None
        
        self.dum_text.insert(tk.END, f"{message}\n", tag)
        self.dum_text.see(tk.END)  # Auto-scroll to bottom

