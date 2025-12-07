#!/usr/bin/env python3
"""
Script Manager - Wrapper for executing existing automation scripts
Handles script execution with progress callbacks and threading
"""

import subprocess
import threading
import logging
import os
import sys

logger = logging.getLogger(__name__)

class ScriptManager:
    """Manages execution of automation scripts"""
    
    def __init__(self, app):
        self.app = app
        self.current_process = None
        self.is_running = False
    
    def run_preparation(self, folder_path, progress_callback=None, completion_callback=None, selected_ltas=None):
        """
        Execute preparation scripts (fuzzy matching and validation)
        
        Args:
            folder_path: Path to folder containing LTA folders
            progress_callback: Function to call with (percent, message)
            completion_callback: Function to call when complete
            selected_ltas: List of LTA folder names to process (None = all)
        """
        def execute():
            try:
                self.is_running = True
                
                if selected_ltas:
                    logger.info(f"Starting preparation for {len(selected_ltas)} selected LTAs: {folder_path}")
                else:
                    logger.info(f"Starting preparation for all LTAs: {folder_path}")
                
                if progress_callback:
                    progress_callback(10, "Démarrage du script de préparation...")
                
                # Find script location - support both development and .exe modes
                # In .exe mode, scripts should be in parent directory of selected folder
                # or bundled with the executable
                if getattr(sys, 'frozen', False):
                    # Running as .exe - scripts should be with the .exe or in selected folder's parent
                    exe_dir = os.path.dirname(sys.executable)
                    # First check if scripts are in the parent of the selected folder
                    parent_of_selected = os.path.dirname(folder_path)
                    
                    # Try parent of selected folder first (most common case)
                    if os.path.exists(os.path.join(parent_of_selected, "script_all_fuzy_match.py")):
                        project_root = parent_of_selected
                    # Then try exe directory
                    elif os.path.exists(os.path.join(exe_dir, "script_all_fuzy_match.py")):
                        project_root = exe_dir
                    # Finally try the selected folder itself
                    else:
                        project_root = folder_path
                else:
                    # Running in development - use normal path resolution
                    current_file = os.path.abspath(__file__)
                    utils_dir = os.path.dirname(current_file)  # gui/utils
                    gui_dir = os.path.dirname(utils_dir)  # gui
                    project_root = os.path.dirname(gui_dir)  # project root
                
                # Step 1: Run script_all_fuzy_match.py
                fuzzy_script = os.path.join(project_root, "script_all_fuzy_match.py")
                
                if not os.path.exists(fuzzy_script):
                    raise FileNotFoundError(f"Script not found: {fuzzy_script}")
                
                logger.info(f"Executing: {fuzzy_script}")
                if progress_callback:
                    progress_callback(20, "Exécution de script_all_fuzy_match.py...")
                
                # Change to folder directory for script execution
                original_cwd = os.getcwd()
                os.chdir(folder_path)
                
                try:
                    # If selected_ltas is provided, temporarily move unselected folders
                    moved_folders = []
                    temp_dir = None
                    
                    if selected_ltas:
                        import shutil
                        temp_dir = os.path.join(folder_path, ".temp_unselected")
                        os.makedirs(temp_dir, exist_ok=True)
                        
                        # Move unselected LTA folders to temp
                        for item in os.listdir(folder_path):
                            item_path = os.path.join(folder_path, item)
                            if os.path.isdir(item_path) and item not in selected_ltas and item != ".temp_unselected":
                                if "LTA" in item or "lta" in item:
                                    shutil.move(item_path, os.path.join(temp_dir, item))
                                    moved_folders.append(item)
                        
                        logger.info(f"Temporarily moved {len(moved_folders)} unselected folders")
                    
                    # Execute fuzzy match script with real-time output
                    env = os.environ.copy()
                    env['PYTHONIOENCODING'] = 'utf-8'
                    env['PYTHONUNBUFFERED'] = '1'  # Disable buffering for real-time output
                    
                    # Determine the correct Python executable
                    if getattr(sys, 'frozen', False):
                        # Running as .exe - use system Python
                        python_exe = 'python'  # Use system Python from PATH
                    else:
                        # Running in development - use current Python
                        python_exe = sys.executable
                    
                    # Use Popen for real-time output
                    # CREATE_NO_WINDOW hides the terminal window on Windows
                    creation_flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                    process = subprocess.Popen(
                        [python_exe, fuzzy_script],
                        cwd=folder_path,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,  # Merge stderr into stdout for combined output
                        text=True,
                        encoding='utf-8',
                        env=env,
                        stdin=subprocess.DEVNULL,
                        bufsize=1,  # Line buffered
                        creationflags=creation_flags
                    )
                    
                    # Read output in real-time
                    while True:
                        output = process.stdout.readline()
                        if output == '' and process.poll() is not None:
                            break
                        if output:
                            self.app.logs_screen.add_script_output(output, "fuzzy")
                            logger.info(output.strip())
                    
                    # Wait for process to complete
                    process.wait()
                    
                    if process.returncode != 0:
                        error_msg = f"Script failed with code {process.returncode}"
                        self.app.logs_screen.add_script_output(f"\n❌ ERROR: {error_msg}\n", "fuzzy")
                        logger.error(error_msg)
                        raise Exception(error_msg)
                    
                    logger.info("Fuzzy match script completed successfully")
                    if progress_callback:
                        progress_callback(60, "Script de préparation terminé")
                    
                    # Step 2: Run validation.py
                    validation_script = os.path.join(project_root, "validation.py")
                    
                    if os.path.exists(validation_script):
                        logger.info(f"Executing: {validation_script}")
                        if progress_callback:
                            progress_callback(70, "Exécution de validation.py...")
                        
                        # Use same environment with UTF-8 encoding
                        # CREATE_NO_WINDOW hides the terminal window on Windows
                        creation_flags_val = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                        result = subprocess.run(
                            [python_exe, validation_script],
                            cwd=folder_path,
                            capture_output=True,
                            text=True,
                            encoding='utf-8',
                            env=env,
                            stdin=subprocess.DEVNULL,
                            timeout=60,  # 1 minute timeout
                            creationflags=creation_flags_val
                        )
                        
                        if result.returncode != 0:
                            logger.warning(f"Validation script warnings: {result.stderr}")
                        
                        logger.info("Validation script completed")
                        if progress_callback:
                            progress_callback(90, "Validation terminée")
                    else:
                        logger.warning(f"Validation script not found: {validation_script}")
                    
                finally:
                    # Restore moved folders if any
                    if moved_folders and temp_dir:
                        import shutil
                        for folder_name in moved_folders:
                            src = os.path.join(temp_dir, folder_name)
                            dst = os.path.join(folder_path, folder_name)
                            if os.path.exists(src):
                                shutil.move(src, dst)
                        # Remove temp directory
                        if os.path.exists(temp_dir):
                            try:
                                os.rmdir(temp_dir)
                            except:
                                pass
                        logger.info(f"Restored {len(moved_folders)} folders")
                    
                    # Restore original working directory
                    os.chdir(original_cwd)
                
                if progress_callback:
                    progress_callback(100, "✅ Traitement terminé")
                
                if completion_callback:
                    completion_callback(success=True)
                
            except Exception as e:
                logger.error(f"Preparation failed: {e}", exc_info=True)
                if progress_callback:
                    progress_callback(0, f"❌ Erreur: {str(e)[:100]}")
                if completion_callback:
                    completion_callback(success=False, error=str(e))
            finally:
                self.is_running = False
        
        # Run in separate thread
        thread = threading.Thread(target=execute, daemon=True)
        thread.start()
    
    def run_phase1(self, folder_path, credentials, lta_selection=None, progress_callback=None, completion_callback=None, selected_lta_names=None):
        """
        Execute Phase 1 of badr_login_test.py
        
        Args:
            folder_path: Path to folder containing LTA folders
            credentials: Dict with 'username' and 'password'
            lta_selection: List of LTA indices to process, or "all" for all LTAs
            progress_callback: Function to call with progress updates
            completion_callback: Function to call when complete
            selected_lta_names: List of LTA folder names to process (for filtering)
        """
        def execute():
            try:
                self.is_running = True
                logger.info("Starting Phase 1 automation...")
                
                if progress_callback:
                    progress_callback(10, "Démarrage Phase 1...")
                
                # Find script location - support both development and .exe modes
                if getattr(sys, 'frozen', False):
                    # Running as .exe - scripts should be with the .exe or in selected folder's parent
                    exe_dir = os.path.dirname(sys.executable)
                    parent_of_selected = os.path.dirname(folder_path)
                    
                    # Try parent of selected folder first
                    if os.path.exists(os.path.join(parent_of_selected, "badr_login_test.py")):
                        project_root = parent_of_selected
                    # Then try exe directory
                    elif os.path.exists(os.path.join(exe_dir, "badr_login_test.py")):
                        project_root = exe_dir
                    # Finally try the selected folder itself
                    else:
                        project_root = folder_path
                else:
                    # Running in development - use normal path resolution
                    current_file = os.path.abspath(__file__)
                    utils_dir = os.path.dirname(current_file)
                    gui_dir = os.path.dirname(utils_dir)
                    project_root = os.path.dirname(gui_dir)
                
                script_path = os.path.join(project_root, "badr_login_test.py")
                
                if not os.path.exists(script_path):
                    raise FileNotFoundError(f"Script not found: {script_path}")
                
                # Determine the correct Python executable
                if getattr(sys, 'frozen', False):
                    # Running as .exe - use system Python
                    python_exe = 'python'
                else:
                    # Running in development - use current Python
                    python_exe = sys.executable
                
                # If specific LTAs selected, temporarily move unselected folders
                moved_folders = []
                temp_dir = None
                
                try:
                    if selected_lta_names:
                        import shutil
                        temp_dir = os.path.join(folder_path, ".temp_unselected")
                        os.makedirs(temp_dir, exist_ok=True)
                        
                        # Move unselected LTA folders to temp
                        for item in os.listdir(folder_path):
                            item_path = os.path.join(folder_path, item)
                            if os.path.isdir(item_path) and item not in selected_lta_names and item != ".temp_unselected":
                                if "LTA" in item or "lta" in item:
                                    shutil.move(item_path, os.path.join(temp_dir, item))
                                    moved_folders.append(item)
                        
                        logger.info(f"Phase 1: Temporarily moved {len(moved_folders)} unselected folders")
                    
                    try:
                        # Construct command with phase "1" and LTA selection
                        cmd = [
                            python_exe,
                            script_path,
                            "1"  # Phase 1 selection
                        ]
                        
                        # If we moved folders (folder filtering), always use "all"
                        # because we've already filtered by hiding unselected folders
                        if moved_folders:
                            cmd.append("all")
                        # Otherwise use the original lta_selection indices
                        elif lta_selection is None or lta_selection == "all":
                            cmd.append("all")
                        elif isinstance(lta_selection, list):
                            # Convert list of indices to comma-separated string
                            indices_str = ",".join(str(i) for i in lta_selection)
                            cmd.append(indices_str)
                        else:
                            cmd.append("all")
                        
                        logger.info(f"Executing Phase 1: {script_path} with selection: {lta_selection}")
                        
                        # Set environment for UTF-8 and disable buffering
                        env = os.environ.copy()
                        env['PYTHONIOENCODING'] = 'utf-8'
                        env['PYTHONUNBUFFERED'] = '1'  # Disable Python buffering for real-time output
                        
                        # Run script - CREATE_NO_WINDOW hides the terminal window on Windows
                        creation_flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                        process = subprocess.Popen(
                            cmd,
                            cwd=project_root,  # Run from project root
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            text=True,
                            encoding='utf-8',
                            env=env,
                            stdin=subprocess.DEVNULL,
                            bufsize=1,  # Line buffered
                            creationflags=creation_flags
                        )
                        self.current_process = process
                        
                        # Open log file for BADR script output
                        log_file_path = os.path.join(project_root, "badr_login_test_logs.txt")
                        with open(log_file_path, 'a', encoding='utf-8') as log_file:
                            # Write session header
                            from datetime import datetime
                            log_file.write(f"\n{'='*70}\n")
                            log_file.write(f"PHASE 1 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                            log_file.write(f"{'='*70}\n\n")
                            
                            # Read output in real-time
                            stdout_lines = []
                            stderr_lines = []
                            
                            while True:
                                output = process.stdout.readline()
                                if output == '' and process.poll() is not None:
                                    break
                                if output:
                                    line = output.strip()
                                    stdout_lines.append(output)
                                    logger.info(f"Phase 1: {line}")
                                    
                                    # Write to log file
                                    log_file.write(output)
                                    log_file.flush()  # Ensure immediate write
                                    
                                    # Send to logs screen in real-time
                                    self.app.logs_screen.add_script_output(output, "badr")
                                    
                                    if progress_callback and line:
                                        if "Traitement du dossier" in line:
                                            progress_callback(50, f"Traitement: {line[:30]}...")
                                        elif "CONNEXION: Authentification réussie" in line:
                                            progress_callback(30, "Connexion réussie")
                            
                            rc = process.poll()
                            
                            # Read any remaining stderr
                            stderr_output = process.stderr.read()
                            if stderr_output:
                                stderr_lines.append(stderr_output)
                                log_file.write("\n=== ERRORS ===\n")
                                log_file.write(stderr_output + "\n")
                                self.app.logs_screen.add_script_output("\n=== ERRORS ===\n" + stderr_output + "\n", "badr")
                        
                        logger.info(f"BADR logs saved to: {log_file_path}")
                        
                        if rc != 0:
                            logger.error(f"Phase 1 failed with code {rc}")
                            raise Exception(f"Script failed with exit code {rc}")
                        
                        logger.info("Phase 1 completed successfully")
                        if progress_callback:
                            progress_callback(100, "Phase 1 terminée avec succès")
                        
                        if completion_callback:
                            completion_callback(success=True)
                    
                    finally:
                        # Restore moved folders (always, even on error)
                        if moved_folders and temp_dir:
                            import shutil
                            for folder in moved_folders:
                                src = os.path.join(temp_dir, folder)
                                dst = os.path.join(folder_path, folder)
                                try:
                                    if os.path.exists(src):
                                        shutil.move(src, dst)
                                except Exception as restore_err:
                                    logger.error(f"Failed to restore folder {folder}: {restore_err}")
                            
                            # Remove temp directory
                            try:
                                if os.path.exists(temp_dir) and not os.listdir(temp_dir):
                                    os.rmdir(temp_dir)
                            except Exception as cleanup_err:
                                logger.error(f"Failed to cleanup temp directory: {cleanup_err}")
                            
                            logger.info(f"Phase 1: Restored {len(moved_folders)} folders")
                
                except Exception as e:
                    logger.error(f"Phase 1 failed: {e}", exc_info=True)
                    if completion_callback:
                        completion_callback(success=False, error=str(e))
            finally:
                self.is_running = False
                self.current_process = None
        
        thread = threading.Thread(target=execute, daemon=True)
        thread.start()
    
    def run_phase2(self, folder_path, lta_selection, credentials, progress_callback=None, completion_callback=None):
        """
        Execute Phase 2 of badr_login_test.py
        
        Args:
            folder_path: Path to folder containing LTA folders
            lta_selection: List of LTA indices to process, or "all" for all LTAs
            credentials: Dict with 'username' and 'password'
            progress_callback: Function to call with progress updates
            completion_callback: Function to call when complete
        """
        def execute():
            try:
                self.is_running = True
                logger.info("Starting Phase 2 automation...")
                
                if progress_callback:
                    progress_callback(10, "Démarrage Phase 2...")
                
                # Find script location - support both development and .exe modes
                if getattr(sys, 'frozen', False):
                    # Running as .exe - scripts should be with the .exe or in selected folder's parent
                    exe_dir = os.path.dirname(sys.executable)
                    parent_of_selected = os.path.dirname(folder_path)
                    
                    # Try parent of selected folder first
                    if os.path.exists(os.path.join(parent_of_selected, "badr_login_test.py")):
                        project_root = parent_of_selected
                    # Then try exe directory
                    elif os.path.exists(os.path.join(exe_dir, "badr_login_test.py")):
                        project_root = exe_dir
                    # Finally try the selected folder itself
                    else:
                        project_root = folder_path
                else:
                    # Running in development - use normal path resolution
                    current_file = os.path.abspath(__file__)
                    utils_dir = os.path.dirname(current_file)
                    gui_dir = os.path.dirname(utils_dir)
                    project_root = os.path.dirname(gui_dir)
                
                script_path = os.path.join(project_root, "badr_login_test.py")
                
                if not os.path.exists(script_path):
                    raise FileNotFoundError(f"Script not found: {script_path}")
                
                # Determine the correct Python executable
                if getattr(sys, 'frozen', False):
                    # Running as .exe - use system Python
                    python_exe = 'python'
                else:
                    # Running in development - use current Python
                    python_exe = sys.executable
                
                # Construct command with phase "2" and LTA selection
                cmd = [
                    python_exe,
                    script_path,
                    "2"  # Phase 2 selection
                ]
                
                # Add LTA selection argument
                if lta_selection is None or lta_selection == "all":
                    cmd.append("all")
                elif isinstance(lta_selection, list):
                    # Convert list of indices to comma-separated string
                    indices_str = ",".join(str(i) for i in lta_selection)
                    cmd.append(indices_str)
                else:
                    cmd.append("all")
                
                logger.info(f"Executing Phase 2: {script_path} with selection: {lta_selection}")
                
                # Set environment for UTF-8 and disable buffering
                env = os.environ.copy()
                env['PYTHONIOENCODING'] = 'utf-8'
                env['PYTHONUNBUFFERED'] = '1'  # Disable Python buffering for real-time output
                
                # Run script - CREATE_NO_WINDOW hides the terminal window on Windows
                creation_flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                process = subprocess.Popen(
                    cmd,
                    cwd=project_root,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    env=env,
                    stdin=subprocess.DEVNULL,
                    bufsize=1,  # Line buffered
                    creationflags=creation_flags
                )
                self.current_process = process
                
                # Open log file for BADR script output
                log_file_path = os.path.join(project_root, "badr_login_test_logs.txt")
                with open(log_file_path, 'a', encoding='utf-8') as log_file:
                    # Write session header
                    from datetime import datetime
                    log_file.write(f"\n{'='*70}\n")
                    log_file.write(f"PHASE 2 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    log_file.write(f"{'='*70}\n\n")
                    
                    # Read output in real-time
                    stdout_lines = []
                    stderr_lines = []
                    
                    while True:
                        output = process.stdout.readline()
                        if output == '' and process.poll() is not None:
                            break
                        if output:
                            line = output.strip()
                            stdout_lines.append(output)
                            logger.info(f"Phase 2: {line}")
                            
                            # Write to log file
                            log_file.write(output)
                            log_file.flush()  # Ensure immediate write
                            
                            # Send to logs screen in real-time
                            self.app.logs_screen.add_script_output(output, "badr")
                            
                            if progress_callback and line:
                                if "DUMs traités" in line:
                                    progress_callback(50, f"Progression: {line}")
                                elif "CONNEXION: Authentification réussie" in line:
                                    progress_callback(20, "Connexion réussie")
                    
                    rc = process.poll()
                    
                    # Read any remaining stderr
                    stderr_output = process.stderr.read()
                    if stderr_output:
                        stderr_lines.append(stderr_output)
                        log_file.write("\n=== ERRORS ===\n")
                        log_file.write(stderr_output + "\n")
                        self.app.logs_screen.add_script_output("\n=== ERRORS ===\n" + stderr_output + "\n", "badr")
                
                logger.info(f"BADR logs saved to: {log_file_path}")
                
                if rc != 0:
                    logger.error(f"Phase 2 failed with code {rc}")
                    raise Exception(f"Script failed with exit code {rc}")
                
                logger.info("Phase 2 completed successfully")
                if progress_callback:
                    progress_callback(100, "Phase 2 terminée avec succès")
                
                if completion_callback:
                    completion_callback(success=True)
                
            except Exception as e:
                logger.error(f"Phase 2 failed: {e}", exc_info=True)
                if completion_callback:
                    completion_callback(success=False, error=str(e))
            finally:
                self.is_running = False
                self.current_process = None
        
        thread = threading.Thread(target=execute, daemon=True)
        thread.start()
    
    def stop(self):
        """Stop current execution"""
        if self.current_process:
            self.current_process.terminate()
            logger.info("Script execution terminated")
