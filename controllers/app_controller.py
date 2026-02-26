import os
import sys
import subprocess
import threading
from typing import List

from views import MainView
from services import ConversionService

# half the shit is AI can't be asked to properly write code sry ;(
# This is very simple idk why u would wanna sit and waste hours 

class AppController:    
    def __init__(self, view: MainView):
        self.view = view
        self.service = ConversionService()
        self.files: List[str] = []
        
        self._bind_callbacks()
        self._check_backend_availability()

    def _open_path(self, path: str) -> None:
        if sys.platform == "win32":
            os.startfile(path)
            return
        if sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
            return
        subprocess.run(["xdg-open", path], check=False)

    def _check_backend_availability(self):
        try:
            available = self.service.get_available_backends()
            if not available:
                self.view.show_warning("No Backend Found", self.service.get_install_message())
                self.view.update_status("No backend found - install required app", 0)
                return

            backend_name = self.service.get_active_backend_name()
            self.view.update_status(f"Ready - Backend: {backend_name}", 0)
        except Exception as e:
            self.view.show_warning("Backend Check Failed", str(e))
            self.view.update_status("Backend check failed", 0)
    
    def _bind_callbacks(self):
        self.view.on_add_files = self.add_files
        self.view.on_remove_selected = self.remove_selected
        self.view.on_clear_all = self.clear_all
        self.view.on_move_up = self.move_up
        self.view.on_move_down = self.move_down
        self.view.on_sort_files = self.sort_files
        self.view.on_convert = self.start_conversion
        self.view.on_convert_separate = self.start_separate_conversion
        self.view.on_drag_reorder = self.drag_reorder
    
    def add_files(self):
        new_files = self.view.ask_open_files()
        
        for f in new_files:
            if f not in self.files:
                self.files.append(f)
        
        self._refresh_list()
    
    def remove_selected(self):
        i = self.view.get_selected_index()
        if i is not None:
            self.files.pop(i)
            self._refresh_list()
    
    def clear_all(self):
        self.files.clear()
        self._refresh_list()
    
    def move_up(self):
        i = self.view.get_selected_index()
        if i is not None and i > 0:
            self.files[i], self.files[i-1] = self.files[i-1], self.files[i]
            self._refresh_list()
            self.view.set_selection(i - 1)
    
    def move_down(self):
        i = self.view.get_selected_index()
        if i is not None and i < len(self.files) - 1:
            self.files[i], self.files[i+1] = self.files[i+1], self.files[i]
            self._refresh_list()
            self.view.set_selection(i + 1)
    
    def drag_reorder(self, from_index: int, to_index: int):
        if from_index is not None and 0 <= from_index < len(self.files):
            i = self.files.pop(from_index)
            self.files.insert(to_index, i)
            self._refresh_list()
            self.view.set_selection(to_index)

    def sort_files(self):
        self.files.sort(key=lambda path: os.path.basename(path).lower())
        self._refresh_list()
    
    def start_conversion(self):
        if not self.files:
            self.view.show_warning("No Files", "Add PowerPoint files first.")
            return
        
        dPath = self.view.ask_save_file()
        if not dPath:
            return
        
        # Disable button during conversion
        self.view.set_convert_button_enabled(False)
        
        # Run conversion in separate thread
        th = threading.Thread(
            target=self._do_conversion,
            args=(dPath,),
            daemon=True
        )
        th.start()
    
    def _do_conversion(self, output_path: str):
        try:
            def progress_callback(message: str, progress: float):
                self.view.schedule(self.view.update_status, message, progress)
            
            self.service.convert_and_merge(
                ppt_files=self.files,
                output_path=output_path,
                progress_callback=progress_callback,
                delete_temp=self.view.delete_temp_files
            )
            
            self.view.schedule(
                self.view.show_info,
                "Success",
                f"PDF created successfully!\n{output_path}"
            )
            
            # Open the PDF if requested
            if self.view.open_after_conversion:
                self._open_path(output_path)
        
        except Exception as e:
            self.view.schedule(
                self.view.show_error,
                "Error",
                f"Conversion failed:\n{str(e)}"
            )
            self.view.schedule(self.view.update_status, "Error", 0)
        
        finally:
            self.view.schedule(self.view.set_convert_button_enabled, True)
    
    def start_separate_conversion(self):
        if not self.files:
            self.view.show_warning("No Files", "Please add PPT files first.")
            return
        
        output_dir = self.view.ask_directory()
        if not output_dir:
            return
        
        # Disable buttons during conversion
        self.view.set_convert_button_enabled(False)
        
        # Run conversion in separate thread
        th = threading.Thread(
            target=self._do_separate_conversion,
            args=(output_dir,),
            daemon=True
        )
        th.start()
    
    def _do_separate_conversion(self, output_dir: str):
        try:
            def progress_callback(message: str, progress: float):
                self.view.schedule(self.view.update_status, message, progress)
            
            created_files = self.service.convert_separate(
                ppt_files=self.files,
                output_dir=output_dir,
                progress_callback=progress_callback
            )
            
            self.view.schedule(
                self.view.show_info,
                "Done",
                f"{len(created_files)} PDF files created in:\n{output_dir}"
            )
            
            # Open the folder if requested
            if self.view.open_after_conversion:
                self._open_path(output_dir)
        
        except Exception as e:
            self.view.schedule(
                self.view.show_error,
                "Error",
                f"Conversion failed:\n{str(e)}"
            )
            self.view.schedule(self.view.update_status, "Error", 0)
        
        finally:
            self.view.schedule(self.view.set_convert_button_enabled, True)
    
    def _refresh_list(self):
        self.view.update_file_list(self.files)
