"""
PDF to DOCX Converter GUI Application
Provides intuitive file selection and batch conversion capabilities
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import List
import threading
import os
from pdf_to_docx_converter import convert_pdf_to_docx, PDFToDOCXError, get_pdf_converter_info


class PDFToDOCXConverterGUI:
    """GUI application for PDF to DOCX conversion with batch support"""
    
    def __init__(self, root):
        """Initialize the GUI application"""
        self.root = root
        self.root.title("PDF to DOCX Converter")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Application state
        self.selected_files: List[str] = []
        self.output_directory: str = ""
        self.conversion_method = tk.StringVar(value="auto")
        self.is_converting = False
        
        # Create GUI elements
        self._create_widgets()
        self._load_available_methods()
        
    def _create_widgets(self):
        """Create all GUI widgets"""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # ═══ Title ═══
        title_label = ttk.Label(
            main_frame,
            text="PDF to DOCX Converter",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky=tk.W)
        
        # ═══ File Selection Section ═══
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        file_frame.columnconfigure(1, weight=1)
        
        # Single file button
        single_btn = ttk.Button(
            file_frame,
            text="📄 Select Single PDF",
            command=self._select_single_file
        )
        single_btn.grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # Folder button
        folder_btn = ttk.Button(
            file_frame,
            text="📁 Select Folder",
            command=self._select_folder
        )
        folder_btn.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Clear button
        clear_btn = ttk.Button(
            file_frame,
            text="🗑️ Clear Selection",
            command=self._clear_selection
        )
        clear_btn.grid(row=0, column=2, sticky=tk.E, padx=5)
        
        # Selected files label
        ttk.Label(file_frame, text="Selected Files:", font=("Arial", 10, "bold")).grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 5)
        )
        
        # Files listbox with scrollbar
        list_frame = ttk.Frame(file_frame)
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.files_listbox = tk.Listbox(
            list_frame,
            height=6,
            yscrollcommand=scrollbar.set,
            font=("Courier", 9)
        )
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.config(command=self.files_listbox.yview)
        
        # File count label
        self.file_count_label = ttk.Label(
            file_frame,
            text="Total files: 0",
            font=("Arial", 9)
        )
        self.file_count_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # ═══ Output Directory Section ═══
        output_frame = ttk.LabelFrame(main_frame, text="Output Directory", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        output_frame.columnconfigure(1, weight=1)
        
        # Output directory label
        ttk.Label(output_frame, text="Output Path:").grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # Output directory display
        self.output_label = ttk.Label(
            output_frame,
            text="(Same as source files)",
            font=("Courier", 9),
            foreground="gray"
        )
        self.output_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # Browse output button
        output_btn = ttk.Button(
            output_frame,
            text="📁 Browse",
            command=self._select_output_directory
        )
        output_btn.grid(row=0, column=2, sticky=tk.E, padx=5)
        
        # ═══ Conversion Settings Section ═══
        settings_frame = ttk.LabelFrame(main_frame, text="Conversion Settings", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        settings_frame.columnconfigure(1, weight=1)
        
        # Method selection
        ttk.Label(settings_frame, text="Conversion Method:").grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # Create method radio buttons
        methods_frame = ttk.Frame(settings_frame)
        methods_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5)
        
        self.method_radios = {}
        for idx, method in enumerate(['auto', 'pdfplumber', 'ocr', 'libreoffice', 'pymupdf']):
            radio = ttk.Radiobutton(
                methods_frame,
                text=method,
                variable=self.conversion_method,
                value=method
            )
            radio.grid(row=0, column=idx, padx=5)
            self.method_radios[method] = radio
        
        # Available methods info
        ttk.Label(settings_frame, text="Available Methods:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(10, 0))
        self.available_methods_label = ttk.Label(
            settings_frame,
            text="",
            font=("Courier", 9),
            foreground="green"
        )
        self.available_methods_label.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=(10, 0))
        
        # Progress Section
        progress_frame = ttk.LabelFrame(settings_frame, text="Progress", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Progress label
        self.progress_label = ttk.Label(
            progress_frame,
            text="Ready to convert",
            font=("Arial", 9)
        )
        self.progress_label.grid(row=1, column=0, sticky=tk.W, padx=5)
        
        # ═══ Action Buttons ═══
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        button_frame.columnconfigure(1, weight=1)
        
        # Convert button
        self.convert_btn = ttk.Button(
            button_frame,
            text="🚀 Start Conversion",
            command=self._start_conversion
        )
        self.convert_btn.grid(row=0, column=0, padx=5)
        
        # Stop button
        self.stop_btn = ttk.Button(
            button_frame,
            text="⏹️ Stop",
            command=self._stop_conversion,
            state=tk.DISABLED
        )
        self.stop_btn.grid(row=0, column=1, padx=5)
        
        # Open Output button
        open_btn = ttk.Button(
            button_frame,
            text="📂 Open Output Folder",
            command=self._open_output_folder
        )
        open_btn.grid(row=0, column=2, sticky=tk.E, padx=5)
        
        # Status bar at bottom
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=1, pady=1)
        self.root.columnconfigure(0, weight=1)
    
    def _load_available_methods(self):
        """Load and display available conversion methods"""
        try:
            info = get_pdf_converter_info()
            available = [method for method, available in info.items() if available]
            
            if available:
                self.available_methods_label.config(
                    text=f"✅ {', '.join(available)}",
                    foreground="green"
                )
            else:
                self.available_methods_label.config(
                    text="❌ No methods available - please install dependencies",
                    foreground="red"
                )
                self.convert_btn.config(state=tk.DISABLED)
                
        except Exception as e:
            self.available_methods_label.config(
                text=f"❌ Error checking methods: {str(e)}",
                foreground="red"
            )
    
    def _select_single_file(self):
        """Select a single PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            self.selected_files = [file_path]
            self._update_file_list()
    
    def _select_folder(self):
        """Select a folder and add all PDFs in it"""
        folder_path = filedialog.askdirectory(
            title="Select Folder with PDF Files"
        )
        
        if folder_path:
            folder = Path(folder_path)
            pdf_files = list(folder.glob("*.pdf"))
            
            if not pdf_files:
                messagebox.showwarning(
                    "No PDFs Found",
                    f"No PDF files found in:\n{folder_path}"
                )
                return
            
            self.selected_files = [str(f) for f in sorted(pdf_files)]
            self._update_file_list()
    
    def _clear_selection(self):
        """Clear all selected files"""
        self.selected_files = []
        self.files_listbox.delete(0, tk.END)
        self.file_count_label.config(text="Total files: 0")
        self.status_var.set("Selection cleared")
    
    def _select_output_directory(self):
        """Select output directory"""
        folder_path = filedialog.askdirectory(
            title="Select Output Directory"
        )
        
        if folder_path:
            self.output_directory = folder_path
            self.output_label.config(
                text=folder_path,
                foreground="black"
            )
            self.status_var.set(f"Output directory: {folder_path}")
    
    def _update_file_list(self):
        """Update the file listbox"""
        self.files_listbox.delete(0, tk.END)
        
        for file_path in self.selected_files:
            display_path = Path(file_path).name
            self.files_listbox.insert(tk.END, display_path)
        
        count = len(self.selected_files)
        self.file_count_label.config(text=f"Total files: {count}")
        self.status_var.set(f"{count} file(s) selected")
    
    def _start_conversion(self):
        """Start conversion process in a separate thread"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select PDF file(s) to convert")
            return
        
        # Disable buttons during conversion
        self.convert_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.is_converting = True
        
        # Start conversion in background thread
        thread = threading.Thread(target=self._conversion_worker)
        thread.daemon = True
        thread.start()
    
    def _stop_conversion(self):
        """Stop conversion process"""
        self.is_converting = False
        self.convert_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("Conversion stopped by user")
        self.progress_var.set(0)
    
    def _conversion_worker(self):
        """Worker thread for batch conversion"""
        method = self.conversion_method.get()
        total = len(self.selected_files)
        successful = 0
        failed = 0
        errors = []
        
        try:
            for idx, pdf_file in enumerate(self.selected_files, 1):
                if not self.is_converting:
                    break
                
                try:
                    # Determine output path
                    pdf_path = Path(pdf_file)
                    if self.output_directory:
                        output_path = Path(self.output_directory) / (pdf_path.stem + '.docx')
                    else:
                        output_path = pdf_path.parent / (pdf_path.stem + '.docx')
                    
                    # Update progress
                    progress = int((idx - 1) / total * 100)
                    self.progress_var.set(progress)
                    self.progress_label.config(
                        text=f"Converting {idx}/{total}: {pdf_path.name}"
                    )
                    self.status_var.set(f"Converting: {pdf_path.name}")
                    self.root.update()
                    
                    # Perform conversion
                    result = convert_pdf_to_docx(
                        pdf_file,
                        str(output_path),
                        method=method,
                        timeout=300
                    )
                    
                    successful += 1
                    
                except PDFToDOCXError as e:
                    failed += 1
                    error_msg = f"{Path(pdf_file).name}: {str(e)}"
                    errors.append(error_msg)
                except Exception as e:
                    failed += 1
                    error_msg = f"{Path(pdf_file).name}: {str(e)}"
                    errors.append(error_msg)
            
            # Conversion complete
            self.progress_var.set(100)
            self.progress_label.config(
                text=f"Conversion Complete: {successful} succeeded, {failed} failed"
            )
            self.status_var.set("Conversion complete")
            
            # Show results dialog
            self.root.after(0, self._show_conversion_results, successful, failed, errors)
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Conversion Error", f"Unexpected error: {str(e)}")
        
        finally:
            self.convert_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_converting = False
    
    def _show_conversion_results(self, successful, failed, errors):
        """Show conversion results dialog"""
        total = successful + failed
        
        if failed == 0:
            messagebox.showinfo(
                "✅ Conversion Complete",
                f"Successfully converted {successful} PDF file(s) to DOCX!"
            )
        else:
            error_details = "\n".join(errors[:5])  # Show first 5 errors
            if len(errors) > 5:
                error_details += f"\n... and {len(errors) - 5} more errors"
            
            messagebox.showwarning(
                "⚠️ Conversion Complete with Errors",
                f"Successful: {successful}/{total}\n"
                f"Failed: {failed}/{total}\n\n"
                f"Errors:\n{error_details}"
            )
    
    def _open_output_folder(self):
        """Open output folder in file explorer"""
        folder = self.output_directory if self.output_directory else str(Path.home())
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(folder)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{folder}"' if sys.platform == 'darwin' else f'xdg-open "{folder}"')
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {str(e)}")


def main():
    """Main entry point for the GUI application"""
    root = tk.Tk()
    app = PDFToDOCXConverterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    import sys
    main()
