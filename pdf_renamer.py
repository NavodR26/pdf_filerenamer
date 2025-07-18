import os
import sys
import subprocess
import importlib.util
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path

class DependencyManager:
    """Manages and installs required dependencies"""
    
    REQUIRED_PACKAGES = {
        'pandas': 'pandas',
        'PyMuPDF': 'fitz',
        'ttkthemes': 'ttkthemes'
    }
    
    @staticmethod
    def check_package(package_name, import_name=None):
        """Check if a package is installed"""
        if import_name is None:
            import_name = package_name
        
        try:
            importlib.import_module(import_name)
            return True
        except ImportError:
            return False
    
    @staticmethod
    def install_package(package_name):
        """Install a package using pip"""
        try:
            subprocess.check_call([
                sys.executable, '-m', 'pip', 'install', package_name, '--user'
            ])
            return True
        except subprocess.CalledProcessError:
            return False
    
    @classmethod
    def ensure_dependencies(cls):
        """Ensure all required dependencies are installed"""
        missing_packages = []
        
        for package_name, import_name in cls.REQUIRED_PACKAGES.items():
            if not cls.check_package(package_name, import_name):
                missing_packages.append(package_name)
        
        if missing_packages:
            # Show installation dialog
            root = tk.Tk()
            root.withdraw()  # Hide main window
            
            response = messagebox.askyesno(
                "Missing Dependencies",
                f"The following packages are required but not installed:\n\n"
                f"• {chr(10).join(missing_packages)}\n\n"
                f"Would you like to install them automatically?\n"
                f"(This may take a few minutes)",
                icon='warning'
            )
            
            if response:
                progress_window = tk.Toplevel()
                progress_window.title("Installing Dependencies")
                progress_window.geometry("400x200")
                progress_window.resizable(False, False)
                
                # Center the window
                progress_window.transient(root)
                progress_window.grab_set()
                
                ttk.Label(progress_window, text="Installing required packages...").pack(pady=20)
                
                progress_text = scrolledtext.ScrolledText(progress_window, height=6, width=45)
                progress_text.pack(pady=10, padx=20, fill='both', expand=True)
                
                progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
                progress_bar.pack(pady=10, padx=20, fill='x')
                progress_bar.start()
                
                def install_packages():
                    success_count = 0
                    for package in missing_packages:
                        progress_text.insert(tk.END, f"Installing {package}...\n")
                        progress_text.see(tk.END)
                        progress_window.update()
                        
                        if cls.install_package(package):
                            progress_text.insert(tk.END, f"✓ {package} installed successfully\n")
                            success_count += 1
                        else:
                            progress_text.insert(tk.END, f"✗ Failed to install {package}\n")
                        
                        progress_text.see(tk.END)
                        progress_window.update()
                    
                    progress_bar.stop()
                    
                    if success_count == len(missing_packages):
                        progress_text.insert(tk.END, "\n✓ All packages installed successfully!")
                        messagebox.showinfo("Success", "All dependencies installed successfully!\nPlease restart the application.")
                    else:
                        progress_text.insert(tk.END, f"\n⚠ {len(missing_packages) - success_count} packages failed to install.")
                        messagebox.showwarning("Partial Success", "Some packages failed to install.\nPlease install them manually or check your internet connection.")
                    
                    progress_window.after(2000, progress_window.destroy)
                
                # Start installation in a separate thread
                install_thread = threading.Thread(target=install_packages)
                install_thread.daemon = True
                install_thread.start()
                
                progress_window.mainloop()
                return False  # Indicate restart needed
            else:
                messagebox.showerror("Error", "Required dependencies are not installed.\nThe application cannot run without them.")
                return False
        
        return True


class PDFRenamer:
    """Main PDF Renaming Application"""
    
    def __init__(self):
        self.pdf_files = []
        self.csv_file = ""
        self.output_dir = ""
        self.setup_ui()
    
    def setup_ui(self):
        """Set up the user interface"""
        # Try to use themed tkinter, fall back to regular tkinter
        try:
            from ttkthemes import ThemedTk
            self.root = ThemedTk(theme="arc")
        except ImportError:
            self.root = tk.Tk()
        
        self.root.title("PDF File Renamer v2.0")
        self.root.geometry("700x600")
        
        # Set icon if available
        self.set_icon()
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # PDF files selection
        ttk.Label(file_frame, text="PDF Files:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.pdf_files_label = ttk.Label(file_frame, text="No files selected", foreground="gray")
        self.pdf_files_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        ttk.Button(file_frame, text="Select PDF Files", command=self.select_pdf_files).grid(row=0, column=2, padx=(10, 0), pady=5)
        
        # CSV file selection
        ttk.Label(file_frame, text="Name List (CSV):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.csv_file_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.csv_file_label.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        ttk.Button(file_frame, text="Select CSV File", command=self.select_csv_file).grid(row=1, column=2, padx=(10, 0), pady=5)
        
        # Output directory selection
        ttk.Label(file_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_dir_label = ttk.Label(file_frame, text="No directory selected", foreground="gray")
        self.output_dir_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        ttk.Button(file_frame, text="Select Output Directory", command=self.select_output_dir).grid(row=2, column=2, padx=(10, 0), pady=5)
        
        # Options section
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.fully_rename_var = tk.BooleanVar()
        fully_rename_checkbox = ttk.Checkbutton(
            options_frame, 
            text="Fully Rename Files (replace original name entirely)", 
            variable=self.fully_rename_var
        )
        fully_rename_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        self.case_sensitive_var = tk.BooleanVar()
        case_sensitive_checkbox = ttk.Checkbutton(
            options_frame, 
            text="Case Sensitive Search", 
            variable=self.case_sensitive_var
        )
        case_sensitive_checkbox.grid(row=1, column=0, sticky=tk.W)
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        self.start_button = ttk.Button(button_frame, text="Start Renaming Process", command=self.start_renaming)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT)
        
        # Output section
        output_frame = ttk.LabelFrame(main_frame, text="Process Output", padding="10")
        output_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Text output with scrollbar
        self.text_output = scrolledtext.ScrolledText(output_frame, wrap="word", height=15, width=70)
        self.text_output.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Progress section
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to start...")
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def set_icon(self):
        """Set application icon if available"""
        try:
            # Look for icon file in the same directory as the script
            script_dir = Path(__file__).parent if not getattr(sys, 'frozen', False) else Path(sys.executable).parent
            icon_path = script_dir / "file_renamer.ico"
            
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
            else:
                # Try alternative locations
                alternative_paths = [
                    Path.cwd() / "file_renamer.ico",
                    Path.cwd() / "icon.ico",
                    Path.cwd() / "app.ico"
                ]
                
                for path in alternative_paths:
                    if path.exists():
                        self.root.iconbitmap(str(path))
                        break
        except Exception:
            pass  # If icon setting fails, continue without icon
    
    def select_pdf_files(self):
        """Handle PDF file selection"""
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if files:
            self.pdf_files = list(files)
            self.pdf_files_label.config(text=f"{len(files)} PDF files selected", foreground="black")
            self.log_message(f"Selected {len(files)} PDF files")
        else:
            self.pdf_files = []
            self.pdf_files_label.config(text="No files selected", foreground="gray")
    
    def select_csv_file(self):
        """Handle CSV file selection"""
        file_path = filedialog.askopenfilename(
            title="Select CSV file containing names",
            filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            self.csv_file = file_path
            filename = Path(file_path).name
            self.csv_file_label.config(text=filename, foreground="black")
            self.log_message(f"Selected CSV file: {filename}")
        else:
            self.csv_file = ""
            self.csv_file_label.config(text="No file selected", foreground="gray")
    
    def select_output_dir(self):
        """Handle output directory selection"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        
        if directory:
            self.output_dir = directory
            self.output_dir_label.config(text=directory, foreground="black")
            self.log_message(f"Selected output directory: {directory}")
        else:
            self.output_dir = ""
            self.output_dir_label.config(text="No directory selected", foreground="gray")
    
    def clear_all(self):
        """Clear all selections and output"""
        self.pdf_files = []
        self.csv_file = ""
        self.output_dir = ""
        
        self.pdf_files_label.config(text="No files selected", foreground="gray")
        self.csv_file_label.config(text="No file selected", foreground="gray")
        self.output_dir_label.config(text="No directory selected", foreground="gray")
        
        self.text_output.delete(1.0, tk.END)
        self.progress_label.config(text="Ready to start...")
        self.progress_bar.config(value=0)
        
        self.log_message("All selections cleared")
    
    def log_message(self, message):
        """Add message to output text widget"""
        self.text_output.insert(tk.END, f"{message}\n")
        self.text_output.see(tk.END)
        self.root.update_idletasks()
    
    def read_pdf_content(self, pdf_file):
        """Read content from PDF file"""
        try:
            import fitz  # PyMuPDF
            with fitz.open(pdf_file) as doc:
                text = ""
                for page in doc:
                    text += page.get_text()
                return text
        except Exception as e:
            self.log_message(f"Error reading PDF {pdf_file}: {str(e)}")
            return ""
    
    def load_name_list(self):
        """Load names from CSV file"""
        try:
            names = []
            with open(self.csv_file, 'r', encoding='utf-8') as file:
                for line in file:
                    name = line.strip()
                    if name:  # Skip empty lines
                        names.append(name)
            return names
        except Exception as e:
            self.log_message(f"Error reading CSV file: {str(e)}")
            return []
    
    def generate_unique_filename(self, base_path, extension):
        """Generate a unique filename if file already exists"""
        counter = 1
        new_path = base_path
        
        while os.path.exists(new_path + extension):
            new_path = f"{base_path}_{counter}"
            counter += 1
        
        return new_path + extension
    
    def rename_files(self):
        """Perform the actual file renaming"""
        if not all([self.pdf_files, self.csv_file, self.output_dir]):
            messagebox.showerror("Error", "Please select PDF files, CSV file, and output directory.")
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Load names from CSV
        name_list = self.load_name_list()
        if not name_list:
            messagebox.showerror("Error", "No names found in CSV file.")
            return
        
        self.log_message(f"Loaded {len(name_list)} names from CSV file")
        self.log_message("Starting renaming process...")
        
        renamed_count = 0
        total_files = len(self.pdf_files)
        
        self.progress_bar.config(maximum=total_files)
        
        for index, pdf_file in enumerate(self.pdf_files):
            try:
                self.progress_label.config(text=f"Processing {index + 1}/{total_files}: {Path(pdf_file).name}")
                self.progress_bar.config(value=index + 1)
                self.root.update_idletasks()
                
                # Read PDF content
                pdf_content = self.read_pdf_content(pdf_file)
                if not pdf_content:
                    self.log_message(f"⚠ Could not read content from {Path(pdf_file).name}")
                    continue
                
                # Search for names in PDF content
                found_name = None
                for name in name_list:
                    if self.case_sensitive_var.get():
                        if name in pdf_content:
                            found_name = name
                            break
                    else:
                        if name.lower() in pdf_content.lower():
                            found_name = name
                            break
                
                if found_name:
                    # Generate new filename
                    original_filename = Path(pdf_file).stem
                    
                    if self.fully_rename_var.get():
                        new_filename = found_name
                    else:
                        new_filename = f"{found_name}_{original_filename}"
                    
                    # Create full path
                    new_filepath = os.path.join(self.output_dir, new_filename + '.pdf')
                    
                    # Handle duplicate filenames
                    if os.path.exists(new_filepath):
                        new_filepath = self.generate_unique_filename(
                            os.path.join(self.output_dir, new_filename), '.pdf'
                        )
                    
                    # Rename file
                    os.rename(pdf_file, new_filepath)
                    self.log_message(f"✓ Renamed: {Path(pdf_file).name} → {Path(new_filepath).name}")
                    renamed_count += 1
                else:
                    self.log_message(f"⚠ No matching name found for: {Path(pdf_file).name}")
                
            except Exception as e:
                self.log_message(f"✗ Error processing {Path(pdf_file).name}: {str(e)}")
        
        self.progress_label.config(text=f"Completed: {renamed_count}/{total_files} files renamed")
        self.log_message(f"\nRenaming process completed!")
        self.log_message(f"Successfully renamed {renamed_count} out of {total_files} files.")
        
        messagebox.showinfo("Success", f"Renaming completed!\n{renamed_count} files renamed successfully.")
    
    def start_renaming(self):
        """Start the renaming process in a separate thread"""
        if not all([self.pdf_files, self.csv_file, self.output_dir]):
            messagebox.showerror("Error", "Please select PDF files, CSV file, and output directory.")
            return
        
        # Disable start button during processing
        self.start_button.config(state='disabled')
        
        # Start renaming in a separate thread to prevent UI freezing
        rename_thread = threading.Thread(target=self.rename_files_thread)
        rename_thread.daemon = True
        rename_thread.start()
    
    def rename_files_thread(self):
        """Thread wrapper for rename_files"""
        try:
            self.rename_files()
        finally:
            # Re-enable start button
            self.start_button.config(state='normal')
    
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    """Main function"""
    print("PDF File Renamer v2.0")
    print("Checking dependencies...")
    
    # Check and install dependencies
    if not DependencyManager.ensure_dependencies():
        print("Dependency installation failed or incomplete. Please restart the application.")
        return
    
    print("All dependencies are available. Starting application...")
    
    # Create and run the application
    app = PDFRenamer()
    app.run()


if __name__ == "__main__":
    main()
