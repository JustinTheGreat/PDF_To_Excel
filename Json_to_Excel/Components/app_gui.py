import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
from Components.json_processor import JsonProcessor
from Components.excel_generator import ExcelGenerator
from Components.text_filters import TextFilter
from Components.business_rules import BusinessRules

class JsonToExcelApp:
    def __init__(self, root, debug_mode=False):
        self.root = root
        self.root.title("JSON to Excel Converter")
        self.debug_mode = debug_mode
        
        # Set window size - larger if debug mode is enabled
        if debug_mode:
            self.root.geometry("800x700")  # Larger height for debug area
        else:
            self.root.geometry("600x500")
        
        self.root.resizable(True, True)
        
        # Initialize the Excel generator
        self.excel_generator = ExcelGenerator()
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#ccc")
        self.style.configure("TLabel", padding=6)
        self.style.configure("TFrame", padding=10)
        
        # Main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Variables
        self.json_dir_var = tk.StringVar()
        self.excel_dir_var = tk.StringVar()
        self.file_name_var = tk.StringVar(value="output.xlsx")
        self.filter_text_var = tk.StringVar()
        self.filter_units_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.progress_var = tk.DoubleVar()
        
        # Create the UI
        self.create_ui(main_frame)
    
    def create_ui(self, main_frame):
        """Create the user interface elements."""
        # JSON directory selection
        ttk.Label(main_frame, text="JSON Files Directory:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        dir_frame = ttk.Frame(main_frame)
        dir_frame.grid(row=0, column=1, sticky=tk.EW, pady=5)
        
        ttk.Entry(dir_frame, textvariable=self.json_dir_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(dir_frame, text="Browse...", command=self.browse_json_dir).pack(side=tk.RIGHT, padx=5)
        
        # Excel directory selection
        ttk.Label(main_frame, text="Output Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        excel_dir_frame = ttk.Frame(main_frame)
        excel_dir_frame.grid(row=1, column=1, sticky=tk.EW, pady=5)
        
        ttk.Entry(excel_dir_frame, textvariable=self.excel_dir_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(excel_dir_frame, text="Browse...", command=self.browse_excel_dir).pack(side=tk.RIGHT, padx=5)
        
        # Output file name
        ttk.Label(main_frame, text="Output File Name:").grid(row=2, column=0, sticky=tk.W, pady=5)
        file_name_frame = ttk.Frame(main_frame)
        file_name_frame.grid(row=2, column=1, sticky=tk.EW, pady=5)
        
        ttk.Entry(file_name_frame, textvariable=self.file_name_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(file_name_frame, text=".xlsx").pack(side=tk.RIGHT)
        
        # Filename filter
        ttk.Label(main_frame, text="Filter Text to Remove:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.filter_text_var, width=40).grid(row=3, column=1, sticky=tk.EW, pady=5)
        ttk.Label(main_frame, text="(Remove common text from filenames)").grid(row=4, column=1, sticky=tk.W)
        
        # Filter units option
        ttk.Checkbutton(
            main_frame,
            text="Remove units from values (e.g., [ms], [V])",
            variable=self.filter_units_var
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Recursive search option
        ttk.Checkbutton(
            main_frame,
            text="Search in subdirectories",
            variable=self.recursive_var
        ).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Process button
        ttk.Button(main_frame, text="Process JSON Files", command=self.process_files).grid(row=7, column=0, columnspan=2, pady=15)
        
        # Progress bar
        ttk.Label(main_frame, text="Progress:").grid(row=8, column=0, sticky=tk.W, pady=5)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, length=300, mode='determinate')
        self.progress_bar.grid(row=8, column=1, sticky=tk.EW, pady=5)
        
        # Status
        ttk.Label(main_frame, text="Status:").grid(row=9, column=0, sticky=tk.NW, pady=5)
        self.status_label = ttk.Label(main_frame, text="Ready", wraplength=400)
        self.status_label.grid(row=9, column=1, sticky=tk.W, pady=5)
        
        # Debug Text Area (only shown in debug mode)
        if self.debug_mode:
            ttk.Label(main_frame, text="Debug Log:").grid(row=10, column=0, sticky=tk.NW, pady=5)
            self.debug_text = scrolledtext.ScrolledText(main_frame, width=70, height=15)
            self.debug_text.grid(row=10, column=1, sticky=tk.NSEW, pady=5)
            
            # Add clear log button
            ttk.Button(main_frame, text="Clear Log", command=self.clear_debug_log).grid(row=11, column=0, columnspan=2, pady=5)
            
            # Make debug area expandable
            main_frame.rowconfigure(10, weight=1)
        
        # Configure grid weight
        main_frame.columnconfigure(1, weight=1)
        
        # Add padding to all widgets
        for child in main_frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
    
    def clear_debug_log(self):
        """Clear the debug log area."""
        if self.debug_mode and hasattr(self, 'debug_text'):
            self.debug_text.delete(1.0, tk.END)
    
    def log_debug(self, message):
        """Add a message to the debug log area."""
        if self.debug_mode and hasattr(self, 'debug_text'):
            self.debug_text.insert(tk.END, f"{message}\n")
            self.debug_text.see(tk.END)  # Scroll to the end
            self.root.update_idletasks()
        print(message)  # Also print to console
    
    def browse_json_dir(self):
        """Browse for JSON directory."""
        directory = filedialog.askdirectory(title="Select Directory with JSON Files")
        if directory:
            self.json_dir_var.set(directory)
            if self.debug_mode:
                self.log_debug(f"Selected JSON directory: {directory}")
                # List JSON files in the directory
                json_files = [f for f in os.listdir(directory) if f.lower().endswith('.json')]
                self.log_debug(f"JSON files in directory: {json_files}")
    
    def browse_excel_dir(self):
        """Browse for Excel output directory."""
        directory = filedialog.askdirectory(title="Select Directory to Save Excel File")
        if directory:
            self.excel_dir_var.set(directory)
            if self.debug_mode:
                self.log_debug(f"Selected Excel output directory: {directory}")
    
    def update_ui(self, update_type, value):
        """Update UI elements based on processing status."""
        if update_type == "progress":
            self.progress_var.set(value)
        elif update_type == "status":
            self.status_label.config(text=value)
            if self.debug_mode:
                self.log_debug(f"STATUS: {value}")
        elif update_type == "debug" and self.debug_mode:
            self.log_debug(value)
        self.root.update_idletasks()
    
    def process_files(self):
        """Process JSON files and create Excel."""
        json_dir = self.json_dir_var.get()
        excel_dir = self.excel_dir_var.get()
        file_name = self.file_name_var.get()
        filter_text = self.filter_text_var.get()
        filter_units = self.filter_units_var.get()
        recursive = self.recursive_var.get()
        
        # Clear the debug log
        if self.debug_mode:
            self.clear_debug_log()
            self.log_debug("=== Starting JSON to Excel Conversion ===")
            self.log_debug(f"JSON Directory: {json_dir}")
            self.log_debug(f"Excel Directory: {excel_dir}")
            self.log_debug(f"File Name: {file_name}")
            self.log_debug(f"Filter Text: {filter_text}")
            self.log_debug(f"Filter Units: {filter_units}")
            self.log_debug(f"Recursive: {recursive}")
        
        # Ensure file name has .xlsx extension
        if not file_name.lower().endswith('.xlsx'):
            file_name += '.xlsx'
        
        if not json_dir or not excel_dir:
            messagebox.showerror("Error", "Please select both JSON directory and Excel output directory")
            return
        
        if not file_name:
            messagebox.showerror("Error", "Please enter a file name")
            return
        
        # Combine directory and filename
        excel_file = os.path.join(excel_dir, file_name)
        
        # Reset progress
        self.progress_var.set(0)
        self.status_label.config(text="Processing...")
        
        # Run processing in a separate thread to avoid freezing UI
        threading.Thread(
            target=self.run_processing, 
            args=(json_dir, excel_file, filter_text, filter_units, recursive),
            daemon=True
        ).start()
    
    def run_processing(self, json_dir, excel_file, filter_text, filter_units, recursive):
        """Run the JSON processing and Excel creation in a background thread."""
        try:
            # Update status
            if recursive:
                self.update_ui("status", "Reading JSON files in directory and subdirectories...")
            else:
                self.update_ui("status", "Reading JSON files in main directory...")
            
            # Read all JSON files
            self.update_ui("debug", f"Starting to read JSON files from: {json_dir}")
            all_json_data = JsonProcessor.read_json_files(json_dir, recursive, self.debug_mode)
            
            # Log how many files were found
            file_count = len(all_json_data)
            self.update_ui("debug", f"Found {file_count} JSON files")
            
            # Apply business rules to transform the data
            self.update_ui("status", "Applying business rules...")
            self.update_ui("debug", "Starting to apply business rules transformations")
            transformed_data = BusinessRules.transform_all_data(all_json_data)
            self.update_ui("debug", "Business rules applied successfully")
            
            if not transformed_data:
                self.update_ui("status", "No JSON files found or all files were empty")
                self.root.after(0, lambda: messagebox.showerror("Error", "No valid JSON files found in the selected directory"))
                return
            
            # Show file count
            self.update_ui("status", f"Found {file_count} JSON files. Creating Excel file...")
            
            # Create Excel file with the transformed data
            self.update_ui("debug", f"Starting Excel file creation at: {excel_file}")
            success = self.excel_generator.create_excel_file(
                transformed_data,  # Use transformed data instead of original
                excel_file,
                filter_text,
                filter_units,
                self.update_ui
            )
            
            if success:
                self.update_ui("debug", "Excel creation completed successfully")
                self.root.after(0, lambda: messagebox.showinfo("Success", f"Excel file created successfully at:\n{excel_file}"))
            else:
                self.update_ui("debug", "Excel creation failed")
                self.root.after(0, lambda: messagebox.showerror("Error", "Failed to process JSON files. Check the status for details."))
                
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.update_ui("status", error_msg)
            self.update_ui("debug", f"EXCEPTION: {str(e)}")
            import traceback
            self.update_ui("debug", f"STACK TRACE: {traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}"))