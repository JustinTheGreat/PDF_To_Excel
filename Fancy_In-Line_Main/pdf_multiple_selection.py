import tkinter as tk
from tkinter import filedialog, messagebox
import os
from Components.pdf_processor import create_document_json

def launch_single_file_dialog(extraction_params):
    """
    Launch a dialog to select and process a single PDF file.
    
    Args:
        extraction_params (list): Extraction parameters for document processing
        
    Returns:
        str: Path to the created JSON file or None if failed
    """
    # Open file dialog to select a single PDF
    pdf_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )
    
    if not pdf_path:
        return None
        
    json_path = create_document_json(pdf_path, extraction_params)
    
    if json_path:
        messagebox.showinfo("Success", f"JSON file created: {json_path}")
        return json_path
    else:
        messagebox.showerror("Error", "Failed to create JSON file.")
        return None

def launch_multiple_files_dialog(extraction_params, process_multiple_function):
    """
    Launch a dialog to select and process multiple PDF files.
    
    Args:
        extraction_params (list): Extraction parameters for document processing
        process_multiple_function (function): Function to process multiple files
        
    Returns:
        str: Path to the created combined JSON file or None if failed
    """
    # Open file dialog to select multiple PDFs
    pdf_paths = filedialog.askopenfilenames(
        title="Select PDF Files",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )
    
    if not pdf_paths:
        return None
        
    combined_json_path = process_multiple_function(pdf_paths, extraction_params)
    
    if combined_json_path:
        messagebox.showinfo("Success", f"Combined JSON file created: {combined_json_path}")
        return combined_json_path
    else:
        messagebox.showerror("Error", "Failed to create combined JSON file.")
        return None

def launch_mode_selection_dialog(extraction_params, process_multiple_function):
    """
    Launch a mode selection dialog and process files based on user choice.
    
    Args:
        extraction_params (list): Extraction parameters for document processing
        process_multiple_function (function): Function to process multiple files
        
    Returns:
        str: Path to the created JSON file or None if failed
    """
    # Initialize main Tkinter root
    root = tk.Tk()
    root.title("PDF Processor")
    root.withdraw()  # Hide the main window
    
    try:
        # Prompt for mode selection using a dialog
        mode_choice = messagebox.askquestion("PDF Processor Mode", 
                                            "Do you want to process multiple PDF files?\n\n"
                                            "Click 'Yes' for multiple files or 'No' for a single file.")
        
        result = None
        
        # Process based on selected mode
        if mode_choice == 'no':  # Single file
            result = launch_single_file_dialog(extraction_params)
        
        elif mode_choice == 'yes':  # Multiple files
            result = launch_multiple_files_dialog(extraction_params, process_multiple_function)
        
        return result
    
    finally:
        # Clean up and close Tkinter
        root.destroy()

def create_custom_dialog(title="PDF Processor", message="Select an option:"):
    """
    Create a custom dialog with buttons for different processing options.
    
    Args:
        title (str): Dialog window title
        message (str): Message to display in the dialog
        
    Returns:
        tk.Toplevel: Dialog window object
    """
    # Initialize main Tkinter root
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Create dialog window
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.geometry("400x200")
    dialog.resizable(False, False)
    
    # Center the dialog on screen
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # Add message label
    label = tk.Label(dialog, text=message, padx=20, pady=20)
    label.pack(fill=tk.BOTH, expand=True)
    
    return dialog, root

# Example of a custom dialog with specific options
def launch_custom_options_dialog(extraction_params, process_multiple_function, custom_options=None):
    """
    Launch a custom dialog with multiple options for PDF processing.
    
    Args:
        extraction_params (list): Extraction parameters for document processing
        process_multiple_function (function): Function to process multiple files
        custom_options (dict): Dictionary of option names and handler functions
        
    Returns:
        Any: Result of the selected option's handler function
    """
    if custom_options is None:
        custom_options = {
            "Process Single File": lambda: launch_single_file_dialog(extraction_params),
            "Process Multiple Files": lambda: launch_multiple_files_dialog(extraction_params, process_multiple_function),
            "Exit": lambda: None
        }
    
    dialog, root = create_custom_dialog(
        title="PDF Processor Options",
        message="Select a processing option:"
    )
    
    result = [None]  # Use a list to store the result for access in button callbacks
    
    # Create buttons for each option
    for i, (option_name, handler_func) in enumerate(custom_options.items()):
        def create_handler(func):
            def handler():
                dialog.destroy()
                root.destroy()
                result[0] = func()
            return handler
        
        button = tk.Button(dialog, text=option_name, width=20, command=create_handler(handler_func))
        button.pack(pady=5)
    
    # Wait for the dialog to be closed before returning
    root.wait_window(dialog)
    
    return result[0]