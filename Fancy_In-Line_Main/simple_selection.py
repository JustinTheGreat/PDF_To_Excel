import os
import glob
from Components.pdf_processor import create_document_json

def process_single_file_terminal(extraction_params):
    """
    Simple terminal-based interface to process a single PDF file.
    Uses the current directory as the default location with options to navigate directories.
    
    Args:
        extraction_params (list): List of extraction parameter dictionaries
        
    Returns:
        str: Path to the created JSON file or None if failed
    """
    # Start with the current working directory
    current_dir = os.getcwd()
    
    print("\n===== PDF Processor =====")
    print("Special commands:")
    print("  - '..' : Go to parent directory")
    print("  - 'ls' : List PDF files in current directory")
    print("  - 'cd [path]' : Change to directory")
    print("  - 'q' : Quit")
    
    while True:
        print(f"\nCurrent directory: {current_dir}")
        pdf_path = input("PDF file path or command: ").strip()
        
        # Process special commands
        if pdf_path.lower() in ['q', 'quit', 'exit']:
            print("Operation cancelled.")
            return None
        
        elif pdf_path == '..':
            # Navigate to parent directory
            parent_dir = os.path.dirname(current_dir)
            if parent_dir == current_dir:  # Root directory check
                print("Already at root directory.")
            else:
                current_dir = parent_dir
                print(f"Changed to parent directory: {current_dir}")
            continue
        
        elif pdf_path.lower() == 'ls':
            # List PDF files in current directory
            pdf_files = glob.glob(os.path.join(current_dir, "*.pdf"))
            if pdf_files:
                print("\nPDF files in current directory:")
                for i, file_path in enumerate(pdf_files, 1):
                    file_name = os.path.basename(file_path)
                    print(f"  {i}. {file_name}")
            else:
                print("No PDF files found in current directory.")
            continue
        
        elif pdf_path.lower().startswith('cd '):
            # Change directory
            target_dir = pdf_path[3:].strip()
            
            # Handle relative paths
            if not os.path.isabs(target_dir):
                target_dir = os.path.join(current_dir, target_dir)
                
            # Check if directory exists
            if os.path.isdir(target_dir):
                current_dir = os.path.abspath(target_dir)
                print(f"Changed to directory: {current_dir}")
            else:
                print(f"Error: Directory '{target_dir}' not found.")
            continue
        
        # Handle numeric selection from 'ls' command
        elif pdf_path.isdigit():
            pdf_files = glob.glob(os.path.join(current_dir, "*.pdf"))
            try:
                index = int(pdf_path) - 1
                if 0 <= index < len(pdf_files):
                    pdf_path = pdf_files[index]
                else:
                    print(f"Error: Invalid selection. Choose between 1 and {len(pdf_files)}.")
                    continue
            except (ValueError, IndexError):
                print("Invalid selection.")
                continue
        
        # Process the file path
        if not os.path.isabs(pdf_path):
            pdf_path = os.path.join(current_dir, pdf_path)
        
        # Check if file exists
        if not os.path.isfile(pdf_path):
            print(f"Error: File '{pdf_path}' not found. Please try again.")
            continue
        
        # Check file extension
        if not pdf_path.lower().endswith('.pdf'):
            confirm = input("Warning: This doesn't appear to be a PDF file. Process anyway? (y/n): ").strip().lower()
            if confirm != 'y':
                continue
        
        # File exists, process it
        print(f"\nProcessing: {pdf_path}")
        try:
            json_path = create_document_json(pdf_path, extraction_params)
            
            if json_path:
                print(f"\nSuccess! JSON file created: {json_path}")
                return json_path
            else:
                print("\nError: Failed to create JSON file.")
                retry = input("Would you like to try another file? (y/n): ").strip().lower()
                if retry != 'y':
                    return None
        except Exception as e:
            print(f"\nError processing file: {str(e)}")
            retry = input("Would you like to try another file? (y/n): ").strip().lower()
            if retry != 'y':
                return None