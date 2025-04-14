import tkinter as tk
import argparse
import sys
from app_gui import JsonToExcelApp

def main():
    """Main entry point for the application."""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='JSON to Excel Converter')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    
    # Check if running in a frozen environment (e.g., PyInstaller)
    if getattr(sys, 'frozen', False):
        # Running as executable, parse args carefully
        args, _ = parser.parse_known_args()
    else:
        # Running as script, parse args normally
        args = parser.parse_args()
    
    # Configure debug mode if requested
    debug_mode = args.debug
    
    if debug_mode:
        print("Debug mode enabled - detailed logs will be shown in the application")
    
    # Start the application
    root = tk.Tk()
    app = JsonToExcelApp(root, debug_mode=debug_mode)
    root.mainloop()

if __name__ == "__main__":
    main()