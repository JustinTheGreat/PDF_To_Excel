"""
Configuration module for PDF Extraction system.

This module contains global configuration settings that can be 
imported and used throughout the application.
"""
import re
import os
import datetime

# Debug configuration
DEBUG_MODE = False  # Default value, can be changed at runtime
DEBUG_FILE = None   # File handle for debug log

def set_debug_mode(enabled, log_dir="logs"):
    """
    Enable or disable debug mode globally.
    
    Args:
        enabled (bool): True to enable debugging, False to disable
        log_dir (str): Directory to store debug log files (default: logs)
    """
    global DEBUG_MODE, DEBUG_FILE
    
    # Close any existing debug file
    if DEBUG_FILE is not None:
        DEBUG_FILE.close()
        DEBUG_FILE = None
    
    DEBUG_MODE = enabled
    
    if enabled:
        # Create logs directory if it doesn't exist
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        # Create timestamped log filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file_path = os.path.join(log_dir, f"debug_{timestamp}.log")
        
        # Open log file for writing
        try:
            DEBUG_FILE = open(log_file_path, 'w', encoding='utf-8')
            print(f"Debug mode enabled. Logging to: {log_file_path}")
            debug_print("=== Debug Session Started ===")
            debug_print(f"Timestamp: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        except Exception as e:
            print(f"Warning: Could not create debug log file: {str(e)}")
            print("Debug messages will be printed to console instead.")
            DEBUG_FILE = None
            DEBUG_MODE = True
    else:
        print("Debug mode disabled.")

def debug_print(*args, **kwargs):
    """
    Write debug messages to log file or print to console if DEBUG_MODE is enabled.
    
    Args:
        *args: Arguments to log/print
        **kwargs: Keyword arguments for print function
    """
    if not DEBUG_MODE:
        return
        
    # Format the message
    try:
        message = " ".join(str(arg) for arg in args)
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        formatted_message = f"[{timestamp}] {message}"
        
        # Write to file if available, otherwise print to console
        if DEBUG_FILE is not None:
            try:
                DEBUG_FILE.write(formatted_message + "\n")
                DEBUG_FILE.flush()  # Ensure message is written immediately
            except Exception as e:
                print(f"Error writing to debug file: {str(e)}")
                print(formatted_message)
        else:
            print(formatted_message)
    except Exception as e:
        print(f"Error in debug_print: {str(e)}")