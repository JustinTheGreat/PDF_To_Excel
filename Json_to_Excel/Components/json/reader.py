import os
import json
import traceback

class JsonReader:
    """
    Class for reading and loading JSON files from directories.
    """
    
    @staticmethod
    def read_json_files(directory_path, recursive=True, print_debug=True):
        """
        Read all JSON files in the directory and its subdirectories and return their data.
        
        Args:
            directory_path: Path to the root directory
            recursive: Whether to search in subdirectories as well (default: True)
            print_debug: Whether to print debug information to console (default: True)
            
        Returns:
            Dictionary mapping file paths to their JSON content
        """
        json_data = {}
        error_files = []
        processed_files = 0
        
        def debug_print(message):
            if print_debug:
                print(message)
        
        def process_directory(dir_path, relative_path=""):
            """Process a directory and its subdirectories recursively."""
            nonlocal processed_files
            
            # Get all items in the directory
            try:
                items = os.listdir(dir_path)
                debug_print(f"Scanning directory: {dir_path} - Found {len(items)} items")
                
                # Create list of JSON files to process (to ensure deterministic ordering)
                json_files = []
                for item in items:
                    item_path = os.path.join(dir_path, item)
                    rel_item_path = os.path.join(relative_path, item) if relative_path else item
                    
                    # If it's a directory and recursive is enabled, process it
                    if os.path.isdir(item_path) and recursive:
                        debug_print(f"Entering subdirectory: {item_path}")
                        process_directory(item_path, rel_item_path)
                    
                    # If it's a JSON file, add to list
                    elif item.lower().endswith('.json'):
                        json_files.append((item_path, rel_item_path))
                
                # Process each JSON file
                for item_path, rel_item_path in json_files:
                    processed_files += 1
                    try:
                        debug_print(f"Reading JSON file {processed_files}: {item_path}")
                        with open(item_path, 'r', encoding='utf-8') as file:
                            file_content = file.read()
                            try:
                                file_data = json.loads(file_content)
                                json_data[rel_item_path] = file_data
                                
                                # Print some info about the data
                                if isinstance(file_data, dict):
                                    debug_print(f"  - Successfully loaded as dictionary with {len(file_data)} keys")
                                    some_keys = list(file_data.keys())[:5]
                                    debug_print(f"  - Sample keys: {some_keys}")
                                elif isinstance(file_data, list):
                                    debug_print(f"  - Successfully loaded as list with {len(file_data)} items")
                                    if file_data and isinstance(file_data[0], dict):
                                        some_keys = list(file_data[0].keys())[:5]
                                        debug_print(f"  - First item sample keys: {some_keys}")
                                else:
                                    debug_print(f"  - Successfully loaded as {type(file_data).__name__}")
                            except json.JSONDecodeError as json_err:
                                error_msg = f"JSON decode error in {rel_item_path}: {str(json_err)}"
                                debug_print(error_msg)
                                error_files.append((rel_item_path, error_msg))
                    except Exception as e:
                        error_msg = f"Error reading {rel_item_path}: {str(e)}"
                        debug_print(error_msg)
                        error_files.append((rel_item_path, error_msg))
                        debug_print(traceback.format_exc())
            
            except Exception as e:
                error_msg = f"Error accessing directory {dir_path}: {str(e)}"
                debug_print(error_msg)
                debug_print(traceback.format_exc())
        
        # Start processing from the root directory
        process_directory(directory_path)
        
        # Print summary
        debug_print(f"\nJSON Processing Summary:")
        debug_print(f"Total files processed: {processed_files}")
        debug_print(f"Successfully loaded files: {len(json_data)}")
        debug_print(f"Files with errors: {len(error_files)}")
        
        if error_files:
            debug_print("\nFiles with errors:")
            for file_path, error in error_files:
                debug_print(f"- {file_path}: {error}")
        
        return json_data