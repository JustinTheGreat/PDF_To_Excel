import os
import json
import traceback

class JsonProcessor:
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
    
    @staticmethod
    def analyze_json_structure(json_data, print_debug=True):
        """
        Analyze the structure of the JSON data to determine how to format the Excel sheet.
        
        Args:
            json_data: JSON data to analyze
            print_debug: Whether to print debug information
            
        Returns a dict with information about:
        - All unique keys
        - Maximum nesting depth for each key
        - Whether subtitles are needed
        """
        def debug_print(message):
            if print_debug:
                print(message)
        
        structure_info = {
            'keys': set(),
            'nesting_depth': {},
            'needs_subtitles': False
        }
        
        # Debug the input
        if isinstance(json_data, list):
            debug_print(f"analyze_json_structure: Input is a list with {len(json_data)} items")
        else:
            debug_print(f"analyze_json_structure: Input is a {type(json_data).__name__}")
        
        # Ensure we're working with a list of objects
        data_list = json_data if isinstance(json_data, list) else [json_data]
        
        for i, report in enumerate(data_list):
            debug_print(f"Analyzing item {i+1} of {len(data_list)}")
            
            # Handle different JSON structures
            fields = {}
            if isinstance(report, dict):
                # If report has a 'fields' key, use that, otherwise treat the whole report as fields
                if 'fields' in report:
                    debug_print(f"  - Found 'fields' key with {len(report['fields'])} fields")
                    fields = report.get('fields', {})
                else:
                    debug_print(f"  - No 'fields' key found, treating entire object as fields with {len(report)} keys")
                    fields = report
            else:
                debug_print(f"  - Item is not a dictionary, it's a {type(report).__name__}")
                continue
            
            # Process each field
            for key, value in fields.items():
                structure_info['keys'].add(key)
                
                # Check for lists that need subtitles
                if isinstance(value, list) and len(value) > 1:
                    structure_info['nesting_depth'][key] = len(value)
                    structure_info['needs_subtitles'] = True
                    debug_print(f"  - Field '{key}' has a list with {len(value)} items (needs subtitles)")
                elif key not in structure_info['nesting_depth']:
                    structure_info['nesting_depth'][key] = 0
                    debug_print(f"  - Field '{key}' has type {type(value).__name__}")
        
        debug_print(f"Analysis result: {len(structure_info['keys'])} unique keys, needs_subtitles={structure_info['needs_subtitles']}")
        return structure_info
    
    @staticmethod
    def process_filename(filename, filter_text=""):
        """Process filename to remove extension and filter text."""
        # Remove extension
        display_filename = os.path.splitext(filename)[0]
        
        # Remove filter text if provided
        if filter_text and filter_text in display_filename:
            display_filename = display_filename.replace(filter_text, "").strip()
        
        return display_filename