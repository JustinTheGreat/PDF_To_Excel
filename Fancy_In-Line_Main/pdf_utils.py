import os
import json
import re
from Components.pdf_processor import create_document_json

def find_common_name(filenames):
    """
    Find common substring across multiple filenames.
    
    Args:
        filenames (list): List of filenames to compare
        
    Returns:
        str: Common substring or default name if no common part found
    """
    if not filenames:
        return "combined_output"
    
    # Remove file extensions and directory paths
    base_names = [os.path.splitext(os.path.basename(f))[0] for f in filenames]
    
    # Find the common part using regex
    common_chars = []
    for chars in zip(*base_names):
        if all(c == chars[0] for c in chars):
            common_chars.append(chars[0])
        else:
            break
    
    common_part = ''.join(common_chars).strip('_- ')
    
    # If common part is too short, try matching alphanumeric patterns
    if len(common_part) < 3:
        # Extract alphanumeric patterns that appear in all filenames
        patterns = []
        for name in base_names:
            # Find all alphanumeric sequences
            alphanumeric = re.findall(r'[A-Za-z0-9]+', name)
            patterns.append(set(alphanumeric))
        
        # Find common patterns across all files
        if patterns:
            common_patterns = set.intersection(*patterns)
            if common_patterns:
                # Use the longest common pattern
                common_part = max(common_patterns, key=len)
    
    # If still no meaningful common part, use default name
    if not common_part or len(common_part) < 2:
        return "combined_output"
    
    return common_part

def merge_json_data(json_data_list):
    """
    Merge multiple JSON data lists into a single list.
    
    Args:
        json_data_list (list): List of JSON data lists to merge
        
    Returns:
        list: Merged JSON data
    """
    # Dictionary to hold merged data, keyed by field_name
    merged_data = {}
    
    # Process each file's JSON data
    for file_data in json_data_list:
        for entry in file_data:
            field_name = entry["title"]
            
            # If this field already exists, merge the content
            if field_name in merged_data:
                # Append raw text with separator
                merged_data[field_name]["raw_text"] += "\n\n--- From Another File ---\n\n" + entry["raw_text"]
                
                # Append formatted text with separator
                merged_data[field_name]["formatted_text"] += "\n\n--- From Another File ---\n\n" + entry["formatted_text"]
                
                # Merge fields dictionaries
                for key, value in entry["fields"].items():
                    if key in merged_data[field_name]["fields"]:
                        # Check if both are lists
                        if isinstance(merged_data[field_name]["fields"][key], list) and isinstance(value, list):
                            merged_data[field_name]["fields"][key].extend(value)
                        # Check if existing is list but new is not
                        elif isinstance(merged_data[field_name]["fields"][key], list):
                            merged_data[field_name]["fields"][key].append(value)
                        # Check if new is list but existing is not
                        elif isinstance(value, list):
                            merged_data[field_name]["fields"][key] = [merged_data[field_name]["fields"][key]] + value
                        # Both are single values and different
                        elif merged_data[field_name]["fields"][key] != value:
                            merged_data[field_name]["fields"][key] = [merged_data[field_name]["fields"][key], value]
                    else:
                        # Key doesn't exist in merged data, add it
                        merged_data[field_name]["fields"][key] = value
            else:
                # Field doesn't exist yet, add it
                merged_data[field_name] = entry
    
    # Convert merged_data dictionary back to a list
    return list(merged_data.values())

def process_multiple_files(file_paths, extraction_params):
    """
    Process multiple PDF files and create a combined JSON file.
    
    Args:
        file_paths (list): List of PDF file paths to process
        extraction_params (list): Extraction parameters
        
    Returns:
        str: Path to the created combined JSON file
    """
    all_json_data = []
    processed_files = []
    
    for pdf_path in file_paths:
        # Skip if file doesn't exist
        if not os.path.isfile(pdf_path):
            print(f"Warning: File '{pdf_path}' not found. Skipping.")
            continue
        
        print(f"Processing: {pdf_path}")
        
        # Create JSON for this file
        json_path = create_document_json(pdf_path, extraction_params)
        
        if json_path:
            # Read the JSON data
            with open(json_path, 'r', encoding='utf-8') as file:
                content = json.load(file)
                all_json_data.append(content)
            processed_files.append(pdf_path)
        else:
            print(f"Failed to process '{pdf_path}'. Skipping.")
    
    if not all_json_data:
        print("No files were successfully processed.")
        return None
    
    # Merge all JSON data
    merged_data = merge_json_data(all_json_data)
    
    # Find common name across files
    common_name = find_common_name(processed_files)
    
    # Create output path in the same directory as the first file
    output_dir = os.path.dirname(processed_files[0])
    output_path = os.path.join(output_dir, f"{common_name}_combined.json")
    
    # Save merged JSON data
    with open(output_path, 'w', encoding='utf-8') as json_file:
        json.dump(merged_data, json_file, indent=2)
    
    print(f"Combined JSON file created: {output_path}")
    return output_path