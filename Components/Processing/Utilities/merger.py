"""
Field merging utilities for PDF processing.

This module provides functions for merging related fields 
in data extracted from PDF documents.
"""
import os
# os.chdir('Components/Processing/Utilities')
from Components.Processing.Utilities.cleaner import clean_empty_keys

def process_field_merging(extracted_data):
    """
    Process and merge fields that have a (+1) suffix in their names.
    Ensures proper handling of duplicate keys during merging.
    
    Args:
        extracted_data (dict): Dictionary containing extracted field data
        
    Returns:
        dict: Processed data with merged fields
    """
    merged_data = {}
    merge_candidates = {}
    
    # First pass: identify fields that need merging
    for field_name in extracted_data.keys():
        if "(+1)" in field_name:
            # Extract the base field name (without the +1)
            base_field_name = field_name.replace("(+1)", "").strip()
            
            # Check if we have the base field
            if base_field_name in extracted_data:
                # Add to merge candidates
                if base_field_name not in merge_candidates:
                    merge_candidates[base_field_name] = []
                merge_candidates[base_field_name].append(field_name)
            else:
                # If base field doesn't exist, keep as is
                merged_data[field_name] = extracted_data[field_name]
        else:
            # Regular field, add to merged data
            merged_data[field_name] = extracted_data[field_name]
    
    # Second pass: perform the merging
    for base_field, extension_fields in merge_candidates.items():
        # Start with the base field data
        base_data = extracted_data[base_field]
        merged_raw_text = base_data["raw_text"]
        merged_formatted_text = base_data["formatted_text"]
        merged_parsed_data = base_data["parsed_data"].copy()
        
        # Merge each extension field
        for ext_field in extension_fields:
            ext_data = extracted_data[ext_field]
            
            # Concatenate raw and formatted text with a separator
            merged_raw_text += "\n\n--- Additional Data ---\n\n" + ext_data["raw_text"]
            merged_formatted_text += "\n\n--- Additional Data ---\n\n" + ext_data["formatted_text"]
            
            # Merge the parsed data dictionaries
            for key, value in ext_data["parsed_data"].items():
                # Skip empty values
                if value == "" or value is None:
                    continue
                    
                if isinstance(value, list) and all(v == "" or v is None for v in value):
                    continue
                
                # If key exists in merged data
                if key in merged_parsed_data:
                    base_value = merged_parsed_data[key]
                    
                    # Both values are lists
                    if isinstance(base_value, list) and isinstance(value, list):
                        merged_parsed_data[key].extend([v for v in value if v != "" and v is not None])
                    # Base value is a list, new value is not
                    elif isinstance(base_value, list):
                        if value != "" and value is not None:
                            merged_parsed_data[key].append(value)
                    # New value is a list, base value is not
                    elif isinstance(value, list):
                        non_empty_values = [v for v in value if v != "" and v is not None]
                        if non_empty_values:
                            merged_parsed_data[key] = [base_value] + non_empty_values
                    # Neither value is a list, but they're different
                    elif base_value != value and value != "" and value is not None:
                        merged_parsed_data[key] = [base_value, value]
                    # Values are identical or new value is empty, keep as is
                else:
                    # Key doesn't exist in base, simply add it if it's not empty
                    if isinstance(value, list):
                        non_empty_values = [v for v in value if v != "" and v is not None]
                        if non_empty_values:
                            merged_parsed_data[key] = non_empty_values if len(non_empty_values) > 1 else non_empty_values[0]
                    else:
                        merged_parsed_data[key] = value
        
        # Clean out any empty values from the merged data
        merged_parsed_data = clean_empty_keys(merged_parsed_data)
        
        # Update the base field with merged data
        merged_data[base_field] = {
            "raw_text": merged_raw_text,
            "formatted_text": merged_formatted_text,
            "parsed_data": merged_parsed_data
        }
    
    return merged_data