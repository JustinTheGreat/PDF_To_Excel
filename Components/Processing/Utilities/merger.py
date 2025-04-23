"""
Field merging utilities for PDF processing with automatic unit preservation.

This module provides functions for merging related fields 
in data extracted from PDF documents, with intelligent detection
for unit values that should be preserved separately.
"""
import os
import re
# os.chdir('Components/Processing/Utilities')
from Components.Processing.Utilities.cleaner import clean_empty_keys
from Components.config import debug_print

def process_field_merging(extracted_data):
    """
    Process and merge fields that have a (+1) suffix in their names.
    Ensures proper handling of duplicate keys during merging.
    Automatically detects and preserves unit measurements.
    
    Args:
        extracted_data (dict): Dictionary containing extracted field data
        
    Returns:
        dict: Processed data with merged fields
    """
    merged_data = {}
    merge_candidates = {}
    
    debug_print(f"[MERGE] Starting field merging process")
    debug_print(f"[MERGE] Found {len(extracted_data)} total fields")
    
    # First pass: identify fields that need merging
    for field_name in extracted_data.keys():
        if "(+1)" in field_name:
            # Extract the base field name (without the +1)
            base_field_name = field_name.replace("(+1)", "").strip()
            
            debug_print(f"[MERGE] Found extension field: '{field_name}' -> base: '{base_field_name}'")
            
            # Check if we have the base field
            if base_field_name in extracted_data:
                # Add to merge candidates
                if base_field_name not in merge_candidates:
                    merge_candidates[base_field_name] = []
                merge_candidates[base_field_name].append(field_name)
                debug_print(f"[MERGE] Added '{field_name}' to merge candidates for '{base_field_name}'")
            else:
                # If base field doesn't exist, keep as is
                merged_data[field_name] = extracted_data[field_name]
                debug_print(f"[MERGE] Base field '{base_field_name}' not found, keeping extension field as-is")
        else:
            # Regular field, add to merged data
            merged_data[field_name] = extracted_data[field_name]
            debug_print(f"[MERGE] Added regular field: '{field_name}'")
    
    debug_print(f"[MERGE] Found {len(merge_candidates)} base fields with extensions to merge")
    
    # List of fields that should preserve duplicate values
    preserve_duplicates_fields = [
        "E-Storage_EOL_Report_Stats",
        "E-Storage_EOL_Report_General_Info"
        # Add more fields here as needed
    ]
    
    # Second pass: perform the merging
    for base_field, extension_fields in merge_candidates.items():
        debug_print(f"\n[MERGE] Merging extensions for base field: '{base_field}'")
        debug_print(f"[MERGE] Extensions to merge: {extension_fields}")
        
        # Check if this field should preserve duplicates
        should_preserve_duplicates = any(base_field.startswith(field) for field in preserve_duplicates_fields)
        if should_preserve_duplicates:
            debug_print(f"[MERGE] Field '{base_field}' is configured to preserve duplicate values")
        
        # Start with the base field data
        base_data = extracted_data[base_field]
        merged_raw_text = base_data["raw_text"]
        merged_formatted_text = base_data["formatted_text"]
        merged_parsed_data = base_data["parsed_data"].copy()
        
        # Check for unit keys in the base field
        unit_pattern = r'\[([^\]]+)\]'
        base_unit_keys = [k for k in merged_parsed_data.keys() if re.search(unit_pattern, k)]
        if base_unit_keys:
            debug_print(f"[MERGE] Base field has unit keys: {base_unit_keys}")
            
        debug_print(f"[MERGE] Base field '{base_field}' parsed data keys: {list(merged_parsed_data.keys())}")
        
        # Track all page numbers for extensions
        ext_page_numbers = {}
        for i, ext_field in enumerate(extension_fields):
            # Try to extract page number from parameters
            ext_data = extracted_data[ext_field]
            ext_page_numbers[ext_field] = i+1  # Default numbering if we can't extract page
        
        # Merge each extension field
        for ext_field in extension_fields:
            debug_print(f"\n[MERGE] Processing extension field: '{ext_field}'")
            ext_data = extracted_data[ext_field]
            
            # Get the extension page number 
            page_id = ext_page_numbers[ext_field]
            
            # Concatenate raw and formatted text with a separator
            merged_raw_text += "\n\n--- Additional Data ---\n\n" + ext_data["raw_text"]
            merged_formatted_text += "\n\n--- Additional Data ---\n\n" + ext_data["formatted_text"]
            
            debug_print(f"[MERGE] Extension field '{ext_field}' parsed data keys: {list(ext_data['parsed_data'].keys())}")
            
            # Check for unit keys in this extension
            ext_unit_keys = [k for k in ext_data["parsed_data"].keys() if re.search(unit_pattern, k)]
            if ext_unit_keys:
                debug_print(f"[MERGE] Found unit keys in extension: {ext_unit_keys}")
            
            # Merge the parsed data dictionaries
            for key, value in ext_data["parsed_data"].items():
                # Skip empty values
                if value == "" or value is None:
                    debug_print(f"[MERGE] Skipping empty value for key: '{key}'")
                    continue
                    
                if isinstance(value, list) and all(v == "" or v is None for v in value):
                    debug_print(f"[MERGE] Skipping list with all empty values for key: '{key}'")
                    continue
                
                # Check if this is a unit key (contains [...])
                is_unit_key = bool(re.search(unit_pattern, key))
                
                # Additional logging for keys that might be units
                if is_unit_key:
                    debug_print(f"[MERGE] Processing unit key: '{key}' with value: {value}")
                
                # If key exists in merged data
                if key in merged_parsed_data:
                    base_value = merged_parsed_data[key]
                    debug_print(f"[MERGE] Key '{key}' exists in base with value: {base_value}")
                    
                    # Special handling for unit keys or fields configured to preserve duplicates
                    if is_unit_key or should_preserve_duplicates:
                        # Don't merge unit keys or preserve duplicates field values - always keep them separate
                        debug_print(f"[MERGE] Preserving value as separate entry for key: {key}")
                        
                        # Already have this value? Convert to a list if needed
                        if isinstance(base_value, list):
                            # It's already a list, so just add the new value even if it's a duplicate
                            debug_print(f"[MERGE] Adding {value} to existing list")
                            if isinstance(value, list):
                                merged_parsed_data[key].extend(value)
                            else:
                                merged_parsed_data[key].append(value)
                        else:
                            # Convert single value to list and add new value
                            debug_print(f"[MERGE] Converting key {key} to list with values: [{base_value}, {value}]")
                            if isinstance(value, list):
                                merged_parsed_data[key] = [base_value] + value
                            else:
                                merged_parsed_data[key] = [base_value, value]
                    
                    # Normal merging for non-unit keys that don't need to preserve duplicates
                    else:
                        # Both values are lists
                        if isinstance(base_value, list) and isinstance(value, list):
                            debug_print(f"[MERGE] Both base and extension values are lists")
                            # Only add values that don't already exist
                            for item in value:
                                if item not in base_value and item != "" and item is not None:
                                    merged_parsed_data[key].append(item)
                                    debug_print(f"[MERGE] Added new item: {item} to list")
                                else:
                                    debug_print(f"[MERGE] Skipping duplicate item: {item}")
                                
                        # Base value is a list, new value is not
                        elif isinstance(base_value, list):
                            debug_print(f"[MERGE] Base value is a list, extension value is not")
                            if value != "" and value is not None and value not in base_value:
                                debug_print(f"[MERGE] Adding single value to existing list")
                                merged_parsed_data[key].append(value)
                            else:
                                debug_print(f"[MERGE] Skipping duplicate value: {value}")
                                
                        # New value is a list, base value is not
                        elif isinstance(value, list):
                            debug_print(f"[MERGE] Base value is not a list, extension value is")
                            non_empty_values = [v for v in value if v != "" and v is not None and v != base_value]
                            if non_empty_values:
                                debug_print(f"[MERGE] Converting base value to list and adding extension values")
                                merged_parsed_data[key] = [base_value] + non_empty_values
                        
                        # Neither value is a list, but they're different
                        elif base_value != value and value != "" and value is not None:
                            debug_print(f"[MERGE] Neither value is a list and they differ")
                            merged_parsed_data[key] = [base_value, value]
                        
                        # Values are identical or new value is empty
                        else:
                            debug_print(f"[MERGE] Values are identical or new value is empty, keeping as is")
                else:
                    # Key doesn't exist in base, simply add it if it's not empty
                    debug_print(f"[MERGE] Key '{key}' does not exist in base data")
                    if isinstance(value, list):
                        non_empty_values = [v for v in value if v != "" and v is not None]
                        if non_empty_values:
                            debug_print(f"[MERGE] Adding list with {len(non_empty_values)} non-empty values")
                            merged_parsed_data[key] = non_empty_values if len(non_empty_values) > 1 else non_empty_values[0]
                    else:
                        debug_print(f"[MERGE] Adding single value: {value}")
                        merged_parsed_data[key] = value
        
        # Clean out any empty values from the merged data
        debug_print(f"[MERGE] Cleaning empty values from merged data")
        debug_print(f"[MERGE] BEFORE cleaning: {list(merged_parsed_data.keys())}")
        merged_parsed_data = clean_empty_keys(merged_parsed_data)
        debug_print(f"[MERGE] AFTER cleaning: {list(merged_parsed_data.keys())}")
        
        # Update the base field with merged data
        merged_data[base_field] = {
            "raw_text": merged_raw_text,
            "formatted_text": merged_formatted_text,
            "parsed_data": merged_parsed_data
        }
        debug_print(f"[MERGE] Updated base field '{base_field}' with merged data")
        
        # Final check for unit keys in the merged data
        merged_unit_keys = [k for k in merged_parsed_data.keys() if re.search(unit_pattern, k)]
        if merged_unit_keys:
            debug_print(f"[MERGE] Final unit keys in merged data: {merged_unit_keys}")
            for uk in merged_unit_keys:
                debug_print(f"[MERGE] Unit key '{uk}' final value: {merged_parsed_data[uk]}")
    
    debug_print(f"[MERGE] Field merging complete, processed {len(merged_data)} fields")
    return merged_data