"""
Core extraction functionality for PDF processing.

This module contains the main logic for extracting data from PDFs
based on extraction parameters and coordinating the use of parsers
and utilities.
"""
import os
from Components.pdf_extractor import (
    parse_text_to_key_value, 
    format_raw_text, 
    apply_special_formatting
)
from Components.GeneralInfo import extract_serial_data
from Components.Processing.Utilities.cleaner import clean_empty_keys
from Components.Processing.Parsers.keywords import handle_same_start_end_keyword, handle_different_start_end_keyword
from Components.Processing.Parsers.table import process_table_data
from Components.config import debug_print

def extract_pdf_data(pdf_path, extraction_params):
    """
    Extract various data from PDF based on a list of extraction parameters.
    Parse the extracted text into key-value pairs.
    
    Args:
        pdf_path (str): Path to the PDF file
        extraction_params (list): List of parameter dictionaries for extraction
        
    Returns:
        dict: Dictionary with all extracted data
    """
    extracted_data = {}
    
    debug_print(f"\n[DEBUG] extract_pdf_data called with {len(extraction_params)} parameter sets")
    
    for i, param_set in enumerate(extraction_params):
        debug_print(f"\n[DEBUG] Processing parameter set {i+1}")
        
        # Extract the parameters from the dictionary
        field_name = param_set.get('field_name', 'N/A')
        start_keyword = param_set.get('start_keyword', 'N/A')
        start_keyword_occurrence = param_set.get('start_keyword_occurrence', 1)  # Default to first occurrence
        end_keyword = param_set.get('end_keyword', None)  # Default to None to allow line break mode only
        page_num = param_set.get('page_num', 0)
        horiz_margin = param_set.get('horiz_margin', 200)
        vertical_margin = param_set.get('vertical_margin', None)  # New parameter for vertical margin
        end_keyword_occurrence = param_set.get('end_keyword_occurrence', 1)
        left_move = param_set.get('left_move', 0)
        
        # Get the list of forced keywords (if any)
        forced_keywords = param_set.get('forced_keywords', None)
        
        # Get the list of words to remove line breaks before (if any)
        remove_breaks_before = param_set.get('remove_breaks_before', None)
        
        # Get the list of words to remove line breaks after (if any)
        remove_breaks_after = param_set.get('remove_breaks_after', None)
        
        # Get the list of keywords to remove colons after (if any)
        remove_colon_after = param_set.get('remove_colon_after', None)
        
        # Get the end break line count (if specified)
        end_break_line_count = param_set.get('end_break_line_count', None)
        
        # Debug the parameters
        debug_print(f"[DEBUG] Parameter details for '{field_name}':")
        debug_print(f"  start_keyword: '{start_keyword}'")
        debug_print(f"  start_keyword_occurrence: {start_keyword_occurrence}")
        debug_print(f"  end_keyword: '{end_keyword}'")
        debug_print(f"  page_num: {page_num}")
        debug_print(f"  horiz_margin: {horiz_margin}")
        debug_print(f"  vertical_margin: {vertical_margin}")
        debug_print(f"  end_keyword_occurrence: {end_keyword_occurrence}")
        debug_print(f"  left_move: {left_move}")
        debug_print(f"  end_break_line_count: {end_break_line_count}")
        debug_print(f"  forced_keywords: {forced_keywords}")
        debug_print(f"  remove_breaks_before: {remove_breaks_before}")
        debug_print(f"  remove_breaks_after: {remove_breaks_after}")
        debug_print(f"  remove_colon_after: {remove_colon_after}")
        
        # Extract the data using the specified parameters
        debug_print(f"[DEBUG] Calling extract_serial_data")
        original_raw_text = extract_serial_data(
            pdf_path,
            start_keyword=start_keyword,
            start_keyword_occurrence=start_keyword_occurrence,
            end_keyword=end_keyword,
            page_num=page_num,
            horiz_margin=horiz_margin,
            vertical_margin=vertical_margin,  # Pass the new parameter
            end_keyword_occurrence=end_keyword_occurrence,
            left_move=left_move,
            end_break_line_count=end_break_line_count
        )
        
        # Debug the raw text length
        debug_print(f"[DEBUG] Raw text length after extraction: {len(original_raw_text) if original_raw_text else 0}")
        
        # Format the raw text based on field name (but keep original for the raw_text field)
        debug_print(f"[DEBUG] Formatting raw text")
        formatted_raw_text = format_raw_text(
            field_name, 
            original_raw_text, 
            forced_keywords,
            remove_breaks_before,
            remove_breaks_after,
            remove_colon_after  # Pass the new parameter
        )
        
        # Debug the formatted text length
        debug_print(f"[DEBUG] Formatted text length: {len(formatted_raw_text) if formatted_raw_text else 0}")
        
        # Check if table processing is requested
        is_table_processing = (
            param_set.get('table_top_labeling', False) or 
            param_set.get('table_left_labeling', False)
        )
        
        if is_table_processing:
            # Process as table data
            debug_print(f"[DEBUG] Using table processing")
            table_params = {
                'table_top_labeling': param_set.get('table_top_labeling', False),
                'table_left_labeling': param_set.get('table_left_labeling', False),
                'table_labeling_priority': param_set.get('table_labeling_priority', 'top')
            }
            
            # Process the table data
            processed_result = process_table_data(formatted_raw_text, table_params)
            
            # Store the results
            extracted_data[field_name] = {
                "raw_text": original_raw_text,
                "formatted_text": formatted_raw_text,
                "parsed_data": processed_result
            }
        else:
            # Standard text processing
            debug_print(f"[DEBUG] Using standard text processing")
            parsed_result, unparsed_lines = parse_text_to_key_value(formatted_raw_text)
            
            # Debug the parsed result
            debug_print(f"[DEBUG] Parsed {len(parsed_result) if parsed_result else 0} key-value pairs")
            debug_print(f"[DEBUG] Unparsed lines: {len(unparsed_lines) if unparsed_lines else 0}")
            
            # Apply special formatting to handle unparsed lines and business-specific rules
            formatted_result = apply_special_formatting(field_name, parsed_result, unparsed_lines)
            
            # Process the results based on keyword configuration
            processed_result = process_parsed_result(
                formatted_result,
                start_keyword, 
                end_keyword,
                original_raw_text,
                end_keyword_occurrence,
                forced_keywords
            )
            
            # Check if this field already exists (duplicate extraction parameters)
            if field_name in extracted_data:
                debug_print(f"[DEBUG] Field '{field_name}' already exists. Merging data.")
                # Append to raw text
                existing_raw_text = extracted_data[field_name]["raw_text"]
                new_raw_text = existing_raw_text + "\n\n--- Additional Data ---\n\n" + original_raw_text
                
                # Append to formatted text
                existing_formatted_text = extracted_data[field_name]["formatted_text"]
                new_formatted_text = existing_formatted_text + "\n\n--- Additional Data ---\n\n" + formatted_raw_text
                
                # Merge parsed data dictionaries
                existing_parsed_data = extracted_data[field_name]["parsed_data"]
                new_parsed_data = {}
                
                # First copy all existing data
                for key, value in existing_parsed_data.items():
                    new_parsed_data[key] = value
                
                # Then merge in new data, handling duplicates
                for key, value in processed_result.items():
                    if key in new_parsed_data:
                        existing_value = new_parsed_data[key]
                        if isinstance(existing_value, list):
                            if isinstance(value, list):
                                new_parsed_data[key].extend(value)
                            else:
                                new_parsed_data[key].append(value)
                        else:
                            if isinstance(value, list):
                                new_parsed_data[key] = [existing_value] + value
                            elif existing_value != value:
                                new_parsed_data[key] = [existing_value, value]
                    else:
                        new_parsed_data[key] = value
                
                # Update the extracted data with merged content
                extracted_data[field_name] = {
                    "raw_text": new_raw_text,
                    "formatted_text": new_formatted_text,
                    "parsed_data": new_parsed_data
                }
            else:
                # Store both the original raw text and the formatted result in the dictionary
                extracted_data[field_name] = {
                    "raw_text": original_raw_text,
                    "formatted_text": formatted_raw_text,
                    "parsed_data": processed_result
                }
            
            debug_print(f"[DEBUG] Completed processing for field '{field_name}'")
    
    return extracted_data

def process_parsed_result(parsed_result, start_keyword, end_keyword, original_raw_text, end_keyword_occurrence, forced_keywords=None):
    """
    Process the parsed result based on keyword configuration.
    
    Args:
        parsed_result (dict): Dictionary of parsed key-value pairs
        start_keyword (str): Start keyword for extraction
        end_keyword (str): End keyword for extraction (can be None)
        original_raw_text (str): Original raw text
        end_keyword_occurrence (int): Occurrence number of end keyword to use
        forced_keywords (list): List of keywords that should be treated as keys
        
    Returns:
        dict: Processed parsed result
    """
    debug_print(f"[DEBUG] process_parsed_result called")
    debug_print(f"[DEBUG] start_keyword: '{start_keyword}'")
    debug_print(f"[DEBUG] end_keyword: '{end_keyword}'")
    debug_print(f"[DEBUG] end_keyword_occurrence: {end_keyword_occurrence}")
    
    # Create a copy to avoid modifying the original
    processed_result = parsed_result.copy()
    
    # Handle forced keywords to ensure they weren't included in previous values
    if forced_keywords and isinstance(forced_keywords, list):
        debug_print(f"[DEBUG] Processing {len(forced_keywords)} forced keywords")
        # For each forced keyword, check if it's contained in any value
        for keyword in forced_keywords:
            for key, value in list(processed_result.items()):
                # Handle both single values and lists of values
                if isinstance(value, str) and keyword in value:
                    # Split the value at the keyword
                    parts = value.split(keyword, 1)
                    
                    # Update the original value to stop at the keyword
                    processed_result[key] = parts[0].strip()
                    
                    # If there's content after the keyword, create a new key-value pair
                    if len(parts) > 1 and parts[1].strip():
                        # Set the keyword as a new key, and the remaining part as its value
                        # Make sure to handle the case where there's a colon already
                        if ':' in parts[1][:5]:  # Check first few characters for a colon
                            rest_parts = parts[1].split(':', 1)
                            new_value = rest_parts[1].strip() if len(rest_parts) > 1 else ""
                            # Only add if the new value is not empty
                            if new_value:
                                if keyword in processed_result:
                                    if isinstance(processed_result[keyword], list):
                                        processed_result[keyword].append(new_value)
                                    else:
                                        processed_result[keyword] = [processed_result[keyword], new_value]
                                else:
                                    processed_result[keyword] = new_value
                        else:
                            new_value = parts[1].strip()
                            # Only add if the new value is not empty
                            if new_value:
                                if keyword in processed_result:
                                    if isinstance(processed_result[keyword], list):
                                        processed_result[keyword].append(new_value)
                                    else:
                                        processed_result[keyword] = [processed_result[keyword], new_value]
                                else:
                                    processed_result[keyword] = new_value
                elif isinstance(value, list):
                    # Process each item in the list
                    new_values = []
                    for item in value:
                        if keyword in item:
                            # Split the value at the keyword
                            parts = item.split(keyword, 1)
                            
                            # Add the part before the keyword to the new values list
                            new_values.append(parts[0].strip())
                            
                            # If there's content after the keyword, create a new key-value pair
                            if len(parts) > 1 and parts[1].strip():
                                # Handle the case where there's a colon already
                                if ':' in parts[1][:5]:
                                    rest_parts = parts[1].split(':', 1)
                                    new_value = rest_parts[1].strip() if len(rest_parts) > 1 else ""
                                    # Only add if the new value is not empty
                                    if new_value:
                                        if keyword in processed_result:
                                            # Add to existing list or convert to list
                                            if isinstance(processed_result[keyword], list):
                                                processed_result[keyword].append(new_value)
                                            else:
                                                processed_result[keyword] = [processed_result[keyword], new_value]
                                        else:
                                            processed_result[keyword] = new_value
                                else:
                                    new_value = parts[1].strip()
                                    # Only add if the new value is not empty
                                    if new_value:
                                        if keyword in processed_result:
                                            # Add to existing list or convert to list
                                            if isinstance(processed_result[keyword], list):
                                                processed_result[keyword].append(new_value)
                                            else:
                                                processed_result[keyword] = [processed_result[keyword], new_value]
                                        else:
                                            processed_result[keyword] = new_value
                        else:
                            # Keep unchanged values
                            new_values.append(item)
                    
                    # Update with the filtered values
                    if new_values:
                        processed_result[key] = new_values
                    else:
                        # If all values were processed, remove the original key
                        del processed_result[key]
    
    # Special handling for when start_keyword == end_keyword
    if end_keyword and start_keyword == end_keyword:
        debug_print(f"[DEBUG] Using same_start_end_keyword handler")
        result = handle_same_start_end_keyword(processed_result, start_keyword)
    # Handle normal case with different start and end keywords
    elif end_keyword:
        debug_print(f"[DEBUG] Using different_start_end_keyword handler")
        result = handle_different_start_end_keyword(
            processed_result, 
            end_keyword, 
            original_raw_text, 
            end_keyword_occurrence
        )
    else:
        # When no end_keyword is provided (using only line breaks), 
        # don't apply any special keyword processing
        debug_print(f"[DEBUG] No end_keyword provided, skipping keyword processing")
        result = processed_result
    
    # Clean out any empty values
    final_result = clean_empty_keys(result)
    debug_print(f"[DEBUG] Keys after processing: {list(final_result.keys())}")
    
    return final_result