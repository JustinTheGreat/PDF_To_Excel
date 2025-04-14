"""
Keyword parsing functionality for PDF extraction.

This module provides functions for handling special keyword logic
in PDF text extraction, particularly for managing start and end keywords.
"""

def handle_same_start_end_keyword(parsed_result, keyword):
    """
    Handle case where start and end keywords are the same.
    
    Args:
        parsed_result (dict): Dictionary of parsed key-value pairs
        keyword (str): The keyword
        
    Returns:
        dict: Processed parsed result
    """
    # When start and end keywords are the same, we'll keep all entries
    # Just ensure we have at least one key with the keyword
    if not keyword:
        return parsed_result
        
    has_keyword = False
    for key in parsed_result.keys():
        if keyword in key:
            has_keyword = True
            break
    
    # If no keys contain the keyword, add one
    if not has_keyword and parsed_result:
        # Try to find a key that might contain part of the keyword
        matching_key = None
        for key in parsed_result.keys():
            if any(word in key.lower() for word in keyword.lower().split()):
                matching_key = key
                break
        
        # If found a partial match, rename it
        if matching_key:
            parsed_result[keyword] = parsed_result[matching_key]
            del parsed_result[matching_key]
        else:
            # Otherwise, add a new key
            parsed_result[keyword] = ""
    
    return parsed_result

def handle_different_start_end_keyword(parsed_result, end_keyword, original_raw_text, end_keyword_occurrence):
    """
    Handle case where start and end keywords are different.
    Prevents adding any empty keys or duplicates.
    
    Args:
        parsed_result (dict): Dictionary of parsed key-value pairs
        end_keyword (str): End keyword for extraction (can be None)
        original_raw_text (str): Original raw text
        end_keyword_occurrence (int): Occurrence number of end keyword to use
        
    Returns:
        dict: Processed parsed result
    """
    # If no end_keyword is provided, return the parsed result as is
    if not end_keyword:
        return parsed_result
        
    # Remove the colon from the end_keyword if it exists
    clean_end_keyword = end_keyword.rstrip(':')
    
    # Find all occurrences of the end_keyword in the text
    end_keyword_lines = []
    lines = original_raw_text.split('\n')
    
    for i, line in enumerate(lines):
        if end_keyword in line:
            end_keyword_lines.append((i, line))
    
    # If we have the expected occurrence of end_keyword
    if end_keyword_lines and end_keyword_occurrence <= len(end_keyword_lines):
        # Use the specified occurrence
        end_line_index, end_line = end_keyword_lines[end_keyword_occurrence - 1]
        
        # Find all keys that appear after this end_keyword line
        keys_to_remove = []
        for i, line in enumerate(lines):
            if i > end_line_index and ':' in line:
                parts = line.split(':', 1)
                key = parts[0].strip()
                if key in parsed_result:
                    keys_to_remove.append(key)
        
        # Remove the identified keys
        for key in keys_to_remove:
            if key in parsed_result:
                del parsed_result[key]
        
        # Check if the end_keyword is part of an existing key
        is_part_of_existing_key = False
        for key in parsed_result.keys():
            if clean_end_keyword in key:
                is_part_of_existing_key = True
                break
        
        # Only add the end_keyword as a key if it's not already present in any form
        # and it has actual content in the end line
        if not is_part_of_existing_key and ':' in end_line:
            parts = end_line.split(':', 1)
            key = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
            
            # Only add if the value is not empty
            if value:
                parsed_result[key] = value
    
    return parsed_result