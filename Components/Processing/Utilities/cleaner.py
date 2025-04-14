"""
Data cleaning utilities for PDF processing.

This module provides functions for cleaning and sanitizing data 
extracted from PDF documents.
"""

def clean_empty_keys(data_dict):
    """
    Clean a dictionary by removing any keys with empty values.
    
    Args:
        data_dict (dict): Dictionary to clean
        
    Returns:
        dict: Cleaned dictionary with no empty values
    """
    # Create a copy to avoid modifying during iteration
    cleaned_dict = {}
    
    for key, value in data_dict.items():
        # Skip empty values and empty lists
        if value == "" or value == [] or value is None:
            continue
        
        # For lists, filter out any empty values
        if isinstance(value, list):
            filtered_values = [v for v in value if v != "" and v is not None]
            if filtered_values:  # Only add if there are non-empty values
                cleaned_dict[key] = filtered_values if len(filtered_values) > 1 else filtered_values[0]
        else:
            cleaned_dict[key] = value
    
    return cleaned_dict