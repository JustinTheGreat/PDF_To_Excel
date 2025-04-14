"""
Business rules for PDF extraction and formatting.

This module contains business-specific rules for formatting extracted data.
Add your own custom formatting rules for specific field types below.
"""
import re


def format_custom_field_data(data_dict, unparsed_lines):
    """
    Format data for your custom field type.
    
    Implement your custom formatting logic here. This function should
    transform the provided data dictionary according to specific business rules.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Formatted data dictionary
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # TODO: Implement your custom formatting logic here
    # Example:
    # - Fix inconsistent keys
    # - Transform data formats
    # - Extract additional data from unparsed lines
    # - Combine related fields
    
    return result_dict


def apply_business_rules(field_name, data_dict, unparsed_lines):
    """
    Apply all business-specific rules based on the field name.
    
    Args:
        field_name (str): Name of the field
        data_dict (dict): Dictionary of parsed data
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with business rules applied
    """
    # Create a copy to avoid modifying the original
    result_dict = data_dict.copy()

    # Apply field-specific formatting rules
    if field_name == "Your Field Name":
        result_dict = format_custom_field_data(result_dict, unparsed_lines)
    # Add more field-specific rules as needed
    # elif field_name == "Another Field":
    #     result_dict = format_another_field_data(result_dict, unparsed_lines)
    
    return result_dict