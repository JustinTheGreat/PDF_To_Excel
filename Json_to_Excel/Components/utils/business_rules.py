import re
from typing import Dict, List, Any, Tuple, Union, Optional

class BusinessRules:
    """
    Class for implementing custom business rules that transform JSON data
    before it is written to Excel.
    
    This module allows for application-specific transformations that are 
    separate from the core Excel generation functionality.
    """
    
    def transform_key_value_lists(data: Dict[str, Any], debug=False) -> Dict[str, Any]:
        """
        Transform appropriate fields into key-value lists for better display in Excel.
        
        This transformation looks for fields containing lists of dictionaries with 
        consistent keys and converts them into a format that the Excel generator
        will recognize as key-value lists.
        
        Example:
            "Parameters": [
                {"name": "Frequency", "value": "60 Hz"},
                {"name": "Voltage", "value": "230 V"}
            ]
            
        Will be processed specially to show as subtitles in Excel.
        
        Args:
            data: The data dictionary to transform
            debug: Whether to print debug messages
            
        Returns:
            Transformed data dictionary
        """
        if not isinstance(data, dict):
            if debug:
                print(f"  Not processing non-dictionary data of type {type(data)}")
            return data
        
        result = data.copy()
        
        # Look for fields dictionary if it exists
        fields = result.get('fields', result)
        
        # Verify fields is a dictionary
        if not isinstance(fields, dict):
            if debug:
                print(f"  'fields' is not a dictionary, it's a {type(fields)}")
            return result
                
        if debug:
            print(f"  Examining {len(fields)} fields for key-value lists")
        
        # Find potential key-value list fields (lists of dictionaries)
        for key, value in fields.items():
            if isinstance(value, list) and value and all(isinstance(item, dict) for item in value):
                if debug:
                    print(f"  Found potential key-value list in field '{key}'")
                
                # Check if all dictionaries have the same keys
                if value:
                    first_keys = set(value[0].keys())
                    if all(set(item.keys()) == first_keys for item in value):
                        if debug:
                            print(f"  Confirmed key-value list with keys: {first_keys}")
                        
                        # This field is already in the right format for the enhanced Excel generator
                        # No transformation needed, but we can mark this for debug purposes
                        if debug:
                            print(f"  Field '{key}' will be processed as a key-value list")
        
        # If we were working with a nested 'fields' dictionary, update it
        if 'fields' in result and result['fields'] is not fields:
            result['fields'] = fields
                
        return result
    @staticmethod
    def transform_data(json_data: Dict[str, Any], debug=False) -> Dict[str, Any]:
        """
        Apply all business rules transformations to a single JSON data object.
        
        Args:
            json_data: A dictionary containing the JSON data to transform
            debug: Whether to print debug messages
            
        Returns:
            The transformed JSON data
        """
        result = json_data.copy()  # Create a copy to avoid modifying the original
        
        # Apply specific transformations with debug flag
        result = BusinessRules.transform_overshoot_values(result, debug)
        
        # Add the new transformation
        result = BusinessRules.transform_key_value_lists(result, debug)
        
        # Add more transformations here as needed
        
        return result
    
    @staticmethod
    def transform_all_data(all_json_data: Dict[str, Any], debug=True) -> Dict[str, Any]:
        """
        Apply business rules to all JSON data entries.
        
        Args:
            all_json_data: Dictionary mapping file paths to their JSON content
            debug: Whether to print debug messages
            
        Returns:
            Transformed data dictionary
        """
        if debug:
            print("\n==== Starting Business Rules Transformation ====")
            print(f"Processing {len(all_json_data)} JSON files")
        
        transformed_data = {}
        
        for file_name, file_json_data in all_json_data.items():
            if debug:
                print(f"\nProcessing file: {file_name}")
            
            # Handle different data structures (list or dict)
            if isinstance(file_json_data, list):
                if debug:
                    print(f"  File contains a list with {len(file_json_data)} items")
                transformed_data[file_name] = [
                    BusinessRules.transform_data(item, debug) 
                    for item in file_json_data
                ]
            else:
                if debug:
                    print(f"  File contains a dictionary with {len(file_json_data.keys()) if hasattr(file_json_data, 'keys') else 0} keys")
                transformed_data[file_name] = BusinessRules.transform_data(file_json_data, debug)
                
        if debug:
            print("\n==== Business Rules Transformation Complete ====")
            
        return transformed_data
    
    @staticmethod
    def transform_overshoot_values(data: Dict[str, Any], debug=False) -> Dict[str, Any]:
        """
        Transform "Overshoot [V]" values into an array for proper subtitle display.
        
        This function looks for "Overshoot" keys that contain brackets and converts the space-separated
        values into an array with the following order:
        [min_value, max_value]
        
        This array representation causes ExcelGenerator to create subtitles in the Excel output,
        where the first column will be for the minimum value and the second column for the maximum.
        
        Example:
            "Overshoot [V]": "3.5 2.1" becomes:
            "Overshoot [V]": ["2.1", "3.5"]
            
        Args:
            data: The data dictionary to transform
            debug: Whether to print debug messages
            
        Returns:
            Transformed data dictionary
        """
        if not isinstance(data, dict):
            if debug:
                print(f"  Not processing non-dictionary data of type {type(data)}")
            return data
        
        result = data.copy()
        
        # Look for fields dictionary if it exists
        fields = result.get('fields', result)
        
        # Verify fields is a dictionary
        if not isinstance(fields, dict):
            if debug:
                print(f"  'fields' is not a dictionary, it's a {type(fields)}")
            return result
            
        if debug:
            print(f"  Examining {len(fields)} fields for overshoot values")
        
        # Find overshoot keys (including any with units)
        overshoot_keys = [
            key for key in fields.keys() 
            if isinstance(key, str) and key.lower().startswith('overshoot') and '[' in key
        ]
        
        if debug:
            print(f"  Found {len(overshoot_keys)} potential overshoot keys: {overshoot_keys}")
        
        transformations_made = 0
        
        for key in overshoot_keys:
            value = fields.get(key)
            
            if debug:
                print(f"  Examining key '{key}' with value: {value!r} (type: {type(value)})")
            
            # Skip if not a string or not containing a space
            if not isinstance(value, str):
                if debug:
                    print(f"  - Skipping: value is not a string, it's a {type(value)}")
                continue
                
            if ' ' not in value:
                if debug:
                    print(f"  - Skipping: value does not contain a space")
                continue
                
            # Parse the values
            try:
                # Split by space and extract numbers
                parts = value.split()
                
                if debug:
                    print(f"  - Split value into parts: {parts}")
                
                # Handle the case of two numbers
                if len(parts) == 2:
                    max_val = parts[0]
                    min_val = parts[1]
                    
                    if debug:
                        print(f"  - Transforming to array [min={min_val}, max={max_val}]")
                    
                    # Replace the original field with an array [min, max]
                    # This array approach ensures proper subtitling in Excel
                    # with min before max, as requested
                    fields[key] = [min_val, max_val]  # Min first, then Max
                    transformations_made += 1
                else:
                    if debug:
                        print(f"  - Value has {len(parts)} parts, expected 2")
            except Exception as e:
                # If parsing fails, keep the original field
                if debug:
                    print(f"  - Error parsing value: {str(e)}")
                continue
        
        if debug:
            print(f"  Made {transformations_made} overshoot transformations")
        
        # If we were working with a nested 'fields' dictionary, update it
        if 'fields' in result and result['fields'] is not fields:
            result['fields'] = fields
            
        return result