import re
from typing import Dict, List, Any, Tuple, Union, Optional

class BusinessRules:
    """
    Class for implementing custom business rules that transform JSON data
    before it is written to Excel.
    
    This module allows for application-specific transformations that are 
    separate from the core Excel generation functionality.
    """
    
    @staticmethod
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
        
        # Apply nested key-value list transformations
        result = BusinessRules.transform_nested_key_value_lists(result, debug)
        
        # Convert dictionary fields to key-value list format
        result = BusinessRules.transform_dict_fields(result, debug)
        
        # Apply key-value list transformation (should be done after the other transformations)
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
    def transform_nested_key_value_lists(data: Dict[str, Any], debug=False) -> Dict[str, Any]:
        """
        Flatten nested key-value lists that are inside arrays to remove the extra level of nesting.
        Also, handle cases where a dictionary is being passed directly to a field.
        
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
            print(f"  Examining {len(fields)} fields for nested key-value lists")
        
        # Process all dictionary values recursively
        for key, value in list(fields.items()):
            # Process lists
            if isinstance(value, list):
                # If it's a list with exactly one dictionary item, flatten it
                if len(value) == 1 and isinstance(value[0], dict):
                    if debug:
                        print(f"  Flattening single-item key-value list in '{key}'")
                    fields[key] = value[0]
                
                # Process each list item recursively for complex nested structures
                new_list = []
                for i, item in enumerate(value):
                    if isinstance(item, dict):
                        # Create a recursively transformed version of this dict
                        transformed_item = item.copy()
                        for sub_key, sub_value in list(transformed_item.items()):
                            if isinstance(sub_value, list) and len(sub_value) == 1 and isinstance(sub_value[0], dict):
                                if debug:
                                    print(f"  Flattening nested key-value list in '{key}[{i}].{sub_key}'")
                                transformed_item[sub_key] = sub_value[0]
                            elif isinstance(sub_value, dict):
                                # Handle deeply nested dictionaries
                                for nested_key, nested_value in list(sub_value.items()):
                                    if isinstance(nested_value, list) and len(nested_value) == 1 and isinstance(nested_value[0], dict):
                                        if debug:
                                            print(f"  Flattening deeply nested key-value list in '{key}[{i}].{sub_key}.{nested_key}'")
                                        sub_value[nested_key] = nested_value[0]
                        new_list.append(transformed_item)
                    else:
                        new_list.append(item)
                
                # Only update if we made changes
                if new_list and value is not new_list:
                    fields[key] = new_list
                    
            # Process nested dictionaries
            elif isinstance(value, dict):
                # Create a transformed copy
                transformed_dict = value.copy()
                
                # Process each nested field
                for sub_key, sub_value in list(transformed_dict.items()):
                    if isinstance(sub_value, list) and len(sub_value) == 1 and isinstance(sub_value[0], dict):
                        if debug:
                            print(f"  Flattening nested key-value list in '{key}.{sub_key}'")
                        transformed_dict[sub_key] = sub_value[0]
                    elif isinstance(sub_value, dict):
                        # Process deeply nested dictionaries
                        for nested_key, nested_value in list(sub_value.items()):
                            if isinstance(nested_value, list) and len(nested_value) == 1 and isinstance(nested_value[0], dict):
                                if debug:
                                    print(f"  Flattening deeply nested key-value list in '{key}.{sub_key}.{nested_key}'")
                                sub_value[nested_key] = nested_value[0]
                
                # Update if we made changes
                if transformed_dict != value:
                    fields[key] = transformed_dict
        
        # If we were working with a nested 'fields' dictionary, update it
        if 'fields' in result and result['fields'] is not fields:
            result['fields'] = fields
        
        return result
    
    @staticmethod
    def transform_dict_fields(data: Dict[str, Any], debug=False) -> Dict[str, Any]:
        """
        Transform dictionary fields into a format that can be properly written to Excel.
        
        This function looks for fields that are dictionaries and converts them to strings
        or other formats that Excel can handle.
        
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
            print(f"  Examining {len(fields)} fields for dictionary values")
        
        # Find fields that are dictionaries but not lists
        for key, value in list(fields.items()):
            if isinstance(value, dict):
                # Convert the dictionary to a key-value list format which the Excel generator can handle
                if debug:
                    print(f"  Converting dictionary field '{key}' to key-value list format")
                
                # Create a key-value list with just one item - this is a format the Excel generator understands
                fields[key] = [value]
        
        # If we were working with a nested 'fields' dictionary, update it
        if 'fields' in result and result['fields'] is not fields:
            result['fields'] = fields
        
        return result
    
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