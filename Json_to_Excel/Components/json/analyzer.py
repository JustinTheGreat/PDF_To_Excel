class JsonAnalyzer:
    """
    Enhanced class for analyzing the structure of JSON data to determine formatting needs.
    Support for multiple levels of nesting and hierarchical key-value lists.
    """
    
    @staticmethod
    def analyze_json_structure(json_data, print_debug=True):
        """
        Analyze the structure of the JSON data to determine how to format the Excel sheet.
        Now supports multiple levels of nesting, key-value pair lists, and nested key-value lists.
        
        Args:
            json_data: JSON data to analyze
            print_debug: Whether to print debug information
            
        Returns a dict with information about:
        - All unique keys
        - Maximum nesting depth for each key with nested dimensions
        - Whether subtitles are needed
        - Key-value list information
        - Multi-level key-value list information
        """
        def debug_print(message):
            if print_debug:
                print(message)
        
        structure_info = {
            'keys': set(),
            'nesting_depth': {},  # Will store nested dimensions for each key
            'nesting_structure': {},  # Will store the structure of nested arrays
            'needs_subtitles': False,
            'kv_lists': {},  # Store info about key-value pair lists
            'multilevel_kv': {},  # NEW: Store info about multi-level key-value structures
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
                
                # NEW: Check for multi-level key-value hierarchies
                if JsonAnalyzer._is_multilevel_key_value(value):
                    debug_print(f"  - Field '{key}' appears to be a multi-level key-value structure")
                    
                    # Analyze the multi-level structure
                    multilevel_structure = JsonAnalyzer._analyze_multilevel_key_value(value)
                    
                    if multilevel_structure['is_multilevel']:
                        debug_print(f"  - Confirmed as multi-level key-value with {len(multilevel_structure['level_info'])} levels")
                        structure_info['multilevel_kv'][key] = multilevel_structure
                        structure_info['needs_subtitles'] = True
                        
                        # Set nesting depth for multi-level KV structures
                        max_level = multilevel_structure['max_level']
                        structure_info['nesting_depth'][key] = max_level
                        
                        # Calculate dimensions based on number of keys at each level
                        dimensions = []
                        for level in range(1, max_level + 1):
                            if level in multilevel_structure['level_info']:
                                dimensions.append(len(multilevel_structure['level_info'][level]['unique_keys']))
                            else:
                                dimensions.append(1)
                        
                        structure_info['nesting_structure'][key] = dimensions
                        debug_print(f"  - Multi-level structure has dimensions: {dimensions}")
                        continue
                
                # Check for standard key-value lists
                elif JsonAnalyzer._is_key_value_list(value):
                    debug_print(f"  - Field '{key}' appears to be a key-value list")
                    
                    # Analyze the list structure
                    kv_structure = JsonAnalyzer._analyze_key_value_list(value)
                    
                    if kv_structure['is_kv_list']:
                        debug_print(f"  - Confirmed as key-value list with keys: {kv_structure['unique_keys']}")
                        structure_info['kv_lists'][key] = kv_structure
                        structure_info['needs_subtitles'] = True
                        
                        # Set nesting depth and structure for KV lists
                        depth = 1  # We treat KV lists as having a depth of 1
                        dimensions = [len(kv_structure['unique_keys'])]  # Number of unique keys as dimension
                        
                        structure_info['nesting_depth'][key] = depth
                        structure_info['nesting_structure'][key] = dimensions
                        continue
                
                # Standard analysis for regular nested lists
                depth, dimensions, is_nested = JsonAnalyzer._analyze_list_depth(value)
                
                # If it has any nesting, update the structure info
                if depth > 0:
                    current_max_depth = structure_info['nesting_depth'].get(key, 0)
                    
                    # Update nesting depth if this is deeper
                    if depth > current_max_depth:
                        structure_info['nesting_depth'][key] = depth
                        structure_info['nesting_structure'][key] = dimensions
                        debug_print(f"  - Field '{key}' has nested lists with dimensions: {dimensions}")
                    
                    # If we have at least one level of nesting, we need subtitles
                    if is_nested or dimensions[0] > 1:
                        structure_info['needs_subtitles'] = True
                        debug_print(f"  - Field '{key}' needs subtitles (nested: {is_nested}, dimensions: {dimensions})")
                elif key not in structure_info['nesting_depth']:
                    structure_info['nesting_depth'][key] = 0
                    structure_info['nesting_structure'][key] = []
                    debug_print(f"  - Field '{key}' has type {type(value).__name__}")
        
        debug_print(f"Analysis result: {len(structure_info['keys'])} unique keys, needs_subtitles={structure_info['needs_subtitles']}")
        return structure_info
    
    @staticmethod
    def _is_multilevel_key_value(value):
        """
        Determine if a value is a multi-level key-value structure.
        This could be:
        1. A list of dictionaries where values are also dictionaries or lists of dictionaries
        2. A dictionary with nested dictionaries or lists of dictionaries
        
        Args:
            value: The value to check
            
        Returns:
            Boolean indicating if this appears to be a multi-level key-value structure
        """
        # Case 1: List of dictionaries with dictionaries as values
        if isinstance(value, list) and len(value) > 0 and all(isinstance(item, dict) for item in value):
            # Check if any values in these dictionaries are also dictionaries or lists of dictionaries
            for item in value:
                for k, v in item.items():
                    if isinstance(v, dict) or (isinstance(v, list) and len(v) > 0 and all(isinstance(x, dict) for x in v)):
                        return True
            
        # Case 2: Dictionary with values that are lists of dictionaries
        elif isinstance(value, dict):
            for k, v in value.items():
                if isinstance(v, list) and len(v) > 0 and all(isinstance(item, dict) for item in v):
                    return True
        
        return False
    
    @staticmethod
    def _analyze_multilevel_key_value(value):
        """
        Analyze a multi-level key-value structure to extract metadata.
        
        Args:
            value: A multi-level key-value structure to analyze
            
        Returns:
            Dictionary with analysis results containing level information
        """
        result = {
            'is_multilevel': False,
            'max_level': 0,
            'level_info': {},  # Will contain info for each level: {level: {unique_keys: [...], ...}}
            'structure_type': None  # 'list_of_dicts', 'dict_of_lists', etc.
        }
        
        # Determine the structure type
        if isinstance(value, list) and len(value) > 0 and all(isinstance(item, dict) for item in value):
            result['structure_type'] = 'list_of_dicts'
            
            # Level 1: Keys from the dictionaries in the list
            level1_keys = set()
            for item in value:
                for k in item.keys():
                    level1_keys.add(k)
            
            result['level_info'][1] = {
                'unique_keys': sorted(level1_keys),
                'has_consistent_keys': all(set(item.keys()) == level1_keys for item in value)
            }
            
            # Check for level 2: Look for dictionaries or lists of dictionaries as values
            level2_keys = {}  # Mapping from level1 key to its level2 keys
            
            for item in value:
                for k, v in item.items():
                    if isinstance(v, dict):
                        if k not in level2_keys:
                            level2_keys[k] = set()
                        for k2 in v.keys():
                            level2_keys[k].add(k2)
                    elif isinstance(v, list) and len(v) > 0 and all(isinstance(x, dict) for x in v):
                        if k not in level2_keys:
                            level2_keys[k] = set()
                        # Collect all unique keys from all dictionaries in the list
                        for dict_item in v:
                            for k2 in dict_item.keys():
                                level2_keys[k].add(k2)
            
            # If we found any level 2 keys
            if level2_keys:
                result['level_info'][2] = {
                    'parent_keys': sorted(level2_keys.keys()),
                    'key_mapping': {k: sorted(keys) for k, keys in level2_keys.items()}
                }
                result['max_level'] = 2
                result['is_multilevel'] = True
            else:
                result['max_level'] = 1
                result['is_multilevel'] = True
                
        elif isinstance(value, dict):
            result['structure_type'] = 'dict_of_lists'
            
            # Level 1: Keys from the main dictionary
            level1_keys = sorted(value.keys())
            result['level_info'][1] = {
                'unique_keys': level1_keys
            }
            
            # Check for level 2
            level2_keys = {}
            
            for k, v in value.items():
                if isinstance(v, list) and len(v) > 0 and all(isinstance(item, dict) for item in v):
                    # Collect all unique keys from all dictionaries in the list
                    if k not in level2_keys:
                        level2_keys[k] = set()
                    for dict_item in v:
                        for k2 in dict_item.keys():
                            level2_keys[k].add(k2)
            
            # If we found any level 2 keys
            if level2_keys:
                result['level_info'][2] = {
                    'parent_keys': sorted(level2_keys.keys()),
                    'key_mapping': {k: sorted(keys) for k, keys in level2_keys.items()}
                }
                result['max_level'] = 2
                result['is_multilevel'] = True
            else:
                result['max_level'] = 1
                result['is_multilevel'] = True
        
        return result
    
    @staticmethod
    def _is_key_value_list(value):
        """
        Determine if a value is a list of dictionaries that could be treated as a key-value list.
        
        Args:
            value: The value to check
            
        Returns:
            Boolean indicating if this appears to be a key-value list
        """
        # Must be a list
        if not isinstance(value, list):
            return False
            
        # List must not be empty
        if len(value) == 0:
            return False
            
        # All items must be dictionaries
        return all(isinstance(item, dict) for item in value)
    
    @staticmethod
    def _analyze_key_value_list(value):
        """
        Analyze a potential key-value list structure to extract metadata.
        
        Args:
            value: A list of dictionaries to analyze
            
        Returns:
            Dictionary with analysis results
        """
        result = {
            'is_kv_list': False,
            'unique_keys': set(),
            'item_count': len(value),
            'has_consistent_keys': False
        }
        
        # Collect all unique keys
        for item in value:
            for k in item.keys():
                result['unique_keys'].add(k)
        
        # Check if all dictionaries have the same keys
        result['has_consistent_keys'] = all(
            set(item.keys()) == result['unique_keys']
            for item in value
        )
        
        # Convert unique_keys to a sorted list for consistent ordering
        result['unique_keys'] = sorted(result['unique_keys'])
        
        # Only consider it a key-value list if it has consistent keys
        result['is_kv_list'] = result['has_consistent_keys']
        
        return result
    
    @staticmethod
    def _analyze_list_depth(value, current_depth=0):
        """
        Recursively analyze a value to determine its nesting depth and dimensions.
        
        Args:
            value: The value to analyze
            current_depth: Current depth in the recursion
            
        Returns:
            Tuple of (max_depth, dimensions, is_nested)
            - max_depth: Maximum nesting depth
            - dimensions: List of sizes at each nesting level
            - is_nested: Boolean indicating if the structure has multiple levels of nesting
        """
        if isinstance(value, list):
            # This level is a list
            list_length = len(value)
            
            # If it's an empty list, return current info
            if list_length == 0:
                return current_depth, [0], current_depth > 1
            
            # Check if any items in this list are also lists
            has_nested_list = any(isinstance(item, list) for item in value)
            
            if has_nested_list:
                # We have nested lists, recurse to find max depth
                max_depth = current_depth
                sub_dimensions = []
                
                for item in value:
                    if isinstance(item, list):
                        sub_depth, item_dimensions, _ = JsonAnalyzer._analyze_list_depth(
                            item, current_depth + 1
                        )
                        max_depth = max(max_depth, sub_depth)
                        
                        # Merge dimensions if needed
                        if not sub_dimensions:
                            sub_dimensions = item_dimensions
                        else:
                            # Keep maximum dimension at each level
                            for i in range(min(len(sub_dimensions), len(item_dimensions))):
                                sub_dimensions[i] = max(sub_dimensions[i], item_dimensions[i])
                            
                            # Add any additional dimensions
                            if len(item_dimensions) > len(sub_dimensions):
                                sub_dimensions.extend(item_dimensions[len(sub_dimensions):])
                    
                # Prepend the current level's length
                dimensions = [list_length] + sub_dimensions
                return max_depth, dimensions, True
            else:
                # This is a simple, non-nested list
                return current_depth + 1, [list_length], current_depth > 0
        else:
            # Not a list, return current depth
            return current_depth, [], current_depth > 1