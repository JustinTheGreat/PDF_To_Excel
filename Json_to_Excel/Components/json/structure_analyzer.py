class JsonStructureAnalyzer:
    """
    Class for analyzing the structure of JSON data specifically for Excel format determination.
    This is separate from JsonAnalyzer to focus only on structure analysis for Excel formatting.
    """
    
    @staticmethod
    def analyze_for_excel(json_data, print_debug=True):
        """
        Analyze JSON structure to determine Excel formatting requirements.
        
        Args:
            json_data: JSON data to analyze
            print_debug: Whether to print debug information
            
        Returns:
            Structure information dictionary
        """
        def debug_print(message):
            if print_debug:
                print(message)
        
        structure_info = {
            'keys': set(),
            'nesting_depth': {},  # Will store nested dimensions for each key
            'nesting_structure': {},  # Will store the structure of nested arrays
            'needs_subtitles': False
        }
        
        # Debug the input
        if isinstance(json_data, list):
            debug_print(f"analyze_for_excel: Input is a list with {len(json_data)} items")
        else:
            debug_print(f"analyze_for_excel: Input is a {type(json_data).__name__}")
        
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
                
                # Analyze the depth and structure of nested lists
                depth, dimensions, is_nested = JsonStructureAnalyzer._analyze_list_depth(value)
                
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
                        sub_depth, item_dimensions, _ = JsonStructureAnalyzer._analyze_list_depth(
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