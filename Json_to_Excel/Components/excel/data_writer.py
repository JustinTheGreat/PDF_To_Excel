from Components.utils.text_filters import TextFilter
from Components.utils.file_utils import FileUtils

class ExcelDataWriter:
    """
    Enhanced class for writing data to Excel worksheets with support for complex data structures
    including nested lists and key-value pair lists.
    """
    
    def add_data_row(self, worksheet, row_num, file_name, fields, structure_info, max_list_lengths, 
                     filter_text="", apply_value_filters=True):
        """
        Add a row of data to the worksheet with support for nested lists and key-value lists.
        
        Args:
            worksheet: The worksheet to write to
            row_num: Row number to write at
            file_name: Name of the source file
            fields: Dictionary of field key-value pairs to write
            structure_info: Dictionary with structure information about the data
            max_list_lengths: Dictionary of maximum list lengths (deprecated)
            filter_text: Text to remove from filenames
            apply_value_filters: Whether to apply text filters to values
        """
        # Process filename to remove extension and filter text
        display_filename = FileUtils.process_filename(file_name, filter_text)
        
        # Write the processed filename
        worksheet.cell(row=row_num, column=1, value=display_filename)
        
        # Start with column 2 (after file name column)
        current_column = 2
        
        # Process each field
        for key in sorted(structure_info['keys']):
            value = fields.get(key, "")
            nesting_structure = structure_info['nesting_structure'].get(key, [])
            
            # Check if this is a key-value list field
            if 'kv_lists' in structure_info and key in structure_info['kv_lists']:
                # Handle key-value list type fields
                column_increment = self._add_key_value_list_data(
                    worksheet,
                    row_num,
                    current_column,
                    value,
                    structure_info['kv_lists'][key],
                    apply_value_filters
                )
                current_column += column_increment
            
            # Handle the different value types for regular lists
            elif nesting_structure:
                # This field might have nested lists
                column_increment = self._add_nested_data(
                    worksheet, 
                    row_num, 
                    current_column, 
                    value, 
                    nesting_structure, 
                    apply_value_filters
                )
                current_column += column_increment
            else:
                # This field has a single value or is not a list
                # Set the value (single value or first item of a list)
                if isinstance(value, list) and value:
                    value_to_set = value[0]
                else:
                    value_to_set = value
                
                # Apply text filtering if needed
                if apply_value_filters and isinstance(value_to_set, str):
                    value_to_set = TextFilter.remove_units(value_to_set)
                    
                # Set the processed value
                worksheet.cell(row=row_num, column=current_column, value=value_to_set)
                current_column += 1
    
    def _add_key_value_list_data(self, worksheet, row_num, start_column, value, kv_list_info, apply_value_filters):
        """
        Add key-value list data to a worksheet row.
        
        Args:
            worksheet: The worksheet to add data to
            row_num: The row number
            start_column: The starting column
            value: The value (list of dictionaries)
            kv_list_info: Information about the key-value list structure
            apply_value_filters: Whether to apply text filters to values
        
        Returns:
            The number of columns used
        """
        # Get the ordered list of keys
        ordered_keys = kv_list_info['unique_keys']
        total_columns = len(ordered_keys)
        
        # Initialize all cells to empty
        for col in range(start_column, start_column + total_columns):
            worksheet.cell(row=row_num, column=col, value="")
        
        # Handle if value is not a list or is empty
        if not isinstance(value, list) or not value:
            return total_columns
        
        # Extract values for each key from the first item in the list
        # (We assume the first item has the information we want)
        first_item = value[0]
        if not isinstance(first_item, dict):
            return total_columns
            
        # Add value for each key in the order specified
        for i, key in enumerate(ordered_keys):
            if key in first_item:
                item_value = first_item[key]
                
                # Apply filters if needed
                if apply_value_filters and isinstance(item_value, str):
                    item_value = TextFilter.remove_units(item_value)
                
                # Set the cell value
                col = start_column + i
                worksheet.cell(row=row_num, column=col, value=item_value)
        
        return total_columns
    
    def _add_nested_data(self, worksheet, row_num, start_column, value, dimensions, apply_value_filters):
        """
        Add nested data to a worksheet row.
        
        Args:
            worksheet: The worksheet to add data to
            row_num: The row number
            start_column: The starting column
            value: The value (possibly nested list)
            dimensions: List of dimensions for the nested structure
            apply_value_filters: Whether to apply text filters to values
        
        Returns:
            The number of columns used
        """
        if not dimensions:
            # Handle non-list value
            if apply_value_filters and isinstance(value, str):
                value = TextFilter.remove_units(value)
            worksheet.cell(row=row_num, column=start_column, value=value)
            return 1
        
        # Calculate the total number of columns
        total_columns = self._calculate_total_columns(dimensions)
        
        # Initialize all cells to empty
        for col in range(start_column, start_column + total_columns):
            worksheet.cell(row=row_num, column=col, value="")
        
        # Flatten the nested list structure to map to columns
        flattened_values = []
        self._flatten_nested_list(value, flattened_values, dimensions)
        
        # Add values to cells
        for i, item in enumerate(flattened_values):
            if i < total_columns:
                col = start_column + i
                if apply_value_filters and isinstance(item, str):
                    item = TextFilter.remove_units(item)
                worksheet.cell(row=row_num, column=col, value=item)
        
        return total_columns
    
    def _calculate_total_columns(self, dimensions):
        """
        Calculate the total number of columns needed for a nested structure.
        
        Args:
            dimensions: List of dimensions at each nesting level [d1, d2, d3, ...]
        
        Returns:
            Total number of columns needed
        """
        if not dimensions:
            return 1
        
        # Multiply all dimensions together
        total = 1
        for dim in dimensions:
            total *= max(1, dim)  # Ensure at least 1 column even for empty dimensions
        
        return total
    
    def _flatten_nested_list(self, value, result, dimensions, current_dim=0):
        """
        Recursively flatten a nested list structure.
        
        Args:
            value: The value to flatten (may be a nested list)
            result: List to store flattened values
            dimensions: List of dimensions for the nested structure
            current_dim: Current dimension being processed
        """
        if not isinstance(value, list):
            # Base case: not a list, add the value
            result.append(value)
            return
        
        # If we've reached the end of our dimensions but still have a list
        if current_dim >= len(dimensions):
            # Just add the first item if available
            if value:
                result.append(value[0])
            else:
                result.append("")
            return
        
        # Current dimension size
        dim_size = dimensions[current_dim]
        
        # Process each item in the current dimension
        for i in range(dim_size):
            if i < len(value):
                # Recurse with the nested item
                self._flatten_nested_list(value[i], result, dimensions, current_dim + 1)
            else:
                # Fill in with blanks for missing items
                # Calculate how many empty slots to add
                if current_dim < len(dimensions) - 1:
                    empties_to_add = self._calculate_total_columns(dimensions[current_dim + 1:])
                else:
                    empties_to_add = 1
                
                for _ in range(empties_to_add):
                    result.append("")