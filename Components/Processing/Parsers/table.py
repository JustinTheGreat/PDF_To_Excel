"""
Table parsing functionality for PDF extraction.

This module handles the detection and extraction of tabular data from
raw text extracted from PDFs.
"""

import re

def process_table_data(raw_text, table_params):
    """
    Process extracted text as tabular data with specified labeling options.
    Dynamically handles tables of any size and content.
    
    Args:
        raw_text (str): Raw text extracted from PDF
        table_params (dict): Parameters for table processing with keys:
                           - table_top_labeling: Use top row as keys (boolean)
                           - table_left_labeling: Use left column as keys (boolean)
                           - table_structure: Structure type ('top_only', 'left_only', 'top_main', 'left_main')
                           - delimiter: Optional delimiter for cell separation (default: whitespace)
                           - min_column_width: Minimum width for columns when using space delimiter
                           - header_row: Optional specific row to use as header (default: 0)
                           - key_column: Optional specific column to use as row keys (default: 0)
        
    Returns:
        dict: Processed table data as nested dictionary
    """
    # Extract parameters
    top_labeling = table_params.get('table_top_labeling', False)
    left_labeling = table_params.get('table_left_labeling', False)
    structure_type = table_params.get('table_structure', 'top_only')
    delimiter = table_params.get('delimiter', None)  # None means use whitespace
    min_column_width = table_params.get('min_column_width', 3)
    header_row = table_params.get('header_row', 0)
    key_column = table_params.get('key_column', 0)
    
    # If no labeling is enabled, return empty dict
    if not top_labeling and not left_labeling:
        return {}
    
    # Split text into lines and filter empty lines
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
    
    # Handle empty text
    if not lines:
        return {}
    
    # Parse the table
    table_data = []
    
    # If a specific delimiter is provided, use it to split lines
    if delimiter:
        for line in lines:
            cells = [cell.strip() for cell in line.split(delimiter)]
            table_data.append(cells)
    else:
        # Use a more sophisticated approach for space-delimited tables
        # First, detect column positions
        positions = detect_column_positions(lines, min_column_width)
        
        if positions:
            # Use detected positions to extract cells
            for line in lines:
                cells = extract_cells_by_position(line, positions)
                table_data.append(cells)
        else:
            # Fallback to simple whitespace splitting
            for line in lines:
                cells = re.findall(r'(?:"[^"]*"|\S+)', line)
                cells = [cell.strip('"') for cell in cells]
                table_data.append(cells)
    
    # Ensure we have enough data
    if len(table_data) <= header_row:
        return {}
    
    # Normalize the table (ensure all rows have the same number of columns)
    max_cols = max(len(row) for row in table_data)
    for i in range(len(table_data)):
        while len(table_data[i]) < max_cols:
            table_data[i].append("")
    
    # Get headers and row keys
    headers = table_data[header_row][1:] if left_labeling else table_data[header_row]
    
    # Process according to structure type
    result = {}
    
    # Mode 1: Top row is the key and the rest is stored in tuples/lists
    if structure_type == 'top_only':
        # Skip header row
        data_rows = [row for i, row in enumerate(table_data) if i != header_row] if top_labeling else table_data
        
        # For each column header
        for i, header in enumerate(headers):
            if not header.strip():  # Skip empty headers
                continue
                
            values = []
            # Collect values for this column from all rows
            for row in data_rows:
                if i < len(row):  # Make sure column exists in this row
                    values.append(row[i])
            
            # Store in result
            if values:
                result[header] = values
    
    # Mode 2: Left column is the key and the rest is stored in tuples/lists
    elif structure_type == 'left_only':
        # Determine which rows to process
        data_rows = table_data[header_row+1:] if top_labeling else table_data
        
        for row in data_rows:
            if len(row) < 2:  # Need at least a label and one value
                continue
                
            row_label = row[key_column]
            if not row_label.strip():  # Skip rows with empty labels
                continue
                
            row_values = row[key_column+1:] if left_labeling else row[1:]
            
            # Add to result as a list of values for this row label
            result[row_label] = row_values
    
    # Mode 3: Top is main key and left is nested key for values
    elif structure_type == 'top_main':
        # Determine which rows to process
        data_rows = table_data[header_row+1:] if top_labeling else table_data[1:]
        
        # For each column header (starting from index 1 to skip row labels)
        col_start = key_column + 1 if left_labeling else 1
        
        for j, header in enumerate(headers):
            if not header.strip():  # Skip empty headers
                continue
                
            col_index = j + col_start  # Adjust index to account for row label column
            result[header] = {}
            
            # For each data row
            for row in data_rows:
                if len(row) <= key_column:  # Skip rows that don't have a label
                    continue
                    
                row_label = row[key_column]
                if not row_label.strip():  # Skip rows with empty labels
                    continue
                
                if col_index < len(row):  # Make sure column exists in this row
                    value = row[col_index]
                    result[header][row_label] = value
    
    # Mode 4: Left is main key and top is nested key for values
    elif structure_type == 'left_main':
        # Determine which rows to process
        data_rows = table_data[header_row+1:] if top_labeling else table_data[1:]
        
        # For each data row
        for row in data_rows:
            if len(row) <= key_column:  # Skip rows that don't have a label
                continue
                
            row_label = row[key_column]
            if not row_label.strip():  # Skip rows with empty labels
                continue
                
            result[row_label] = {}
            
            # For each column (skip the row label column)
            col_start = key_column + 1 if left_labeling else 1
            
            for j, header in enumerate(headers):
                if not header.strip():  # Skip empty headers
                    continue
                    
                col_index = j + col_start  # Adjust index
                if col_index < len(row):  # Make sure column exists
                    value = row[col_index]
                    result[row_label][header] = value
    
    return result

def detect_column_positions(lines, min_width=3):
    """
    Detect column positions in a space-delimited table.
    
    Args:
        lines (list): List of text lines
        min_width (int): Minimum width of a column
        
    Returns:
        list: List of column boundary positions
    """
    if not lines:
        return []
    
    # Find spaces that are consistent across multiple lines
    space_counts = {}
    for line in lines:
        for i, char in enumerate(line):
            if char.isspace():
                space_counts[i] = space_counts.get(i, 0) + 1
    
    # Find ranges of spaces that appear in most lines
    line_count = len(lines)
    threshold = line_count * 0.7  # Consider spaces that appear in at least 70% of lines
    
    # Group consecutive spaces
    space_groups = []
    current_group = []
    
    sorted_positions = sorted(space_counts.keys())
    if not sorted_positions:
        return []
    
    for i, pos in enumerate(sorted_positions):
        # Check if this position has enough occurrences
        if space_counts[pos] >= threshold:
            # Check if it's part of the current group
            if not current_group or pos == current_group[-1] + 1:
                current_group.append(pos)
            else:
                # Start a new group if the gap is too large
                if current_group:
                    space_groups.append(current_group)
                current_group = [pos]
        else:
            # This position doesn't have enough occurrences
            if current_group:
                space_groups.append(current_group)
                current_group = []
    
    # Add the last group if it exists
    if current_group:
        space_groups.append(current_group)
    
    # Filter groups that are wide enough to be column separators
    column_separators = []
    for group in space_groups:
        if len(group) >= min_width:
            # Take the middle of the group as the separator position
            column_separators.append(group[len(group) // 2])
    
    # Add start and end positions
    column_positions = [0] + column_separators
    
    # Add the right edge
    max_length = max(len(line) for line in lines)
    column_positions.append(max_length)
    
    return column_positions

def extract_cells_by_position(line, positions):
    """
    Extract cells from a line using the detected column positions.
    
    Args:
        line (str): Text line
        positions (list): List of column positions
        
    Returns:
        list: List of cell values
    """
    cells = []
    for i in range(len(positions) - 1):
        start = positions[i]
        end = positions[i+1]
        
        # Handle lines that are shorter than the table width
        if start >= len(line):
            cells.append("")
        else:
            cell_value = line[start:end].strip()
            cells.append(cell_value)
    
    return cells