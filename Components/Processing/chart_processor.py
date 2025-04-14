"""
Improved chart processing module for PDF extraction that works with formatted text.

This module transforms structured data from PDFs into chart-friendly formats
based on parameters that control how keys and values are organized.
"""

def process_chart_data(extracted_data, extraction_params=None):
    """
    Process extracted data and identify fields marked for chart conversion.
    Uses formatted_text instead of raw_text to ensure proper formatting is preserved.
    
    Args:
        extracted_data (dict): Dictionary containing extracted field data
        extraction_params (list, optional): Original extraction parameters
        
    Returns:
        dict: Processed data with chart fields converted to chart format
    """
    processed_data = {}
    chart_candidates = {}
    
    # First pass: identify fields that need chart conversion
    for field_name in extracted_data.keys():
        if "(Chart)" in field_name:
            # Extract the base field name (without the Chart indicator)
            base_field_name = field_name.replace("(Chart)", "").strip()
            
            # Find all related fields (base field and extensions)
            related_fields = [
                f for f in extracted_data.keys() 
                if f == base_field_name or 
                (f.startswith(base_field_name) and "(+1)" in f)
            ]
            
            if related_fields:
                chart_candidates[base_field_name] = {
                    "chart_field": field_name,
                    "related_fields": related_fields
                }
            else:
                # If no related fields found, keep as is but under the base name
                processed_data[base_field_name] = {
                    "raw_text": extracted_data[field_name]["raw_text"],
                    "formatted_text": extracted_data[field_name]["formatted_text"],
                    "parsed_data": extracted_data[field_name]["parsed_data"],
                    # Keep track that this was originally a chart field
                    "was_chart": True
                }
        elif not any(field_name == info["chart_field"] for info in chart_candidates.values()):
            # Regular field (not a chart field), add to processed data
            # Only add if it's not already identified as a chart field
            processed_data[field_name] = extracted_data[field_name]
    
    # Second pass: convert chart candidates and merge into the base field name
    for base_field_name, chart_info in chart_candidates.items():
        chart_field = chart_info["chart_field"]
        related_fields = chart_info["related_fields"]
        
        # Extract chart parameters from multiple possible locations
        params = {}
        
        # First check if parameters are in the parsed_data dictionary
        if isinstance(extracted_data[chart_field].get("parsed_data"), dict):
            params = extracted_data[chart_field].get("parsed_data", {})
            
        # Now check for chart parameters in the extraction_params
        chart_params_found = False
        if extraction_params:
            # First look specifically for params with "(Chart)" in field_name
            for param_set in extraction_params:
                param_field_name = param_set.get("field_name", "")
                if "(Chart)" in param_field_name and param_field_name.replace("(Chart)", "").strip() == base_field_name:
                    # Found chart-specific parameters, add them
                    for key, value in param_set.items():
                        if key not in ["raw_text", "formatted_text", "parsed_data", "field_name", 
                                      "start_keyword", "end_keyword", "page_num", "horiz_margin", 
                                      "end_keyword_occurrence", "left_move"]:
                            params[key] = value
                    chart_params_found = True
                    break
            
            # Only if we didn't find chart-specific params, look for base field params
            if not chart_params_found:
                for param_set in extraction_params:
                    param_field_name = param_set.get("field_name", "")
                    if param_field_name == base_field_name:
                        # Found base field parameters, add them
                        for key, value in param_set.items():
                            if key not in ["raw_text", "formatted_text", "parsed_data", "field_name", 
                                          "start_keyword", "end_keyword", "page_num", "horiz_margin", 
                                          "end_keyword_occurrence", "left_move"]:
                                params[key] = value
                        break
        
        # Also check if parameters are at the top level of the chart field data
        for key, value in extracted_data[chart_field].items():
            # Skip certain known keys that are not parameters
            if key not in ["raw_text", "formatted_text", "parsed_data"]:
                params[key] = value
        
        # Set chart parameters - handle case-insensitive keys and multiple formats
        # Check for both camelCase, PascalCase and snake_case variations
        left_title = False
        for key in ["left_title", "Left Title", "LeftTitle", "lefttitle", "LEFTTITLE"]:
            if key in params:
                val = params[key]
                if isinstance(val, str):
                    # More robust string conversion - also handle "True" with case variation
                    left_title = val.lower() in ["true", "yes", "1", "t", "y"]
                else:
                    left_title = bool(val)
                break  # Stop once we've found and processed a matching key
                    
        top_title = False
        for key in ["top_title", "Top Title", "TopTitle", "toptitle", "TOPTITLE"]:
            if key in params:
                val = params[key]
                if isinstance(val, str):
                    # More robust string conversion - also handle "True" with case variation
                    top_title = val.lower() in ["true", "yes", "1", "t", "y"]
                else:
                    top_title = bool(val)
                break  # Stop once we've found and processed a matching key
                    
        priority_side = "top"  # Default priority
        for key in ["priority_side", "Priority Side", "PrioritySide", "priorityside", "PRIORITYSIDE"]:
            if key in params:
                val = str(params[key]).lower()
                if val in ["left", "l"]:
                    priority_side = "left"
                break  # Stop once we've found and processed a matching key
        
        chart_params = {
            "left_title": left_title,
            "top_title": top_title,
            "priority_side": priority_side
        }
        
        # Combine related fields data for chart processing
        # USE FORMATTED TEXT INSTEAD OF RAW TEXT
        chart_data = _combine_chart_data(extracted_data, related_fields, use_formatted_text=True)
        
        # Structure the data based on chart parameters
        structured_data = _structure_chart_data(chart_data, chart_params)
        
        # Check if base field already exists in processed data
        if base_field_name in processed_data:
            # Keep the original raw_text and formatted_text
            raw_text = processed_data[base_field_name]["raw_text"]
            formatted_text = processed_data[base_field_name]["formatted_text"]
            
            # Update only the parsed_data with the structured chart data
            parsed_data = processed_data[base_field_name]["parsed_data"]
            
            # Merge the parsed data with the structured chart data
            for key, value in structured_data.items():
                if key in parsed_data:
                    # Handle merging if the key already exists
                    if isinstance(parsed_data[key], list) and isinstance(value, list):
                        parsed_data[key].extend(value)
                    elif isinstance(parsed_data[key], list):
                        parsed_data[key].append(value)
                    elif isinstance(value, list):
                        parsed_data[key] = [parsed_data[key]] + value
                    elif parsed_data[key] != value:
                        parsed_data[key] = [parsed_data[key], value]
                else:
                    parsed_data[key] = value
            
            # Update the processed data with merged parsed data
            processed_data[base_field_name] = {
                "raw_text": raw_text,
                "formatted_text": formatted_text,
                "parsed_data": parsed_data,
                "has_chart_data": True
            }
        else:
            # Base field doesn't exist yet, create it with chart data
            # For the raw_text and formatted_text, we'll use the data from the chart field
            # but keep the structured data as the parsed_data
            processed_data[base_field_name] = {
                "raw_text": extracted_data[chart_field]["raw_text"],
                "formatted_text": extracted_data[chart_field]["formatted_text"],
                "parsed_data": structured_data,
                "has_chart_data": True
            }
    
    return processed_data

def _combine_chart_data(extracted_data, related_fields, use_formatted_text=True):
    """
    Combine data from related fields for chart processing.
    Can use either formatted_text or raw_text based on the use_formatted_text parameter.
    
    Args:
        extracted_data (dict): Dictionary containing extracted field data
        related_fields (list): List of field names to combine
        use_formatted_text (bool): Whether to use formatted_text instead of raw_text
        
    Returns:
        dict: Combined data structure for chart processing
    """
    # Sort fields to ensure base field comes first
    related_fields.sort(key=lambda x: 0 if "(+1)" not in x else 1)
    
    # Start with the first field's data
    text_type = "formatted_text" if use_formatted_text else "raw_text"
    
    combined = {
        "raw_text": extracted_data[related_fields[0]]["raw_text"],
        "formatted_text": extracted_data[related_fields[0]]["formatted_text"],
        "used_text_type": text_type,
        "columns": []
    }
    
    # Process each field as a column
    for field_name in related_fields:
        field_data = extracted_data[field_name]
        
        # Use the appropriate text type (formatted or raw)
        text_to_process = field_data[text_type]
        
        # Split by the separator
        sections = text_to_process.split("\n\n--- Additional Data ---\n\n")
        
        # Process each section as a column
        for i, section in enumerate(sections):
            lines = [line.strip() for line in section.split('\n') if line.strip()]
            
            # Make sure we have meaningful data
            if not lines:
                continue
                
            if i >= len(combined["columns"]):
                combined["columns"].append(lines)
            else:
                # When extending an existing column, make sure we preserve the header
                # if the column already has data
                if len(combined["columns"][i]) > 0:
                    # Keep the existing header and add new data
                    header = combined["columns"][i][0]
                    # Use the new lines but keep the header from the existing column
                    if len(lines) > len(combined["columns"][i]):
                        new_column = [header] + lines[1:] if len(lines) > 1 else [header]
                        combined["columns"][i] = new_column
                else:
                    # No existing data, just add the new lines
                    combined["columns"][i] = lines
    
    # Ensure we have consistent column lengths
    max_length = max([len(col) for col in combined["columns"]]) if combined["columns"] else 0
    for i, column in enumerate(combined["columns"]):
        # If a column is empty or doesn't have enough rows, pad it
        while len(column) < max_length:
            column.append("")
    
    return combined

def _structure_chart_data(chart_data, chart_params):
    """
    Structure chart data based on specified parameters.
    
    Args:
        chart_data (dict): Combined chart data
        chart_params (dict): Chart parameters
        
    Returns:
        dict: Structured chart data
    """
    columns = chart_data["columns"]
    
    # Ensure we have data to process
    if not columns or len(columns) == 0:
        return {}
    
    # Extract parameters
    left_title = chart_params["left_title"]
    top_title = chart_params["top_title"]
    priority_side = chart_params["priority_side"]
    
    result = {}
    
    # Case 1: Only Top Title is enabled
    if top_title and not left_title:
        # Each column header (first item) becomes a key, and its values are the rest of the column
        for column in columns:
            if column and len(column) > 0:
                key = column[0]
                values = column[1:] if len(column) > 1 else []
                # Store only the values without the header
                result[key] = values
    
    # Case 2: Only Left Title is enabled
    elif left_title and not top_title:
        # First column contains keys, other columns are values for each key
        if len(columns) > 0 and columns[0]:
            # Extract keys from first column (including the first item)
            keys = columns[0]
            
            # For each key in the first column, collect values from corresponding rows in other columns
            for i, key in enumerate(keys):
                if not key.strip():  # Skip empty keys
                    continue
                    
                values = []
                # Collect values from all other columns at the same position
                for j in range(1, len(columns)):
                    if j < len(columns) and i < len(columns[j]):
                        values.append(columns[j][i])
                
                # Store the values for this key
                result[key] = values if len(values) > 1 else (values[0] if values else "")
    
    # Case 3: Both Top and Left Titles are enabled
    elif top_title and left_title:
        # Extract headers - use first row of each column as the header name
        # Avoid using generic "Column X" names when we have actual header data
        headers = []
        for i, col in enumerate(columns):
            if col and len(col) > 0:
                headers.append(col[0])  # Use the first item in the column as the header
            else:
                headers.append(f"Column {i+1}")  # Fall back to generic name only if needed
        
        # Extract row keys from first column (skip header)
        row_keys = columns[0][1:] if columns[0] and len(columns[0]) > 1 else []
        
        # Priority determines nesting structure
        if priority_side == "top":
            # Get the top-left cell (first cell of first column) as the main key
            main_key = columns[0][0] if columns[0] and len(columns[0]) > 0 else "Data"
            
            # Create the nested structure
            nested_structure = {}
            
            # Process each column starting from the second column (skip row headers column)
            for col_idx in range(1, len(columns)):
                col_header = headers[col_idx]
                if not col_header.strip():  # Skip columns with empty headers
                    continue
                
                # Create a dictionary for this column
                nested_structure[col_header] = {}
                
                # For each row key from the first column (skip the header row)
                for row_idx, row_key in enumerate(row_keys):
                    if not row_key.strip():  # Skip rows with empty keys
                        continue
                    
                    # Add value if it exists
                    row_idx_adjusted = row_idx + 1  # Adjust for header row
                    if row_idx_adjusted < len(columns[col_idx]):
                        value = columns[col_idx][row_idx_adjusted]
                        if value.strip():  # Only add non-empty values
                            nested_structure[col_header][row_key] = value
            
            # Add the nested structure under the main key
            if nested_structure:
                result[main_key] = nested_structure
            else:
                # Fallback if no nested structure was created
                result[main_key] = {}
        else:  # priority_side == "left"
            # Left column items are the main keys
            
            # First row of first column is the main key (e.g., "Parts")
            if not columns[0] or len(columns[0]) == 0:
                return {}
                
            main_key = columns[0][0]
            if not main_key.strip():
                main_key = "Data"
            
            # Row keys are the rest of the first column (e.g., "Storage", "Test", "Test Load")
            row_keys = columns[0][1:] if len(columns[0]) > 1 else []
            
            # Column headers are the first row of other columns (e.g., "Typ", "Comment")
            col_headers = []
            for i in range(1, len(columns)):
                if columns[i] and len(columns[i]) > 0:
                    col_headers.append(columns[i][0])
                else:
                    col_headers.append(f"Column {i}")
            
            # Create the main nested object
            nested_obj = {}
            
            # For each row key (e.g., "Storage", "Test", "Test Load")
            for row_idx, row_key in enumerate(row_keys):
                if not row_key.strip():  # Skip empty keys
                    continue
                
                # Create an object with column values for this row
                col_values = {}
                
                # Gather all column values for this row
                for col_idx, col_header in enumerate(col_headers):
                    # Actual column index (skip first column)
                    actual_col_idx = col_idx + 1
                    
                    # Actual row index (skip header row)
                    actual_row_idx = row_idx + 1  # +1 because we skip the header row
                    
                    # Get the value if it exists
                    if (actual_col_idx < len(columns) and 
                        actual_row_idx < len(columns[actual_col_idx])):
                        value = columns[actual_col_idx][actual_row_idx]
                        if value.strip():  # Only add non-empty values
                            col_values[col_header] = value
                
                # Add the column values object to the nested object
                if col_values:
                    nested_obj[row_key] = [col_values]
            
            # Add the nested object under the main key
            if nested_obj:
                result[main_key] = [nested_obj]
    
    # Default case: No titles
    else:
        # Create a simple 2D array representation
        data_array = []
        for i, column in enumerate(columns):
            for j, value in enumerate(column):
                if j >= len(data_array):
                    data_array.append({})
                data_array[j][f"Column {i+1}"] = value
                
        # Store in the result
        result["Data"] = data_array
    
    return result