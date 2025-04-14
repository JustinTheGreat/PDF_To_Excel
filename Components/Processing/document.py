"""
Document processing module for PDF to JSON conversion.

This module provides the main functionality for creating structured
JSON documents from extracted PDF data.
"""

import os
import json
from Components.Processing.Core.extraction import extract_pdf_data
from Components.Processing.Utilities.merger import process_field_merging
from Components.Processing.chart_processor import process_chart_data  # Import the chart processor


def create_document_json(pdf_path, extraction_params):
    """
    Process a PDF file, extract document details and additional data,
    and create a JSON file with the extracted data. Supports merging
    fields marked with (+1) suffix and chart conversion for fields
    marked with (Chart) suffix.
    
    Args:
        pdf_path (str): Path to the PDF file
        extraction_params (list): List of parameter dictionaries for extraction
        
    Returns:
        str: Path to the created JSON file
    """
    # Check if file exists
    if not os.path.isfile(pdf_path):
        print(f"Error: File '{pdf_path}' not found.")
        return None
        
    # Extract data based on parameters
    extracted_data = extract_pdf_data(pdf_path, extraction_params)
    
    # Process field merging for fields with (+1) suffix
    merged_data = process_field_merging(extracted_data)
    
    # Process chart conversion for fields with (Chart) suffix
    # Chart fields will be merged into their base field names
    # Pass the original extraction_params to help with parameter lookup
    processed_data = process_chart_data(merged_data, extraction_params)
    
    # Prepare the JSON structure
    json_data = []
    
    # Add all extracted data to JSON
    for field_name, content_dict in processed_data.items():
        # Skip any internal tracking fields
        has_chart_data = content_dict.pop("has_chart_data", False)
        was_chart = content_dict.pop("was_chart", False)
        
        # Post-process to ensure no duplicate keys in the final output and no empty values
        final_fields = {}
        for key, value in content_dict["parsed_data"].items():
            # Skip any empty values
            if value == "" or value is None:
                continue
                
            if isinstance(value, list) and all(v == "" or v is None for v in value):
                continue
                
            # Handle duplicates by merging into lists
            if key in final_fields:
                if isinstance(final_fields[key], list):
                    if isinstance(value, list):
                        final_fields[key].extend([v for v in value if v != "" and v is not None])
                    elif value != "" and value is not None:
                        final_fields[key].append(value)
                else:
                    if isinstance(value, list):
                        non_empty_values = [v for v in value if v != "" and v is not None]
                        if non_empty_values:
                            final_fields[key] = [final_fields[key]] + non_empty_values
                    elif value != "" and value is not None and final_fields[key] != value:
                        final_fields[key] = [final_fields[key], value]
            else:
                # Filter out empty values from lists
                if isinstance(value, list):
                    non_empty_values = [v for v in value if v != "" and v is not None]
                    if non_empty_values:
                        final_fields[key] = non_empty_values if len(non_empty_values) > 1 else non_empty_values[0]
                else:
                    final_fields[key] = value
        
        data_entry = {
            "title": field_name,
            "raw_text": content_dict["raw_text"],
            "formatted_text": content_dict["formatted_text"],
            "fields": final_fields
        }
        
        json_data.append(data_entry)
    
    # Create JSON file name based on the PDF name
    pdf_filename = os.path.basename(pdf_path)
    pdf_name_without_ext = os.path.splitext(pdf_filename)[0]
    json_filename = f"{pdf_name_without_ext}.json"
    json_path = os.path.join(os.path.dirname(pdf_path), json_filename)
    
    # Save JSON data to file
    with open(json_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=2)
    
    print(f"JSON file created: {json_path}")
    return json_path