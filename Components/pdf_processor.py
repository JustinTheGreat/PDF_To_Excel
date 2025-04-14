"""
PDF Processor - Main module that coordinates PDF data extraction and processing

This module serves as the entry point for PDF processing functionality,
importing all necessary components from their dedicated modules.
"""

import os
import json

# Import from the reorganized structure
from Components.Processing.document import create_document_json
from Components.Processing.Utilities.text import get_bbox_coordinates

def process_pdf(pdf_path, extraction_params):
    """
    Process a PDF file and create a JSON with extracted data.
    
    Args:
        pdf_path (str): Path to the PDF file
        extraction_params (list): List of parameter dictionaries for extraction
        
    Returns:
        str: Path to the created JSON file or None if processing failed
    """
    # Validate inputs
    if not os.path.isfile(pdf_path):
        print(f"Error: PDF file '{pdf_path}' not found.")
        return None
        
    if not extraction_params or not isinstance(extraction_params, list):
        print("Error: Invalid extraction parameters.")
        return None
    
    # Process the PDF file
    try:
        # Create JSON from the PDF data
        json_path = create_document_json(pdf_path, extraction_params)
        
        if json_path:
            print(f"Successfully processed PDF: {pdf_path}")
            print(f"JSON output saved to: {json_path}")
            return json_path
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
    
    return None

# Main execution example
if __name__ == "__main__":
    # Define test PDF path
    pdf_path = "sample.pdf"
    
    # Define extraction parameters for different sections (page_num starts at 0)
    extraction_params = [
        {
            "field_name": "Customer Sheet",
            "start_keyword": "BA:",
            "end_keyword": "FW package",
            "page_num": 0,
            "horiz_margin": 500,
            "end_keyword_occurrence": 1
        },
        {
            "field_name": "Technical Data",
            "start_keyword": "Serial no.:",
            "end_keyword": "[kW]",
            "page_num": 0,
            "horiz_margin": 200,
            "end_keyword_occurrence": 3
        }
    ]
    
    # Process the PDF
    json_path = process_pdf(pdf_path, extraction_params)
    
    if json_path:
        print(f"Successfully created JSON at: {json_path}")
    else:
        print("Failed to process PDF")