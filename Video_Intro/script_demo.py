#!/usr/bin/env python3
"""
Sample PDF Processor Script

This script demonstrates how to create a custom processor
that can be loaded by the PDF Processor application.
"""

import os
import json
from Components.Processing.document import create_document_json

# Define extraction parameters - customize these for your specific PDFs
extraction_params = [
    # Basic text extraction with simple start/end keywords
    {
        "field_name": "Test Info General Info",
        "start_keyword": "APPLICANT",
        "end_keyword": "ADDRESS",
        "page_num": 0,
        "horiz_margin": 500,
        "end_keyword_occurrence": 1
    }
]

def process_pdf_file(pdf_path):
    """
    Process a PDF file and create a JSON with extracted data.
    
    Args:
        pdf_path (str): Path to the PDF file
        
    Returns:
        str: Path to the created JSON file or None if processing failed
    """
    # Input validation
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        return None
    
    if not pdf_path.lower().endswith('.pdf'):
        print(f"Error: File is not a PDF: {pdf_path}")
        return None
    
    try:
        # Create JSON from the PDF data using the document module
        json_path = create_document_json(pdf_path, extraction_params)
        
        if json_path:
            print(f"Successfully processed PDF: {pdf_path}")
            print(f"JSON output saved to: {json_path}")
            return json_path
        else:
            print(f"Failed to process PDF: {pdf_path}")
            return None
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return None

# This allows the script to be run directly or imported as a module
if __name__ == "__main__":
    # When run directly, prompt for a PDF file
    pdf_path = input("Enter the path to the PDF file: ")
    process_pdf_file(pdf_path)