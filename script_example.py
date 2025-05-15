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
        "field_name": "General Info",
        "start_keyword": "Customer Information:",
        "end_keyword": "Equipment Details",
        "page_num": 0,
        "horiz_margin": 300,
        "end_keyword_occurrence": 1
    },
    
    # Using vertical margin to limit extraction height
    {
        "field_name": "Equipment Specs",
        "start_keyword": "Equipment Details",
        "end_keyword": None,  # No end keyword
        "page_num": 0,
        "horiz_margin": 350,
        "vertical_margin": 200,  # Limit vertical extraction to 200 points
        "left_move": 20  # Shift extraction box 20 points to the left
    },
    
    # Using forced keywords to help with parsing
    {
        "field_name": "Technical Parameters",
        "start_keyword": "Parameters",
        "end_keyword": "Test Results",
        "page_num": 1,
        "horiz_margin": 400,
        "forced_keywords": ["Serial Number", "Model", "Voltage", "Current", "Power"]
    },
    
    # Using line break handling for better formatting
    {
        "field_name": "Test Results",
        "start_keyword": "Test Results",
        "end_keyword": "Certification",
        "page_num": 1,
        "horiz_margin": 450,
        "remove_breaks_before": ["passed", "failed", "warning"],
        "remove_breaks_after": ["Test:", "Result:"]
    },
    
    # Using limited line breaks instead of end keyword
    {
        "field_name": "Certification",
        "start_keyword": "Certification",
        "end_keyword": None,  # No end keyword
        "page_num": 1,
        "horiz_margin": 350,
        "end_break_line_count": 10  # Stop after 10 line breaks
    },
    
    # Field with continuation on another page
    {
        "field_name": "Maintenance Records",
        "start_keyword": "Maintenance History",
        "end_keyword": "End of Records",
        "page_num": 2,
        "horiz_margin": 500,
    },
    
    # Continuation field (will be merged with the base field)
    {
        "field_name": "Maintenance Records(+1)",
        "start_keyword": "Continued from previous page",
        "end_keyword": "Recommendations",
        "page_num": 3,
        "horiz_margin": 500,
    },
    
    # Chart field with specific chart parameters
    {
        "field_name": "Performance Data(Chart)",
        "start_keyword": "Performance Metrics",
        "end_keyword": "Graph Data End",
        "page_num": 4,
        "horiz_margin": 450,
        "top_title": True,  # Use top row as column titles
        "left_title": True,  # Use left column as row titles
        "priority_side": "left"  # Make left titles the primary grouping
    },
    
    # Table processing with customized structure
    {
        "field_name": "Components List",
        "start_keyword": "Component Inventory",
        "end_keyword": "Inventory End",
        "page_num": 5,
        "horiz_margin": 500,
        "table_top_labeling": True,
        "table_left_labeling": True,
        "table_structure": "top_main",
        "min_column_width": 5
    },
    
    # Using specific occurrence of keywords for targeted extraction
    {
        "field_name": "Additional Notes",
        "start_keyword": "Notes:",
        "start_keyword_occurrence": 2,  # Use the 2nd occurrence of "Notes:"
        "end_keyword": "End of Notes",
        "end_keyword_occurrence": 1,
        "page_num": 6,
        "horiz_margin": 300
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