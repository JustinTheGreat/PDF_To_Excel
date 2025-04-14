"""
Text utility functions for PDF processing.

This module provides utilities for extracting and manipulating text 
from PDF documents.
"""

import pdfplumber

def get_bbox_coordinates(pdf_path, extraction_params=None):
    """
    Get bounding box coordinates for specified fields in a PDF based on extraction parameters.
    
    Args:
        pdf_path (str): Path to the PDF file
        extraction_params (list): List of parameter dictionaries for extraction,
                                 following the same format as in main.py.
        
    Returns:
        dict: Dictionary containing bounding box information for each field or key
    """
    bbox_data = {}
    
    # Validate input
    if not extraction_params:
        raise ValueError("extraction_params must be provided")
    
    try:
        # Open the PDF file
        with pdfplumber.open(pdf_path) as pdf:
            # Iterate through each page in the PDF
            for page_num, page in enumerate(pdf.pages):
                # Extract words with their bounding boxes
                words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
                
                # Extract keywords from extraction parameters
                keywords = []
                for param in extraction_params:
                    # Add start and end keywords from each parameter set
                    if 'start_keyword' in param:
                        keywords.append(param['start_keyword'])
                    if 'end_keyword' in param and param['end_keyword'] != param.get('start_keyword'):
                        keywords.append(param['end_keyword'])
                
                # Find positions of all specified keywords
                for keyword in keywords:
                    # Find all occurrences of the keyword
                    keyword_positions = []
                    
                    for i, word in enumerate(words):
                        if keyword.lower() in word["text"].lower():
                            # Get the position of the keyword
                            position = {
                                "page": page_num,
                                "x0": word["x0"],
                                "top": word["top"],
                                "x1": word["x1"],
                                "bottom": word["bottom"],
                                "text": word["text"]
                            }
                            
                            # Try to get the value if it's on the same line or next to the keyword
                            value = ""
                            value_bbox = None
                            
                            # Check if there are words to the right on the same line
                            same_line_words = [w for w in words if abs(w["top"] - word["top"]) < 5 and w["x0"] > word["x1"]]
                            if same_line_words:
                                # Take the first word to the right
                                value = same_line_words[0]["text"]
                                value_bbox = {
                                    "x0": same_line_words[0]["x0"],
                                    "top": same_line_words[0]["top"],
                                    "x1": same_line_words[0]["x1"],
                                    "bottom": same_line_words[0]["bottom"]
                                }
                            
                            # Add the value and its position if found
                            position["value"] = value
                            position["value_bbox"] = value_bbox
                            
                            # Add to the list of positions for this keyword
                            keyword_positions.append(position)
                    
                    # Add all positions found for this keyword
                    if keyword_positions:
                        # Use a compound key with keyword and page to avoid overwriting
                        key = f"{keyword}_{page_num}"
                        bbox_data[key] = keyword_positions
    
    except Exception as e:
        print(f"Error extracting bounding box coordinates: {str(e)}")
    
    return bbox_data