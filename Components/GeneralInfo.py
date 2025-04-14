import pdfplumber
import re

def find_keyword_position(words, keyword):
    """
    Find the position of a keyword in the extracted words.
    
    Args:
        words (list): List of word dictionaries from pdfplumber
        keyword (str): Keyword to search for
        
    Returns:
        dict: Position information or None if not found
    """
    for word in words:
        if keyword in word['text']:
            return {
                'x0': word['x0'],
                'y0': word['top'],
                'x1': word['x1'],
                'y1': word['bottom']
            }
    return None

def find_nth_occurrence_position(words, keyword, n):
    """
    Find the position of the nth occurrence of a keyword.
    If fewer than n occurrences are found, returns the last occurrence.
    
    Args:
        words (list): List of word dictionaries from pdfplumber
        keyword (str): Keyword to search for
        n (int): Which occurrence to find (1-based indexing)
        
    Returns:
        dict: Position information or None if not found at all
    """
    occurrences = []
    for word in words:
        if keyword in word['text']:
            occurrences.append({
                'x0': word['x0'],
                'y0': word['top'],
                'x1': word['x1'],
                'y1': word['bottom']
            })
    
    if len(occurrences) > 0:
        # If we have fewer occurrences than requested, return the last one
        if len(occurrences) < n:
            return occurrences[-1]
        else:
            return occurrences[n - 1]
    return None

def extract_serial_data(
    pdf_path,
    # Step 1: Initial Extraction parameters
    page_num=0,
    start_keyword="Serial no.:",
    start_keyword_occurrence=1,  # New parameter - which occurrence of start_keyword to use
    end_keyword=None,
    end_keyword_occurrence=3,
    horiz_margin=100,
    vertical_margin=None,  # New parameter for vertical margin
    left_move=0,
    end_break_line_count=None,
    # Step 2: Basic Formatting parameters
    forced_keywords=None,
    remove_breaks_before=None,
    # Step 3-7: Parsing, Special Formatting, Keyword Processing, Merging, and Chart Processing
    # don't have direct parameters in extract_serial_data
):
    """
    Extracts text data starting from a specified keyword within specified horizontal and vertical boundaries.
    Parameters are ordered according to the processing flow.
    
    Flow:
    1. Initial Extraction - Extract raw text from PDF
    2. Basic Formatting - Format raw text with keywords and line break handling
    3. Parsing to Key-Value - Convert to structured data (handled in outer functions)
    4. Special Field Formatting - Apply field-specific rules (handled in outer functions)
    5. Keyword Processing - Refine data based on keywords (handled in outer functions)
    6. Field Merging - Combine related fields (handled in outer functions)
    7. Chart Processing - Convert to chart format if needed (handled in outer functions)
    8. Final Cleanup - Remove empty values (handled in outer functions)
    9. JSON Structure Creation - Build final output (handled in outer functions)
    
    Args:
        pdf_path (str): Path to the PDF file
        
        # Step 1: Initial Extraction parameters
        page_num (int): Page number to extract from (default: 0 - first page)
        start_keyword (str): Keyword to start extraction from (default: "Serial no.:")
        start_keyword_occurrence (int): Which occurrence of start_keyword to use (default: 1 - first occurrence)
        end_keyword (str, optional): Keyword to end extraction at. If None, only end_break_line_count will be used
        end_keyword_occurrence (int): Which occurrence of end_keyword to stop at (default: 3)
        horiz_margin (float): Horizontal distance in points to expand from start position
        vertical_margin (float, optional): Vertical distance in points to expand from keyword position
        left_move (float): Distance to move the left coordinate leftward (default: 0)
        end_break_line_count (int, optional): Stop extraction after this many newlines
        
        # Step 2: Basic Formatting parameters
        forced_keywords (list, optional): List of keywords to add colons to if missing
        remove_breaks_before (list, optional): List of words to remove line breaks before them
        
    Returns:
        str: Extracted text data
    """
    try:
        # Step 1: Initial Extraction
        # Open the PDF file
        with pdfplumber.open(pdf_path) as pdf:
            # Get the specified page
            if page_num >= len(pdf.pages):
                return f"Error: Page number {page_num} out of range"
            
            page = pdf.pages[page_num]
            
            # Extract page text with bounding boxes for positioning
            words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
            
            # Find starting keyword position - now using the specified occurrence
            start_pos = find_nth_occurrence_position(words, start_keyword, start_keyword_occurrence)
            if not start_pos:
                if start_keyword_occurrence > 1:
                    return f"Occurrence {start_keyword_occurrence} of {start_keyword} not found on page {page_num}"
                else:
                    return f"{start_keyword} not found on page {page_num}"
            
            # If end_keyword is provided, try to find its position
            end_pos = None
            if end_keyword:
                end_pos = find_nth_occurrence_position(words, end_keyword, end_keyword_occurrence)
            
            # Define the initial bounding box
            left = start_pos['x0'] - left_move  # Apply left_move here
            right = left + horiz_margin
            top = start_pos['y0']
            
            # If vertical_margin is specified, adjust the bottom boundary accordingly
            if vertical_margin is not None and vertical_margin > 0:
                bottom = top + vertical_margin
            # Otherwise, use end position or page bottom
            elif end_pos:
                bottom = end_pos['y1'] + 5  # Add small margin
            else:
                # Use the bottom of the page if no end keyword is provided
                bottom = page.height
            
            # Ensure left doesn't go negative
            left = max(0, left)
            
            # Check for negative dimensions and fix if needed
            if bottom < top:
                # Swap top and bottom to ensure positive height
                top, bottom = bottom, top
            
            if right < left:
                # Swap left and right to ensure positive width
                left, right = right, left
            
            # Create the corrected bounding box
            bbox = (left, top, right, bottom)
            
            # Crop the page to our bounding box
            cropped_page = page.crop(bbox)
            
            # Extract text from the cropped area
            extracted_text = cropped_page.extract_text()
            
            # If no text was extracted or end_keyword wasn't found, try a different approach
            if (not extracted_text or (end_keyword and end_keyword not in extracted_text)):
                # Try extracting the whole page text and filtering
                full_text = page.extract_text()
                
                if start_keyword in full_text:
                    # Find the correct occurrence of start_keyword
                    start_idx = -1
                    for _ in range(start_keyword_occurrence):
                        next_idx = full_text.find(start_keyword, start_idx + 1 if start_idx >= 0 else 0)
                        if next_idx == -1:
                            # If we can't find the next occurrence, use the last found one or report error
                            if start_idx >= 0:
                                break
                            else:
                                return f"Occurrence {start_keyword_occurrence} of {start_keyword} not found in text"
                        start_idx = next_idx
                    
                    # If end_keyword is provided, use it to extract between keywords
                    if end_keyword and end_keyword in full_text[start_idx:]:
                        # Find the correct occurrence of end_keyword
                        curr_idx = start_idx
                        end_idx = -1
                        for _ in range(end_keyword_occurrence):
                            next_idx = full_text.find(end_keyword, curr_idx)
                            if next_idx == -1:
                                # If we can't find the next occurrence, use the last found one
                                break
                            end_idx = next_idx
                            curr_idx = end_idx + len(end_keyword)
                        
                        if end_idx > start_idx:
                            # Extract from start to end (including the end keyword)
                            extracted_text = full_text[start_idx:end_idx + len(end_keyword)]
                    else:
                        # Just extract from start to the end of the page
                        extracted_text = full_text[start_idx:]
            
            # Apply end_break_line_count limit if specified
            if end_break_line_count is not None and extracted_text:
                extracted_text = limit_by_newline_count(extracted_text, end_break_line_count)
            
            return extracted_text or "No text found in the specified area"
            
    except Exception as e:
        # If the error is specifically about negative width/height, try alternate extraction
        if "negative width or height" in str(e):
            try:
                # Try extracting the whole page text and filtering
                full_text = page.extract_text()
                lines = full_text.split('\n')
                
                # Find lines containing our keywords
                start_line_idx = -1
                end_line_idx = -1
                
                # Find the specified occurrence of start_keyword
                start_keyword_count = 0
                for i, line in enumerate(lines):
                    if start_keyword in line:
                        start_keyword_count += 1
                        if start_keyword_count == start_keyword_occurrence:
                            start_line_idx = i
                            break
                
                # Find the specified occurrence of end_keyword
                if end_keyword:
                    end_keyword_count = 0
                    for i, line in enumerate(lines[start_line_idx:], start=start_line_idx):
                        if end_keyword in line:
                            end_keyword_count += 1
                            if end_keyword_count == end_keyword_occurrence:
                                end_line_idx = i
                                break
                
                # If we found a start line
                if start_line_idx != -1:
                    # If we also found an end line, extract between them
                    if end_line_idx != -1 and end_line_idx >= start_line_idx:
                        result = '\n'.join(lines[start_line_idx:end_line_idx+1])
                    else:
                        # Just extract from start keyword to end of page
                        result = '\n'.join(lines[start_line_idx:])
                    
                    # Apply end_break_line_count limit if specified
                    if end_break_line_count is not None:
                        result = limit_by_newline_count(result, end_break_line_count)
                    
                    return result
                
                return f"Fallback extraction failed: Keywords not found in text"
            except Exception as fallback_error:
                return f"Error in fallback extraction: {str(fallback_error)}"
        
        return f"Error extracting data: {str(e)}"

def limit_by_newline_count(text, max_newlines):
    """
    Limit text extraction after encountering a specified number of newline characters,
    regardless of whether they are consecutive or not.
    
    Args:
        text (str): The extracted text
        max_newlines (int): Maximum number of newlines before stopping extraction
        
    Returns:
        str: The limited text
    """
    if not text or max_newlines is None:
        return text
    
    # Count newlines as we go through the text
    newline_count = 0
    result = ""
    
    # Read character by character
    for char in text:
        result += char
        if char == '\n':
            newline_count += 1
            
            # Check if we've reached the limit
            if newline_count >= max_newlines:
                break
    
    return result