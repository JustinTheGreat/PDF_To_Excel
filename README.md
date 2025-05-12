# PDF Extraction System (To JSON): Usage Guide

This guide focuses on using the PDF Extraction System to extract structured data from PDF documents and explains how both the command-line interface (`main.py`) and the graphical user interface (`directory_processor.py`) work.

## System Overview

This PDF Extraction System provides a variety of processes to help extract info from PDF files:

1. **Visual GUI used to process PDFs** (`main.py`) - Process Group PDF files
2. **Preview your script finding visually** ('Previewer_main.py) - Preview your script files that you use on your reports
3. **Convert JSON to Excel Files** ('PDF_To_Excel/Excel_main.py) - Convert your json files to excel files

## Getting Started

### Installation and Setup

Ensure you have Python installed on your system (Python 3.6 or later recommended).

Install the required dependencies:
```
pip install pdfplumber tkinter
```

## Using the Graphical Interface (main.py)

The GUI application allows you to scan directories for PDF files containing specific keywords and batch process them.

### Running the Application

```bash
python main.py
```

This launches the PDF Processor application.

### Step-by-Step Usage Guide

1. **Configure Input and Output**:
   - **Input Directory**: Click "Browse..." to select the directory containing PDF files to process.
   - **Output Directory**: Click "Browse..." to select where processed files should be saved.
   - **Subfolder Options**: 
     - Check "Create subfolder for output" to save results in a dedicated subfolder.
     - Enter a name for this subfolder (default: "Processed_Files").
   - **Search Keyword**: Enter the keyword to search for in PDF files (default: "BA:").
   - **Processor Script**: Click "Browse..." to select the Python script that will process the PDFs.

2. **Scan for Files**:
   - Click the "Scan Directory" button to find PDFs containing your keyword.
   - Files matching the criteria will appear in the "Found Files" section.
   - Use "Select All" or "Deselect All" to manage the file selection.

3. **Process Files**:
   - Select the files you want to process from the list.
   - Click "Process Selected Files" to begin processing.
   - The Processing Log will show progress and results.
   - Processed JSON files will be saved to your specified output location.

## Creating Custom Processor Scripts

You can create your own processor scripts that extract specific information from PDFs. Here's a basic template:

```python
import os
from Components.Processing.document import create_document_json

# Define extraction parameters
extraction_params = [
    {
        "field_name": "Technical Data",
        "start_keyword": "Serial no.:",
        "end_keyword": "[kW]",
        "page_num": 0,
        "horiz_margin": 200,
        "end_keyword_occurrence": 3
    },
    # Add more parameter sets as needed
]

def process_pdf_file(pdf_path):
    """
    Process a PDF file and create a JSON with extracted data.
    
    Args:
        pdf_path (str): Path to the PDF file
        
    Returns:
        str: Path to the created JSON file or None if processing failed
    """
    # Create JSON from the PDF data
    json_path = create_document_json(pdf_path, extraction_params)
    
    if json_path:
        print(f"Successfully processed PDF: {pdf_path}")
        print(f"JSON output saved to: {json_path}")
        return json_path
    else:
        print(f"Failed to process PDF: {pdf_path}")
        return None

# This allows the script to be run directly or imported as a module
if __name__ == "__main__":
    # When run directly, prompt for a PDF file
    pdf_path = input("Enter the path to the PDF file: ")
    process_pdf_file(pdf_path)
```

## Command-line Interface (main.py)

For processing individual files, you can use the `main.py` script:

```python
import os
import json
from Components.pdf_processor import create_document_json

if __name__ == "__main__":
    # Get PDF file path from user
    pdf_path = input("Enter the path to the PDF file: ")
    pdf_path = pdf_path.strip('"\'')
    
    # Define extraction parameters
    extraction_params = [
        {
            "field_name": "General Info",
            "start_keyword": "Customer Information:",
            "end_keyword": "Equipment Details",
            "page_num": 0,
            "horiz_margin": 300
        }
    ]
    
    # Process the PDF
    json_path = create_document_json(pdf_path, extraction_params)
    
    # Display the results
    if json_path:
        with open(json_path, 'r', encoding='utf-8') as file:
            content = json.load(file)
            print("\nJSON Content:")
            print(json.dumps(content, indent=2))
    else:
        print("Failed to create JSON file.")
```

## Default Auto-Formatting

The system automatically performs these actions without additional parameters:

1. **Key-Value Detection**: Identifies "Key: Value" patterns and structures them
2. **List Conversion**: Single values are kept as strings, multiple values become arrays
3. **Empty Value Removal**: Keys with empty values are automatically removed
4. **Line Break Normalization**: Ensures consistent line breaks in the formatted text
5. **Duplicate Key Handling**: Merges duplicate keys into arrays
6. **Whitespace Trimming**: Removes extra whitespace from keys and values

## Available Extraction Parameters

All available parameters that can be used in your extraction definition:

| Parameter | Type | Description |
|-----------|------|-------------|
| `field_name` | string | **Required**. Name for the extracted field (use (+1) suffix for continuation fields and (Chart) suffix for chart conversion) |
| `start_keyword` | string | **Required**. Keyword to start extraction from |
| `end_keyword` | string or null | Keyword to end extraction at (can be null if using line breaks) |
| `page_num` | integer | Page number to extract from (0-based, default: 0) |
| `horiz_margin` | float | Horizontal distance to expand from start position (default: 200) |
| `start_keyword_occurrence` | integer | Which occurrence of start_keyword to use (default: 1) |
| `end_keyword_occurrence` | integer | Which occurrence of end_keyword to stop at (default: 1) |
| `vertical_margin` | float or null | Vertical distance to expand from keyword position |
| `left_move` | float | Distance to move the left coordinate leftward (default: 0) |
| `end_break_line_count` | integer or null | Stop extraction after this many newlines |
| `forced_keywords` | array or null | List of keywords to add colons to if missing (Make sure it's a list data type even if it's single item ("[]")) |
| `remove_colon_after` | array or null | List of keywords to remove colons after if not intended (Make sure it's a list data type even if it's single item ("[]")) |
| `remove_breaks_before` | array or null | List of words to remove line breaks ("\n") before them (Make sure it's a list data type even if it's single item ("[]"))|
| `remove_breaks_after` | array or null | List of words to remove line breaks ("\n") after them (Make sure it's a list data type even if it's single item ("[]"))|

### Chart-Specific Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `top_title` | boolean | Use top row as column titles for chart processing |
| `left_title` | boolean | Use left column as row titles for chart processing |
| `priority_side` | string | Primary grouping for chart ('top' or 'left') |

### Table-Specific Parameters (Not recommended)

| Parameter | Type | Description |
|-----------|------|-------------|
| `table_top_labeling` | boolean | Use top row as keys for table processing |
| `table_left_labeling` | boolean | Use left column as keys for table processing |
| `table_structure` | string | Structure type for table ('top_only', 'left_only', 'top_main', 'left_main') |
| `min_column_width` | integer | Minimum width for columns when using space delimiter |

## How the Extraction Process Works

The PDF extraction system follows this pipeline:

### 1. Initial Extraction
- Locates the `start_keyword` on the specified page
- Creates a bounding box extending horizontally by `horiz_margin`
- Extracts text within this area until the `end_keyword` or other limits

   ```
    <--Horizontal Margin Manually Set-->
   ┌─────────────────────────────────┐
   │ Start Keyword                   │
   │ Content to extract...           │
   │ ...                             │
   │ ...                             │
   │ End Keyword                     │
   └─────────────────────────────────┘
   ```

### 2. Text Formatting
- Adds colons to `forced_keywords` if they're missing
- Removes line breaks before/after specified words
- Prepares the text for structured parsing

### 3. Key-Value Parsing
- Identifies key-value pairs based on colons (e.g., "Name: John")
- Creates a dictionary with these pairs
- Collects unparsed lines for further processing

### 4. Field Merging
- Combines fields marked with (+1) suffix (e.g., "Report" and "Report(+1)")
- Maintains proper structure for duplicate keys

### 5. Chart Processing
- Converts fields marked with (Chart) suffix to structured format
- Organizes data based on chart parameters (top_title, left_title, etc.)

### 6. JSON Structure Creation
- Creates the final JSON structure with:
  - Field title
  - Raw extracted text
  - Formatted text
  - Structured field data

## Extraction Parameter Examples

### Two-Pass Approach for Conditional Parameters

This approach allows you to first extract document metadata and then apply different extraction parameters based on the content. For example, applying different formatting rules based on document version:

```python
import os
import json
import re
from Components.pdf_processor import create_document_json
from Components.GeneralInfo import extract_serial_data

if __name__ == "__main__":
    # Get PDF file path
    pdf_path = input("Enter the path to the PDF file: ")
    pdf_path = pdf_path.strip('"\'')
    
    # FIRST PASS: Extract document details to check version
    doc_details_text = extract_serial_data(
        pdf_path,
        start_keyword="Documentnumber",
        end_keyword="Version:",
        page_num=0,
        horiz_margin=200
    )
    
    # Check if version 06 is present
    is_version_06 = False
    if doc_details_text and re.search(r'Version:\s*06\b', doc_details_text):
        is_version_06 = True
        print("Version 06 detected. Using specific parameters.")
    
    # SECOND PASS: Extract with version-specific parameters
    extraction_params = [
        # Common parameters for all versions
        {
            "field_name": "Document Details",
            "start_keyword": "Documentnumber",
            "end_keyword": "Version:",
            "page_num": 0,
            "horiz_margin": 200
        }
    ]
    
    # Add version-specific parameters
    safety_params = {
        "field_name": "Safety Data",
        "start_keyword": "Safety tests",
        "end_keyword": "Test Results",
        "page_num": 1,
        "horiz_margin": 300
    }
    
    # Only include forced_keywords for version 06
    if is_version_06:
        safety_params["forced_keywords"] = ["High voltage test II"]
    
    extraction_params.append(safety_params)
    
    # Process the PDF with the conditional parameters
    json_path = create_document_json(pdf_path, extraction_params)
    
    if json_path:
        print("Created JSON File Successfully")
    else:
        print("Failed to create JSON file.")
```

### Basic Text Extraction

```python
{
    "field_name": "General Info",
    "start_keyword": "Customer Information:",
    "end_keyword": "Equipment Details",
    "page_num": 0,
    "horiz_margin": 300
}
```

### Using Vertical Margin

```python
{
    "field_name": "Equipment Specs",
    "start_keyword": "Equipment Details",
    "end_keyword": None,  # No end keyword
    "page_num": 0,
    "horiz_margin": 350,
    "vertical_margin": 200  # Limit vertical extraction to 200 points
}
```

### Using Forced Keywords

```python
{
    "field_name": "Technical Parameters",
    "start_keyword": "Parameters",
    "end_keyword": "Test Results",
    "page_num": 1,
    "horiz_margin": 400,
    "forced_keywords": ["Serial Number", "Model", "Voltage", "Current", "Power"]
}
```

### Line Break Handling

```python
{
    "field_name": "Test Results",
    "start_keyword": "Test Results",
    "end_keyword": "Certification",
    "page_num": 1,
    "horiz_margin": 450,
    "remove_breaks_before": ["passed", "failed", "warning"],
    "remove_breaks_after": ["Test:", "Result:"]
}
```

### Limited Line Breaks

```python
{
    "field_name": "Certification",
    "start_keyword": "Certification",
    "end_keyword": None,  # No end keyword
    "page_num": 1,
    "horiz_margin": 350,
    "end_break_line_count": 10  # Stop after 10 line breaks
}
```

### Multi-Page Content with Continuation

```python
# First page
{
    "field_name": "Maintenance Records",
    "start_keyword": "Maintenance History",
    "end_keyword": "End of Records",
    "page_num": 2,
    "horiz_margin": 500
},
# Continuation on next page
{
    "field_name": "Maintenance Records(+1)",
    "start_keyword": "Continued from previous page",
    "end_keyword": "Recommendations",
    "page_num": 3,
    "horiz_margin": 500
}
```

### Chart Data Processing

```python
{
    "field_name": "Performance Data(Chart)",
    "start_keyword": "Performance Metrics",
    "end_keyword": "Graph Data End",
    "page_num": 4,
    "horiz_margin": 450,
    "top_title": True,  # Use top row as column titles
    "left_title": True,  # Use left column as row titles
    "priority_side": "left"  # Make left titles the primary grouping
}
```

### Table Processing

```python
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
}
```

## Tips for Effective Extraction

1. **Start broad, then narrow down**: Begin with wider margins and refine based on results
2. **Use visual inspection**: Check PDF layouts to identify start/end keywords
3. **Try multiple approaches**: If one set of parameters doesn't work, try different keywords or boundaries
4. **Handle multi-page content**: Use the (+1) suffix for content that spans pages
5. **Forced keywords**: Use when a PDF has inconsistent formatting
6. **Line break handling**: Use when text layout affects parsing
7. **Two-pass approach**: When you need to apply different parameters based on document content

## Troubleshooting

### GUI Application Issues

If you encounter issues with the directory processor application:

1. **Processing Log Not Showing**: If log messages aren't appearing, try resizing the application window
2. **Duplicate JSON Files**: The application is designed to create JSON files only in the specified output directory
3. **File Access Errors**: Ensure you have write permission in both the source and output directories

### Extraction Issues

If extraction isn't working as expected:

1. Check that the `start_keyword` exists exactly as specified in the PDF
2. Verify the page number (remember page numbering starts at 0)
3. Try increasing the `horiz_margin` to capture more content
4. Use `start_keyword_occurrence` if the keyword appears multiple times
5. For content spanning pages, ensure the continuation field uses the (+1) suffix
6. Check for invisible characters or formatting that might affect keyword matching

*If you encounter persistent issues, consider using AI tools like Claude.ai to help develop specific extraction parameters for your documents.*
