# PDF Processor

A modular PDF data extraction and processing system that extracts structured data from PDF documents based on configured extraction parameters.

## Project Structure

```
project_root/
├── pdf_processor.py                    # Main entry point
├── Components/                         # Original components folder
│   ├── __init__.py
│   ├── pdf_extractor.py                # Existing file
│   ├── GeneralInfo.py                  # Existing file
│   └── business_rules.py               # Existing file
└── Processing/                         # New processing modules folder
    ├── __init__.py
    ├── document.py                     # Document handling
    ├── Core/                           # Core processing functionality
    │   ├── __init__.py
    │   └── extraction.py               # Core extraction logic
    ├── Parsers/                        # Text and data parsing modules
    │   ├── __init__.py
    │   ├── table.py                    # Table parsing
    │   └── keywords.py                 # Keyword handling
    └── Utilities/                      # Helper utilities
        ├── __init__.py
        ├── text.py                     # Text extraction utilities
        ├── cleaner.py                  # Data cleaning utilities
        └── merger.py                   # Field merging utilities
    └── Documentation/                  # Documentation on the codebase 
```

## Module Descriptions

### Main Module

- **pdf_processor.py**: The main entry point for PDF processing. Provides a simplified interface to all PDF processing functionality.

### Components (Original)

Contains original components that handle core PDF reading and business-specific formatting:

- **pdf_extractor.py**: Handles the initial extraction and formatting of PDF text
- **GeneralInfo.py**: Extracts serial data and manages keyword positioning
- **business_rules.py**: Applies business-specific formatting rules

### Processing (New Modular Structure)

#### Core

Core processing modules that coordinate the extraction process:

- **extraction.py**: Main extraction functionality that coordinates the use of parsers and utilities
- **document.py**: Handles creating structured JSON documents from extracted data

#### Parsers

Specialized parsers for different data formats:

- **table.py**: Handles detection and extraction of tabular data from PDF text
- **keywords.py**: Provides functionality for managing start/end keywords in text extraction

#### Utilities

Helper utilities for processing:

- **text.py**: Text extraction utilities for PDF processing
- **cleaner.py**: Data cleaning utilities for removing empty values
- **merger.py**: Field merging utilities for combining related fields

## Usage Example

```python
from pdf_processor import process_pdf

# Define PDF path
pdf_path = "sample.pdf"

# Define extraction parameters
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
```

## Benefits of the Structure

This modular approach offers several advantages:

1. **Improved maintainability** - Each module has a clear, specific responsibility
2. **Better testability** - Functions can be tested in isolation
3. **Enhanced readability** - Smaller, focused files are easier to understand
4. **Easier extensibility** - New features can be added by extending specific modules
5. **Better collaboration** - Multiple developers can work on different modules simultaneously

## Dependencies

- pdfplumber
- re
- os
- json
