# JSON to Excel Converter

A robust Python application that converts JSON data files into well-formatted Excel workbooks with advanced formatting and filtering capabilities.

## Overview

This application processes JSON files from a selected directory (including subdirectories if specified) and converts them into a structured Excel workbook. It intelligently analyzes the JSON structure and creates appropriate worksheet layouts with headers, subtitles, and data columns. The application handles complex nested structures, key-value pairs, units, and special data types like dates and numbers.

## Features

- **User-friendly GUI interface** with directory selection and conversion options
- **Intelligent JSON structure analysis** to determine optimal Excel layout
- **Support for nested data structures** including arrays and key-value pairs
- **Multi-sheet Excel output** organized by JSON document titles
- **Advanced data processing**:
  - Filter specific text from filenames
  - Remove units from values (e.g., [ms], [V])
  - Replace commas with periods in numeric values
  - Date detection and formatting
  - Number format standardization
- **Recursive directory scanning** for batch processing
- **Debug mode** for detailed logging

## Getting Started

### Prerequisites

- Python 3.6 or higher
- Required packages:
  - tkinter (GUI)
  - openpyxl (Excel generation)

### Installation

1. Clone this repository or download the source code
2. Install required dependencies:

```bash
pip install openpyxl
```

Note: tkinter is included in most Python installations, but if needed:

```bash
# For Debian/Ubuntu
sudo apt-get install python3-tk

# For Fedora
sudo dnf install python3-tkinter

# For macOS (using Homebrew)
brew install python-tk
```

### Running the Application

Launch the application using:

```bash
python excel_main.py
```

For debug mode with detailed logging:

```bash
python excel_main.py --debug
```

### Usage Instructions

1. **Select JSON Files Directory**: Choose the directory containing your JSON files.
2. **Select Output Directory**: Choose where to save the generated Excel file.
3. **Configure Options**:
   - **Output File Name**: Set a name for the Excel file.
   - **Filter Text to Remove**: Specify text to remove from filenames.
   - **Remove units**: Toggle removal of units from values.
   - **Replace commas with periods**: Toggle decimal separator standardization.
   - **Search in subdirectories**: Toggle whether to process files in subdirectories.
4. **Process JSON Files**: Click to start the conversion process.
5. **View Progress**: Monitor the progress bar and status updates.

## Excel Formatting

The application applies sophisticated formatting to make the Excel output clear and well-organized:

### Sheet Organization
- Each unique report title from the JSON data gets its own worksheet
- Sheet names are sanitized to comply with Excel's requirements

### Header Formatting
- Column headers use bold text with a light gray background
- Headers for nested data are merged across multiple columns
- Multi-level nested data generates appropriate subtitle rows

### Data Formatting
- **Filenames**: The first column contains the processed filenames (with extensions and filtered text removed)
- **Dates**: Automatically detected and formatted using Excel's date format
- **Numbers**: 
  - Properly converted from string to actual numeric values
  - Comma decimal separators optionally converted to periods
  - Displayed with appropriate number formatting
- **Lists**: 
  - Simple lists are displayed across multiple columns
  - Key-value lists create subtitled columns with the keys as subtitles
  - Nested lists generate hierarchical column structures

### Visual Enhancements
- Column widths are automatically adjusted based on content
- Cells have light borders for better readability
- Headers and subtitles are centered when spanning multiple columns

### Special Processing
- Unit notations like [ms], [V], etc. are optionally stripped from values
- Number-separated values (e.g., "123 & 456") can be split into separate values
- Single-item lists in nested structures are flattened for cleaner display

## Project Structure

The application is organized in a modular structure for maintainability:

```
├── excel_main.py                 # Main entry point
├── Components/                   # Main components directory
│   ├── app_gui.py               # GUI implementation
│   ├── excel/                   # Excel generation components
│   │   ├── formatter.py         # Excel formatting utilities
│   │   ├── generator.py         # Excel file generation logic
│   │   └── data_writer.py       # Excel data writing utilities
│   ├── json/                    # JSON processing components
│   │   ├── analyzer.py          # JSON structure analysis
│   │   ├── reader.py            # JSON file reading utilities
│   │   ├── processor.py         # JSON processing facade
│   │   └── structure_analyzer.py # Supplementary structure analysis
│   └── utils/                   # Common utilities
│       ├── business_rules.py    # Business logic for data transformation
│       ├── text_filters.py      # Text filtering and processing utilities
│       └── file_utils.py        # File handling utilities
```

### Key Components

- **app_gui.py**: Implements the user interface with tkinter
- **generator.py**: Orchestrates the Excel creation process
- **analyzer.py**: Analyzes JSON structure to determine Excel formatting
- **business_rules.py**: Contains business-specific data transformation rules
- **data_writer.py**: Handles writing data to Excel with proper formatting

## Extending the Application

### Adding New Units to Filter

To add support for additional units, modify the `remove_units` method in `Components/utils/text_filters.py`.

### Adding New Business Rules

Add custom data transformations in `Components/utils/business_rules.py`. The existing methods demonstrate the pattern to follow.

### Supporting New Data Types

Extend the data detection and conversion logic in `Components/excel/data_writer.py`.

## License

MIT License

## Acknowledgments

I would like to acknowledge the Javaid Baksh for being such a great engineer to work with to push the program to it's limits