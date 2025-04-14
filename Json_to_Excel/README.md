# JSON to Excel Converter

## Overview

The JSON to Excel Converter is a desktop application that allows you to convert multiple JSON files into a structured Excel workbook. It's particularly useful for analyzing reports that are stored in JSON format but need to be viewed and manipulated in Excel.

## Features

- Convert multiple JSON files into a single Excel workbook
- Combine data from files with the same report title into single worksheets
- Support for nested data structures and lists
- Filter out unwanted units from values (e.g., [ms], [V])
- Remove common text from filenames
- Option to search in subdirectories
- Debug mode for troubleshooting

## File Structure and Functionality

The application consists of several Python modules, each with specific responsibilities:

### excel_main.py
- Run this script to start the program

### app_gui.py

This module provides the graphical user interface for the application. It:
- Creates the main window with all UI elements
- Handles user interactions (button clicks, directory selection)
- Manages the progress display and status updates
- Coordinates the conversion process

### json_processor.py

This module is responsible for processing JSON files. It:
- Reads JSON files from directories and subdirectories
- Parses the JSON data structure
- Analyzes the structure of reports to determine how to format the Excel output
- Processes filenames to remove unwanted text

### excel_generator.py

This module generates the Excel output from the processed JSON data. It:
- Creates a new Excel workbook
- Sets up worksheets based on report titles
- Formats headers and cells with appropriate styles
- Handles lists of varying lengths to ensure consistent column alignment
- Adjusts column widths for better readability
- Manages data that should be combined across multiple files

### text_filters.py

This module provides text filtering capabilities. It:
- Removes units from values (e.g., [ms], [V])
- Cleanses numeric values
- Processes different types of values (strings, lists, dictionaries)
- Applies custom replacements to text

## Using the Application

1. **Select JSON Files Directory**: Choose the directory containing your JSON files.
2. **Select Output Directory**: Choose where to save the Excel file.
3. **Specify Output Filename**: Enter a name for the output Excel file.
4. **Configure Options**:
   - Filter Text to Remove: Remove common text from filenames
   - Remove units: Clean values by removing unit notations
   - Search in subdirectories: Include files from nested folders
5. **Process Files**: Click the "Process JSON Files" button to start the conversion.

## Data Handling

### JSON Structure

The application expects JSON files to contain report data. Each report should ideally have:
- A `title` field identifying the report type
- A `fields` section containing the data to be processed

However, it can also handle files without this structure by treating the entire JSON object as fields.

### Worksheets Organization

- Reports with the same title are combined into a single worksheet
- Each row represents data from one JSON file
- The first column shows the source filename
- Header rows identify each data field
- For fields with multiple values (lists), subtitles show item numbers

### List Handling

- Lists of values are displayed across multiple columns
- When lists have different lengths, the application maintains column alignment
- Missing values in shorter lists result in empty cells
- This ensures data integrity and consistent structure

## Debugging

When run in debug mode (`python excel_main.py --debug`), the application:
- Shows a detailed log of the processing steps
- Displays file counts and structures
- Reports any errors or warnings
- Helps identify issues with JSON files or data structure

## Technical Considerations

- The application uses threading to prevent the UI from freezing during processing
- Error handling is comprehensive to prevent crashes
- Memory usage is optimized for handling multiple files
- Column widths are automatically adjusted based on content