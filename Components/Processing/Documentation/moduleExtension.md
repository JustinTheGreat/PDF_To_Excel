# Module Extension Patterns

This document outlines common patterns for extending the PDF processing framework. Understanding these patterns will help when requesting new features from an AI assistant, ensuring that new code maintains the architectural integrity of the system.

## 1. New Parser Implementation Pattern

When adding a new parser for a specific type of content within PDFs:

### File Structure
```
Processing/
└── Parsers/
    └── new_parser.py
```

### Key Implementation Elements
- Parser function with clear input/output documentation
- Helper functions for specific parsing tasks
- Error handling for common parsing failures
- Integration with the extraction pipeline

### Integration Point
```python
# In Processing/Core/extraction.py

from Processing.Parsers.new_parser import process_new_content_type

# Add to extract_pdf_data function
if param_set.get('new_content_type_processing', False):
    # Configure parameters
    new_content_params = {
        'param1': param_set.get('param1', default_value),
        'param2': param_set.get('param2', default_value)
    }
    
    # Process with new parser
    processed_result = process_new_content_type(formatted_raw_text, new_content_params)
    
    # Store results
    extracted_data[field_name] = {
        "raw_text": original_raw_text,
        "formatted_text": formatted_raw_text,
        "parsed_data": processed_result
    }
```

## 2. New Utility Function Pattern

When adding new utilities for data manipulation or processing:

### File Structure
```
Processing/
└── Utilities/
    └── new_utility.py
```

### Key Implementation Elements
- Pure functions with minimal side effects
- Clear input/output documentation
- Comprehensive error handling
- Unit testable functionality

### Integration Point
```python
# In appropriate module that needs the utility

from Processing.Utilities.new_utility import perform_utility_function

# Use in processing pipeline
processed_data = perform_utility_function(input_data, params)
```

## 3. New Export Format Pattern

When adding support for exporting to new file formats:

### File Structure
```
Processing/
└── export_format.py
```

### Key Implementation Elements
- Export function that takes processed data and outputs desired format
- Configuration options for customizing output
- Error handling and validation
- Clear documentation of output format

### Integration Point
```python
# In pdf_processor.py

from Processing.export_format import export_to_format

def convert_to_format(json_path, output_path=None, config=None):
    """
    Convert processed JSON data to new format.
    
    Args:
        json_path (str): Path to JSON file
        output_path (str, optional): Output file path
        config (dict, optional): Export configuration
        
    Returns:
        str: Path to created file
    """
    # Validate inputs
    if not os.path.isfile(json_path):
        print(f"Error: JSON file '{json_path}' not found.")
        return None
    
    # Set default output path if not provided
    if output_path is None:
        base_path = os.path.splitext(json_path)[0]
        output_path = f"{base_path}.new_format"
    
    # Load JSON data
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Export to new format
    try:
        result_path = export_to_format(data, output_path, config)
        return result_path
    except Exception as e:
        print(f"Error converting to new format: {str(e)}")
        return None
```

## 4. New Core Feature Pattern

When adding a significant new capability that enhances the core functionality:

### File Structure
```
Processing/
└── Core/
    └── new_feature.py
```

### Key Implementation Elements
- Well-defined module with clear purpose
- Integration with existing workflow
- Extension of extraction parameters
- Comprehensive error handling

### Integration Point
```python
# In Processing/__init__.py

from Processing.Core.new_feature import new_feature_function

# Export at package level
__all__ = ['create_document_json', 'new_feature_function']
```

## 5. Configuration Management Pattern

When adding features that require complex configuration:

### File Structure
```
Processing/
└── config/
    ├── __init__.py
    └── feature_config.py
```

### Key Implementation Elements
- Default configuration definitions
- Configuration validation functions
- Helper functions for merging user configuration with defaults
- Clear documentation of all configuration options

### Integration Point
```python
# In module that needs configuration

from Processing.config.feature_config import get_default_config, validate_config

def process_with_config(data, user_config=None):
    # Get default configuration
    config = get_default_config()
    
    # Merge with user configuration if provided
    if user_config:
        config.update(user_config)
    
    # Validate merged configuration
    if not validate_config(config):
        raise ValueError("Invalid configuration")
    
    # Process using configuration
    result = process_data_with_config(data, config)
    return result
```

By following these established patterns, new features developed with AI assistance will integrate seamlessly with the existing codebase, maintaining architectural consistency and code quality.