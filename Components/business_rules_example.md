# Business Rules Implementation

The enhanced `business_rules.py` demonstrates how to implement customized data formatting rules:

1. **Device Serial Data Formatting**:
   - Standardize serial numbers to format ABC-12345 or ABC-12345-XYZ
   - Ensure model numbers follow a consistent format (MDL-xxxxx)

2. **Date Formatting**:
   - Standardize all dates to ISO format (YYYY-MM-DD)
   - Handle different input formats (MM/DD/YYYY, DD/MM/YYYY, Month DD, YYYY, etc.)

3. **Measurement Value Standardization**:
   - Convert measurements to consistent units (m, kg, l, °C)
   - Ensure proper formatting of values with appropriate precision

4. **Test Result Standardization**:
   - Normalize pass/fail results ("pass", "ok", "acceptable" → "PASS")
   - Extract additional test details from unparsed lines

5. **Maintenance Record Formatting**:
   - Organize maintenance dates and actions
   - Extract technician information and notes
   - Sort entries by date

6. **Technical Parameter Standardization**:
   - Consistent naming of technical parameters
   - Group related parameters into a single Technical Specifications object

7. **Customer Information Formatting**:
   - Standardize contact details
   - Format phone numbers, email addresses, and postal addresses consistently

## Using Business Rules

To implement your own business rules, follow these steps:

1. Create specialized formatting functions for each field type
2. Update the `apply_business_rules` function to call the appropriate formatter based on field name
3. Implement pattern recognition to extract information from unparsed lines

Example:

```python
def apply_business_rules(field_name, data_dict, unparsed_lines):
    # Create a copy to avoid modifying the original
    result_dict = data_dict.copy()

    # Apply field-specific formatting rules
    if field_name == "General Info":
        result_dict = format_customer_information(result_dict, unparsed_lines)
    elif field_name == "Technical Parameters":
        result_dict = format_technical_parameters(result_dict, unparsed_lines)
    
    return result_dict
```

## Tips for Effective Extraction

1. **Start with broad extraction parameters** and refine them based on results
2. **Use `forced_keywords`** to improve key-value parsing for fields missing colons
3. **Use `remove_breaks_before` and `remove_breaks_after`** to fix formatting issues
4. **Implement detailed business rules** for each field type to handle edge cases
5. **Use field continuation with (+1) suffix** for data that spans multiple pages
6. **Use chart processing with (Chart) suffix** for tabular data
7. **Leverage unparsed_lines** to extract additional information that wasn't parsed as key-value pairs

## Practical Example

Here's a practical example of using the system to extract data from a technical report:

```python
extraction_params = [
    # Extract customer information from the cover page
    {
        "field_name": "Customer Info",
        "start_keyword": "Client:",
        "end_keyword": "Project Details",
        "page_num": 0,
        "horiz_margin": 300,
        "forced_keywords": ["Name", "Company", "Contact", "Phone", "Email"]
    },
    
    # Extract device specifications from the technical section
    {
        "field_name": "Device Specifications",
        "start_keyword": "Technical Specifications",
        "end_keyword": "Test Protocol",
        "page_num": 1,
        "horiz_margin": 400,
        "forced_keywords": ["Model", "Serial Number", "Firmware"]
    },
    
    # Extract test results that span multiple pages
    {
        "field_name": "Test Results",
        "start_keyword": "Test Results",
        "end_keyword": "End of Test Data",
        "page_num": 2,
        "horiz_margin": 500
    },
    
    # Continuation of test results on next page
    {
        "field_name": "Test Results(+1)",
        "start_keyword": "Additional Tests",
        "end_keyword": "Conclusion",
        "page_num": 3,
        "horiz_margin": 500
    },
    
    # Extract chart data
    {
        "field_name": "Performance Metrics(Chart)",
        "start_keyword": "Performance Data",
        "end_keyword": "End of Performance Data",
        "page_num": 4,
        "horiz_margin": 500,
        "top_title": True,
        "left_title": True,
        "priority_side": "top"
    }
]
```

The system will extract these fields, apply appropriate business rules to format the data, and output a structured JSON document with the organized information.