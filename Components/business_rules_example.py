"""
Business rules for PDF extraction and formatting.

This module contains business-specific rules for formatting extracted data.
Add your own custom formatting rules for specific field types below.
"""
import re
import datetime


def format_device_serial_data(data_dict, unparsed_lines):
    """
    Format device serial number and model information.
    
    This function standardizes device serial numbers and model identifiers
    to ensure consistent formatting across all extractions.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Formatted data dictionary
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Standardize serial number format (ABC-12345 or ABC-12345-XYZ)
    if "Serial Number" in result_dict:
        serial = result_dict["Serial Number"]
        
        # Remove any spaces in the serial number
        serial = serial.replace(" ", "")
        
        # Apply standard format if it's not already formatted
        if not re.match(r'^[A-Z]{3}-\d{5}(-[A-Z]{3})?$', serial):
            # Extract letters and numbers
            letters = ''.join([c for c in serial if c.isalpha()]).upper()
            numbers = ''.join([c for c in serial if c.isdigit()])
            
            # Format as ABC-12345 or ABC-12345-XYZ
            if len(letters) <= 3:
                # Pad with X if needed
                letters = letters.ljust(3, 'X')
                result_dict["Serial Number"] = f"{letters}-{numbers.zfill(5)}"
            else:
                # Format as ABC-12345-XYZ
                result_dict["Serial Number"] = f"{letters[:3]}-{numbers.zfill(5)}-{letters[3:6]}"
    
    # Standardize model numbers
    if "Model" in result_dict:
        model = result_dict["Model"]
        
        # Ensure model always starts with the product line code
        if not model.startswith("MDL-"):
            model = "MDL-" + model
        
        result_dict["Model"] = model
    
    # Try to extract model information from unparsed lines if not already present
    if "Model" not in result_dict:
        for line in unparsed_lines:
            # Look for patterns like "Model: XYZ123" or "Model XYZ123"
            match = re.search(r'[Mm]odel\s*:?\s*([A-Za-z0-9-]+)', line)
            if match:
                model = match.group(1)
                if not model.startswith("MDL-"):
                    model = "MDL-" + model
                result_dict["Model"] = model
                break
    
    # Standardize field names
    for standard_name, aliases in customer_fields.items():
        for field in list(result_dict.keys()):
            # Check if the field name is an alias for a standard name
            if field.lower() in aliases:
                # If it's not already the standard name, rename it
                if field != standard_name:
                    result_dict[standard_name] = result_dict[field]
                    del result_dict[field]
    
    # Format phone numbers consistently: (XXX) XXX-XXXX for US or international format
    if "Phone Number" in result_dict:
        phone = result_dict["Phone Number"]
        # Remove any non-digit characters
        digits = ''.join([c for c in phone if c.isdigit()])
        
        # Format based on length
        if len(digits) == 10:  # US number
            result_dict["Phone Number"] = f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"
        elif len(digits) > 10:  # International
            country_code = digits[0:len(digits)-10]
            result_dict["Phone Number"] = f"+{country_code} ({digits[-10:-7]}) {digits[-7:-4]}-{digits[-4:]}"
    
    # Format email address consistently (lowercase)
    if "Email" in result_dict:
        result_dict["Email"] = result_dict["Email"].lower().strip()
    
    # Normalize address format
    if "Address" in result_dict:
        address = result_dict["Address"]
        # Remove double spaces
        address = re.sub(r'\s+', ' ', address).strip()
        # Capitalize first letter of each word
        address = ' '.join(word.capitalize() if not word.startswith("#") else word for word in address.split())
        result_dict["Address"] = address
    
    # Extract address components from unparsed lines
    address_pattern = re.compile(r'([A-Za-z0-9\s,#\.-]+),?\s*([A-Za-z\s]+),?\s*([A-Za-z\s]+),?\s*(\d{5}(?:-\d{4})?)', re.IGNORECASE)
    
    for line in unparsed_lines:
        match = address_pattern.search(line)
        if match:
            street = match.group(1).strip()
            city = match.group(2).strip()
            state = match.group(3).strip()
            zipcode = match.group(4).strip()
            
            # Only add if not already present
            if "Address" not in result_dict:
                result_dict["Address"] = street
            if "City" not in result_dict:
                result_dict["City"] = city
            if "State/Province" not in result_dict:
                result_dict["State/Province"] = state
            if "ZIP/Postal Code" not in result_dict:
                result_dict["ZIP/Postal Code"] = zipcode
    
    return result_dict


def format_text_date(month_str, day_str, year_str):
    """
    Convert a text date (with month name) to ISO format.
    
    Args:
        month_str (str): Month name or abbreviation
        day_str (str): Day of month
        year_str (str): Year
        
    Returns:
        str: Date in ISO format (YYYY-MM-DD)
    """
    month_names = {
        "jan": 1, "january": 1,
        "feb": 2, "february": 2,
        "mar": 3, "march": 3,
        "apr": 4, "april": 4,
        "may": 5, "may": 5,
        "jun": 6, "june": 6,
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12
    }
    
    month_num = month_names.get(month_str.lower(), 1)  # Default to January if not found
    day = int(day_str)
    year = int(year_str)
    
    return f"{year}-{month_num:02d}-{day:02d}"


def format_date_values(data_dict, unparsed_lines):
    """
    Standardize all date formats to ISO format (YYYY-MM-DD).
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with standardized dates
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Common date patterns to match
    date_patterns = [
        # MM/DD/YYYY
        (r'(\d{1,2})/(\d{1,2})/(\d{4})', lambda m: f"{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"),
        # DD/MM/YYYY
        (r'(\d{1,2})\.(\d{1,2})\.(\d{4})', lambda m: f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"),
        # Month DD, YYYY
        (r'([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})', lambda m: format_text_date(m.group(1), m.group(2), m.group(3))),
        # DD Month YYYY
        (r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', lambda m: format_text_date(m.group(2), m.group(1), m.group(3))),
    ]
    
    # Date-related field names to check
    date_fields = [
        "Date", "Test Date", "Manufactured Date", "Installation Date", "Maintenance Date", 
        "Calibration Date", "Expiry Date", "Certification Date"
    ]
    
    # Process each field that might contain a date
    for field in list(result_dict.keys()):
        # Check if this is a date field or contains the word "date"
        if field in date_fields or "date" in field.lower():
            value = result_dict[field]
            
            # Skip if already in ISO format
            if re.match(r'^\d{4}-\d{2}-\d{2}$', str(value)):
                continue
                
            # Try each pattern
            for pattern, formatter in date_patterns:
                match = re.search(pattern, str(value), re.IGNORECASE)
                if match:
                    try:
                        result_dict[field] = formatter(match)
                        break
                    except ValueError:
                        # If date conversion fails, keep original
                        pass
    
    # Check unparsed lines for dates and add as new fields if found
    for line in unparsed_lines:
        # First, check if line contains a date field
        date_field_match = re.search(r'([A-Za-z\s]+Date[A-Za-z\s]*):?\s*(.*)', line, re.IGNORECASE)
        if date_field_match:
            field_name = date_field_match.group(1).strip()
            value = date_field_match.group(2).strip()
            
            # Skip if already processed
            if field_name in result_dict:
                continue
                
            # Try each pattern
            for pattern, formatter in date_patterns:
                match = re.search(pattern, value)
                if match:
                    try:
                        result_dict[field_name] = formatter(match)
                        break
                    except ValueError:
                        # If date conversion fails, keep original
                        result_dict[field_name] = value
    
    return result_dict


def format_measurement_values(data_dict, unparsed_lines):
    """
    Standardize measurement values to ensure consistent units and formatting.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with standardized measurements
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Unit conversion factors
    conversions = {
        # Length
        "mm_to_m": 0.001,
        "cm_to_m": 0.01,
        "in_to_m": 0.0254,
        "ft_to_m": 0.3048,
        # Weight
        "g_to_kg": 0.001,
        "oz_to_kg": 0.0283495,
        "lb_to_kg": 0.453592,
        # Volume
        "ml_to_l": 0.001,
        "cl_to_l": 0.01,
        "pt_to_l": 0.473176,
        "gal_to_l": 3.78541,
    }
    
    # Mapping of field names to expected unit types
    unit_expectations = {
        # Length fields (convert to meters)
        "Length": "m",
        "Width": "m",
        "Height": "m",
        "Diameter": "m",
        "Depth": "m",
        # Weight fields (convert to kg)
        "Weight": "kg",
        "Mass": "kg",
        # Volume fields (convert to liters)
        "Volume": "l",
        "Capacity": "l",
        # Electrical fields (no conversion, just standardization)
        "Voltage": "V",
        "Current": "A",
        "Power": "W",
        "Resistance": "Ω",
        # Temperature (convert to Celsius)
        "Temperature": "°C",
        "Ambient Temperature": "°C",
        "Operating Temperature": "°C",
    }
    
    # Process each field based on its expected unit
    for field, target_unit in unit_expectations.items():
        if field in result_dict:
            value = result_dict[field]
            
            # Try to extract the numeric value and unit
            match = re.search(r'([-+]?\d*\.?\d+)\s*([a-zA-Z°]+)', str(value))
            if match:
                try:
                    num_value = float(match.group(1))
                    unit = match.group(2).lower()
                    
                    # Apply specific conversions based on target unit
                    if target_unit == "m":
                        if unit in ["mm", "millimeter", "millimeters"]:
                            num_value *= conversions["mm_to_m"]
                        elif unit in ["cm", "centimeter", "centimeters"]:
                            num_value *= conversions["cm_to_m"]
                        elif unit in ["in", "inch", "inches"]:
                            num_value *= conversions["in_to_m"]
                        elif unit in ["ft", "foot", "feet"]:
                            num_value *= conversions["ft_to_m"]
                            
                    elif target_unit == "kg":
                        if unit in ["g", "gram", "grams"]:
                            num_value *= conversions["g_to_kg"]
                        elif unit in ["oz", "ounce", "ounces"]:
                            num_value *= conversions["oz_to_kg"]
                        elif unit in ["lb", "pound", "pounds"]:
                            num_value *= conversions["lb_to_kg"]
                            
                    elif target_unit == "l":
                        if unit in ["ml", "milliliter", "milliliters"]:
                            num_value *= conversions["ml_to_l"]
                        elif unit in ["cl", "centiliter", "centiliters"]:
                            num_value *= conversions["cl_to_l"]
                        elif unit in ["pt", "pint", "pints"]:
                            num_value *= conversions["pt_to_l"]
                        elif unit in ["gal", "gallon", "gallons"]:
                            num_value *= conversions["gal_to_l"]
                            
                    elif target_unit == "°C":
                        if unit in ["f", "°f", "fahrenheit"]:
                            # Convert Fahrenheit to Celsius
                            num_value = (num_value - 32) * 5/9
                    
                    # Format the value with appropriate precision
                    if num_value < 0.01:
                        formatted_value = f"{num_value:.6f} {target_unit}"
                    elif num_value < 1:
                        formatted_value = f"{num_value:.4f} {target_unit}"
                    else:
                        formatted_value = f"{num_value:.2f} {target_unit}"
                    
                    # Update the dictionary
                    result_dict[field] = formatted_value
                    
                except ValueError:
                    # Keep original if conversion fails
                    pass
            elif isinstance(value, str) and re.match(r'^[-+]?\d*\.?\d+$', value):
                # If we only have a number without a unit, add the expected unit
                try:
                    num_value = float(value)
                    
                    # Format the value with appropriate precision
                    if num_value < 0.01:
                        formatted_value = f"{num_value:.6f} {target_unit}"
                    elif num_value < 1:
                        formatted_value = f"{num_value:.4f} {target_unit}"
                    else:
                        formatted_value = f"{num_value:.2f} {target_unit}"
                    
                    # Update the dictionary
                    result_dict[field] = formatted_value
                except ValueError:
                    # Keep original if conversion fails
                    pass
    
    return result_dict


def format_test_results(data_dict, unparsed_lines):
    """
    Standardize test result data and extract additional test information from unparsed lines.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with standardized test results
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Standardize pass/fail results
    for field in list(result_dict.keys()):
        if "test" in field.lower() or "result" in field.lower():
            value = str(result_dict[field]).lower()
            
            # Standardize Pass results
            if value in ["pass", "passed", "ok", "success", "acceptable", "yes", "y", "✓", "√"]:
                result_dict[field] = "PASS"
                
            # Standardize Fail results
            elif value in ["fail", "failed", "not ok", "unacceptable", "no", "n", "×", "x"]:
                result_dict[field] = "FAIL"
                
            # Standardize Warning results
            elif value in ["warning", "caution", "attention", "marginal"]:
                result_dict[field] = "WARNING"
    
    # Extract test data from unparsed lines
    test_pattern = re.compile(r'([A-Za-z\s]+)(?:\s+[Tt]est)?\s*[:]?\s*(pass|fail|warning|[\d.]+)(?:\s+([A-Za-z]+))?', re.IGNORECASE)
    
    for line in unparsed_lines:
        match = test_pattern.search(line)
        if match:
            test_name = match.group(1).strip()
            result = match.group(2).strip().lower()
            unit = match.group(3) if match.group(3) else ""
            
            # Skip if already processed
            if test_name in result_dict:
                continue
                
            # Process test results
            if result in ["pass", "passed", "ok", "success", "acceptable", "yes", "y"]:
                result_dict[f"{test_name} Test"] = "PASS"
            elif result in ["fail", "failed", "not ok", "unacceptable", "no", "n"]:
                result_dict[f"{test_name} Test"] = "FAIL"
            elif result in ["warning", "caution", "attention", "marginal"]:
                result_dict[f"{test_name} Test"] = "WARNING"
            else:
                # Assume it's a numeric measurement
                try:
                    value = float(result)
                    result_dict[f"{test_name} Test"] = f"{value} {unit}".strip()
                except ValueError:
                    # If not a recognized value, store as is
                    result_dict[f"{test_name} Test"] = f"{result} {unit}".strip()
    
    return result_dict


def format_maintenance_records(data_dict, unparsed_lines):
    """
    Format maintenance record data with proper categorization.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with formatted maintenance records
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Organize maintenance dates and actions
    maintenance_entries = []
    
    # First, extract any existing maintenance-related fields
    for field in list(result_dict.keys()):
        if "maintenance" in field.lower() and "date" in field.lower():
            date_value = result_dict[field]
            
            # Look for matching action based on date
            action_field = None
            for action_key in list(result_dict.keys()):
                if "maintenance" in action_key.lower() and "action" in action_key.lower():
                    action_field = action_key
                    break
            
            # Create maintenance entry
            entry = {
                "date": date_value,
                "action": result_dict.get(action_field, "Regular maintenance"),
                "technician": "",
                "notes": ""
            }
            
            # Look for technician information
            for tech_key in list(result_dict.keys()):
                if "technician" in tech_key.lower() or "service" in tech_key.lower() and "provider" in tech_key.lower():
                    entry["technician"] = result_dict[tech_key]
                    break
            
            maintenance_entries.append(entry)
            
            # Remove processed fields
            result_dict.pop(field, None)
            if action_field:
                result_dict.pop(action_field, None)
                
    # Parse unparsed lines for additional maintenance information
    maintenance_pattern = re.compile(
        r'(?:Maintenance|Service)(?:\s+Date)?:\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}|[A-Za-z]+\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4})'
        r'(?:.*?Action(?:s)?:\s*([^,;]+))?'
        r'(?:.*?(?:Technician|Service Provider):\s*([^,;]+))?'
        r'(?:.*?Notes?:\s*(.+))?',
        re.IGNORECASE
    )
    
    for line in unparsed_lines:
        match = maintenance_pattern.search(line)
        if match:
            # Extract date
            date_str = match.group(1).strip()
            
            # Try to standardize the date format
            for pattern, formatter in [
                # MM/DD/YYYY
                (r'(\d{1,2})/(\d{1,2})/(\d{4})', lambda m: f"{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"),
                # DD/MM/YYYY
                (r'(\d{1,2})\.(\d{1,2})\.(\d{4})', lambda m: f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"),
                # YYYY/MM/DD
                (r'(\d{4})/(\d{1,2})/(\d{1,2})', lambda m: f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"),
                # Month DD, YYYY
                (r'([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})', lambda m: format_text_date(m.group(1), m.group(2), m.group(3))),
                # DD Month YYYY
                (r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', lambda m: format_text_date(m.group(2), m.group(1), m.group(3))),
            ]:
                date_match = re.search(pattern, date_str)
                if date_match:
                    try:
                        date_str = formatter(date_match)
                        break
                    except ValueError:
                        # If date conversion fails, keep original
                        pass
            
            # Extract other information
            action = match.group(2).strip() if match.group(2) else "Regular maintenance"
            technician = match.group(3).strip() if match.group(3) else ""
            notes = match.group(4).strip() if match.group(4) else ""
            
            # Add to maintenance entries
            maintenance_entries.append({
                "date": date_str,
                "action": action,
                "technician": technician,
                "notes": notes
            })
    
    # Sort maintenance entries by date (newest first)
    try:
        maintenance_entries.sort(key=lambda x: datetime.datetime.strptime(x["date"], "%Y-%m-%d"), reverse=True)
    except:
        # If date sorting fails, keep original order
        pass
    
    # Add the maintenance entries to the result dictionary
    if maintenance_entries:
        result_dict["Maintenance History"] = maintenance_entries
    
    return result_dict


def format_technical_parameters(data_dict, unparsed_lines):
    """
    Format technical parameter data with consistent naming and units.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with formatted technical parameters
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Standard parameter names and their aliases
    param_aliases = {
        "Voltage": ["voltage", "potential difference", "v", "volt", "volts"],
        "Current": ["current", "amperage", "a", "amp", "amps", "ampere", "amperes"],
        "Power": ["power", "wattage", "w", "watt", "watts"],
        "Power Factor": ["power factor", "pf", "cos phi", "cosφ"],
        "Frequency": ["frequency", "freq", "hz", "hertz"],
        "Temperature": ["temperature", "temp", "°c", "°f", "celsius", "fahrenheit"],
        "Pressure": ["pressure", "press", "bar", "psi", "kpa", "mpa"],
        "Flow Rate": ["flow rate", "flow", "lpm", "gpm", "m³/h", "cfm"],
        "Efficiency": ["efficiency", "eff", "%", "percent"],
        "RPM": ["rpm", "rotation", "rotation speed", "speed", "rev/min"],
        "Humidity": ["humidity", "rh", "relative humidity"],
        "Noise Level": ["noise", "noise level", "sound", "db", "dba", "decibel", "decibels"],
    }
    
    # Standardize parameter names
    for standard_name, aliases in param_aliases.items():
        for field in list(result_dict.keys()):
            # Check if the field name is an alias for a standard name
            if field.lower() in aliases:
                # If it's not already the standard name, rename it
                if field != standard_name:
                    result_dict[standard_name] = result_dict[field]
                    del result_dict[field]
    
    # Group related parameters into a single Technical Specifications object
    tech_spec_fields = [
        "Voltage", "Current", "Power", "Power Factor", "Frequency", 
        "Temperature", "Pressure", "Flow Rate", "Efficiency", 
        "RPM", "Humidity", "Noise Level"
    ]
    
    tech_specs = {}
    for field in tech_spec_fields:
        if field in result_dict:
            tech_specs[field] = result_dict.pop(field)
    
    if tech_specs:
        result_dict["Technical Specifications"] = tech_specs
    
    # Extract additional parameters from unparsed lines
    param_pattern = re.compile(r'([A-Za-z\s]+)(?:\(([A-Za-z%°/]+)\))?:\s*([-+]?\d*\.?\d+)(?:\s*([A-Za-z%°/]+))?', re.IGNORECASE)
    
    for line in unparsed_lines:
        match = param_pattern.search(line)
        if match:
            param_name = match.group(1).strip()
            unit_in_name = match.group(2).strip() if match.group(2) else ""
            value = match.group(3).strip()
            unit_after_value = match.group(4).strip() if match.group(4) else ""
            
            # Determine the unit (prefer unit after value, then unit in name)
            unit = unit_after_value if unit_after_value else unit_in_name
            
            # Standardize parameter name
            standard_name = param_name
            for std_name, aliases in param_aliases.items():
                if param_name.lower() in aliases:
                    standard_name = std_name
                    break
            
            # Skip if already processed
            if standard_name in result_dict:
                continue
                
            # Format the value with the unit
            if unit:
                result_dict[standard_name] = f"{value} {unit}"
            else:
                result_dict[standard_name] = value
    
    return result_dict


def format_customer_information(data_dict, unparsed_lines):
    """
    Format customer information with consistent naming and formats.
    
    Args:
        data_dict (dict): Dictionary of data to format
        unparsed_lines (list): Lines that couldn't be parsed with simple key-value rules
        
    Returns:
        dict: Dictionary with formatted customer information
    """
    # Make a copy of the dictionary to avoid modifying the original
    result_dict = data_dict.copy()
    
    # Standard customer field names and their aliases
    customer_fields = {
        "Customer Name": ["customer", "client", "name", "customer name", "client name"],
        "Company": ["company", "organization", "business", "enterprise", "firm"],
        "Contact Person": ["contact", "contact person", "representative", "point of contact", "poc"],
        "Phone Number": ["phone", "tel", "telephone", "contact number", "mobile"],
        "Email": ["email", "e-mail", "electronic mail", "email address"],
        "Address": ["address", "location", "site", "physical address"],
        "ZIP/Postal Code": ["zip", "postal code", "post code", "zip code"],
        "Country": ["country", "nation"],
        "Customer ID": ["customer id", "client id", "account number", "account id", "customer number"]