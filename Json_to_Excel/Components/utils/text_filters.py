import re

class TextFilter:
    """
    Class for handling text filtering operations on values and strings.
    Provides methods to clean and standardize different types of data.
    """
    
    @staticmethod
    def remove_units(text, unit_patterns=None):
        """
        Remove unit patterns from text strings.
        
        Args:
            text: The text to process
            unit_patterns: List of unit patterns to remove (e.g., "[ms]", "V", etc.)
                           If None, default patterns will be used
        
        Returns:
            Cleaned text with units removed
        """
        if text is None:
            return None
            
        # Convert to string if not already
        text = str(text)
        
        # Default patterns to remove common units
        if unit_patterns is None:
            unit_patterns = [
                r"\[ms\]",       # milliseconds
                r"\[s\]",        # seconds
                r"\[V\]",        # volts
                r"\[mV\]",       # millivolts
                r"\[A\]",        # amps
                r"\[mA\]",       # milliamps
                r"\[Hz\]",      # hertz
                r"\[kHz\]",      # kilohertz
                r"\[MHz\]",      # megahertz
                r"\[°C\]",       # celsius
                r"\[mm\]",       # millimeters
                r"\[cm\]",       # centimeters
                r"\[m\]",        # meters
                r"\[\w+\]"       # catch-all for other bracketed units
            ]
        
        # Process each pattern
        for pattern in unit_patterns:
            text = re.sub(pattern, "", text)
        
        # Trim any whitespace
        return text.strip()
    
    @staticmethod
    def clean_numeric_value(text):
        """
        Extract numeric values from text, removing units and converting to a numeric type if possible.
        
        Args:
            text: The text to process
        
        Returns:
            Numeric value (float or int) if possible, otherwise cleaned string
        """
        if text is None:
            return None
            
        # Convert to string if not already
        text = str(text)
        
        # First remove units
        cleaned_text = TextFilter.remove_units(text)
        
        # Try to convert to numeric
        try:
            # Check if it's an integer
            if cleaned_text.isdigit() or (cleaned_text.startswith('-') and cleaned_text[1:].isdigit()):
                return int(cleaned_text)
            # Otherwise try float
            return float(cleaned_text)
        except ValueError:
            # If conversion fails, return the cleaned string
            return cleaned_text
    
    @staticmethod
    def process_value(value, remove_units=True, convert_numeric=False):
        """
        Process a value based on specified filters.
        
        Args:
            value: The value to process (can be string, list, etc.)
            remove_units: Whether to remove unit notations
            convert_numeric: Whether to convert to numeric values when possible
        
        Returns:
            Processed value
        """
        if value is None:
            return None
            
        # Handle lists recursively
        if isinstance(value, list):
            return [TextFilter.process_value(item, remove_units, convert_numeric) for item in value]
        
        # Handle dictionaries recursively
        if isinstance(value, dict):
            return {k: TextFilter.process_value(v, remove_units, convert_numeric) for k, v in value.items()}
        
        # Handle string values
        if isinstance(value, str):
            processed = value
            
            if remove_units:
                processed = TextFilter.remove_units(processed)
            
            if convert_numeric:
                processed = TextFilter.clean_numeric_value(processed)
                
            return processed
            
        # Return other types unchanged
        return value
    
    @staticmethod
    def custom_replace(text, replacements):
        """
        Apply custom text replacements.
        
        Args:
            text: The text to process
            replacements: Dictionary of {pattern: replacement}
        
        Returns:
            Text with replacements applied
        """
        if text is None or not isinstance(text, str):
            return text
            
        result = text
        for pattern, replacement in replacements.items():
            result = re.sub(pattern, replacement, result)
            
        return result.strip()