import os

class FileUtils:
    """
    Utility class for file operations, particularly for processing filenames.
    """
    
    @staticmethod
    def process_filename(filename, filter_text=""):
        """
        Process filename to remove extension and filter text.
        
        Args:
            filename: The original filename to process
            filter_text: Text to remove from the filename (optional)
            
        Returns:
            Processed filename without extension and filtered text
        """
        # Remove extension
        display_filename = os.path.splitext(filename)[0]
        
        # Remove filter text if provided
        if filter_text and filter_text in display_filename:
            display_filename = display_filename.replace(filter_text, "").strip()
        
        return display_filename
    
    @staticmethod
    def sanitize_sheet_name(sheet_name, max_length=31):
        """
        Sanitize a sheet name for use in Excel.
        
        Args:
            sheet_name: The original sheet name to sanitize
            max_length: Maximum length for Excel sheet names (default: 31)
            
        Returns:
            Sanitized sheet name
        """
        # Remove invalid characters
        safe_name = ''.join(c for c in sheet_name if c not in '\\/:*?[]')
        
        # Truncate to maximum length
        safe_name = safe_name[:max_length]
        
        return safe_name