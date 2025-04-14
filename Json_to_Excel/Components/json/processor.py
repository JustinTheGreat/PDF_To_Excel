from Components.json.reader import JsonReader
from Components.json.analyzer import JsonAnalyzer
from Components.utils.file_utils import FileUtils

class JsonProcessor:
    """
    Main processor class for JSON operations, serving as a facade for the modularized functionality.
    This class delegates to specialized functions for specific operations.
    """
    
    @staticmethod
    def read_json_files(directory_path, recursive=True, print_debug=True):
        """
        Read all JSON files in the directory and its subdirectories and return their data.
        Delegates to JsonReader class.
        
        Args:
            directory_path: Path to the root directory
            recursive: Whether to search in subdirectories as well (default: True)
            print_debug: Whether to print debug information to console (default: True)
            
        Returns:
            Dictionary mapping file paths to their JSON content
        """
        return JsonReader.read_json_files(directory_path, recursive, print_debug)
    
    @staticmethod
    def analyze_json_structure(json_data, print_debug=True):
        """
        Analyze the structure of the JSON data to determine how to format the Excel sheet.
        Delegates to JsonAnalyzer class.
        
        Args:
            json_data: JSON data to analyze
            print_debug: Whether to print debug information
            
        Returns a dict with information about:
        - All unique keys
        - Maximum nesting depth for each key
        - Whether subtitles are needed
        """
        return JsonAnalyzer.analyze_json_structure(json_data, print_debug)
    
    @staticmethod
    def process_filename(filename, filter_text=""):
        """
        Process filename to remove extension and filter text.
        Delegates to FileUtils class.
        
        Args:
            filename: The original filename to process
            filter_text: Text to remove from the filename (optional)
            
        Returns:
            Processed filename without extension and filtered text
        """
        return FileUtils.process_filename(filename, filter_text)