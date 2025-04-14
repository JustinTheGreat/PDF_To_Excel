"""
Processing package for PDF data extraction and processing.

This package contains modules organized in a hierarchical structure:
- Core: Core extraction and document handling functionality
- Parsers: Text and data parsing modules
- Utilities: Helper utilities for processing PDFs
"""

# Version information
__version__ = '1.0.0'

# Expose main functionality at package level for convenience
from Components.Processing.document import create_document_json