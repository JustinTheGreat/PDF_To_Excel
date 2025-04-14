# AI-Assisted Development Guide for PDF Processing Framework

This guide outlines which files to provide to an AI assistant when developing new features or extending the PDF processing framework. Selecting the right context files helps the AI understand the system architecture and provide more accurate and useful code.
(Link)[https://claude.site/artifacts/d85b1387-ab68-4563-97d9-c198fabe2282]

## Core Files to Always Provide

When working with an AI assistant on this framework, always provide these key files for context:

1. **pdf_processor.py**
   - Main entry point that shows how components interact
   - Provides overview of the processing workflow

2. **Processing/__init__.py**
   - Shows how modules are exported and organized
   - Reveals the public API of the Processing package

3. **Processing/document.py**
   - Contains core document processing logic
   - Shows how JSON files are generated from extracted data

## Feature-Specific Files

Depending on the type of feature you're developing, include additional files based on the category:

### For Extraction-Related Features

If developing new ways to extract data from PDFs, include:

- **Processing/Core/extraction.py**
- **Components/GeneralInfo.py**
- **Components/pdf_extractor.py**

### For Parsing-Related Features

If developing new data format parsers, include:

- **Processing/Parsers/table.py** (for table-based features)
- **Processing/Parsers/keywords.py** (for keyword detection features)

### For Data Processing Features

If developing features related to processing or transforming extracted data:

- **Processing/Utilities/cleaner.py**
- **Processing/Utilities/merger.py**

### For PDF Text Analysis Features

If developing features related to PDF text analysis:

- **Processing/Utilities/text.py**
- **Components/GeneralInfo.py**

## Example Scenarios

### Scenario 1: Adding a New Extraction Method

Files to provide:
- pdf_processor.py
- Processing/__init__.py
- Processing/document.py
- Processing/Core/extraction.py
- Components/GeneralInfo.py
- Components/pdf_extractor.py

### Scenario 2: Adding a New Table Parsing Technique

Files to provide:
- pdf_processor.py
- Processing/__init__.py
- Processing/document.py
- Processing/Parsers/table.py
- Processing/Core/extraction.py

### Scenario 3: Adding PDF Image Extraction Capability

Files to provide:
- pdf_processor.py
- Processing/__init__.py
- Processing/document.py
- Processing/Core/extraction.py
- Processing/Utilities/text.py
- Example of the image-containing PDF (if possible)

### Scenario 4: Creating a New Data Field Merging Strategy

Files to provide:
- pdf_processor.py
- Processing/__init__.py
- Processing/document.py
- Processing/Utilities/merger.py
- Processing/Utilities/cleaner.py

## Tips for Effective AI Collaboration

1. **Provide Working Examples**: Include small sample PDFs and expected output JSONs when possible

2. **Specify Integration Points**: Clearly explain where the new functionality should connect to existing code

3. **Explain Business Logic**: Give context about the business rules or domain-specific knowledge related to the feature

4. **Include Relevant Types**: If you have type definitions or interface specifications, include those to ensure type compliance

5. **Start with Structure**: Ask the AI to outline the approach and structure first before implementing details

6. **Use Iterative Development**: Break complex features into smaller chunks that build on each other

## Documentation Template for New Feature Requests

When requesting a new feature, use this template to communicate with the AI:

```
Feature Name: [Name of the feature]

Description: [Brief description of what the feature should do]

Integration Points: [Where/how this feature connects to existing code]

Files Provided:
- [List of files provided for context]

Expected Input: [Description of input data/format]

Expected Output: [Description of output data/format]

Business Rules: [Any specific business rules or constraints]

Additional Context: [Any other relevant information]
```

By following this guide, you'll ensure the AI has sufficient context to develop high-quality, well-integrated extensions to your PDF processing framework.