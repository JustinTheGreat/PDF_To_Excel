import os
import pdfplumber

def detect_pdf_rotation(pdf_path):
    """
    Detects and reports the rotation of all pages in a PDF document.
    
    Args:
        pdf_path (str): Path to the PDF file
    
    Returns:
        list: A list of rotation values for each page
    """
    try:
        # Check if file exists
        if not os.path.isfile(pdf_path):
            print(f"Error: File '{pdf_path}' not found.")
            return None
        
        # Open the PDF file
        with pdfplumber.open(pdf_path) as pdf:
            rotations = []
            page_dimensions = []
            
            # Iterate through each page
            for i, page in enumerate(pdf.pages):
                rotations.append(page.rotation)
                
                # Get page dimensions
                width = page.width
                height = page.height
                page_dimensions.append((width, height))
                
                # Determine orientation based on dimensions
                orientation = "Portrait" if height > width else "Landscape"
                
                # Display info about each page
                print(f"Page {i+1}:")
                print(f"  - Rotation: {page.rotation} degrees")
                print(f"  - Dimensions: {width} x {height} points")
                print(f"  - Physical orientation: {orientation}")
                print(f"  - Metadata: {page.bbox}")
                
                # For the first page, print more detailed info about PDF structure
                if i == 0:
                    print(f"  - PDF Structure Info:")
                    if hasattr(page, 'stream') and page.stream is not None:
                        print("    - Has stream data")
                    if hasattr(page, 'resources') and page.resources is not None:
                        print(f"    - Resources: {list(page.resources.keys()) if page.resources else 'None'}")
                    if hasattr(page, 'mediabox') and page.mediabox is not None:
                        print(f"    - MediaBox: {page.mediabox}")
            
            # Summary
            print("\nSummary:")
            print(f"Total pages: {len(pdf.pages)}")
            rotation_set = set(rotations)
            if len(rotation_set) == 1:
                print(f"All pages have the same rotation: {rotations[0]} degrees")
            else:
                print("The document has mixed page rotations:")
                for rotation in sorted(rotation_set):
                    count = rotations.count(rotation)
                    print(f"  - {rotation} degrees: {count} page(s)")
            
            # Check if dimensions suggest a different orientation than rotation
            for i, (width, height) in enumerate(page_dimensions):
                physical_orientation = "Portrait" if height > width else "Landscape"
                rotated_orientation = "Landscape" if rotations[i] in [90, 270] else "Portrait"
                
                if physical_orientation != rotated_orientation and rotations[i] != 0:
                    print(f"\nNote: Page {i+1} appears to be in {physical_orientation} orientation")
                    print(f"but has a rotation value of {rotations[i]} degrees, which suggests {rotated_orientation}.")
                    print(f"This could cause coordinate system issues when extracting content.")
                
            return rotations
            
    except Exception as e:
        print(f"Error analyzing PDF: {str(e)}")
        return None

if __name__ == "__main__":
    # Get PDF path from user input
    print("PDF Rotation Detector")
    print("--------------------")
    pdf_path = input("Enter the path to the PDF file: ")
    
    # Remove any quotes that might have been included in the path
    pdf_path = pdf_path.strip('"\'')
    
    # Run the rotation detection
    detect_pdf_rotation(pdf_path)
    
    # Wait for user to read results
    input("\nPress Enter to exit...")