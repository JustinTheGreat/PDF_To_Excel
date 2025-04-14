from simple_selection import process_single_file_terminal
from pdf_multiple_selection import launch_mode_selection_dialog,launch_custom_options_dialog

# Main execution
if __name__ == "__main__":
    # Define extraction parameters for different sections (page_num starts at 0)
    extraction_params = [
        {
            "field_name": "Customer Sheet",
            "start_keyword": "Customer Name",
            "end_keyword": "Package",
            "page_num": 0,
            "horiz_margin": 500,
            "end_keyword_occurrence": 1
        },
    ]

    process_single_file_terminal(extraction_params)
    
    # Use the multiple GUI function from the GUI module
    #launch_mode_selection_dialog(extraction_params, process_multiple_files)
    
    # Alternatively, you could use the custom options dialog:
    # launch_custom_options_dialog(extraction_params, process_multiple_files)