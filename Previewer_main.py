import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os
import sys
import importlib.util
import json
import traceback
try:
    import pdfplumber
except ImportError:
    messagebox.showerror("Error", "pdfplumber is required. Please install it with: pip install pdfplumber")
    sys.exit(1)

class CombinedBoundingBoxPreviewer:
    def __init__(self, master):
        self.master = master
        self.master.title("Combined PDF Bounding Box Previewer")
        self.master.geometry("1200x900")
        
        self.pdf_document = None  # PyMuPDF document
        self.pdfplumber_doc = None  # pdfplumber document
        self.current_page = 0
        self.total_pages = 0
        self.zoom_level = 1.0
        self.image_tk = None
        
        self.extraction_params = []
        self.bounding_boxes = {}
        self.current_displayed_boxes = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Top frame for controls
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=5)
        
        # Left controls frame for PDF
        pdf_frame = ttk.LabelFrame(top_frame, text="PDF Controls")
        pdf_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # PDF selection
        ttk.Button(pdf_frame, text="Open PDF", command=self.open_pdf).grid(row=0, column=0, padx=5, pady=5)
        self.pdf_label = ttk.Label(pdf_frame, text="No PDF selected")
        self.pdf_label.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Page navigation
        page_frame = ttk.Frame(pdf_frame)
        page_frame.grid(row=0, column=2, padx=20, pady=5)
        
        ttk.Button(page_frame, text="<", command=self.prev_page).pack(side=tk.LEFT)
        self.page_label = ttk.Label(page_frame, text="Page 0 of 0")
        self.page_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(page_frame, text=">", command=self.next_page).pack(side=tk.LEFT)
        
        # Zoom controls
        zoom_frame = ttk.Frame(pdf_frame)
        zoom_frame.grid(row=0, column=3, padx=20, pady=5)
        
        ttk.Button(zoom_frame, text="-", command=self.zoom_out).pack(side=tk.LEFT)
        self.zoom_label = ttk.Label(zoom_frame, text="100%")
        self.zoom_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(zoom_frame, text="+", command=self.zoom_in).pack(side=tk.LEFT)
        
        # Right controls frame for extraction parameters
        params_frame = ttk.LabelFrame(top_frame, text="Extraction Parameters")
        params_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
        
        # Parameters selection
        ttk.Button(params_frame, text="Load Parameters", command=self.load_parameters).grid(row=0, column=0, padx=5, pady=5)
        self.params_label = ttk.Label(params_frame, text="No parameters loaded")
        self.params_label.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Button(params_frame, text="Generate Boxes", command=self.generate_boxes).grid(row=0, column=2, padx=5, pady=5)
        
        # Manual bounding box controls
        bbox_frame = ttk.LabelFrame(main_frame, text="Manual Bounding Box")
        bbox_frame.pack(fill=tk.X, pady=10)
        
        # Input fields for manual bounding box coordinates
        ttk.Label(bbox_frame, text="x0:").grid(row=0, column=0, padx=5, pady=5)
        self.x0_var = tk.StringVar(value="0")
        ttk.Entry(bbox_frame, textvariable=self.x0_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(bbox_frame, text="x1:").grid(row=0, column=2, padx=5, pady=5)
        self.x1_var = tk.StringVar(value="100")
        ttk.Entry(bbox_frame, textvariable=self.x1_var, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(bbox_frame, text="top:").grid(row=0, column=4, padx=5, pady=5)
        self.top_var = tk.StringVar(value="0")
        ttk.Entry(bbox_frame, textvariable=self.top_var, width=10).grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Label(bbox_frame, text="bottom:").grid(row=0, column=6, padx=5, pady=5)
        self.bottom_var = tk.StringVar(value="100")
        ttk.Entry(bbox_frame, textvariable=self.bottom_var, width=10).grid(row=0, column=7, padx=5, pady=5)
        
        ttk.Button(bbox_frame, text="Add Manual Box", command=self.add_manual_box).grid(row=0, column=8, padx=10, pady=5)
        
        # Bounding box list frame (left side)
        list_display_frame = ttk.Frame(main_frame)
        list_display_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Left panel for bounding box list
        bbox_list_frame = ttk.LabelFrame(list_display_frame, text="Bounding Boxes")
        bbox_list_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        
        # Create a treeview for bounding boxes
        self.bbox_tree = ttk.Treeview(bbox_list_frame, columns=("Page", "Left", "Right", "Top", "Bottom"))
        self.bbox_tree.heading("#0", text="Field Name")
        self.bbox_tree.heading("Page", text="Page")
        self.bbox_tree.heading("Left", text="Left")
        self.bbox_tree.heading("Right", text="Right")
        self.bbox_tree.heading("Top", text="Top")
        self.bbox_tree.heading("Bottom", text="Bottom")
        
        self.bbox_tree.column("#0", width=150)
        self.bbox_tree.column("Page", width=50)
        self.bbox_tree.column("Left", width=50)
        self.bbox_tree.column("Right", width=50)
        self.bbox_tree.column("Top", width=50)
        self.bbox_tree.column("Bottom", width=50)
        
        self.bbox_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add button to view selected bounding box
        ttk.Button(bbox_list_frame, text="Show Selected Box", command=self.show_selected_box).pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(bbox_list_frame, text="Show All Boxes on Current Page", command=self.show_all_boxes_on_page).pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(bbox_list_frame, text="Clear All Boxes", command=self.clear_all_boxes).pack(fill=tk.X, padx=5, pady=5)
        
        # Right panel for PDF display
        display_frame = ttk.LabelFrame(list_display_frame, text="PDF Preview")
        display_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Canvas for displaying the PDF
        self.canvas = tk.Canvas(display_frame, bg="lightgray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Add scrollbars for the canvas
        h_scrollbar = ttk.Scrollbar(display_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(display_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
        
        # Set up event bindings for coordinate tracking
        self.canvas.bind("<Motion>", self.mouse_move)
        self.canvas.bind("<ButtonPress-1>", self.mouse_click)
        
        # Bind double-click on treeview to show the box
        self.bbox_tree.bind("<Double-1>", lambda event: self.show_selected_box())
        
    def open_pdf(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            # Close previous documents if open
            if self.pdf_document:
                self.pdf_document.close()
            if self.pdfplumber_doc:
                self.pdfplumber_doc.close()
                
            # Open with PyMuPDF for display
            self.pdf_document = fitz.open(file_path)
            self.total_pages = len(self.pdf_document)
            self.current_page = 0
            
            # Also open with pdfplumber for extraction
            self.pdfplumber_doc = pdfplumber.open(file_path)
            
            self.pdf_label.config(text=os.path.basename(file_path))
            self.update_page_label()
            self.render_page()
            
            self.status_var.set(f"Loaded: {file_path}")
            
            # Clear any existing bounding boxes
            self.clear_bbox_tree()
            self.current_displayed_boxes = []
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")
            self.status_var.set("Error loading PDF")
    
    def load_parameters(self):
        params_path = filedialog.askopenfilename(
            filetypes=[("Parameter files", "*.py;*.json"), ("Python files", "*.py"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not params_path:
            return
            
        try:
            self.extraction_params = self.load_params_from_file(params_path)
            if self.extraction_params:
                self.params_label.config(text=os.path.basename(params_path))
                self.status_var.set(f"Loaded {len(self.extraction_params)} parameters from {params_path}")
            else:
                self.params_label.config(text="No parameters found")
                self.status_var.set("No valid parameters found in file")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load parameters: {e}")
            self.params_label.config(text="Error loading parameters")
            self.status_var.set("Error loading parameters")
    
    def load_params_from_file(self, params_path):
        """Load extraction parameters from a file (JSON or Python)"""
        if not os.path.exists(params_path):
            messagebox.showerror("Error", f"Parameters file not found: {params_path}")
            return []
        
        try:
            # Load based on file extension
            file_ext = os.path.splitext(params_path)[1].lower()
            
            if file_ext == '.json':
                # Load from JSON file
                with open(params_path, 'r') as f:
                    extraction_params = json.load(f)
                    
            elif file_ext == '.py':
                # Load from Python module
                try:
                    # Import the module
                    module_name = os.path.basename(params_path).replace('.py', '')
                    spec = importlib.util.spec_from_file_location(module_name, params_path)
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    
                    # Look for extraction_params in the module
                    if hasattr(module, 'extraction_params'):
                        extraction_params = module.extraction_params
                    else:
                        # Try to find any list of dictionaries that might be extraction parameters
                        extraction_params = []
                        for attr_name in dir(module):
                            attr = getattr(module, attr_name)
                            if isinstance(attr, list) and len(attr) > 0 and isinstance(attr[0], dict):
                                if 'field_name' in attr[0] and ('start_keyword' in attr[0] or 'page_num' in attr[0]):
                                    extraction_params = attr
                                    break
                
                except Exception as e:
                    messagebox.showerror("Error", f"Error loading Python module: {str(e)}")
                    raise
            else:
                messagebox.showerror("Error", f"Unsupported file type: {file_ext}")
                return []
            
            # Validate extraction parameters
            if not isinstance(extraction_params, list) or not extraction_params:
                messagebox.showerror("Error", "No valid extraction parameters found")
                return []
            
            # Check that each item has at least field_name and start_keyword or other essential params
            valid_params = []
            for param in extraction_params:
                if isinstance(param, dict) and 'field_name' in param:
                    # Check for required fields based on type
                    # Chart parameters don't need start_keyword
                    if "(Chart)" in param.get('field_name', ''):
                        valid_params.append(param)
                    elif 'start_keyword' in param:
                        valid_params.append(param)
                    
            if not valid_params:
                messagebox.showerror("Error", "No valid extraction parameters found with required fields")
                return []
                
            self.status_var.set(f"Loaded {len(valid_params)} extraction parameters")
            return valid_params
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading extraction parameters: {str(e)}")
            traceback.print_exc()
            raise
    
    def find_keyword_position(self, words, keyword, occurrence=1):
        """Find the position of a keyword in a list of words"""
        if not keyword:
            return None
            
        keyword_count = 0
        keyword = keyword.lower()
        
        for word in words:
            if keyword in word['text'].lower():
                keyword_count += 1
                if keyword_count == occurrence:
                    pos = {
                        'x0': word['x0'],
                        'y0': word['top'],
                        'x1': word['x1'],
                        'y1': word['bottom']
                    }
                    return pos
                    
        return None
    
    def calculate_bounding_box(self, start_pos, end_pos, param_set):
        """Calculate the bounding box based on start/end positions and parameters"""
        # Extract parameters
        horiz_margin = param_set.get('horiz_margin', 200)
        left_move = param_set.get('left_move', 0)
        vertical_margin = param_set.get('vertical_margin', None)
        
        # Calculate the bounding box
        left = start_pos['x0'] - left_move
        right = left + horiz_margin
        top = start_pos['y0']
        
        # Determine the bottom position
        if vertical_margin is not None and vertical_margin > 0:
            # Use the specified vertical margin
            bottom = top + vertical_margin
        elif end_pos:
            # Use the end keyword position plus a small margin
            bottom = end_pos['y1'] + 5
        else:
            # Use a reasonable default vertical distance
            # Based on the end_break_line_count if provided
            end_break_line_count = param_set.get('end_break_line_count')
            if end_break_line_count:
                # Estimate line height as 15 points if not otherwise specified
                line_height = param_set.get('line_height', 15)
                bottom = top + (line_height * end_break_line_count) + 10
            else:
                # Use a default value
                bottom = top + 100
                
        return {
            'left': left,
            'right': right,
            'top': top,
            'bottom': bottom
        }
    
    def generate_boxes(self):
        """Generate bounding boxes based on extraction parameters"""
        if not self.pdf_document or not self.pdfplumber_doc:
            messagebox.showerror("Error", "Please open a PDF file first")
            return
            
        if not self.extraction_params:
            messagebox.showerror("Error", "Please load extraction parameters first")
            return
            
        try:
            self.status_var.set("Generating bounding boxes...")
            self.bounding_boxes = {}
            
            # Clear the tree view
            self.clear_bbox_tree()
            
            # Process each parameter set to generate bounding boxes
            for param_set in self.extraction_params:
                field_name = param_set.get('field_name', 'Unknown Field')
                
                # Skip chart parameters (they don't have physical positions)
                if "(Chart)" in field_name or not param_set.get('start_keyword'):
                    continue
                    
                # Extract parameters needed for bounding box calculation
                page_num = param_set.get('page_num', 0)
                start_keyword = param_set.get('start_keyword', '')
                start_keyword_occurrence = param_set.get('start_keyword_occurrence', 1)
                end_keyword = param_set.get('end_keyword', None)
                end_keyword_occurrence = param_set.get('end_keyword_occurrence', 1)
                
                # Skip if page number is out of range
                if page_num >= len(self.pdfplumber_doc.pages):
                    continue
                    
                # Get the page
                page = self.pdfplumber_doc.pages[page_num]
                
                # Extract words with positions
                words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
                
                # Find the start keyword position (accounting for occurrence)
                start_pos = self.find_keyword_position(words, start_keyword, start_keyword_occurrence)
                
                if not start_pos:
                    continue
                    
                # Find the end keyword position if specified
                end_pos = None
                if end_keyword:
                    end_pos = self.find_keyword_position(words, end_keyword, end_keyword_occurrence)
                
                # Calculate the bounding box
                box = self.calculate_bounding_box(start_pos, end_pos, param_set)
                
                # Store the bounding box with proper field name
                # Remove chart and +1 indicators for display purposes
                display_name = field_name.replace("(+1)", "").replace("(Chart)", "").strip()
                
                # If multiple parameters share the same display name, make them unique
                count = 1
                base_name = display_name
                while display_name in self.bounding_boxes:
                    count += 1
                    display_name = f"{base_name} ({count})"
                
                self.bounding_boxes[display_name] = {
                    'page': page_num,
                    'left': box['left'],
                    'right': box['right'],
                    'top': box['top'],
                    'bottom': box['bottom']
                }
                
                # Add to tree view
                self.bbox_tree.insert(
                    "", tk.END, text=display_name, 
                    values=(page_num, f"{box['left']:.1f}", f"{box['right']:.1f}", f"{box['top']:.1f}", f"{box['bottom']:.1f}")
                )
                
            self.status_var.set(f"Generated {len(self.bounding_boxes)} bounding boxes")
            
            # Show all boxes on the current page
            self.show_all_boxes_on_page()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error generating bounding boxes: {str(e)}")
            traceback.print_exc()
            self.status_var.set("Error generating bounding boxes")
    
    def add_manual_box(self):
        """Add a manually defined bounding box"""
        if not self.pdf_document:
            messagebox.showerror("Error", "Please open a PDF file first")
            return
            
        try:
            # Get coordinates from entry fields
            left = float(self.x0_var.get())
            right = float(self.x1_var.get())
            top = float(self.top_var.get())
            bottom = float(self.bottom_var.get())
            
            # Create a unique name for this box
            box_count = len(self.bounding_boxes) + 1
            box_name = f"Manual Box {box_count}"
            
            # Add to bounding boxes dictionary
            self.bounding_boxes[box_name] = {
                'page': self.current_page,
                'left': left,
                'right': right,
                'top': top,
                'bottom': bottom
            }
            
            # Add to tree view
            self.bbox_tree.insert(
                "", tk.END, text=box_name, 
                values=(self.current_page, f"{left:.1f}", f"{right:.1f}", f"{top:.1f}", f"{bottom:.1f}")
            )
            
            # Show this box
            self.render_page()
            self.draw_bounding_box(box_name, self.bounding_boxes[box_name])
            
            self.status_var.set(f"Added manual box: {box_name}")
            
        except ValueError:
            messagebox.showerror("Error", "Invalid coordinate values. Please enter numbers only.")
        except Exception as e:
            messagebox.showerror("Error", f"Error adding manual box: {str(e)}")
            self.status_var.set("Error adding manual box")
    
    def show_selected_box(self):
        """Show the selected bounding box on the PDF"""
        selected = self.bbox_tree.selection()
        if not selected:
            messagebox.showinfo("Info", "Please select a bounding box from the list")
            return
            
        # Get the field name
        field_name = self.bbox_tree.item(selected[0], "text")
        
        if field_name not in self.bounding_boxes:
            messagebox.showerror("Error", f"Bounding box '{field_name}' not found")
            return
            
        # Get the bounding box
        box = self.bounding_boxes[field_name]
        
        # Navigate to the page if needed
        if box['page'] != self.current_page:
            self.current_page = box['page']
            self.update_page_label()
            
        # Render the page and draw the bounding box
        self.render_page()
        self.draw_bounding_box(field_name, box)
        
        self.status_var.set(f"Showing bounding box: {field_name}")
    
    def show_all_boxes_on_page(self):
        """Show all bounding boxes on the current page"""
        if not self.pdf_document:
            return
            
        # Clear existing boxes
        self.current_displayed_boxes = []
            
        # Render the page
        self.render_page()
        
        # Draw all boxes on the current page
        count = 0
        for field_name, box in self.bounding_boxes.items():
            if box['page'] == self.current_page:
                self.draw_bounding_box(field_name, box)
                count += 1
                
        self.status_var.set(f"Showing {count} bounding boxes on page {self.current_page + 1}")
    
    def clear_all_boxes(self):
        """Clear all displayed bounding boxes"""
        self.current_displayed_boxes = []
        self.render_page()
        self.status_var.set("Cleared all displayed boxes")
    
    def clear_bbox_tree(self):
        """Clear the bounding box tree view"""
        for item in self.bbox_tree.get_children():
            self.bbox_tree.delete(item)
    
    def update_page_label(self):
        self.page_label.config(text=f"Page {self.current_page + 1} of {self.total_pages}")
    
    def prev_page(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.update_page_label()
            self.render_page()
            self.show_all_boxes_on_page()
    
    def next_page(self):
        if self.pdf_document and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.update_page_label()
            self.render_page()
            self.show_all_boxes_on_page()
    
    def zoom_in(self):
        self.zoom_level *= 1.2
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        self.render_page()
        self.show_all_boxes_on_page()
    
    def zoom_out(self):
        self.zoom_level /= 1.2
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        self.render_page()
        self.show_all_boxes_on_page()
    
    def render_page(self):
        if not self.pdf_document:
            return
            
        page = self.pdf_document[self.current_page]
        
        # Get page dimensions and scale factor
        page_rect = page.rect
        
        # Render the page to a pixmap
        mat = fitz.Matrix(2 * self.zoom_level, 2 * self.zoom_level)  # Increase resolution for better quality
        pix = page.get_pixmap(matrix=mat, alpha=False)
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Store image data for reference
        self.image_tk = ImageTk.PhotoImage(img)
        self.img_width = pix.width
        self.img_height = pix.height
        self.page_width = page_rect.width
        self.page_height = page_rect.height
        
        # Clear canvas and display image
        self.canvas.delete("all")
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.image_tk)
        
        # Reset the displayed boxes list
        self.current_displayed_boxes = []
    
    def draw_bounding_box(self, field_name, box_info):
        """Draw a bounding box on the canvas"""
        if not self.pdf_document or box_info['page'] != self.current_page:
            return
            
        # Scale coordinates to match the displayed image
        scale_x = self.img_width / self.page_width
        scale_y = self.img_height / self.page_height
        
        # Convert coordinates to image space
        x0_px = box_info['left'] * scale_x
        x1_px = box_info['right'] * scale_x
        y0_px = box_info['top'] * scale_y
        y1_px = box_info['bottom'] * scale_y
        
        # Generate a random color for this box based on the field name
        # This ensures consistent colors for the same fields
        import hashlib
        hash_val = int(hashlib.md5(field_name.encode()).hexdigest(), 16)
        r = (hash_val & 0xFF0000) >> 16
        g = (hash_val & 0x00FF00) >> 8
        b = hash_val & 0x0000FF
        color = f"#{r:02x}{g:02x}{b:02x}"
        
        # Draw rectangle
        rect_id = self.canvas.create_rectangle(
            x0_px, y0_px, x1_px, y1_px,
            outline=color, width=2, tags=f"bbox_{field_name}"
        )
        
        # Add semi-transparent fill
        fill_id = self.canvas.create_rectangle(
            x0_px, y0_px, x1_px, y1_px,
            fill=color, stipple="gray50", outline="", tags=f"bbox_fill_{field_name}"
        )
        
        # Add label
        label_id = self.canvas.create_text(
            x0_px + 5, y0_px + 15,
            text=field_name, anchor=tk.W, fill="black", 
            tags=f"bbox_label_{field_name}"
        )
        
        # Create white background for the label
        bbox = self.canvas.bbox(label_id)
        bg_id = self.canvas.create_rectangle(
            bbox, fill="white", outline="", tags=f"bbox_label_bg_{field_name}"
        )
        self.canvas.lower(bg_id, label_id)
        
        # Store the displayed box info
        self.current_displayed_boxes.append((field_name, rect_id, fill_id, label_id, bg_id))
    
    def mouse_move(self, event):
        if not self.pdf_document:
            return
            
        # Convert screen coordinates to PDF coordinates
        pdf_x = event.x / (self.img_width / self.page_width)
        pdf_y = event.y / (self.img_height / self.page_height)
        
        self.status_var.set(f"PDF Coordinates: x={pdf_x:.1f}, y={pdf_y:.1f}")
    
    def mouse_click(self, event):
        if not self.pdf_document:
            return
            
        # Convert screen coordinates to PDF coordinates
        pdf_x = event.x / (self.img_width / self.page_width)
        pdf_y = event.y / (self.img_height / self.page_height)
        
        # Auto-fill the nearest coordinate field based on current input focus
        focused = self.master.focus_get()
        if isinstance(focused, tk.Entry):
            focused.delete(0, tk.END)
            focused.insert(0, f"{pdf_x:.1f}" if "x" in str(focused) else f"{pdf_y:.1f}")
            
        self.status_var.set(f"Clicked at PDF Coordinates: x={pdf_x:.1f}, y={pdf_y:.1f}")


def main():
    root = tk.Tk()
    app = CombinedBoundingBoxPreviewer(root)
    root.mainloop()


if __name__ == "__main__":
    main()