import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from pathlib import Path
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

class TransportSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Around Noon Transport Sorter")
        self.root.geometry("1000x800")
        self.root.configure(bg='#f8f9fa')
        
        # Configure modern styling
        self.setup_styles()
        
        # Application data
        self.data_values = []
        self.delivery_data_values = []
        self.json_file = "transport_data.json"
        self.delivery_json_file = "delivery_sequence_data.json"
        
        # Setup the UI
        self.setup_ui()
        
        # Load existing data if available
        self.load_existing_data()
        self.load_existing_delivery_data()
        
        # Update status to show both datasets if available
        self.update_combined_status()
    
    def setup_styles(self):
        """Configure ultra-professional styling for the application"""
        style = ttk.Style()
        
        # Configure modern theme
        style.theme_use('clam')
        
        # Define sophisticated professional color scheme
        colors = {
            'primary': '#1a1d29',        # Deep professional navy
            'primary_light': '#2c3142',   # Lighter navy
            'secondary': '#6c757d',       # Professional gray
            'accent': '#0d6efd',          # Professional blue
            'accent_light': '#6cb2ff',    # Light blue
            'success': '#198754',         # Professional green
            'warning': '#fd7e14',         # Professional orange
            'text': '#212529',            # Rich dark text
            'text_muted': '#6c757d',      # Muted text
            'background': '#f8f9fa',      # Clean light background
            'surface': '#ffffff',         # Pure white surface
            'border': '#dee2e6',          # Subtle border
            'border_light': '#e9ecef',    # Very light border
            'shadow': '#00000010'         # Subtle shadow
        }
        
        # Premium button styles with gradients and shadows
        style.configure('Modern.TButton',
                       background=colors['accent'],
                       foreground=colors['surface'],
                       padding=(18, 10),
                       font=('Segoe UI', 9, 'normal'),
                       borderwidth=0,
                       focuscolor='none',
                       relief='flat')
        
        style.map('Modern.TButton',
                 background=[('active', colors['accent_light']),
                           ('pressed', colors['primary']),
                           ('disabled', colors['border'])])
        
        # Premium primary action button
        style.configure('Primary.TButton',
                       background=colors['primary'],
                       foreground=colors['surface'],
                       padding=(20, 12),
                       font=('Segoe UI', 10, 'bold'),
                       borderwidth=0,
                       focuscolor='none',
                       relief='flat')
        
        style.map('Primary.TButton',
                 background=[('active', colors['primary_light']),
                           ('pressed', colors['text']),
                           ('disabled', colors['secondary'])])
        
        # Premium secondary button
        style.configure('Secondary.TButton',
                       background=colors['surface'],
                       foreground=colors['primary'],
                       padding=(18, 10),
                       font=('Segoe UI', 9, 'normal'),
                       borderwidth=1,
                       focuscolor='none',
                       relief='solid')
        
        style.map('Secondary.TButton',
                 background=[('active', colors['background']),
                           ('pressed', colors['border_light'])],
                 bordercolor=[('active', colors['accent']),
                            ('pressed', colors['primary'])])
        
        # Premium label frame styles
        style.configure('Modern.TLabelframe',
                       background=colors['surface'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=colors['border_light'])
        
        style.configure('Modern.TLabelframe.Label',
                       background=colors['surface'],
                       foreground=colors['primary'],
                       font=('Segoe UI', 11, 'bold'))
        
        # Premium entry styles
        style.configure('Modern.TEntry',
                       fieldbackground=colors['surface'],
                       foreground=colors['text'],
                       bordercolor=colors['border'],
                       padding=(12, 8),
                       font=('Segoe UI', 9),
                       borderwidth=1,
                       relief='solid')
        
        style.map('Modern.TEntry',
                 bordercolor=[('focus', colors['accent']),
                            ('active', colors['accent_light'])])
        
        # Premium combobox styles
        style.configure('Modern.TCombobox',
                       fieldbackground=colors['surface'],
                       foreground=colors['text'],
                       bordercolor=colors['border'],
                       padding=(12, 8),
                       font=('Segoe UI', 9),
                       borderwidth=1,
                       relief='solid',
                       arrowcolor=colors['primary'])
        
        style.map('Modern.TCombobox',
                 bordercolor=[('focus', colors['accent']),
                            ('active', colors['accent_light'])],
                 fieldbackground=[('readonly', colors['background'])])
        
        # Premium treeview styles
        style.configure('Modern.Treeview',
                       background=colors['surface'],
                       foreground=colors['text'],
                       rowheight=32,
                       font=('Segoe UI', 9),
                       borderwidth=1,
                       relief='solid',
                       bordercolor=colors['border_light'])
        
        style.configure('Modern.Treeview.Heading',
                       background=colors['background'],
                       foreground=colors['primary'],
                       font=('Segoe UI', 9, 'bold'),
                       borderwidth=1,
                       relief='solid',
                       bordercolor=colors['border'])
        
        style.map('Modern.Treeview',
                 background=[('selected', colors['accent_light']),
                           ('focus', colors['accent_light'])],
                 foreground=[('selected', colors['surface'])])
        
        # Premium progress bar
        style.configure('Modern.Horizontal.TProgressbar',
                       background=colors['accent'],
                       troughcolor=colors['border_light'],
                       borderwidth=0,
                       lightcolor=colors['accent'],
                       darkcolor=colors['accent'])
        
        # Premium frame styles
        style.configure('Modern.TFrame',
                       background=colors['background'])
        
        # Premium notebook styles with enhanced tabs
        style.configure('Modern.TNotebook',
                       background=colors['background'],
                       borderwidth=0,
                       tabmargins=[0, 0, 0, 0])
        
        style.configure('Modern.TNotebook.Tab',
                       background=colors['surface'],
                       foreground=colors['text_muted'],
                       padding=(24, 16),
                       font=('Segoe UI', 10, 'normal'),
                       borderwidth=1,
                       relief='solid',
                       bordercolor=colors['border_light'])
        
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', colors['primary']),
                           ('active', colors['primary_light'])],
                 foreground=[('selected', colors['surface']),
                           ('active', colors['surface'])],
                 bordercolor=[('selected', colors['primary']),
                            ('active', colors['primary_light'])])
        
        # Premium label styles
        style.configure('Title.TLabel',
                       background=colors['surface'],
                       foreground=colors['primary'],
                       font=('Segoe UI', 12, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=colors['surface'],
                       foreground=colors['text_muted'],
                       font=('Segoe UI', 9, 'normal'))
        
        style.configure('Field.TLabel',
                       background=colors['surface'],
                       foreground=colors['text'],
                       font=('Segoe UI', 9, 'normal'))
    
    def update_combined_status(self):
        """Update status to show both transport and delivery data if loaded"""
        transport_count = len(self.data_values)
        delivery_count = len(self.delivery_data_values)
        
        if transport_count > 0 and delivery_count > 0:
            self.status_var.set(f"Ready: {transport_count} transport records and {delivery_count} delivery sequence records loaded")
            self.status_label.config(fg="#2a9d8f")  # Professional teal
        elif transport_count > 0:
            self.status_var.set(f"Partial: {transport_count} transport records loaded (missing delivery sequence)")
            self.status_label.config(fg="#8d99ae")  # Muted gray
        elif delivery_count > 0:
            self.status_var.set(f"Partial: {delivery_count} delivery sequence records loaded (missing transport data)")
            self.status_label.config(fg="#8d99ae")  # Muted gray
        else:
            self.status_var.set("Ready - Load your data files to get started")
            self.status_label.config(fg="#457b9d")  # Professional blue
    
    def setup_ui(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.configure(style='Modern.TFrame')
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Premium professional header with gradient effect
        header_frame = tk.Frame(main_frame, bg='#1a1d29', height=90)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 25))
        header_frame.grid_propagate(False)
        
        # Create inner header container for better layout
        header_content = tk.Frame(header_frame, bg='#1a1d29')
        header_content.pack(expand=True, fill='both', padx=30, pady=15)
        
        # Premium title with enhanced typography
        title_label = tk.Label(header_content, 
                              text="Transport Sorter", 
                              font=('Segoe UI', 24, 'bold'),
                              fg='#ffffff', 
                              bg='#1a1d29')
        title_label.pack(anchor='w')
        
        # Enhanced subtitle with better spacing
        subtitle_label = tk.Label(header_content,
                                 text="Professional PDF Processing for Delivery Routes",
                                 font=('Segoe UI', 11),
                                 fg='#a8b3cf',
                                 bg='#1a1d29')
        subtitle_label.pack(anchor='w', pady=(2, 0))
        
        # Add subtle accent line
        accent_line = tk.Frame(header_content, bg='#0d6efd', height=3)
        accent_line.pack(fill='x', pady=(8, 0))
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame, style='Modern.TNotebook')
        self.notebook.grid(row=1, column=0, sticky="nsew")
        
        # Create Main tab
        self.main_tab = ttk.Frame(self.notebook, style='Modern.TFrame', padding="20")
        self.notebook.add(self.main_tab, text="Main")
        
        # Create Settings tab  
        self.settings_tab = ttk.Frame(self.notebook, style='Modern.TFrame', padding="20")
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Setup Main tab content
        self.setup_main_tab()
        
        # Setup Settings tab content
        self.setup_settings_tab()
        
        # Premium status section (outside tabs, at bottom)
        status_frame = tk.Frame(main_frame, bg='#f8f9fa')
        status_frame.grid(row=2, column=0, sticky="ew", pady=(25, 0))
        status_frame.columnconfigure(0, weight=1)
        
        # Status header with icon
        status_header = tk.Label(status_frame, text="📊 System Status", 
                                font=('Segoe UI', 11, 'bold'),
                                fg='#1a1d29', bg='#f8f9fa')
        status_header.grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # Premium status content frame with enhanced styling
        status_content_frame = tk.Frame(status_frame, bg='#ffffff', relief='solid', bd=1)
        status_content_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        status_content_frame.columnconfigure(0, weight=1)
        
        # Add subtle top border accent
        status_accent = tk.Frame(status_content_frame, bg='#0d6efd', height=2)
        status_accent.grid(row=0, column=0, sticky="ew")
        
        self.status_var = tk.StringVar()
        self.status_var.set("🟢 Ready - Load your data files to get started")
        self.status_label = tk.Label(status_content_frame, textvariable=self.status_var, 
                                    fg="#1a1d29", bg='#ffffff', 
                                    font=('Segoe UI', 9),
                                    anchor='w', padx=20, pady=15)
        self.status_label.grid(row=1, column=0, sticky="ew")
        
        # Premium progress bar with enhanced styling
        self.progress = ttk.Progressbar(status_content_frame, mode='indeterminate',
                                       style='Modern.Horizontal.TProgressbar')
        self.progress.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 15))
    
    def setup_main_tab(self):
        """Setup the main workflow tab"""
        self.main_tab.columnconfigure(0, weight=1)
        self.main_tab.rowconfigure(1, weight=1)
        self.main_tab.rowconfigure(3, weight=1)
        
        # Clean delivery section
        delivery_frame_section = ttk.LabelFrame(self.main_tab, text="Step 1: Load Delivery Sequence", 
                                               padding="20", style='Modern.TLabelframe')
        delivery_frame_section.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        delivery_frame_section.columnconfigure(1, weight=1)
        
        # File selection row with enhanced styling
        delivery_file_label = ttk.Label(delivery_frame_section, text="Sequence File:", 
                                       style='Field.TLabel')
        delivery_file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.delivery_file_var = tk.StringVar()
        self.delivery_entry = ttk.Entry(delivery_frame_section, textvariable=self.delivery_file_var, 
                                       state="readonly", style='Modern.TEntry')
        self.delivery_entry.grid(row=0, column=1, sticky="ew", padx=(0, 15))
        
        self.browse_delivery_btn = ttk.Button(delivery_frame_section, text="📁 Browse", 
                                             command=self.browse_delivery_file,
                                             style='Modern.TButton',
                                             cursor="hand2")
        self.browse_delivery_btn.grid(row=0, column=2)
        
        # Column selection row with enhanced styling
        delivery_column_label = ttk.Label(delivery_frame_section, text="Column:", 
                                         style='Field.TLabel')
        delivery_column_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 15), pady=(20, 0))
        
        self.delivery_column_var = tk.StringVar()
        self.delivery_column_combo = ttk.Combobox(delivery_frame_section, textvariable=self.delivery_column_var, 
                                                 state="readonly", width=30, style='Modern.TCombobox')
        self.delivery_column_combo.grid(row=1, column=1, sticky="w", padx=(0, 15), pady=(20, 0))
        
        self.preview_delivery_btn = ttk.Button(delivery_frame_section, text="👁️ Preview", 
                                              command=self.preview_delivery_columns,
                                              style='Secondary.TButton',
                                              cursor="hand2")
        self.preview_delivery_btn.grid(row=1, column=2, pady=(20, 0))
        
        self.load_delivery_btn = ttk.Button(delivery_frame_section, text="⚡ Load Sequence Data", 
                                           command=self.load_delivery_file,
                                           style='Primary.TButton',
                                           cursor="hand2")
        self.load_delivery_btn.grid(row=2, column=0, columnspan=3, pady=(20, 0))
        
        # Clean delivery preview section
        delivery_data_frame = ttk.LabelFrame(self.main_tab, text="Delivery Sequence Preview", 
                                            padding="20", style='Modern.TLabelframe')
        delivery_data_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        delivery_data_frame.columnconfigure(0, weight=1)
        delivery_data_frame.rowconfigure(0, weight=1)
        
        # Clean delivery treeview
        self.delivery_tree = ttk.Treeview(delivery_data_frame, columns=("Value",), show="tree headings", 
                                         height=6, style='Modern.Treeview')
        self.delivery_tree.heading("#0", text="Index")
        self.delivery_tree.heading("Value", text="Sequence")
        self.delivery_tree.column("#0", width=80, anchor='center')
        self.delivery_tree.column("Value", width=400)
        
        # Scrollbar for delivery treeview
        delivery_scrollbar = ttk.Scrollbar(delivery_data_frame, orient="vertical", command=self.delivery_tree.yview)
        self.delivery_tree.configure(yscrollcommand=delivery_scrollbar.set)
        
        self.delivery_tree.grid(row=0, column=0, sticky="nsew")
        delivery_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Clean PDF processing section
        pdf_frame = ttk.LabelFrame(self.main_tab, text="Step 2: Process PDF Files", 
                                  padding="20", style='Modern.TLabelframe')
        pdf_frame.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        pdf_frame.columnconfigure(1, weight=1)
        
        # PDF files selection row with enhanced styling
        pdf_label = ttk.Label(pdf_frame, text="PDF Files:", 
                             style='Field.TLabel')
        pdf_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.pdf_files_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_files_var, 
                                  state="readonly", style='Modern.TEntry')
        self.pdf_entry.grid(row=0, column=1, sticky="ew", padx=(0, 15))
        
        self.browse_pdf_btn = ttk.Button(pdf_frame, text="📄 Browse PDFs", 
                                        command=self.browse_pdf_files,
                                        style='Modern.TButton',
                                        cursor="hand2")
        self.browse_pdf_btn.grid(row=0, column=2)
        
        # Output directory selection row with enhanced styling
        save_label = ttk.Label(pdf_frame, text="Output Folder:", 
                              style='Field.TLabel')
        save_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 15), pady=(20, 0))
        
        self.output_dir_var = tk.StringVar()
        self.output_entry = ttk.Entry(pdf_frame, textvariable=self.output_dir_var, 
                                     state="readonly", style='Modern.TEntry')
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=(0, 15), pady=(20, 0))
        
        self.browse_output_btn = ttk.Button(pdf_frame, text="📂 Select Folder", 
                                           command=self.browse_output_directory,
                                           style='Secondary.TButton',
                                           cursor="hand2")
        self.browse_output_btn.grid(row=1, column=2, pady=(20, 0))
        
        self.process_pdf_btn = ttk.Button(pdf_frame, text="🚀 Process PDFs", 
                                         command=self.process_pdf_files,
                                         style='Primary.TButton',
                                         cursor="hand2")
        self.process_pdf_btn.grid(row=2, column=0, columnspan=3, pady=(20, 0))
    
    def setup_settings_tab(self):
        """Setup the settings/configuration tab"""
        self.settings_tab.columnconfigure(0, weight=1)
        self.settings_tab.rowconfigure(1, weight=1)
        
        # Clean data section
        data_frame_section = ttk.LabelFrame(self.settings_tab, text="Transport Data Configuration", 
                                           padding="20", style='Modern.TLabelframe')
        data_frame_section.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        data_frame_section.columnconfigure(1, weight=1)
        
        # File selection row with enhanced styling
        file_label = ttk.Label(data_frame_section, text="Transport Data File:", 
                              style='Field.TLabel')
        file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.data_file_var = tk.StringVar()
        self.data_entry = ttk.Entry(data_frame_section, textvariable=self.data_file_var, 
                                   state="readonly", style='Modern.TEntry')
        self.data_entry.grid(row=0, column=1, sticky="ew", padx=(0, 15))
        
        self.browse_data_btn = ttk.Button(data_frame_section, text="📊 Browse Data", 
                                         command=self.browse_data_file,
                                         style='Modern.TButton',
                                         cursor="hand2")
        self.browse_data_btn.grid(row=0, column=2)
        
        # Column selection row with enhanced styling
        column_label = ttk.Label(data_frame_section, text="Data Column:", 
                                style='Field.TLabel')
        column_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 15), pady=(20, 0))
        
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(data_frame_section, textvariable=self.column_var, 
                                        state="readonly", width=30, style='Modern.TCombobox')
        self.column_combo.grid(row=1, column=1, sticky="w", padx=(0, 15), pady=(20, 0))
        
        self.preview_btn = ttk.Button(data_frame_section, text="👁️ Preview", 
                                     command=self.preview_file_columns,
                                     style='Secondary.TButton',
                                     cursor="hand2")
        self.preview_btn.grid(row=1, column=2, pady=(20, 0))
        
        self.load_data_btn = ttk.Button(data_frame_section, text="💾 Load Transport Data", 
                                       command=self.load_data_file,
                                       style='Primary.TButton',
                                       cursor="hand2")
        self.load_data_btn.grid(row=2, column=0, columnspan=3, pady=(20, 0))
        
        # Clean data preview section
        data_frame = ttk.LabelFrame(self.settings_tab, text="Transport Data Preview", 
                                   padding="20", style='Modern.TLabelframe')
        data_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
        # Clean treeview
        self.tree = ttk.Treeview(data_frame, columns=("Value",), show="tree headings", 
                                height=6, style='Modern.Treeview')
        self.tree.heading("#0", text="Index")
        self.tree.heading("Value", text="Value")
        self.tree.column("#0", width=80, anchor='center')
        self.tree.column("Value", width=400)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
    
    def browse_data_file(self):
        """Browse and select Excel or CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.data_file_var.set(file_path)
            # Clear previous column selection
            self.column_combo['values'] = ()
            self.column_var.set("")
    
    def preview_file_columns(self):
        """Preview the columns in the selected file"""
        data_file = self.data_file_var.get()
        if not data_file:
            messagebox.showerror("Error", "Please select a data file first!")
            return
        
        try:
            # Determine file type and read just the first few rows
            file_extension = os.path.splitext(data_file)[1].lower()
            
            if file_extension in ['.xlsx', '.xls']:
                # Read Excel file (just first row for headers)
                df = pd.read_excel(data_file, nrows=0)
            elif file_extension == '.csv':
                # Read CSV file (just first row for headers)
                df = self.read_csv_with_encoding(data_file)
                df = df.head(0)  # Keep only headers
            else:
                messagebox.showerror("Error", "Unsupported file format!")
                return
            
            # Get column names
            columns = df.columns.tolist()
            
            # Create display names for columns (index + name)
            column_options = []
            for i, col in enumerate(columns):
                # Show both index and column name for clarity
                display_name = f"Column {i+1}: {str(col)}"
                column_options.append(display_name)
            
            # Update combobox
            self.column_combo['values'] = column_options
            
            # Select first column by default
            if column_options:
                self.column_var.set(column_options[0])
            
            self.status_var.set(f"Found {len(columns)} columns in the file")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to preview columns:\n{str(e)}")
            self.status_var.set("Error previewing columns")
    
    def load_data_file(self):
        """Load data from Excel or CSV file from selected column and save as JSON"""
        data_file = self.data_file_var.get()
        if not data_file:
            messagebox.showerror("Error", "Please select a data file first!")
            return
        
        selected_column = self.column_var.get()
        if not selected_column:
            messagebox.showerror("Error", "Please select a column first! Click 'Preview Columns' to see available columns.")
            return
        
        try:
            self.status_var.set("Loading data...")
            self.progress.start()
            
            # Determine file type and read accordingly
            file_extension = os.path.splitext(data_file)[1].lower()
            
            if file_extension in ['.xlsx', '.xls']:
                # Read Excel file
                df = pd.read_excel(data_file)
                file_type = "Excel"
            elif file_extension == '.csv':
                # Read CSV file with encoding detection
                df = self.read_csv_with_encoding(data_file)
                file_type = "CSV"
            else:
                messagebox.showerror("Error", "Unsupported file format! Please select an Excel (.xlsx, .xls) or CSV (.csv) file.")
                return
            
            # Get selected column values
            if df.empty:
                messagebox.showerror("Error", f"{file_type} file is empty!")
                return
            
            # Extract column index from selection (format: "Column X: Name")
            column_index = int(selected_column.split(":")[0].replace("Column ", "")) - 1
            column_name = df.columns[column_index]
            
            # Extract selected column values, skip empty cells
            column_values = df.iloc[:, column_index].dropna().astype(str).tolist()
            
            if not column_values:
                messagebox.showerror("Error", f"No data found in the selected column ({column_name})!")
                return
            
            self.data_values = column_values
            
            # Save to JSON
            data_to_save = {
                "source_file": data_file,
                "file_type": file_type,
                "selected_column": selected_column,
                "column_name": column_name,
                "column_index": column_index,
                "column_values": self.data_values,
                "total_records": len(self.data_values),
                "created_date": pd.Timestamp.now().isoformat()
            }
            
            with open(self.json_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, indent=2, ensure_ascii=False)
            
            # Update display
            self.update_data_display()
            
            self.status_var.set(f"Successfully loaded {len(self.data_values)} records from {file_type} file, column '{column_name}' and saved to JSON")
            messagebox.showinfo("Success", f"Data loaded successfully!\n{len(self.data_values)} records from {file_type} file, column '{column_name}' saved to {self.json_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data file:\n{str(e)}")
            self.status_var.set("Error loading data file")
        finally:
            self.progress.stop()
    
    def read_csv_with_encoding(self, file_path):
        """Read CSV file with automatic encoding detection"""
        encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings_to_try:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                print(f"Successfully read CSV with encoding: {encoding}")
                return df
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        # If all encodings fail, try with error handling
        try:
            df = pd.read_csv(file_path, encoding='utf-8', errors='replace')
            print("Read CSV with error replacement")
            return df
        except Exception as e:
            raise Exception(f"Could not read CSV file with any encoding. Last error: {str(e)}")
    
    def read_csv_with_encoding_no_header(self, file_path, nrows=None):
        """Read CSV file with automatic encoding detection and no headers"""
        encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings_to_try:
            try:
                df = pd.read_csv(file_path, encoding=encoding, header=None, nrows=nrows)
                print(f"Successfully read CSV (no headers) with encoding: {encoding}")
                return df
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        # If all encodings fail, try with error handling
        try:
            df = pd.read_csv(file_path, encoding='utf-8', errors='replace', header=None, nrows=nrows)
            print("Read CSV (no headers) with error replacement")
            return df
        except Exception as e:
            raise Exception(f"Could not read CSV file with any encoding. Last error: {str(e)}")
    
    def update_data_display(self):
        """Update the treeview with loaded data"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add new items
        for i, value in enumerate(self.data_values, 1):
            self.tree.insert("", "end", text=str(i), values=(value,))
    
    def load_existing_data(self):
        """Load existing JSON data if available"""
        if os.path.exists(self.json_file):
            try:
                with open(self.json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Handle both old and new JSON formats
                    self.data_values = data.get('column_values', data.get('column_a_values', []))
                    
                    # Also restore the file path and column selection for convenience
                    source_file = data.get('source_file', '')
                    selected_column = data.get('selected_column', '')
                    
                    if source_file and os.path.exists(source_file):
                        self.data_file_var.set(source_file)
                        
                    if selected_column:
                        # We need to load the columns first to populate the dropdown
                        try:
                            self.preview_file_columns()
                            self.column_var.set(selected_column)
                        except:
                            pass  # If preview fails, just skip setting the column
                    
                    self.update_data_display()
                    file_type = data.get('file_type', 'Unknown')
                    column_info = data.get('column_name', 'Column A')
                    
                    if self.data_values:
                        print(f"Auto-loaded transport data: {len(self.data_values)} records from {file_type} file, column '{column_info}'")
                        self.status_var.set(f"Auto-loaded transport data: {len(self.data_values)} records from {file_type} file, column '{column_info}'")
            except Exception as e:
                print(f"Error loading existing data: {e}")
    
    def browse_delivery_file(self):
        """Browse and select Delivery Sequence Excel or CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select Delivery Sequence File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.delivery_file_var.set(file_path)
            # Clear previous column selection
            self.delivery_column_combo['values'] = ()
            self.delivery_column_var.set("")
    
    def preview_delivery_columns(self):
        """Preview the columns in the selected delivery sequence file"""
        data_file = self.delivery_file_var.get()
        if not data_file:
            messagebox.showerror("Error", "Please select a delivery sequence file first!")
            return
        
        try:
            # Determine file type and read just the first few rows for preview
            file_extension = os.path.splitext(data_file)[1].lower()
            
            if file_extension in ['.xlsx', '.xls']:
                # Read Excel file without headers (header=None) to handle files without headers
                df = pd.read_excel(data_file, header=None, nrows=5)
            elif file_extension == '.csv':
                # Read CSV file without headers (header=None) to handle files without headers
                df = self.read_csv_with_encoding_no_header(data_file, nrows=5)
            else:
                messagebox.showerror("Error", "Unsupported file format!")
                return
            
            if df.empty:
                messagebox.showerror("Error", "File appears to be empty!")
                return
            
            # Get number of columns
            num_columns = len(df.columns)
            
            # Create display names for columns showing first value as sample
            column_options = []
            for i in range(num_columns):
                # Show column number and first value as sample (if available)
                sample_value = "Empty" if df.iloc[0, i] is None or pd.isna(df.iloc[0, i]) else str(df.iloc[0, i])
                if len(sample_value) > 20:
                    sample_value = sample_value[:17] + "..."
                display_name = f"Column {i+1}: {sample_value}"
                column_options.append(display_name)
            
            # Update combobox
            self.delivery_column_combo['values'] = column_options
            
            # Select first column by default
            if column_options:
                self.delivery_column_var.set(column_options[0])
            
            self.status_var.set(f"Found {len(column_options)} columns in the delivery sequence file (no headers detected)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to preview delivery columns:\n{str(e)}")
            self.status_var.set("Error previewing delivery columns")
    
    def load_delivery_file(self):
        """Load data from Delivery Sequence Excel or CSV file from selected column and save as JSON"""
        data_file = self.delivery_file_var.get()
        if not data_file:
            messagebox.showerror("Error", "Please select a delivery sequence file first!")
            return
        
        selected_column = self.delivery_column_var.get()
        if not selected_column:
            messagebox.showerror("Error", "Please select a column first! Click 'Preview Columns' to see available columns.")
            return
        
        try:
            self.status_var.set("Loading delivery sequence data...")
            self.progress.start()
            
            # Determine file type and read accordingly
            file_extension = os.path.splitext(data_file)[1].lower()
            
            if file_extension in ['.xlsx', '.xls']:
                # Read Excel file without headers (header=None) to handle files without headers
                df = pd.read_excel(data_file, header=None)
                file_type = "Excel"
            elif file_extension == '.csv':
                # Read CSV file with encoding detection and no headers
                df = self.read_csv_with_encoding_no_header(data_file)
                file_type = "CSV"
            else:
                messagebox.showerror("Error", "Unsupported file format! Please select an Excel (.xlsx, .xls) or CSV (.csv) file.")
                return
            
            # Get selected column values
            if df.empty:
                messagebox.showerror("Error", f"Delivery sequence {file_type} file is empty!")
                return
            
            # Extract column index from selection (format: "Column X: Name")
            column_index = int(selected_column.split(":")[0].replace("Column ", "")) - 1
            
            # Since we're not using headers, column name is just the column number
            column_name = f"Column_{column_index + 1}"
            
            # Extract selected column values, skip empty cells
            column_values = df.iloc[:, column_index].dropna().astype(str).tolist()
            
            if not column_values:
                messagebox.showerror("Error", f"No data found in the selected delivery sequence column (Column {column_index + 1})!")
                return
            
            self.delivery_data_values = column_values
            
            # Save to JSON
            data_to_save = {
                "source_file": data_file,
                "file_type": file_type,
                "selected_column": selected_column,
                "column_name": column_name,
                "column_index": column_index,
                "column_values": self.delivery_data_values,
                "total_records": len(self.delivery_data_values),
                "created_date": pd.Timestamp.now().isoformat()
            }
            
            with open(self.delivery_json_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, indent=2, ensure_ascii=False)
            
            # Update display
            self.update_delivery_display()
            
            self.status_var.set(f"Successfully loaded {len(self.delivery_data_values)} delivery sequence records from {file_type} file, column {column_index + 1} and saved to JSON")
            messagebox.showinfo("Success", f"Delivery sequence data loaded successfully!\n{len(self.delivery_data_values)} records from {file_type} file, column {column_index + 1} saved to {self.delivery_json_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load delivery sequence file:\n{str(e)}")
            self.status_var.set("Error loading delivery sequence file")
        finally:
            self.progress.stop()
    
    def update_delivery_display(self):
        """Update the delivery treeview with loaded data"""
        # Clear existing items
        for item in self.delivery_tree.get_children():
            self.delivery_tree.delete(item)
        
        # Add new items
        for i, value in enumerate(self.delivery_data_values, 1):
            self.delivery_tree.insert("", "end", text=str(i), values=(value,))
    
    def load_existing_delivery_data(self):
        """Load existing delivery sequence JSON data if available"""
        if os.path.exists(self.delivery_json_file):
            try:
                with open(self.delivery_json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.delivery_data_values = data.get('column_values', [])
                    
                    # Also restore the file path and column selection for convenience
                    source_file = data.get('source_file', '')
                    selected_column = data.get('selected_column', '')
                    
                    if source_file and os.path.exists(source_file):
                        self.delivery_file_var.set(source_file)
                        
                    if selected_column:
                        # We need to load the columns first to populate the dropdown
                        try:
                            self.preview_delivery_columns()
                            self.delivery_column_var.set(selected_column)
                        except:
                            pass  # If preview fails, just skip setting the column
                    
                    self.update_delivery_display()
                    file_type = data.get('file_type', 'Unknown')
                    column_info = data.get('column_name', 'Column A')
                    
                    if self.delivery_data_values:
                        print(f"Auto-loaded delivery data: {len(self.delivery_data_values)} records from {file_type} file, column '{column_info}'")
                        # Update status if this is the primary data load
                        if not self.data_values:  # Only show delivery status if no transport data loaded
                            self.status_var.set(f"Auto-loaded delivery data: {len(self.delivery_data_values)} records from {file_type} file, column '{column_info}'")
            except Exception as e:
                print(f"Error loading existing delivery data: {e}")
    
    def browse_pdf_files(self):
        """Browse and select PDF files"""
        file_paths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_paths:
            self.pdf_files = file_paths
            file_names = [os.path.basename(path) for path in file_paths]
            display_text = f"{len(file_names)} file(s): " + ", ".join(file_names[:3])
            if len(file_names) > 3:
                display_text += f"... and {len(file_names) - 3} more"
            self.pdf_files_var.set(display_text)
    
    def browse_output_directory(self):
        """Browse and select output directory for processed PDFs"""
        directory = filedialog.askdirectory(
            title="Select Output Directory for Processed PDFs"
        )
        if directory:
            self.output_dir_var.set(directory)
    
    def process_pdf_files(self):
        """Process PDF files with OCR"""
        if not hasattr(self, 'pdf_files') or not self.pdf_files:
            messagebox.showerror("Error", "Please select PDF files first!")
            return
        
        if not self.data_values:
            messagebox.showerror("Error", "Please load transport data file first!")
            return
        
        if not self.delivery_data_values:
            messagebox.showerror("Error", "Please load delivery sequence file first! This is required to match PDF content.")
            return
        
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showerror("Error", "Please select an output directory to save the processed PDF files!")
            return
        
        try:
            self.status_var.set("Processing PDF files with OCR...")
            self.progress.start()
            
            # Debug: Show what we're looking for
            print(f"DEBUG: Looking for {len(self.delivery_data_values)} delivery sequence values:")
            for i, val in enumerate(self.delivery_data_values[:10]):  # Show first 10
                print(f"  {i+1}: '{val}'")
            if len(self.delivery_data_values) > 10:
                print(f"  ... and {len(self.delivery_data_values) - 10} more")
            
            # Create results directory in the selected output location with "Delivery" prefix and today's date
            today_date = pd.Timestamp.now().strftime('%Y-%m-%d')
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            results_dir = Path(output_dir) / f"Delivery_{today_date}_{timestamp}"
            results_dir.mkdir(exist_ok=True)
            
            results = []
            
            for pdf_path in self.pdf_files:
                pdf_name = os.path.basename(pdf_path)
                self.status_var.set(f"Processing {pdf_name}...")
                
                # Process PDF with OCR and sequence matching
                pdf_result = self.extract_and_match_pdf_pages(pdf_path)
                
                # Create filtered PDF with only matching pages
                filtered_pdf_path = None
                if pdf_result["matched_pages"]:
                    filtered_pdf_path = self.create_filtered_pdf(pdf_path, pdf_result["matched_pages"], results_dir)
                
                results.append({
                    "file_name": pdf_name,
                    "file_path": pdf_path,
                    "filtered_pdf_path": filtered_pdf_path,
                    "total_pages": pdf_result["total_pages"],
                    "matched_pages": pdf_result["matched_pages"],
                    "matching_page_count": len(pdf_result["matched_pages"]),
                    "found_sequence_values": pdf_result["found_sequence_values"],
                    "processing_date": pd.Timestamp.now().isoformat()
                })
            
            # Save results
            output_file = results_dir / f"processing_results_{timestamp}.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump({
                    "transport_reference_data": self.data_values,
                    "delivery_sequence_data": self.delivery_data_values,
                    "processed_files": results,
                    "total_files": len(results),
                    "has_delivery_data": len(self.delivery_data_values) > 0
                }, f, indent=2, ensure_ascii=False)
            
            # Calculate summary statistics
            total_pages_processed = sum(result["total_pages"] for result in results)
            total_matched_pages = sum(result["matching_page_count"] for result in results)
            filtered_pdfs_created = sum(1 for result in results if result["filtered_pdf_path"])
            all_found_values = set()
            for result in results:
                all_found_values.update(result["found_sequence_values"])
            
            self.status_var.set(f"Successfully processed {len(results)} PDF files - Created {filtered_pdfs_created} filtered PDFs")
            messagebox.showinfo("Success", 
                              f"PDF processing completed!\n"
                              f"Processed {len(results)} files ({total_pages_processed} total pages)\n"
                              f"Found {total_matched_pages} pages with delivery sequence matches\n"
                              f"Created {filtered_pdfs_created} filtered PDF files with matching pages only\n"
                              f"Matched {len(all_found_values)} different sequence values\n"
                              f"Files saved to: {results_dir}\n"
                              f"Processing report: {output_file.name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process PDF files:\n{str(e)}")
            self.status_var.set("Error processing PDF files")
        finally:
            self.progress.stop()
    
    def extract_and_match_pdf_pages(self, pdf_path):
        """Extract text from PDF and find pages matching delivery sequence values"""
        doc = fitz.open(pdf_path)
        matched_pages = []
        found_sequence_values = set()
        total_pages = len(doc)
        
        try:
            for page_num in range(total_pages):
                page = doc.load_page(page_num)
                
                # First try to extract text directly
                text = page.get_text()
                
                # If no text found, use OCR
                if not text.strip():
                    # Convert page to image
                    mat = fitz.Matrix(2, 2)  # 2x zoom for better OCR quality
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("png")
                    
                    # Use PIL to create image
                    image = Image.open(io.BytesIO(img_data))
                    
                    # Use Tesseract OCR
                    try:
                        text = pytesseract.image_to_string(image)
                    except Exception as ocr_error:
                        text = f"OCR Error on page {page_num + 1}: {str(ocr_error)}"
                
                # Check if this page contains any delivery sequence values
                page_matches = self.find_sequence_matches_in_text(text, page_num + 1)
                
                if page_matches["matches_found"]:
                    matched_pages.append(page_matches)
                    found_sequence_values.update(page_matches["matched_values"])
        
        finally:
            doc.close()
        
        return {
            "total_pages": total_pages,
            "matched_pages": matched_pages,
            "found_sequence_values": list(found_sequence_values)
        }
    
    def find_sequence_matches_in_text(self, text, page_number):
        """Find delivery sequence values in the extracted text line by line"""
        matched_values = []
        matched_lines = []
        
        # Split text into lines for line-by-line scanning
        lines = text.split('\n')
        
        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            if not line:  # Skip empty lines
                continue
                
            # Check each delivery sequence value against this line
            for seq_value in self.delivery_data_values:
                # Convert both to string and do case-insensitive comparison
                seq_value_str = str(seq_value).strip()
                if seq_value_str and seq_value_str.lower() in line.lower():
                    matched_values.append(seq_value_str)
                    matched_lines.append({
                        "line_number": line_num,
                        "line_text": line,
                        "matched_sequence_value": seq_value_str
                    })
                    # Debug output
                    print(f"MATCH FOUND on page {page_number}: '{seq_value_str}' in line {line_num}: '{line}'")
        
        if matched_values:
            print(f"Page {page_number}: Found {len(set(matched_values))} unique matches: {list(set(matched_values))}")
        
        return {
            "page_number": page_number,
            "matches_found": len(matched_values) > 0,
            "matched_values": list(set(matched_values)),  # Remove duplicates
            "matched_lines": matched_lines,
            "full_page_text": text
        }
    
    def sort_pages_by_delivery_sequence(self, matched_pages):
        """Sort pages based on the order of delivery sequence values"""
        def get_earliest_sequence_position(page_info):
            """Get the earliest position of any matched value in the delivery sequence"""
            matched_values = page_info["matched_values"]
            earliest_position = float('inf')
            
            for value in matched_values:
                try:
                    # Find the position of this value in the delivery sequence
                    position = self.delivery_data_values.index(value)
                    earliest_position = min(earliest_position, position)
                except ValueError:
                    # Value not found in sequence (shouldn't happen, but just in case)
                    continue
            
            return earliest_position if earliest_position != float('inf') else 999999
        
        # Sort pages by their earliest sequence position
        sorted_pages = sorted(matched_pages, key=get_earliest_sequence_position)
        
        # Debug output
        print("Pages sorted by delivery sequence order:")
        for i, page_info in enumerate(sorted_pages):
            page_num = page_info["page_number"]
            matched_values = page_info["matched_values"]
            earliest_pos = get_earliest_sequence_position(page_info)
            print(f"  {i+1}. Page {page_num}: {matched_values} (earliest position: {earliest_pos})")
        
        return sorted_pages
    
    def create_filtered_pdf(self, original_pdf_path, matched_pages, output_dir):
        """Create a new PDF containing only the pages that matched delivery sequence values"""
        original_doc = None
        filtered_doc = None
        
        try:
            if not matched_pages:
                print(f"No matching pages found for {original_pdf_path}")
                return None
            
            # Sort pages by delivery sequence order instead of page number
            sorted_pages = self.sort_pages_by_delivery_sequence(matched_pages)
            
            # Generate output filename with "Delivery" prefix and today's date
            original_name = os.path.splitext(os.path.basename(original_pdf_path))[0]
            today_date = pd.Timestamp.now().strftime('%Y-%m-%d')
            filtered_filename = f"Delivery_{today_date}_{original_name}_{len(sorted_pages)}pages.pdf"
            filtered_path = output_dir / filtered_filename
            
            print(f"Creating filtered PDF for {original_name} with {len(sorted_pages)} pages in delivery sequence order")
            
            # Open the original PDF with a fresh connection
            original_doc = fitz.open(original_pdf_path)
            
            # Create a new PDF document
            filtered_doc = fitz.open()
            
            # Copy pages in delivery sequence order
            for page_info in sorted_pages:
                page_num = page_info["page_number"]
                sequence_values = page_info["matched_values"]
                
                # PyMuPDF uses 0-based indexing, but our page_number is 1-based
                page_index = page_num - 1
                
                if 0 <= page_index < original_doc.page_count:
                    print(f"Copying page {page_num} (contains: {', '.join(sequence_values)})")
                    # Use insert_pdf to copy the page
                    filtered_doc.insert_pdf(original_doc, from_page=page_index, to_page=page_index)
                else:
                    print(f"Warning: Page {page_num} is out of range (document has {original_doc.page_count} pages)")
            
            # Only save if we actually copied some pages
            if filtered_doc.page_count > 0:
                print(f"Saving filtered PDF with {filtered_doc.page_count} pages to {filtered_path}")
                filtered_doc.save(str(filtered_path))
                success_path = str(filtered_path)
            else:
                print(f"No pages were copied for {original_name}")
                success_path = None
            
            return success_path
            
        except Exception as e:
            print(f"Error creating filtered PDF for {original_pdf_path}: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
        
        finally:
            # Ensure documents are always closed
            if filtered_doc is not None:
                try:
                    filtered_doc.close()
                except:
                    pass
            if original_doc is not None:
                try:
                    original_doc.close()
                except:
                    pass
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF using PyMuPDF and OCR (legacy method)"""
        doc = fitz.open(pdf_path)
        extracted_text = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # First try to extract text directly
            text = page.get_text()
            
            # If no text found, use OCR
            if not text.strip():
                # Convert page to image
                mat = fitz.Matrix(2, 2)  # 2x zoom for better OCR quality
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                
                # Use PIL to create image
                image = Image.open(io.BytesIO(img_data))
                
                # Use Tesseract OCR
                try:
                    text = pytesseract.image_to_string(image)
                except Exception as ocr_error:
                    text = f"OCR Error on page {page_num + 1}: {str(ocr_error)}"
            
            extracted_text.append({
                "page": page_num + 1,
                "text": text.strip()
            })
        
        doc.close()
        
        return {
            "text": extracted_text,
            "page_count": len(doc)
        }

def main():
    root = tk.Tk()
    app = TransportSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()