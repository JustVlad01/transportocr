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
        self.root.title("Transport Sorter")
        self.root.geometry("1200x900")
        self.root.configure(bg='#ffffff')
        
        # Configure premium minimalist styling
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
        """Configure ultra-premium minimalist styling"""
        style = ttk.Style()
        
        # Configure clean theme
        style.theme_use('clam')
        
        # Define sophisticated minimal color palette
        colors = {
            'primary': '#2563eb',         # Clean blue
            'primary_hover': '#1d4ed8',   # Darker blue on hover
            'secondary': '#64748b',       # Subtle gray
            'accent': '#10b981',          # Success green
            'text_primary': '#0f172a',    # Rich black
            'text_secondary': '#64748b',  # Muted gray
            'text_light': '#94a3b8',      # Light gray
            'background': '#ffffff',      # Pure white
            'surface': '#f8fafc',         # Very light gray
            'border': '#e2e8f0',          # Subtle border
            'border_light': '#f1f5f9',    # Very light border
            'error': '#ef4444',           # Clean red
            'warning': '#f59e0b'          # Clean orange
        }
        
        # Premium minimalist button - Primary
        style.configure('Premium.TButton',
                       background=colors['primary'],
                       foreground='white',
                       padding=(24, 14),
                       font=('SF Pro Display', 10, 'normal'),
                       borderwidth=0,
                       focuscolor='none',
                       relief='flat')
        
        style.map('Premium.TButton',
                 background=[('active', colors['primary_hover']),
                           ('pressed', colors['primary_hover']),
                           ('disabled', colors['border'])])
        
        # Premium minimalist button - Secondary  
        style.configure('PremiumSecondary.TButton',
                       background=colors['surface'],
                       foreground=colors['text_primary'],
                       padding=(20, 12),
                       font=('SF Pro Display', 9, 'normal'),
                       borderwidth=1,
                       focuscolor='none',
                       relief='solid')
        
        style.map('PremiumSecondary.TButton',
                 background=[('active', colors['background']),
                           ('pressed', colors['border_light'])],
                 bordercolor=[('active', colors['primary']),
                            ('pressed', colors['primary'])])
        
        # Premium label frame with minimal borders
        style.configure('Premium.TLabelframe',
                       background=colors['background'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=colors['border_light'])
        
        style.configure('Premium.TLabelframe.Label',
                       background=colors['background'],
                       foreground=colors['text_primary'],
                       font=('SF Pro Display', 12, '600'))
        
        # Premium entry fields
        style.configure('Premium.TEntry',
                       fieldbackground=colors['surface'],
                       foreground=colors['text_primary'],
                       bordercolor=colors['border'],
                       padding=(16, 12),
                       font=('SF Pro Text', 10),
                       borderwidth=1,
                       relief='solid',
                       insertcolor=colors['primary'])
        
        style.map('Premium.TEntry',
                 bordercolor=[('focus', colors['primary']),
                            ('active', colors['primary'])])
        
        # Premium combobox
        style.configure('Premium.TCombobox',
                       fieldbackground=colors['surface'],
                       foreground=colors['text_primary'],
                       bordercolor=colors['border'],
                       padding=(16, 12),
                       font=('SF Pro Text', 10),
                       borderwidth=1,
                       relief='solid',
                       arrowcolor=colors['text_secondary'])
        
        style.map('Premium.TCombobox',
                 bordercolor=[('focus', colors['primary']),
                            ('active', colors['primary'])],
                 fieldbackground=[('readonly', colors['surface'])])
        
        # Premium treeview with clean lines
        style.configure('Premium.Treeview',
                       background=colors['background'],
                       foreground=colors['text_primary'],
                       rowheight=40,
                       font=('SF Pro Text', 10),
                       borderwidth=1,
                       relief='solid',
                       bordercolor=colors['border_light'])
        
        style.configure('Premium.Treeview.Heading',
                       background=colors['surface'],
                       foreground=colors['text_secondary'],
                       font=('SF Pro Display', 10, '600'),
                       borderwidth=0,
                       relief='flat')
        
        style.map('Premium.Treeview',
                 background=[('selected', colors['primary']),
                           ('focus', colors['primary'])],
                 foreground=[('selected', 'white')])
        
        # Premium progress bar
        style.configure('Premium.Horizontal.TProgressbar',
                       background=colors['primary'],
                       troughcolor=colors['border_light'],
                       borderwidth=0,
                       lightcolor=colors['primary'],
                       darkcolor=colors['primary'])
        
        # Premium frame
        style.configure('Premium.TFrame',
                       background=colors['background'])
        
        # Premium notebook with clean tabs
        style.configure('Premium.TNotebook',
                       background=colors['background'],
                       borderwidth=0,
                       tabmargins=[0, 0, 0, 0])
        
        style.configure('Premium.TNotebook.Tab',
                       background=colors['surface'],
                       foreground=colors['text_secondary'],
                       padding=(32, 20),
                       font=('SF Pro Display', 11, 'normal'),
                       borderwidth=0,
                       relief='flat')
        
        style.map('Premium.TNotebook.Tab',
                 background=[('selected', colors['background']),
                           ('active', colors['border_light'])],
                 foreground=[('selected', colors['primary']),
                           ('active', colors['text_primary'])])
    
    def update_combined_status(self):
        """Update status with clean minimal indicators"""
        transport_count = len(self.data_values)
        delivery_count = len(self.delivery_data_values)
        
        if transport_count > 0 and delivery_count > 0:
            self.status_var.set(f"Ready • {transport_count} transport • {delivery_count} delivery records")
            self.status_indicator.config(bg="#10b981")  # Success green
        elif transport_count > 0:
            self.status_var.set(f"Partial • {transport_count} transport records • Missing delivery data")
            self.status_indicator.config(bg="#f59e0b")  # Warning orange
        elif delivery_count > 0:
            self.status_var.set(f"Partial • {delivery_count} delivery records • Missing transport data")
            self.status_indicator.config(bg="#f59e0b")  # Warning orange
        else:
            self.status_var.set("Ready to load your data files")
            self.status_indicator.config(bg="#64748b")  # Neutral gray
    
    def setup_ui(self):
        # Main container with generous padding
        main_frame = ttk.Frame(self.root, padding="40")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.configure(style='Premium.TFrame')
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Minimal premium header
        header_frame = tk.Frame(main_frame, bg='#ffffff', height=120)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 40))
        header_frame.grid_propagate(False)
        
        # Header content with clean layout
        header_content = tk.Frame(header_frame, bg='#ffffff')
        header_content.pack(expand=True, fill='both')
        
        # Clean title with better typography
        title_label = tk.Label(header_content, 
                              text="Transport Sorter", 
                              font=('SF Pro Display', 32, '300'),  # Light weight
                              fg='#0f172a', 
                              bg='#ffffff')
        title_label.pack(anchor='w', pady=(20, 0))
        
        # Minimal subtitle
        subtitle_label = tk.Label(header_content,
                                 text="Process delivery routes with precision",
                                 font=('SF Pro Text', 14),
                                 fg='#64748b',
                                 bg='#ffffff')
        subtitle_label.pack(anchor='w', pady=(8, 0))
        
        # Clean divider line
        divider = tk.Frame(header_content, bg='#e2e8f0', height=1)
        divider.pack(fill='x', pady=(24, 0))
        
        # Create clean notebook
        self.notebook = ttk.Notebook(main_frame, style='Premium.TNotebook')
        self.notebook.grid(row=1, column=0, sticky="nsew")
        
        # Create tabs with clean styling
        self.main_tab = ttk.Frame(self.notebook, style='Premium.TFrame', padding="30")
        self.notebook.add(self.main_tab, text="Process")
        
        self.settings_tab = ttk.Frame(self.notebook, style='Premium.TFrame', padding="30")
        self.notebook.add(self.settings_tab, text="Configure")
        
        # Setup tab content
        self.setup_main_tab()
        self.setup_settings_tab()
        
        # Clean status section
        status_frame = tk.Frame(main_frame, bg='#ffffff')
        status_frame.grid(row=2, column=0, sticky="ew", pady=(40, 0))
        status_frame.columnconfigure(1, weight=1)
        
        # Status indicator dot
        self.status_indicator = tk.Frame(status_frame, bg='#64748b', width=12, height=12)
        self.status_indicator.grid(row=0, column=0, padx=(0, 12))
        self.status_indicator.grid_propagate(False)
        
        # Status text
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to load your data files")
        self.status_label = tk.Label(status_frame, textvariable=self.status_var, 
                                    fg="#64748b", bg='#ffffff', 
                                    font=('SF Pro Text', 11),
                                    anchor='w')
        self.status_label.grid(row=0, column=1, sticky="ew")
        
        # Clean progress bar
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate',
                                       style='Premium.Horizontal.TProgressbar')
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(16, 0))
    
    def setup_main_tab(self):
        """Setup the main workflow tab with clean spacing"""
        self.main_tab.columnconfigure(0, weight=1)
        self.main_tab.rowconfigure(1, weight=1)
        self.main_tab.rowconfigure(3, weight=1)
        
        # Step 1: Delivery Sequence - Clean design
        delivery_section = ttk.LabelFrame(self.main_tab, text="Step 1 • Load Delivery Sequence", 
                                         padding="30", style='Premium.TLabelframe')
        delivery_section.grid(row=0, column=0, sticky="ew", pady=(0, 30))
        delivery_section.columnconfigure(1, weight=1)
        
        # File selection with clean layout
        tk.Label(delivery_section, text="Sequence File", 
                font=('SF Pro Text', 11, '500'), fg='#0f172a', bg='#ffffff').grid(
                row=0, column=0, sticky=tk.W, padx=(0, 20), pady=(0, 8))
        
        self.delivery_file_var = tk.StringVar()
        self.delivery_entry = ttk.Entry(delivery_section, textvariable=self.delivery_file_var, 
                                       state="readonly", style='Premium.TEntry')
        self.delivery_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 20))
        
        self.browse_delivery_btn = ttk.Button(delivery_section, text="Browse Files", 
                                             command=self.browse_delivery_file,
                                             style='PremiumSecondary.TButton')
        self.browse_delivery_btn.grid(row=1, column=2)
        
        # Column selection with spacing
        tk.Label(delivery_section, text="Data Column", 
                font=('SF Pro Text', 11, '500'), fg='#0f172a', bg='#ffffff').grid(
                row=2, column=0, sticky=tk.W, padx=(0, 20), pady=(24, 8))
        
        self.delivery_column_var = tk.StringVar()
        self.delivery_column_combo = ttk.Combobox(delivery_section, textvariable=self.delivery_column_var, 
                                                 state="readonly", width=30, style='Premium.TCombobox')
        self.delivery_column_combo.grid(row=3, column=0, sticky="w", padx=(0, 20))
        
        self.preview_delivery_btn = ttk.Button(delivery_section, text="Preview", 
                                              command=self.preview_delivery_columns,
                                              style='PremiumSecondary.TButton')
        self.preview_delivery_btn.grid(row=3, column=1, sticky="w", padx=(20, 0))
        
        self.load_delivery_btn = ttk.Button(delivery_section, text="Load Sequence Data", 
                                           command=self.load_delivery_file,
                                           style='Premium.TButton')
        self.load_delivery_btn.grid(row=4, column=0, columnspan=3, pady=(30, 0), sticky="ew")
        
        # Delivery preview with clean styling
        delivery_preview = ttk.LabelFrame(self.main_tab, text="Sequence Preview", 
                                         padding="30", style='Premium.TLabelframe')
        delivery_preview.grid(row=1, column=0, sticky="nsew", pady=(0, 30))
        delivery_preview.columnconfigure(0, weight=1)
        delivery_preview.rowconfigure(0, weight=1)
        
        # Clean treeview
        self.delivery_tree = ttk.Treeview(delivery_preview, columns=("Value",), show="tree headings", 
                                         height=8, style='Premium.Treeview')
        self.delivery_tree.heading("#0", text="Index")
        self.delivery_tree.heading("Value", text="Sequence")
        self.delivery_tree.column("#0", width=100, anchor='center')
        self.delivery_tree.column("Value", width=500)
        
        delivery_scrollbar = ttk.Scrollbar(delivery_preview, orient="vertical", command=self.delivery_tree.yview)
        self.delivery_tree.configure(yscrollcommand=delivery_scrollbar.set)
        
        self.delivery_tree.grid(row=0, column=0, sticky="nsew")
        delivery_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Step 2: PDF Processing - Clean design
        pdf_section = ttk.LabelFrame(self.main_tab, text="Step 2 • Process PDF Files", 
                                    padding="30", style='Premium.TLabelframe')
        pdf_section.grid(row=2, column=0, sticky="ew", pady=(0, 30))
        pdf_section.columnconfigure(1, weight=1)
        
        # PDF files selection
        tk.Label(pdf_section, text="PDF Files", 
                font=('SF Pro Text', 11, '500'), fg='#0f172a', bg='#ffffff').grid(
                row=0, column=0, sticky=tk.W, padx=(0, 20), pady=(0, 8))
        
        self.pdf_files_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(pdf_section, textvariable=self.pdf_files_var, 
                                  state="readonly", style='Premium.TEntry')
        self.pdf_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 20))
        
        self.browse_pdf_btn = ttk.Button(pdf_section, text="Browse PDFs", 
                                        command=self.browse_pdf_files,
                                        style='PremiumSecondary.TButton')
        self.browse_pdf_btn.grid(row=1, column=2)
        
        # Output directory
        tk.Label(pdf_section, text="Output Folder", 
                font=('SF Pro Text', 11, '500'), fg='#0f172a', bg='#ffffff').grid(
                row=2, column=0, sticky=tk.W, padx=(0, 20), pady=(24, 8))
        
        self.output_dir_var = tk.StringVar()
        self.output_entry = ttk.Entry(pdf_section, textvariable=self.output_dir_var, 
                                     state="readonly", style='Premium.TEntry')
        self.output_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0, 20))
        
        self.browse_output_btn = ttk.Button(pdf_section, text="Select Folder", 
                                           command=self.browse_output_directory,
                                           style='PremiumSecondary.TButton')
        self.browse_output_btn.grid(row=3, column=2)
        
        self.process_pdf_btn = ttk.Button(pdf_section, text="Process PDFs", 
                                         command=self.process_pdf_files,
                                         style='Premium.TButton')
        self.process_pdf_btn.grid(row=4, column=0, columnspan=3, pady=(30, 0), sticky="ew")
    
    def setup_settings_tab(self):
        """Setup the settings tab with clean minimal design"""
        self.settings_tab.columnconfigure(0, weight=1)
        self.settings_tab.rowconfigure(1, weight=1)
        
        # Transport data configuration - Clean design
        data_section = ttk.LabelFrame(self.settings_tab, text="Transport Data Configuration", 
                                     padding="30", style='Premium.TLabelframe')
        data_section.grid(row=0, column=0, sticky="ew", pady=(0, 30))
        data_section.columnconfigure(1, weight=1)
        
        # File selection with clean styling
        tk.Label(data_section, text="Transport Data File", 
                font=('SF Pro Text', 11, '500'), fg='#0f172a', bg='#ffffff').grid(
                row=0, column=0, sticky=tk.W, padx=(0, 20), pady=(0, 8))
        
        self.data_file_var = tk.StringVar()
        self.data_entry = ttk.Entry(data_section, textvariable=self.data_file_var, 
                                   state="readonly", style='Premium.TEntry')
        self.data_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 20))
        
        self.browse_data_btn = ttk.Button(data_section, text="Browse Data", 
                                         command=self.browse_data_file,
                                         style='PremiumSecondary.TButton')
        self.browse_data_btn.grid(row=1, column=2)
        
        # Column selection
        tk.Label(data_section, text="Data Column", 
                font=('SF Pro Text', 11, '500'), fg='#0f172a', bg='#ffffff').grid(
                row=2, column=0, sticky=tk.W, padx=(0, 20), pady=(24, 8))
        
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(data_section, textvariable=self.column_var, 
                                        state="readonly", width=30, style='Premium.TCombobox')
        self.column_combo.grid(row=3, column=0, sticky="w", padx=(0, 20))
        
        self.preview_btn = ttk.Button(data_section, text="Preview", 
                                     command=self.preview_file_columns,
                                     style='PremiumSecondary.TButton')
        self.preview_btn.grid(row=3, column=1, sticky="w", padx=(20, 0))
        
        self.load_data_btn = ttk.Button(data_section, text="Load Transport Data", 
                                       command=self.load_data_file,
                                       style='Premium.TButton')
        self.load_data_btn.grid(row=4, column=0, columnspan=3, pady=(30, 0), sticky="ew")
        
        # Data preview with clean styling
        data_preview = ttk.LabelFrame(self.settings_tab, text="Transport Data Preview", 
                                     padding="30", style='Premium.TLabelframe')
        data_preview.grid(row=1, column=0, sticky="nsew", pady=(0, 30))
        data_preview.columnconfigure(0, weight=1)
        data_preview.rowconfigure(0, weight=1)
        
        # Clean treeview
        self.tree = ttk.Treeview(data_preview, columns=("Value",), show="tree headings", 
                                height=8, style='Premium.Treeview')
        self.tree.heading("#0", text="Index")
        self.tree.heading("Value", text="Value")
        self.tree.column("#0", width=100, anchor='center')
        self.tree.column("Value", width=500)
        
        data_scrollbar = ttk.Scrollbar(data_preview, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=data_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        data_scrollbar.grid(row=0, column=1, sticky="ns")
    
    def browse_data_file(self):
        """Browse and select Excel or CSV file with clean user feedback"""
        file_path = filedialog.askopenfilename(
            title="Select Transport Data File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.data_file_var.set(file_path)
            self.column_combo['values'] = []
            self.column_var.set('')
            self.update_combined_status()
    
    def preview_file_columns(self):
        """Preview file columns with enhanced error handling"""
        if not self.data_file_var.get():
            messagebox.showwarning("No File Selected", "Please select a data file first.")
            return
            
        try:
            file_path = self.data_file_var.get()
            
            if file_path.lower().endswith('.csv'):
                # Try to read with different encodings
                df = self.read_csv_with_encoding(file_path)
                if df is None:
                    return
            else:
                df = pd.read_excel(file_path)
            
            # Show column selection dialog with clean styling
            columns = list(df.columns)
            if columns:
                self.column_combo['values'] = columns
                self.column_combo.current(0) if columns else None
                
                # Show preview in a clean dialog
                preview_window = tk.Toplevel(self.root)
                preview_window.title("Column Preview")
                preview_window.geometry("500x400")
                preview_window.configure(bg='#ffffff')
                preview_window.resizable(True, True)
                
                # Center the window
                preview_window.transient(self.root)
                preview_window.grab_set()
                
                # Clean header
                header_frame = tk.Frame(preview_window, bg='#ffffff', height=60)
                header_frame.pack(fill='x', padx=20, pady=(20, 0))
                header_frame.pack_propagate(False)
                
                tk.Label(header_frame, text="Available Columns",
                        font=('SF Pro Display', 16, '600'),
                        fg='#0f172a', bg='#ffffff').pack(anchor='w', pady=(10, 0))
                
                tk.Label(header_frame, text="Select the column containing your transport data",
                        font=('SF Pro Text', 11),
                        fg='#64748b', bg='#ffffff').pack(anchor='w', pady=(5, 0))
                
                # Clean listbox
                listbox_frame = tk.Frame(preview_window, bg='#ffffff')
                listbox_frame.pack(fill='both', expand=True, padx=20, pady=20)
                
                listbox = tk.Listbox(listbox_frame, 
                                   font=('SF Pro Text', 10),
                                   bg='#f8fafc',
                                   fg='#0f172a',
                                   selectbackground='#2563eb',
                                   selectforeground='white',
                                   relief='solid',
                                   bd=1,
                                   borderwidth=1,
                                   highlightthickness=0)
                listbox.pack(side='left', fill='both', expand=True)
                
                scrollbar = tk.Scrollbar(listbox_frame, orient='vertical', command=listbox.yview)
                scrollbar.pack(side='right', fill='y')
                listbox.configure(yscrollcommand=scrollbar.set)
                
                for col in columns:
                    listbox.insert(tk.END, col)
                
                # Clean button frame
                button_frame = tk.Frame(preview_window, bg='#ffffff')
                button_frame.pack(fill='x', padx=20, pady=(0, 20))
                
                def select_column():
                    selection = listbox.curselection()
                    if selection:
                        selected_column = columns[selection[0]]
                        self.column_var.set(selected_column)
                        preview_window.destroy()
                
                select_btn = tk.Button(button_frame, text="Select Column",
                                     command=select_column,
                                     bg='#2563eb', fg='white',
                                     font=('SF Pro Display', 10, 'normal'),
                                     relief='flat', bd=0,
                                     padx=24, pady=12,
                                     activebackground='#1d4ed8',
                                     activeforeground='white')
                select_btn.pack(side='right')
                
                close_btn = tk.Button(button_frame, text="Cancel",
                                    command=preview_window.destroy,
                                    bg='#f8fafc', fg='#0f172a',
                                    font=('SF Pro Display', 10, 'normal'),
                                    relief='solid', bd=1,
                                    padx=20, pady=12,
                                    activebackground='#ffffff',
                                    activeforeground='#0f172a')
                close_btn.pack(side='right', padx=(0, 12))
                
            else:
                messagebox.showinfo("No Columns", "No columns found in the selected file.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {str(e)}")
    
    def load_data_file(self):
        """Load transport data with clean progress indication"""
        if not self.data_file_var.get():
            messagebox.showwarning("No File Selected", "Please select a transport data file first.")
            return
            
        if not self.column_var.get():
            messagebox.showwarning("No Column Selected", "Please select a column to load data from.")
            return
        
        try:
            # Start progress indication
            self.progress.start(10)
            self.status_var.set("Loading transport data...")
            self.root.update()
            
            file_path = self.data_file_var.get()
            column_name = self.column_var.get()
            
            # Read the file
            if file_path.lower().endswith('.csv'):
                df = self.read_csv_with_encoding(file_path)
                if df is None:
                    self.progress.stop()
                    return
            else:
                df = pd.read_excel(file_path)
            
            # Extract the specified column
            if column_name in df.columns:
                column_data = df[column_name].dropna().astype(str).tolist()
                
                # Clean and filter the data
                self.data_values = [val.strip() for val in column_data if val.strip()]
                
                # Save to JSON
                with open(self.json_file, 'w', encoding='utf-8') as f:
                    json.dump(self.data_values, f, indent=2, ensure_ascii=False)
                
                # Update UI
                self.update_data_display()
                self.update_combined_status()
                
                # Stop progress
                self.progress.stop()
                
                # Show success message
                messagebox.showinfo("Success", 
                                  f"Successfully loaded {len(self.data_values)} transport records.\n"
                                  f"Data saved to {self.json_file}")
                
            else:
                self.progress.stop()
                messagebox.showerror("Error", f"Column '{column_name}' not found in the file.")
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Error", f"Error loading data: {str(e)}")
    
    def read_csv_with_encoding(self, file_path):
        """Read CSV file with multiple encoding attempts"""
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                return pd.read_csv(file_path, encoding=encoding)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                messagebox.showerror("Error", f"Error reading CSV file: {str(e)}")
                return None
        
        messagebox.showerror("Error", "Could not read the CSV file with any supported encoding.")
        return None
    
    def read_csv_with_encoding_no_header(self, file_path, nrows=None):
        """Read CSV file without header with multiple encoding attempts"""
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                return pd.read_csv(file_path, encoding=encoding, header=None, nrows=nrows)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                continue
        
        return None
    
    def update_data_display(self):
        """Update the data display treeview with clean styling"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add new items
        for i, value in enumerate(self.data_values[:100], 1):  # Show first 100 items
            self.tree.insert("", tk.END, text=str(i), values=(value,))
    
    def load_existing_data(self):
        """Load existing transport data from JSON file"""
        try:
            if os.path.exists(self.json_file):
                with open(self.json_file, 'r', encoding='utf-8') as f:
                    self.data_values = json.load(f)
                self.update_data_display()
        except Exception as e:
            print(f"Error loading existing transport data: {e}")
            self.data_values = []
    
    def browse_delivery_file(self):
        """Browse and select delivery sequence file"""
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
            self.delivery_column_combo['values'] = []
            self.delivery_column_var.set('')
            self.update_combined_status()
    
    def preview_delivery_columns(self):
        """Preview delivery file columns with clean interface"""
        if not self.delivery_file_var.get():
            messagebox.showwarning("No File Selected", "Please select a delivery sequence file first.")
            return
            
        try:
            file_path = self.delivery_file_var.get()
            
            if file_path.lower().endswith('.csv'):
                df = self.read_csv_with_encoding(file_path)
                if df is None:
                    return
            else:
                df = pd.read_excel(file_path)
            
            # Show column selection dialog
            columns = list(df.columns)
            if columns:
                self.delivery_column_combo['values'] = columns
                self.delivery_column_combo.current(0) if columns else None
                
                # Show preview dialog
                preview_window = tk.Toplevel(self.root)
                preview_window.title("Delivery Sequence Columns")
                preview_window.geometry("500x400")
                preview_window.configure(bg='#ffffff')
                preview_window.resizable(True, True)
                
                preview_window.transient(self.root)
                preview_window.grab_set()
                
                # Clean header
                header_frame = tk.Frame(preview_window, bg='#ffffff', height=60)
                header_frame.pack(fill='x', padx=20, pady=(20, 0))
                header_frame.pack_propagate(False)
                
                tk.Label(header_frame, text="Delivery Sequence Columns",
                        font=('SF Pro Display', 16, '600'),
                        fg='#0f172a', bg='#ffffff').pack(anchor='w', pady=(10, 0))
                
                tk.Label(header_frame, text="Select the column containing delivery sequence data",
                        font=('SF Pro Text', 11),
                        fg='#64748b', bg='#ffffff').pack(anchor='w', pady=(5, 0))
                
                # Clean listbox
                listbox_frame = tk.Frame(preview_window, bg='#ffffff')
                listbox_frame.pack(fill='both', expand=True, padx=20, pady=20)
                
                listbox = tk.Listbox(listbox_frame,
                                   font=('SF Pro Text', 10),
                                   bg='#f8fafc',
                                   fg='#0f172a',
                                   selectbackground='#2563eb',
                                   selectforeground='white',
                                   relief='solid',
                                   bd=1,
                                   borderwidth=1,
                                   highlightthickness=0)
                listbox.pack(side='left', fill='both', expand=True)
                
                scrollbar = tk.Scrollbar(listbox_frame, orient='vertical', command=listbox.yview)
                scrollbar.pack(side='right', fill='y')
                listbox.configure(yscrollcommand=scrollbar.set)
                
                for col in columns:
                    listbox.insert(tk.END, col)
                
                # Clean buttons
                button_frame = tk.Frame(preview_window, bg='#ffffff')
                button_frame.pack(fill='x', padx=20, pady=(0, 20))
                
                def select_column():
                    selection = listbox.curselection()
                    if selection:
                        selected_column = columns[selection[0]]
                        self.delivery_column_var.set(selected_column)
                        preview_window.destroy()
                
                select_btn = tk.Button(button_frame, text="Select Column",
                                     command=select_column,
                                     bg='#2563eb', fg='white',
                                     font=('SF Pro Display', 10, 'normal'),
                                     relief='flat', bd=0,
                                     padx=24, pady=12,
                                     activebackground='#1d4ed8',
                                     activeforeground='white')
                select_btn.pack(side='right')
                
                close_btn = tk.Button(button_frame, text="Cancel",
                                    command=preview_window.destroy,
                                    bg='#f8fafc', fg='#0f172a',
                                    font=('SF Pro Display', 10, 'normal'),
                                    relief='solid', bd=1,
                                    padx=20, pady=12,
                                    activebackground='#ffffff',
                                    activeforeground='#0f172a')
                close_btn.pack(side='right', padx=(0, 12))
                
            else:
                messagebox.showinfo("No Columns", "No columns found in the selected file.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {str(e)}")
    
    def load_delivery_file(self):
        """Load delivery sequence data with clean progress indication"""
        if not self.delivery_file_var.get():
            messagebox.showwarning("No File Selected", "Please select a delivery sequence file first.")
            return
            
        if not self.delivery_column_var.get():
            messagebox.showwarning("No Column Selected", "Please select a column to load delivery sequence from.")
            return
        
        try:
            # Start progress
            self.progress.start(10)
            self.status_var.set("Loading delivery sequence...")
            self.root.update()
            
            file_path = self.delivery_file_var.get()
            column_name = self.delivery_column_var.get()
            
            # Read the file
            if file_path.lower().endswith('.csv'):
                df = self.read_csv_with_encoding(file_path)
                if df is None:
                    self.progress.stop()
                    return
            else:
                df = pd.read_excel(file_path)
            
            # Extract the specified column
            if column_name in df.columns:
                column_data = df[column_name].dropna().astype(str).tolist()
                
                # Clean and filter the data
                self.delivery_data_values = [val.strip() for val in column_data if val.strip()]
                
                # Save to JSON
                with open(self.delivery_json_file, 'w', encoding='utf-8') as f:
                    json.dump(self.delivery_data_values, f, indent=2, ensure_ascii=False)
                
                # Update UI
                self.update_delivery_display()
                self.update_combined_status()
                
                # Stop progress
                self.progress.stop()
                
                # Show success message
                messagebox.showinfo("Success", 
                                  f"Successfully loaded {len(self.delivery_data_values)} delivery sequence records.\n"
                                  f"Data saved to {self.delivery_json_file}")
                
            else:
                self.progress.stop()
                messagebox.showerror("Error", f"Column '{column_name}' not found in the file.")
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Error", f"Error loading delivery data: {str(e)}")
    
    def update_delivery_display(self):
        """Update the delivery display treeview"""
        # Clear existing items
        for item in self.delivery_tree.get_children():
            self.delivery_tree.delete(item)
        
        # Add new items
        for i, value in enumerate(self.delivery_data_values[:100], 1):  # Show first 100 items
            self.delivery_tree.insert("", tk.END, text=str(i), values=(value,))
    
    def load_existing_delivery_data(self):
        """Load existing delivery data from JSON file"""
        try:
            if os.path.exists(self.delivery_json_file):
                with open(self.delivery_json_file, 'r', encoding='utf-8') as f:
                    self.delivery_data_values = json.load(f)
                self.update_delivery_display()
        except Exception as e:
            print(f"Error loading existing delivery data: {e}")
            self.delivery_data_values = []
    
    def browse_pdf_files(self):
        """Browse and select PDF files for processing"""
        file_paths = filedialog.askopenfilenames(
            title="Select PDF Files to Process",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        if file_paths:
            self.pdf_files_var.set(f"{len(file_paths)} PDF files selected")
            self.pdf_files = file_paths
    
    def browse_output_directory(self):
        """Browse and select output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)
    
    def process_pdf_files(self):
        """Process PDF files with enhanced progress tracking"""
        # Validation
        if not hasattr(self, 'pdf_files') or not self.pdf_files:
            messagebox.showwarning("No PDFs Selected", "Please select PDF files to process first.")
            return
            
        if not self.output_dir_var.get():
            messagebox.showwarning("No Output Directory", "Please select an output directory first.")
            return
            
        if not self.delivery_data_values:
            messagebox.showwarning("No Delivery Data", "Please load delivery sequence data first.")
            return
        
        try:
            # Start processing with clean progress indication
            self.progress.start(10)
            self.status_var.set("Processing PDF files...")
            self.root.update()
            
            output_dir = self.output_dir_var.get()
            processed_count = 0
            total_files = len(self.pdf_files)
            
            # Create results summary
            results = {
                'processed_files': [],
                'failed_files': [],
                'total_pages_processed': 0,
                'total_matches_found': 0
            }
            
            for i, pdf_path in enumerate(self.pdf_files):
                try:
                    # Update progress
                    self.status_var.set(f"Processing {os.path.basename(pdf_path)} ({i+1}/{total_files})...")
                    self.root.update()
                    
                    # Extract and match pages
                    matched_pages = self.extract_and_match_pdf_pages(pdf_path)
                    
                    if matched_pages:
                        # Sort pages by delivery sequence
                        sorted_pages = self.sort_pages_by_delivery_sequence(matched_pages)
                        
                        # Create filtered PDF
                        output_path = self.create_filtered_pdf(pdf_path, sorted_pages, output_dir)
                        
                        results['processed_files'].append({
                            'file': os.path.basename(pdf_path),
                            'output': output_path,
                            'pages': len(sorted_pages),
                            'matches': sum(len(page['matches']) for page in sorted_pages)
                        })
                        
                        results['total_pages_processed'] += len(sorted_pages)
                        results['total_matches_found'] += sum(len(page['matches']) for page in sorted_pages)
                        
                    else:
                        results['failed_files'].append({
                            'file': os.path.basename(pdf_path),
                            'reason': 'No delivery sequence matches found'
                        })
                    
                    processed_count += 1
                    
                except Exception as e:
                    results['failed_files'].append({
                        'file': os.path.basename(pdf_path),
                        'reason': str(e)
                    })
            
            # Stop progress and show results
            self.progress.stop()
            self.update_combined_status()
            
            # Show clean results dialog
            self.show_processing_results(results)
            
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Processing Error", f"An error occurred during processing: {str(e)}")
    
    def show_processing_results(self, results):
        """Show processing results in a clean dialog"""
        results_window = tk.Toplevel(self.root)
        results_window.title("Processing Results")
        results_window.geometry("600x500")
        results_window.configure(bg='#ffffff')
        results_window.resizable(True, True)
        
        results_window.transient(self.root)
        results_window.grab_set()
        
        # Clean header
        header_frame = tk.Frame(results_window, bg='#ffffff', height=80)
        header_frame.pack(fill='x', padx=30, pady=(30, 0))
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="Processing Complete",
                font=('SF Pro Display', 20, '600'),
                fg='#0f172a', bg='#ffffff').pack(anchor='w', pady=(15, 0))
        
        # Summary stats
        stats_frame = tk.Frame(results_window, bg='#f8fafc', relief='solid', bd=1)
        stats_frame.pack(fill='x', padx=30, pady=20)
        
        stats_content = tk.Frame(stats_frame, bg='#f8fafc')
        stats_content.pack(fill='x', padx=20, pady=15)
        
        # Success stats
        success_count = len(results['processed_files'])
        total_files = success_count + len(results['failed_files'])
        
        tk.Label(stats_content, text=f"✓ {success_count} of {total_files} files processed successfully",
                font=('SF Pro Text', 12, '500'),
                fg='#10b981', bg='#f8fafc').pack(anchor='w')
        
        tk.Label(stats_content, text=f"📄 {results['total_pages_processed']} pages processed",
                font=('SF Pro Text', 11),
                fg='#64748b', bg='#f8fafc').pack(anchor='w', pady=(5, 0))
        
        tk.Label(stats_content, text=f"🎯 {results['total_matches_found']} delivery matches found",
                font=('SF Pro Text', 11),
                fg='#64748b', bg='#f8fafc').pack(anchor='w')
        
        # Details frame
        details_frame = tk.Frame(results_window, bg='#ffffff')
        details_frame.pack(fill='both', expand=True, padx=30, pady=(0, 30))
        
        # Create notebook for results
        results_notebook = ttk.Notebook(details_frame, style='Premium.TNotebook')
        results_notebook.pack(fill='both', expand=True)
        
        # Success tab
        if results['processed_files']:
            success_frame = ttk.Frame(results_notebook, style='Premium.TFrame', padding="20")
            results_notebook.add(success_frame, text=f"Successful ({success_count})")
            
            success_text = tk.Text(success_frame, font=('SF Pro Text', 10),
                                 bg='#ffffff', fg='#0f172a',
                                 relief='solid', bd=1,
                                 wrap=tk.WORD)
            success_text.pack(fill='both', expand=True)
            
            for file_info in results['processed_files']:
                success_text.insert(tk.END, f"✓ {file_info['file']}\n")
                success_text.insert(tk.END, f"   Pages: {file_info['pages']} | Matches: {file_info['matches']}\n")
                success_text.insert(tk.END, f"   Output: {file_info['output']}\n\n")
            
            success_text.config(state='disabled')
        
        # Failed tab
        if results['failed_files']:
            failed_frame = ttk.Frame(results_notebook, style='Premium.TFrame', padding="20")
            results_notebook.add(failed_frame, text=f"Failed ({len(results['failed_files'])})")
            
            failed_text = tk.Text(failed_frame, font=('SF Pro Text', 10),
                                bg='#ffffff', fg='#0f172a',
                                relief='solid', bd=1,
                                wrap=tk.WORD)
            failed_text.pack(fill='both', expand=True)
            
            for file_info in results['failed_files']:
                failed_text.insert(tk.END, f"✗ {file_info['file']}\n")
                failed_text.insert(tk.END, f"   Reason: {file_info['reason']}\n\n")
            
            failed_text.config(state='disabled')
        
        # Close button
        close_btn = tk.Button(results_window, text="Close",
                            command=results_window.destroy,
                            bg='#2563eb', fg='white',
                            font=('SF Pro Display', 10, 'normal'),
                            relief='flat', bd=0,
                            padx=24, pady=12,
                            activebackground='#1d4ed8',
                            activeforeground='white')
        close_btn.pack(pady=(0, 30))
    
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
        
        return matched_pages
    
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
        
        return {
            "page_number": page_number,
            "matches_found": len(matched_values) > 0,
            "matched_values": list(set(matched_values)),  # Remove duplicates
            "matches": matched_lines,
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
        return sorted_pages
    
    def create_filtered_pdf(self, original_pdf_path, matched_pages, output_dir):
        """Create a new PDF containing only the pages that matched delivery sequence values"""
        original_doc = None
        filtered_doc = None
        
        try:
            if not matched_pages:
                return None
                
            # Generate output filename with clean naming
            original_name = os.path.splitext(os.path.basename(original_pdf_path))[0]
            today_date = pd.Timestamp.now().strftime('%Y-%m-%d')
            filtered_filename = f"{original_name}_processed_{today_date}_{len(matched_pages)}pages.pdf"
            filtered_path = os.path.join(output_dir, filtered_filename)
            
            # Open the original PDF
            original_doc = fitz.open(original_pdf_path)
            
            # Create a new PDF document
            filtered_doc = fitz.open()
            
            # Copy pages in delivery sequence order
            for page_info in matched_pages:
                page_num = page_info["page_number"]
                
                # PyMuPDF uses 0-based indexing, but our page_number is 1-based
                page_index = page_num - 1
                
                if 0 <= page_index < original_doc.page_count:
                    # Use insert_pdf to copy the page
                    filtered_doc.insert_pdf(original_doc, from_page=page_index, to_page=page_index)
            
            # Only save if we actually copied some pages
            if filtered_doc.page_count > 0:
                filtered_doc.save(filtered_path)
                return filtered_path
            else:
                return None
                
        except Exception as e:
            print(f"Error creating filtered PDF for {original_pdf_path}: {str(e)}")
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
            "page_count": len(extracted_text)
        }

def main():
    root = tk.Tk()
    app = TransportSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()