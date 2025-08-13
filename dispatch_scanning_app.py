import sys
import os
import json
from pathlib import Path
import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
import barcode
from barcode.writer import ImageWriter
import hashlib
import requests
from datetime import datetime, timedelta

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QPushButton, QTextEdit, QLineEdit, 
    QFileDialog, QMessageBox, QProgressBar, QStatusBar, QFrame,
    QScrollArea, QGroupBox, QSplitter, QComboBox, QDialog, 
    QDialogButtonBox, QListWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QPlainTextEdit, QCheckBox, QTabWidget, QDateEdit
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize, QDate
from PySide6.QtGui import QFont, QPalette, QColor, QIcon, QPixmap

# Import Supabase configuration
try:
    from supabase_config import save_generated_barcodes, upload_store_orders_from_excel
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False
    print("Warning: Supabase configuration not available. Some features may be disabled.")


class ProcessingThread(QThread):
    """Background thread for PDF processing operations"""
    progress_signal = Signal(str)
    finished_signal = Signal(bool, dict)
    
    def __init__(self, app_instance):
        super().__init__()
        self.app = app_instance
    
    def run(self):
        try:
            result = self.app.process_picking_dockets_internal()
            self.finished_signal.emit(True, result)
        except Exception as e:
            self.progress_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(False, {"error": str(e)})


class ProcessingResultsDialog(QDialog):
    """Professional dialog for displaying processing results"""
    
    def __init__(self, results, parent=None):
        super().__init__(parent)
        self.results = results
        self.setWindowTitle("Processing Results")
        self.setModal(True)
        self.resize(900, 700)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        header_layout = QHBoxLayout()
        
        # Status icon
        status_icon = "✓"
        status_color = "#10b981"
        
        if results.get('driver_files_created', 0) == 0:
            status_icon = "⚠"
            status_color = "#f59e0b"
        elif results.get('failed_files'):
            status_icon = "⚠"
            status_color = "#f59e0b"
        
        status_label = QLabel(status_icon)
        status_label.setStyleSheet(f"font-size: 32px; color: {status_color}; font-weight: bold;")
        header_layout.addWidget(status_label)
        
        # Title
        title_text = "Picking Dockets Processing Completed Successfully"
        if results.get('driver_files_created', 0) == 0:
            title_text = "Picking Dockets Processing Completed - No Matching Orders Found"
        elif results.get('failed_files'):
            title_text = "Picking Dockets Processing Completed with Some Issues"
        
        title_label = QLabel(title_text)
        title_label.setObjectName("resultTitle")
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # Summary statistics
        stats_frame = QFrame()
        stats_frame.setObjectName("statsFrame")
        stats_layout = QGridLayout(stats_frame)
        
        # Stats
        stats_data = [
            ("Picking PDF Files Processed", str(results.get('processed_files', 0))),
            ("Pages Scanned", str(results.get('total_pages', 0))),
            ("Driver Picking PDFs Created", str(results.get('driver_files_created', 0))),
            ("Barcodes Generated", str(results.get('barcodes_generated', 0))),
            ("Failed Files", str(len(results.get('failed_files', []))))
        ]
        
        for i, (label, value) in enumerate(stats_data):
            label_widget = QLabel(label + ":")
            label_widget.setObjectName("statsLabel")
            value_widget = QLabel(value)
            value_widget.setObjectName("statsValue")
            
            stats_layout.addWidget(label_widget, i, 0)
            stats_layout.addWidget(value_widget, i, 1)
        
        layout.addWidget(stats_frame)
        
        # Create tabbed interface
        tab_widget = QTabWidget()
        
        # Tab 1: Created Files
        if results.get('created_files'):
            files_tab = self.create_files_tab(results.get('created_files', []))
            tab_widget.addTab(files_tab, "Created Files")
        
        # Tab 2: Newly Added Rows (if any)
        if results.get('additional_new_rows_count', 0) > 0:
            added_tab = self.create_added_rows_tab(
                results.get('additional_file', ''),
                results.get('additional_new_rows_count', 0),
                results.get('additional_new_rows', [])
            )
            tab_widget.addTab(added_tab, "New Rows (Additional)")

        # Next: Driver Details
        if results.get('driver_details'):
            driver_tab = self.create_driver_tab(results.get('driver_details', {}))
            tab_widget.addTab(driver_tab, "Driver Details")
        
        # Tab 3: Failed Files (if any)
        if results.get('failed_files'):
            failed_tab = self.create_failed_tab(results.get('failed_files', []))
            tab_widget.addTab(failed_tab, "Failed Files")
        
        # If no tabs were created, show diagnostic information
        if tab_widget.count() == 0:
            if results.get('error') == "No matching orders found in picking docket PDF files":
                diagnostic_tab = self.create_diagnostic_tab(results)
                tab_widget.addTab(diagnostic_tab, "Diagnostic Info")
            else:
                empty_tab = self.create_empty_tab()
                tab_widget.addTab(empty_tab, "Results")
        
        layout.addWidget(tab_widget)
        
        # Output directory info
        output_frame = QFrame()
        output_frame.setObjectName("outputFrame")
        output_layout = QVBoxLayout(output_frame)
        
        output_label = QLabel("Output Directory:")
        output_label.setObjectName("outputLabel")
        output_layout.addWidget(output_label)
        
        output_path = QLabel(results.get('output_dir', ''))
        output_path.setObjectName("outputPath")
        output_path.setWordWrap(True)
        output_layout.addWidget(output_path)
        
        layout.addWidget(output_frame)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        open_folder_btn = QPushButton("Open Output Folder")
        open_folder_btn.setObjectName("primaryButton")
        open_folder_btn.clicked.connect(self.open_output_folder)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        
        button_layout.addWidget(open_folder_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        # Apply styling
        self.apply_results_styling()
    
    def create_files_tab(self, files):
        """Create tab showing created files"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["File Name", "Status"])
        table.setRowCount(len(files))
        
        for i, filename in enumerate(files):
            # File name
            name_item = QTableWidgetItem(filename)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 0, name_item)
            
            # Status
            status_item = QTableWidgetItem("✓ Created")
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 1, status_item)
        
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)
        
        return widget

    def create_added_rows_tab(self, additional_filename: str, count: int, rows: list):
        """Create a tab showing newly added rows from the additional Excel file"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        header = QLabel(
            f"Additional file: {additional_filename} — {count} new row(s) detected and uploaded"
        )
        header.setObjectName("infoText")
        header.setWordWrap(True)
        layout.addWidget(header)

        if not rows:
            empty = QLabel("No new rows")
            layout.addWidget(empty)
            return widget

        # Build a table from the row dicts using union of keys
        columns = []
        for r in rows:
            for k in r.keys():
                if k not in columns:
                    columns.append(k)

        table = QTableWidget()
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels([str(c) for c in columns])
        table.setRowCount(len(rows))

        for i, r in enumerate(rows):
            for j, col in enumerate(columns):
                val = r.get(col, "")
                item = QTableWidgetItem("") if val is None else QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(i, j, item)

        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)

        return widget
    
    def create_driver_tab(self, driver_details):
        """Create tab showing driver details"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Driver", "Total Pages", "Orders"])
        table.setRowCount(len(driver_details))
        
        for i, (driver, details) in enumerate(driver_details.items()):
            # Driver
            driver_item = QTableWidgetItem(f"Driver {driver}")
            driver_item.setFlags(driver_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 0, driver_item)
            
            # Total Pages
            total_pages = details.get('page_count', 0)
            pages_item = QTableWidgetItem(str(total_pages))
            pages_item.setFlags(pages_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 1, pages_item)
            
            # Orders
            orders = details.get('orders', [])
            orders_text = ", ".join(orders[:3])
            if len(orders) > 3:
                orders_text += f" ... (+{len(orders)-3} more)"
            orders_item = QTableWidgetItem(orders_text)
            orders_item.setFlags(orders_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 2, orders_item)
        
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)
        
        return widget
    
    def create_failed_tab(self, failed_files):
        """Create tab showing failed files"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["File Name", "Status"])
        table.setRowCount(len(failed_files))
        
        for i, filename in enumerate(failed_files):
            # File name
            name_item = QTableWidgetItem(filename)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 0, name_item)
            
            # Status
            status_item = QTableWidgetItem("✗ Failed")
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 1, status_item)
        
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)
        
        return widget
    
    def create_empty_tab(self):
        """Create empty tab for when no results are available"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Center the message
        layout.addStretch()
        
        message_label = QLabel("No results to display")
        message_label.setAlignment(Qt.AlignCenter)
        message_label.setStyleSheet("color: #6b7280; font-size: 14px; font-style: italic;")
        layout.addWidget(message_label)
        
        layout.addStretch()
        
        return widget
    
    def create_diagnostic_tab(self, results):
        """Create diagnostic tab when no matching orders are found"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Diagnostic message
        diagnostic_text = QPlainTextEdit()
        diagnostic_text.setReadOnly(True)
        diagnostic_text.setStyleSheet("font-family: 'Consolas', monospace; font-size: 12px; background-color: #f8fafc;")
        
        message = "No matching orders were found in the picking docket PDF files.\n\n"
        message += "Troubleshooting Steps:\n"
        message += "=" * 50 + "\n\n"
        message += "1. Check that your picking docket PDF files contain order IDs\n"
        message += "2. Ensure you have loaded delivery sequence data first\n"
        message += "3. Order ID matching is case-insensitive (AA061B4Y = aa061b4y)\n\n"
        message += f"Processing Summary:\n"
        message += f"- Picking PDF files processed: {results.get('processed_files', 0)}\n"
        message += f"- Total pages scanned: {results.get('total_pages', 0)}\n"
        message += f"- No matching order IDs found\n\n"
        message += "Common Issues:\n"
        message += "- Order IDs in picking PDFs don't match those in delivery data\n"
        message += "- PDF contains images that need OCR processing\n"
        message += "- No delivery sequence data loaded\n"
        message += "- Text extraction failed from PDF pages\n\n"
        message += "Note: You need delivery sequence data loaded before processing picking dockets."
        
        diagnostic_text.setPlainText(message)
        layout.addWidget(diagnostic_text)
        
        return widget
    
    def open_output_folder(self):
        """Open output folder"""
        if hasattr(self.parent(), 'open_output_directory'):
            self.parent().open_output_directory(self.results.get('output_dir', ''))
    
    def apply_results_styling(self):
        """Apply styling to results dialog"""
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
            }
            
            QLabel#resultTitle {
                font-size: 18px;
                font-weight: bold;
                color: #1e293b;
                margin-bottom: 10px;
            }
            
            QFrame#statsFrame {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
            }
            
            QLabel#statsLabel {
                font-weight: 600;
                color: #374151;
                font-size: 14px;
            }
            
            QLabel#statsValue {
                font-size: 14px;
                color: #10b981;
                font-weight: bold;
            }
            
            QFrame#outputFrame {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
            }
            
            QLabel#outputLabel {
                font-weight: 600;
                color: #374151;
                font-size: 14px;
                margin-bottom: 5px;
            }
            
            QLabel#outputPath {
                color: #6b7280;
                font-size: 12px;
                font-family: 'Consolas', monospace;
            }
            
            QTabWidget::pane {
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                background-color: white;
            }
            
            QTabBar::tab {
                background-color: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-bottom: none;
                border-radius: 6px 6px 0 0;
                padding: 8px 16px;
                margin-right: 2px;
                color: #64748b;
                font-weight: 500;
            }
            
            QTabBar::tab:selected {
                background-color: white;
                color: #1e293b;
                border-bottom: 2px solid #2563eb;
            }
            
            QTabBar::tab:hover {
                background-color: #e2e8f0;
                color: #374151;
            }
            
            QTableWidget {
                gridline-color: #f1f5f9;
                border: none;
                background-color: white;
            }
            
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f1f5f9;
            }
            
            QTableWidget::item:selected {
                background-color: #eff6ff;
                color: #1e40af;
            }
            
            QTableWidget QHeaderView::section {
                background-color: #f8fafc;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                padding: 10px;
                font-weight: 600;
                color: #374151;
            }
        """)


class DispatchScanningApp(QMainWindow):
    """Upload Excel Files, Process PDFs, Generate Barcodes"""
    
    def __init__(self):
        super().__init__()
        
        # Application data
        self.delivery_data_values = []
        self.delivery_data_with_drivers = {}
        self.delivery_json_file = "delivery_sequence_data.json"
        self.selected_picking_pdf_files = []

        self.selected_excel_file = ""  # NEW: Excel file with order numbers in column A
        self.selected_output_folder = ""  # NEW: Selected output folder
        self.excel_order_numbers = []  # NEW: Order numbers from Excel column A
        # Additional Excel import for late orders
        self.selected_additional_excel_file = ""
        self.additional_new_rows_data = []
        self.additional_new_rows_count = 0
        self.order_barcodes = {}
        self.processing_thread = None
        
        # Track processing state
        self.picking_dockets_processed = False
        
        # Initialize UI
        self.init_ui()
        self.apply_clean_styling()
        
        # Load existing data
        self.load_existing_delivery_data()
        self.update_status("Ready")
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Dispatch Scanning - Streamlined Processing")
        self.setGeometry(100, 100, 1000, 650)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # Header
        header_frame = self.create_header()
        main_layout.addWidget(header_frame)
        
        # Content area - single column for picking section
        picking_section = self.create_picking_section()
        main_layout.addWidget(picking_section)
        
        # Process button
        self.process_picking_btn = QPushButton("Process PDF Files & Upload")
        self.process_picking_btn.setObjectName("primaryButton")
        self.process_picking_btn.clicked.connect(self.process_picking_dockets)
        self.process_picking_btn.setFixedHeight(35)
        main_layout.addWidget(self.process_picking_btn)
        
        # Output section
        output_section = self.create_output_section()
        main_layout.addWidget(output_section)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status("Ready")
        
        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
    
    def create_header(self):
        """Create application header"""
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_frame.setFixedHeight(55)
        
        layout = QHBoxLayout(header_frame)
        layout.setContentsMargins(0, 10, 0, 10)
        
        title_label = QLabel("Dispatch Scanning")
        title_label.setObjectName("headerTitle")
        layout.addWidget(title_label)
        
        layout.addStretch()
        
       
        
        return header_frame
    
    def create_picking_section(self):
        """Create picking dockets section"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(10)
        
        # Section title
        title = QLabel("Process Picking Dockets & Upload Store Orders")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
       
        
        # Output folder selection subsection (moved to top)
        output_label = QLabel("Output Folder:")
        output_label.setStyleSheet("font-weight: bold; margin-top: 8px;")
        layout.addWidget(output_label)
        
        # Output folder selection row
        output_btn_layout = QHBoxLayout()
        self.browse_output_btn = QPushButton("Select Output Folder")
        self.browse_output_btn.clicked.connect(self.browse_output_folder)
        
        self.clear_output_btn = QPushButton("Clear")
        self.clear_output_btn.setObjectName("secondaryButton")
        self.clear_output_btn.clicked.connect(self.clear_output_folder)
        
        output_btn_layout.addWidget(self.browse_output_btn)
        output_btn_layout.addWidget(self.clear_output_btn)
        layout.addLayout(output_btn_layout)
        
        # Output folder display
        self.output_folder_label = QLabel("Files will be saved in a date-based subfolder (YYYY-MM-DD)")
        self.output_folder_label.setObjectName("infoText")
        self.output_folder_label.setWordWrap(True)
        layout.addWidget(self.output_folder_label)
        
        # Excel file with store orders subsection (moved after output folder)
        excel_label = QLabel("Store Order Excel File (for barcodes & database upload):")
        excel_label.setStyleSheet("font-weight: bold; margin-top: 8px;")
        layout.addWidget(excel_label)
        
        # Excel file selection row
        excel_btn_layout = QHBoxLayout()
        self.browse_excel_btn = QPushButton("Select Store Order Excel File")
        self.browse_excel_btn.clicked.connect(self.browse_excel_file)
        
        self.clear_excel_btn = QPushButton("Clear")
        self.clear_excel_btn.setObjectName("secondaryButton")
        self.clear_excel_btn.clicked.connect(self.clear_excel_file)
        
        excel_btn_layout.addWidget(self.browse_excel_btn)
        excel_btn_layout.addWidget(self.clear_excel_btn)
        layout.addLayout(excel_btn_layout)
        
        # Excel file display
        self.excel_file_label = QLabel("No Excel file selected")
        self.excel_file_label.setObjectName("infoText")
        self.excel_file_label.setWordWrap(True)
        layout.addWidget(self.excel_file_label)

        # Delivery date picker (for created_at override)
        date_label = QLabel("Delivery Date (sets created_at for uploaded orders):")
        date_label.setStyleSheet("font-weight: bold; margin-top: 8px;")
        layout.addWidget(date_label)

        date_row = QHBoxLayout()
        self.delivery_date_edit = QDateEdit()
        self.delivery_date_edit.setCalendarPopup(True)
        self.delivery_date_edit.setDate(QDate.currentDate())
        self.delivery_date_edit.setDisplayFormat("yyyy-MM-dd")
        date_row.addWidget(self.delivery_date_edit)
        layout.addLayout(date_row)

        # Additional Excel import subsection
        add_excel_label = QLabel("Additional Order Excel Import (compare to initial)")
        add_excel_label.setStyleSheet("font-weight: bold; margin-top: 8px;")
        layout.addWidget(add_excel_label)

        add_excel_btn_layout = QHBoxLayout()
        self.browse_additional_excel_btn = QPushButton("Select Additional Excel File")
        self.browse_additional_excel_btn.clicked.connect(self.browse_additional_excel_file)

        self.clear_additional_excel_btn = QPushButton("Clear")
        self.clear_additional_excel_btn.setObjectName("secondaryButton")
        self.clear_additional_excel_btn.clicked.connect(self.clear_additional_excel_file)

        add_excel_btn_layout.addWidget(self.browse_additional_excel_btn)
        add_excel_btn_layout.addWidget(self.clear_additional_excel_btn)
        layout.addLayout(add_excel_btn_layout)

        # Additional Excel file display
        self.additional_excel_file_label = QLabel("No additional Excel file selected")
        self.additional_excel_file_label.setObjectName("infoText")
        self.additional_excel_file_label.setWordWrap(True)
        layout.addWidget(self.additional_excel_file_label)
        
        # PDF files subsection
        pdf_label = QLabel("Picking Docket PDF Files:")
        pdf_label.setStyleSheet("font-weight: bold; margin-top: 8px;")
        layout.addWidget(pdf_label)
        
        # Button row
        btn_layout = QHBoxLayout()
        add_picking_pdf_btn = QPushButton("Add Picking PDFs")
        add_picking_pdf_btn.clicked.connect(self.browse_picking_pdf_files)
        
        clear_picking_pdf_btn = QPushButton("Clear")
        clear_picking_pdf_btn.setObjectName("secondaryButton")
        clear_picking_pdf_btn.clicked.connect(self.clear_picking_pdf_files)
        
        btn_layout.addWidget(add_picking_pdf_btn)
        btn_layout.addWidget(clear_picking_pdf_btn)
        layout.addLayout(btn_layout)
        
        # PDF list
        self.picking_pdf_list = QListWidget()
        self.picking_pdf_list.setMaximumHeight(80)
        layout.addWidget(self.picking_pdf_list)
        
        # Info
        info_label = QLabel("Expected Excel columns:\n• Column A: Order Number (→ ordernumber)\n• Column B: Item Code (→ itemcode)\n• Column C: Product Description (→ product_description)\n• Column D: Barcode (→ barcode)\n• Column E: Customer Type (→ customer_type)\n• Column F: Quantity (→ quantity)")
        info_label.setObjectName("infoText")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        layout.addStretch()
        
        return section
    

    
    def create_output_section(self):
        """Create output section"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(10)
        
       
        
       
        
        return section
    
    # File handling methods
    def browse_picking_pdf_files(self):
        """Browse for picking PDF files to process"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Picking Docket PDF Files",
            str(Path.home()),
            "PDF files (*.pdf);;All files (*.*)"
        )
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.selected_picking_pdf_files:
                    self.selected_picking_pdf_files.append(file_path)
                    self.picking_pdf_list.addItem(Path(file_path).name)
    
    def clear_picking_pdf_files(self):
        """Clear selected picking PDF files"""
        self.selected_picking_pdf_files.clear()
        self.picking_pdf_list.clear()
        
        # Reset processing state since picking PDFs are cleared
        self.picking_dockets_processed = False
    
    # NEW: Excel file handling methods
    def browse_excel_file(self):
        """Browse for Excel file containing store orders for database upload and barcode generation"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Store Order Excel File (for database upload & barcodes)",
            str(Path.home()),
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if file_path:
            self.selected_excel_file = file_path
            try:
                # Load Excel file and read column A
                df = pd.read_excel(file_path)
                if df.empty or len(df.columns) == 0:
                    raise ValueError("Excel file is empty or has no columns")
                
                # Get column A values (first column)
                column_a_values = df.iloc[:, 0].dropna().astype(str).tolist()
                self.excel_order_numbers = [str(val).strip() for val in column_a_values if str(val).strip()]
                
                filename = Path(file_path).name
                self.excel_file_label.setText(f"Selected: {filename} ({len(self.excel_order_numbers)} order numbers)")
                self.excel_file_label.setObjectName("successText")
                self.update_status(f"Loaded {len(self.excel_order_numbers)} order numbers from Excel file")
                
            except Exception as e:
                QMessageBox.critical(self, "Excel File Error", f"Error reading Excel file: {str(e)}")
                self.selected_excel_file = ""
                self.excel_order_numbers = []
                self.excel_file_label.setText("Error reading Excel file")
                self.excel_file_label.setObjectName("warningText")
    
    def clear_excel_file(self):
        """Clear selected Excel file"""
        self.selected_excel_file = ""
        self.excel_order_numbers = []
        self.excel_file_label.setText("No Excel file selected")
        self.excel_file_label.setObjectName("infoText")
        self.update_status("Excel file cleared")

        # Clearing base file invalidates any computed additional diffs
        self.selected_additional_excel_file = ""
        self.additional_new_rows_data = []
        self.additional_new_rows_count = 0
        if hasattr(self, 'additional_excel_file_label'):
            self.additional_excel_file_label.setText("No additional Excel file selected")
            self.additional_excel_file_label.setObjectName("infoText")
    
    # NEW: Output folder handling methods
    def browse_output_folder(self):
        """Browse for output folder"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            str(Path.home())
        )
        if folder_path:
            self.selected_output_folder = folder_path
            self.output_folder_label.setText(f"Selected: {folder_path}")
            self.output_folder_label.setObjectName("successText")
            self.update_status(f"Output folder set to: {folder_path}")
    
    def clear_output_folder(self):
        """Clear selected output folder"""
        self.selected_output_folder = ""
        self.output_folder_label.setText("No output folder selected (will use default: picking_dockets_output)\nFiles will be saved in a date-based subfolder (YYYY-MM-DD)")
        self.output_folder_label.setObjectName("infoText")
        self.update_status("Output folder cleared - will use default location")

    # Additional Excel handling methods
    def browse_additional_excel_file(self):
        """Browse for an additional Excel file and compute new rows vs initial Excel"""
        if not self.selected_excel_file:
            QMessageBox.warning(
                self,
                "Select Initial Excel First",
                "Please select the initial 'Store Order Excel File' first, then choose the additional Excel file to compare."
            )
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Additional Store Order Excel File (late orders)",
            str(Path.home()),
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if not file_path:
            return

        self.selected_additional_excel_file = file_path
        try:
            # Compute new rows
            new_rows = self.compute_new_rows_between_excels(self.selected_excel_file, self.selected_additional_excel_file)
            self.additional_new_rows_data = new_rows
            self.additional_new_rows_count = len(new_rows)

            filename = Path(file_path).name
            if self.additional_new_rows_count > 0:
                self.additional_excel_file_label.setText(
                    f"Selected: {filename} — {self.additional_new_rows_count} new rows will be uploaded"
                )
                self.additional_excel_file_label.setObjectName("successText")
                self.update_status(f"Computed {self.additional_new_rows_count} new rows from additional Excel file")
            else:
                self.additional_excel_file_label.setText(
                    f"Selected: {filename} — 0 new rows (no differences found)"
                )
                self.additional_excel_file_label.setObjectName("warningText")
                self.update_status("No new rows found in additional Excel file")

        except Exception as e:
            QMessageBox.critical(self, "Additional Excel Error", f"Error comparing Excel files: {str(e)}")
            self.selected_additional_excel_file = ""
            self.additional_new_rows_data = []
            self.additional_new_rows_count = 0
            self.additional_excel_file_label.setText("Error reading additional Excel file")
            self.additional_excel_file_label.setObjectName("warningText")

    def clear_additional_excel_file(self):
        """Clear selected additional Excel file and any computed diffs"""
        self.selected_additional_excel_file = ""
        self.additional_new_rows_data = []
        self.additional_new_rows_count = 0
        self.additional_excel_file_label.setText("No additional Excel file selected")
        self.additional_excel_file_label.setObjectName("infoText")
        self.update_status("Additional Excel file cleared")

    def compute_new_rows_between_excels(self, base_excel_path: str, additional_excel_path: str):
        """Return list of dict rows that exist in additional Excel but not in base Excel.

        Comparison is done on normalized full-row content across intersecting columns.
        Preserves the order of the additional Excel file.
        """
        # Read both files
        base_df = pd.read_excel(base_excel_path)
        add_df = pd.read_excel(additional_excel_path)

        if base_df.empty:
            # If base is empty, everything is new
            return add_df.to_dict('records')

        # Align columns: use intersection to avoid mismatches
        common_cols = [c for c in add_df.columns if c in set(base_df.columns)]
        if not common_cols:
            # If no common columns, treat all as new to avoid silent drops
            return add_df.to_dict('records')

        def normalize_value(v):
            if pd.isna(v):
                return ""
            # Keep string form for stable comparison
            return str(v).strip().lower()

        def row_signature(series_like):
            # Signature built from common columns only, sorted by column name for determinism
            items = []
            for col in sorted(common_cols):
                items.append((col.lower(), normalize_value(series_like[col] if col in series_like else None)))
            return tuple(items)

        base_signatures = set()
        for _, row in base_df.iterrows():
            base_signatures.add(row_signature(row))

        new_rows = []
        for _, row in add_df.iterrows():
            if row_signature(row) not in base_signatures:
                new_rows.append({k: row[k] for k in add_df.columns})

        return new_rows
    

    
    def open_output_directory(self, directory_path):
        """Open the output directory in file explorer"""
        try:
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                subprocess.run(["explorer", directory_path], check=True)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", directory_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", directory_path], check=True)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open directory: {str(e)}")
    
    # Data handling methods
    def load_existing_delivery_data(self):
        """Load existing delivery data from JSON file"""
        try:
            if os.path.exists(self.delivery_json_file):
                with open(self.delivery_json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.delivery_data_values = data.get("delivery_sequences", [])
                self.delivery_data_with_drivers = data.get("delivery_data_with_drivers", {})
                
                if self.delivery_data_values:
                    self.update_status(f"Loaded {len(self.delivery_data_values)} delivery sequences from file")
                else:
                    self.update_status("No delivery sequence data found - use OptimoRoute Sorter to load data first")
        except Exception as e:
            print(f"Error loading existing data: {e}")
            self.update_status("Ready - use OptimoRoute Sorter to load delivery sequence data first")
    
    # Processing methods
    def process_picking_dockets(self):
        """Process picking dockets with reversed page order and upload Excel to database"""
        # Check for Excel file with order numbers
        if not self.selected_excel_file or not self.excel_order_numbers:
            QMessageBox.warning(
                self, 
                "No Excel File", 
                "Please select a store order Excel file first.\n\n"
                "The application needs this to:\n"
                "• Upload store orders to database in exact Excel row order\n"
                "• Generate barcodes for order numbers in Column A\n"
                "• Match picking dockets to order numbers"
            )
            return
        
        if not self.selected_picking_pdf_files:
            QMessageBox.warning(self, "No Picking PDFs", "Please select picking docket PDF files to process.")
            return
        
        # Check Supabase availability for database upload
        if not SUPABASE_AVAILABLE:
            reply = QMessageBox.question(
                self, 
                "Supabase Unavailable", 
                "Supabase configuration is not available. The Excel file cannot be uploaded to the database.\n\n"
                "Do you want to continue with picking docket processing only (no database upload)?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        self.show_progress(True)
        self.update_status("Starting processing...")
        self.process_picking_btn.setEnabled(False)
        
        # Start background processing
        self.processing_thread = ProcessingThread(self)
        self.processing_thread.progress_signal.connect(self.update_status)
        self.processing_thread.finished_signal.connect(self.on_picking_processing_finished)
        self.processing_thread.start()
    
    def on_picking_processing_finished(self, success, result):
        """Handle picking processing completion"""
        self.show_progress(False)
        self.process_picking_btn.setEnabled(True)
        
        if success:
            self.update_status("Picking dockets processing completed successfully")
            
            # Mark as processed
            self.picking_dockets_processed = True
            
            # Show professional results dialog
            results_dialog = ProcessingResultsDialog(result, self)
            results_dialog.setWindowTitle("Picking Dockets Processing Results")
            results_dialog.exec()
        else:
            error_msg = result.get("error", "Unknown error occurred")
            self.update_status(f"Picking dockets processing failed: {error_msg}")
            
            # Show professional error dialog with any partial results
            if result.get('processed_files', 0) > 0:
                # Some processing was done, show results dialog but with error status
                results_dialog = ProcessingResultsDialog(result, self)
                results_dialog.setWindowTitle("Picking Dockets Processing Completed with Issues")
                results_dialog.exec()
            else:
                QMessageBox.critical(self, "Processing Error", f"Error during picking dockets processing: {error_msg}")

    def process_picking_dockets_internal(self):
        """Internal method for picking dockets processing with barcode generation and Excel upload"""
        import re
        from barcode import Code128
        from barcode.writer import ImageWriter
        import tempfile
        
        try:
            # STEP 1: Upload to Supabase
            if SUPABASE_AVAILABLE:
                # If user selected an additional file, ONLY upload new rows from it
                if getattr(self, 'selected_additional_excel_file', ""):
                    try:
                        add_name = Path(self.selected_additional_excel_file).name
                        # Ensure we have computed diffs; if not, compute now
                        if not getattr(self, 'additional_new_rows_data', None):
                            new_rows_now = self.compute_new_rows_between_excels(self.selected_excel_file, self.selected_additional_excel_file)
                            self.additional_new_rows_data = new_rows_now
                            self.additional_new_rows_count = len(new_rows_now)

                        count = len(self.additional_new_rows_data or [])
                        if count > 0:
                            # Build created_at from date picker (use start of day in ISO format)
                            date_q = self.delivery_date_edit.date()
                            created_at_iso = f"{date_q.toString('yyyy-MM-dd')}T00:00:00+00:00"
                            self.processing_thread.progress_signal.emit(
                                f"📤 Uploading {count} NEW rows from additional file {add_name} to database..."
                            )
                            add_success = upload_store_orders_from_excel(self.additional_new_rows_data, add_name, created_at_override=created_at_iso)
                            if add_success:
                                self.processing_thread.progress_signal.emit(
                                    f"✅ Successfully uploaded {count} new rows from {add_name}"
                                )
                            else:
                                self.processing_thread.progress_signal.emit(
                                    f"⚠️ Failed to upload new rows from {add_name}"
                                )
                        else:
                            self.processing_thread.progress_signal.emit(
                                f"ℹ️ No new rows detected in additional file {add_name} — nothing to upload"
                            )
                    except Exception as e:
                        self.processing_thread.progress_signal.emit(
                            f"⚠️ Error uploading additional Excel new rows: {str(e)}"
                        )
                else:
                    # No additional file selected — upload the full initial Excel file
                    self.processing_thread.progress_signal.emit("📤 Uploading store order Excel file to database...")
                    try:
                        # Read Excel file maintaining row order
                        self.processing_thread.progress_signal.emit(f"Reading {Path(self.selected_excel_file).name} and preserving Excel row order...")
                        df = pd.read_excel(self.selected_excel_file)
                        
                        # Convert DataFrame to list of dictionaries (preserves row order)
                        store_order_data = df.to_dict('records')
                        
                        self.processing_thread.progress_signal.emit(f"Uploading {len(store_order_data)} rows to dispatch_orders table in picking sequence order...")
                        
                        # Upload to Supabase using the function (order-preserving)
                        date_q = self.delivery_date_edit.date()
                        created_at_iso = f"{date_q.toString('yyyy-MM-dd')}T00:00:00+00:00"
                        success = upload_store_orders_from_excel(store_order_data, Path(self.selected_excel_file).name, created_at_override=created_at_iso)
                        
                        if success:
                            self.processing_thread.progress_signal.emit(f"✅ Successfully uploaded {Path(self.selected_excel_file).name} to database with Excel order preserved!")
                        else:
                            self.processing_thread.progress_signal.emit(f"⚠️ Failed to upload {Path(self.selected_excel_file).name} to database - continuing with picking docket processing")
                    except Exception as e:
                        self.processing_thread.progress_signal.emit(f"⚠️ Error uploading Excel file to database: {str(e)} - continuing with picking docket processing")
            else:
                self.processing_thread.progress_signal.emit("⚠️ Supabase not available - skipping database upload")
            
            # STEP 2: Continue with picking docket processing
            # Determine output directory - use selected folder or default
            if self.selected_output_folder:
                base_output_dir = Path(self.selected_output_folder)
            else:
                base_output_dir = Path.cwd() / "picking_dockets_output"
            
            # Create date-based subfolder (YYYY-MM-DD format)
            current_date = datetime.now().strftime("%Y-%m-%d")
            output_dir = base_output_dir / current_date
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Dictionary to store pages grouped by order number
            order_pages = {}
            processed_files = 0
            total_pages_processed = 0
            
            # Dictionary to store generated barcodes for each order ID
            order_barcodes = {}
            
            self.processing_thread.progress_signal.emit("Starting picking dockets processing...")
            self.processing_thread.progress_signal.emit(f"Processing {len(self.selected_picking_pdf_files)} picking docket PDF files...")
            self.processing_thread.progress_signal.emit(f"Looking for {len(self.excel_order_numbers)} order numbers from Excel file...")
            
            # Debug: Show loaded Excel order numbers
            self.processing_thread.progress_signal.emit(f"Excel order numbers to find: {len(self.excel_order_numbers)}")
            for i, order_num in enumerate(self.excel_order_numbers[:5]):  # Show first 5
                self.processing_thread.progress_signal.emit(f"  {i+1}. '{order_num}'")
            if len(self.excel_order_numbers) > 5:
                self.processing_thread.progress_signal.emit(f"  ... and {len(self.excel_order_numbers) - 5} more order numbers")
            
            # First, generate barcodes for all Excel order numbers
            self.processing_thread.progress_signal.emit("Generating barcodes for Excel order numbers...")
            for order_id in self.excel_order_numbers:
                try:
                    # Create barcode using Code128
                    code128 = Code128(order_id, writer=ImageWriter())
                    
                    # Generate barcode as bytes
                    barcode_buffer = io.BytesIO()
                    code128.write(barcode_buffer)
                    barcode_buffer.seek(0)
                    
                    # Store the barcode image data
                    order_barcodes[order_id] = barcode_buffer.getvalue()
                    
                    self.processing_thread.progress_signal.emit(f"Generated barcode for Order ID: {order_id}")
                    
                except Exception as e:
                    self.processing_thread.progress_signal.emit(f"Error generating barcode for {order_id}: {str(e)}")
                    continue
            
            # Process picking docket PDF files
            for pdf_file in self.selected_picking_pdf_files:
                self.processing_thread.progress_signal.emit(f"Processing picking docket: {Path(pdf_file).name}")
                
                try:
                    # Open PDF
                    pdf_document = fitz.open(pdf_file)
                    
                    # Process each page
                    for page_num in range(len(pdf_document)):
                        page = pdf_document[page_num]
                        
                        # Extract text from page
                        page_text = page.get_text()
                        
                        # If no text found, try OCR
                        if not page_text.strip():
                            try:
                                # Render page as image for OCR
                                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x resolution
                                img_data = pix.tobytes("png")
                                img = Image.open(io.BytesIO(img_data))
                                
                                # Perform OCR
                                page_text = pytesseract.image_to_string(img)
                                self.processing_thread.progress_signal.emit(
                                    f"Used OCR for page {page_num + 1} in {Path(pdf_file).name}"
                                )
                            except Exception as ocr_error:
                                self.processing_thread.progress_signal.emit(
                                    f"OCR failed for page {page_num + 1}: {str(ocr_error)}"
                                )
                                page_text = ""
                        
                        # Search for exact order ID matches from Excel data
                        matched_order_id = None
                        
                        # Search for each order ID from Excel data directly in the PDF text
                        for excel_order_id in self.excel_order_numbers:
                            # Case-insensitive search for the exact order ID
                            if excel_order_id.upper() in page_text.upper():
                                matched_order_id = excel_order_id  # Use the exact case from Excel
                                self.processing_thread.progress_signal.emit(
                                    f"✅ Found exact match: '{excel_order_id}' on page {page_num + 1}"
                                )
                                break
                        
                        # If no exact match found, try word boundary search for more precision
                        if not matched_order_id:
                            for excel_order_id in self.excel_order_numbers:
                                # Use word boundaries to avoid partial matches
                                pattern = r'\b' + re.escape(excel_order_id) + r'\b'
                                if re.search(pattern, page_text, re.IGNORECASE):
                                    matched_order_id = excel_order_id
                                    self.processing_thread.progress_signal.emit(
                                        f"✅ Found word boundary match: '{excel_order_id}' on page {page_num + 1}"
                                    )
                                    break
                        
                        # Debug: Show what we found on this page
                        if matched_order_id:
                            self.processing_thread.progress_signal.emit(
                                f"Found Order ID '{matched_order_id}' on page {page_num + 1} of {Path(pdf_file).name}"
                            )
                            
                            # Initialize order group if not exists
                            if matched_order_id not in order_pages:
                                order_pages[matched_order_id] = []
                            
                            # Store page info for this order
                            order_pages[matched_order_id].append({
                                'source_pdf_path': pdf_file,
                                'page_num': page_num,
                                'order_id': matched_order_id,
                                'source_file': pdf_file
                            })
                            
                            self.processing_thread.progress_signal.emit(
                                f"✓ Added page {page_num + 1} to order '{matched_order_id}' group"
                            )
                        else:
                            # Debug: Show first 200 characters of page text to help troubleshoot
                            if page_text.strip():
                                preview_text = page_text.replace('\n', ' ').strip()[:200]
                                self.processing_thread.progress_signal.emit(
                                    f"Page {page_num + 1} text preview: {preview_text}..."
                                )
                        
                        total_pages_processed += 1
                    
                    processed_files += 1
                    pdf_document.close()
                    
                except Exception as e:
                    self.processing_thread.progress_signal.emit(f"Error processing {pdf_file}: {str(e)}")
                    if 'pdf_document' in locals():
                        pdf_document.close()
                    continue
            
            # Save generated barcodes to Supabase if available
            if SUPABASE_AVAILABLE:
                try:
                    self.processing_thread.progress_signal.emit("Saving barcodes to Supabase database...")
                    
                    # Prepare barcode data for Supabase
                    barcode_data_for_db = []
                    for order_id, pages in order_pages.items():
                        for page_info in pages:
                            if order_id in order_barcodes:
                                barcode_record = {
                                    'order_id': order_id,
                                    'driver_number': 'N/A',  # No driver assignment in this workflow
                                    'pdf_file_name': Path(page_info['source_file']).name,
                                    'page_number': page_info['page_num'] + 1,  # Convert to 1-based indexing
                                    'barcode_type': 'Code128'
                                }
                                barcode_data_for_db.append(barcode_record)
                    
                    # Save to Supabase
                    if barcode_data_for_db:
                        success = save_generated_barcodes(barcode_data_for_db)
                        if success:
                            self.processing_thread.progress_signal.emit(f"✅ Successfully saved {len(barcode_data_for_db)} barcodes to Supabase")
                        else:
                            self.processing_thread.progress_signal.emit("⚠️ Failed to save some barcodes to Supabase")
                    else:
                        self.processing_thread.progress_signal.emit("No barcodes to save to Supabase")
                        
                except Exception as e:
                    self.processing_thread.progress_signal.emit(f"⚠️ Error saving barcodes to Supabase: {str(e)}")
                    # Continue processing even if Supabase save fails
                    pass
            else:
                self.processing_thread.progress_signal.emit("⚠️ Supabase not available - barcodes not saved to database")
            
            # Create separate PDF files for each order number with barcodes
            self.processing_thread.progress_signal.emit("Creating order-specific PDF files with barcodes...")
            
            # Show summary of what was found
            total_matched_pages = sum(len(pages) for pages in order_pages.values())
            self.processing_thread.progress_signal.emit(f"Found {total_matched_pages} picking docket pages with matching Order IDs across {len(order_pages)} orders")
            self.processing_thread.progress_signal.emit("📋 Only including pages with order IDs from Excel file - other pages are filtered out")
            
            created_files = []
            failed_files = []
            
            if not order_pages:
                self.processing_thread.progress_signal.emit("No matching orders found in picking docket PDF files!")
                self.processing_thread.progress_signal.emit("Check that your picking docket PDF files contain order IDs that match those in your Excel file")
                return {
                    "processed_files": processed_files,
                    "total_pages": total_pages_processed,
                    "driver_files_created": 0,
                    "created_files": [],
                    "failed_files": [],
                    "driver_details": {},
                    "output_dir": str(output_dir),
                    "barcodes_generated": len(order_barcodes),
                    "error": "No matching orders found in picking docket PDF files"
                }
            
            for order_id, pages in order_pages.items():
                if not pages:
                    continue
                
                try:
                    # Create new PDF for this order
                    output_filename = f"Order_{order_id}_Combined.pdf"
                    output_path = output_dir / output_filename
                    
                    self.processing_thread.progress_signal.emit(
                        f"Creating {output_filename} with {len(pages)} pages and barcode..."
                    )
                    
                    new_pdf = fitz.open()
                    pages_added = 0
                    
                    # Add all pages for this order with barcodes
                    for page_info in pages:
                        try:
                            # Open source PDF and get the page
                            source_pdf = fitz.open(page_info['source_pdf_path'])
                            source_page = source_pdf[page_info['page_num']]
                            
                            # Create a new page in the output PDF
                            new_page = new_pdf.new_page(width=source_page.rect.width, height=source_page.rect.height)
                            
                            # Copy the original page content
                            new_page.show_pdf_page(new_page.rect, source_pdf, page_info['page_num'])
                            
                            # Add barcode at the top center of the page
                            if order_id in order_barcodes:
                                try:
                                    # Insert barcode image at the top center
                                    barcode_data = order_barcodes[order_id]
                                    
                                    # Calculate position for top center
                                    page_width = new_page.rect.width
                                    barcode_width = 700  # Long barcode
                                    barcode_height = 70  # Shorter barcode
                                    
                                    barcode_x = (page_width - barcode_width) / 2  # Center horizontally
                                    barcode_y = 20  # Top margin
                                    
                                    # Insert barcode image
                                    barcode_rect = fitz.Rect(barcode_x, barcode_y, barcode_x + barcode_width, barcode_y + barcode_height)
                                    new_page.insert_image(barcode_rect, stream=barcode_data)
                                    
                                    self.processing_thread.progress_signal.emit(
                                        f"Added barcode for Order {order_id} to page {pages_added + 1}"
                                    )
                                    
                                except Exception as barcode_error:
                                    self.processing_thread.progress_signal.emit(
                                        f"Error adding barcode to page for Order {order_id}: {str(barcode_error)}"
                                    )
                            
                            source_pdf.close()
                            pages_added += 1
                            
                        except Exception as e:
                            self.processing_thread.progress_signal.emit(
                                f"Error processing page for Order {page_info['order_id']}: {str(e)}"
                            )
                            continue
                    
                    # Only save if we successfully added pages
                    if pages_added > 0:
                        new_pdf.save(str(output_path))
                        new_pdf.close()
                        
                        # Verify the file was created
                        if output_path.exists():
                            created_files.append(output_filename)
                            self.processing_thread.progress_signal.emit(
                                f"✓ Successfully created {output_filename} with {pages_added} pages and barcode"
                            )
                        else:
                            failed_files.append(output_filename)
                            self.processing_thread.progress_signal.emit(
                                f"✗ Failed to create {output_filename} - file not found after save"
                            )
                    else:
                        new_pdf.close()
                        failed_files.append(output_filename)
                        self.processing_thread.progress_signal.emit(
                            f"✗ No pages added to {output_filename}"
                        )
                        
                except Exception as e:
                    failed_files.append(f"Order_{order_id}_Combined.pdf")
                    self.processing_thread.progress_signal.emit(
                        f"✗ Error creating PDF for Order {order_id}: {str(e)}"
                    )
                    continue
            
            # Final summary message
            self.processing_thread.progress_signal.emit("Processing complete!")
            if SUPABASE_AVAILABLE:
                self.processing_thread.progress_signal.emit(f"📤 Uploaded Excel file to dispatch_orders database table with Excel row order preserved")
            self.processing_thread.progress_signal.emit(f"Created {len(created_files)} order-specific PDF files in {output_dir}")
            self.processing_thread.progress_signal.emit(f"📅 Files saved in date folder: {current_date}")
            self.processing_thread.progress_signal.emit(f"🏷️  Generated barcodes for {len(order_barcodes)} order numbers from Excel file")
            self.processing_thread.progress_signal.emit("📋 Only pages with order IDs matching Excel file were included - others were filtered out")
            
            # Generate summary report
            summary_path = output_dir / "picking_dockets_summary.txt"
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write("Dispatch Scanning Processing Summary\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Processing Date: {current_date}\n")
                f.write(f"Output Directory: {output_dir}\n")
                f.write(f"Excel file used: {Path(self.selected_excel_file).name}\n")
                f.write(f"Database upload: {'✅ Success' if SUPABASE_AVAILABLE else '❌ Supabase not available'}\n")
                f.write(f"Order numbers from Excel: {len(self.excel_order_numbers)}\n")
                f.write(f"Total picking docket PDF files processed: {processed_files}\n")
                f.write(f"Total pages scanned: {total_pages_processed}\n")
                f.write(f"Order-specific PDF files created: {len(created_files)}\n")
                f.write(f"Barcodes generated: {len(order_barcodes)}\n")
                if failed_files:
                    f.write(f"Failed PDF files: {len(failed_files)}\n")
                f.write("\n")
                f.write("WORKFLOW COMPLETED:\n")
                f.write(f"1. 📤 Uploaded Excel file to dispatch_orders database table (Excel row order preserved)\n")
                f.write(f"2. 🏷️  Generated barcodes for {len(order_barcodes)} order numbers from Excel Column A\n")
                f.write(f"3. 📄 Created {len(created_files)} order-specific picking docket PDFs with barcodes\n")
                f.write(f"4. 📅 Organized all files in date folder: {current_date}\n\n")
                f.write("Each PDF contains all pages for a specific order number with barcodes at the top.\n")
                f.write("Barcodes are generated for order numbers found in Excel Column A.\n\n")
                
                if created_files:
                    f.write("✓ Successfully Created Order PDF Files:\n")
                    for filename in created_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if failed_files:
                    f.write("✗ Failed PDF Files:\n")
                    for filename in failed_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                f.write("Order Page Counts:\n")
                for order_id, pages in order_pages.items():
                    f.write(f"  Order {order_id}: {len(pages)} pages\n")
                
                f.write("\nExcel Order Numbers:\n")
                for order_id in sorted(self.excel_order_numbers):
                    found_pages = len(order_pages.get(order_id, []))
                    f.write(f"  - {order_id} ({found_pages} pages found)\n")
            
            # Collect order details for results dialog (convert to match expected format)
            order_details = {}
            for order_id, pages in order_pages.items():
                order_details[order_id] = {
                    'page_count': len(pages),
                    'orders': [order_id]  # Single order per group in this workflow
                }
            
            return {
                "processed_files": processed_files,
                "total_pages": total_pages_processed,
                "driver_files_created": len(created_files),
                "created_files": created_files,
                "failed_files": failed_files,
                "driver_details": order_details,  # Use order details instead of driver details
                "output_dir": str(output_dir),
                "barcodes_generated": len(order_barcodes),
                "database_upload": SUPABASE_AVAILABLE,
                "excel_file": Path(self.selected_excel_file).name if self.selected_excel_file else "None",
                "additional_file": Path(self.selected_additional_excel_file).name if getattr(self, 'selected_additional_excel_file', "") else "",
                "additional_new_rows_count": getattr(self, 'additional_new_rows_count', 0),
                "additional_new_rows": getattr(self, 'additional_new_rows_data', [])
            }
            
        except Exception as e:
            self.processing_thread.progress_signal.emit(f"Error: {str(e)}")
            raise e
    
    def update_status(self, message):
        """Update the status bar message"""
        self.status_bar.showMessage(message)
    
    def show_progress(self, show=True):
        """Show or hide the progress bar"""
        self.progress_bar.setVisible(show)
        if show:
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
        else:
            self.progress_bar.setRange(0, 1)
            self.progress_bar.setValue(1)
    
    def apply_clean_styling(self):
        """Apply clean, minimal styling"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8fafc;
            }
            
            QFrame#headerFrame {
                background-color: #2563eb;
                border-radius: 8px;
                margin-bottom: 10px;
            }
            
            QLabel#headerTitle {
                color: white;
                font-size: 22px;
                font-weight: bold;
            }
            
            QLabel#headerSubtitle {
                color: #e2e8f0;
                font-size: 13px;
            }
            
            QFrame#section {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 12px;
            }
            
            QLabel {
                color: #374151;
                font-size: 12px;
            }
            
            QLabel#sectionTitle {
                color: #1e293b;
                font-size: 14px;
                font-weight: bold;
                margin-bottom: 3px;
            }
            
            QLabel#workflowInfo {
                color: #6b7280;
                font-size: 11px;
                font-style: italic;
                margin-bottom: 10px;
            }
            
            QLabel#infoText {
                color: #64748b;
                font-size: 12px;
                padding: 8px;
                background-color: #f1f5f9;
                border-radius: 4px;
            }
            
            QLabel#warningText {
                color: #d97706;
                font-size: 12px;
                padding: 8px;
                background-color: #fef3c7;
                border-radius: 4px;
                font-weight: 500;
            }
            
            QLabel#successText {
                color: #059669;
                font-size: 12px;
                padding: 8px;
                background-color: #d1fae5;
                border-radius: 4px;
                font-weight: 500;
            }
            
            QPushButton {
                background-color: #e2e8f0;
                color: #374151;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
                min-height: 16px;
                max-width: 200px;
                font-size: 12px;
            }
            
            QPushButton:hover {
                background-color: #cbd5e1;
            }
            
            QPushButton#primaryButton {
                background-color: #2563eb;
                color: white;
            }
            
            QPushButton#primaryButton:hover {
                background-color: #1d4ed8;
            }
            
            QPushButton#secondaryButton {
                background-color: #6b7280;
                color: white;
            }
            
            QPushButton#secondaryButton:hover {
                background-color: #4b5563;
            }
            
            QLineEdit {
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 6px 10px;
                background-color: white;
                color: #374151;
                font-size: 12px;
            }
            
            QLineEdit:focus {
                border-color: #2563eb;
            }
            
            QListWidget, QTableWidget {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: white;
                color: #374151;
                alternate-background-color: #f8fafc;
            }
            
            /* Scrollbar Styling */
            QScrollBar:vertical {
                background-color: #f8fafc;
                width: 12px;
                border: none;
                border-radius: 6px;
            }
            
            QScrollBar::handle:vertical {
                background-color: #cbd5e1;
                border-radius: 6px;
                min-height: 20px;
                margin: 2px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: #94a3b8;
            }
            
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }
            
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            
            QScrollBar:horizontal {
                background-color: #f8fafc;
                height: 12px;
                border: none;
                border-radius: 6px;
            }
            
            QScrollBar::handle:horizontal {
                background-color: #cbd5e1;
                border-radius: 6px;
                min-width: 20px;
                margin: 2px;
            }
            
            QScrollBar::handle:horizontal:hover {
                background-color: #94a3b8;
            }
            
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
            }
            
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
            
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f1f5f9;
            }
            
            QTableWidget::item:selected {
                background-color: #eff6ff;
                color: #1e40af;
            }
            
            QTableWidget QHeaderView::section {
                background-color: #f8fafc;
                border: none;
                border-bottom: 1px solid #e2e8f0;
                padding: 8px;
                font-weight: 600;
                color: #374151;
            }
            
            QStatusBar {
                background-color: #f1f5f9;
                border-top: 1px solid #e2e8f0;
                color: #374151;
            }
            
            QProgressBar {
                border: 1px solid #d1d5db;
                border-radius: 4px;
                text-align: center;
                background-color: white;
                color: #374151;
            }
            
            QProgressBar::chunk {
                background-color: #2563eb;
                border-radius: 3px;
            }
            
            /* Message Box Styling */
            QMessageBox {
                background-color: white;
                color: #374151;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 20px;
                font-size: 14px;
            }
            
            QMessageBox QLabel {
                background-color: transparent;
                color: #374151;
                font-size: 14px;
                padding: 10px;
            }
            
            QMessageBox QPushButton {
                background-color: #2563eb;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 14px;
                min-width: 80px;
            }
            
            QMessageBox QPushButton:hover {
                background-color: #1d4ed8;
            }
            
            QMessageBox QPushButton:pressed {
                background-color: #1e40af;
            }
        """)


def main():
    app = QApplication(sys.argv)
    window = DispatchScanningApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 