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

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QPushButton, QTextEdit, QLineEdit, 
    QFileDialog, QMessageBox, QProgressBar, QStatusBar, QFrame,
    QScrollArea, QGroupBox, QSplitter, QComboBox, QDialog, 
    QDialogButtonBox, QListWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QPlainTextEdit, QCheckBox, QTabWidget
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize
from PySide6.QtGui import QFont, QPalette, QColor, QIcon, QPixmap

# Import Supabase configuration
from supabase_config import save_generated_barcodes

# Import OptimoRoute tab
from optimoroute_tab import OptimoRouteTab


class ProcessingThread(QThread):
    """Background thread for PDF processing operations"""
    progress_signal = Signal(str)
    finished_signal = Signal(bool, dict)
    
    def __init__(self, app_instance, operation_type, **kwargs):
        super().__init__()
        self.app = app_instance
        self.operation_type = operation_type
        self.kwargs = kwargs
    
    def run(self):
        try:
            if self.operation_type == "process_all":
                result = self.app.process_all_pdfs_and_packing_internal()
                self.finished_signal.emit(True, result)
            elif self.operation_type == "process_picking":
                result = self.app.process_picking_dockets_internal()
                self.finished_signal.emit(True, result)
            else:
                self.finished_signal.emit(False, {})
        except Exception as e:
            self.progress_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(False, {"error": str(e)})


class DataPreviewDialog(QDialog):
    """Dialog for data validation and column selection"""
    
    def __init__(self, data, column_name, parent=None):
        super().__init__(parent)
        self.data = data
        self.column_name = column_name
        self.selected_column = None
        self.confirmed = False
        
        self.setWindowTitle("Data Validation")
        self.setModal(True)
        self.resize(800, 600)
        
        layout = QVBoxLayout(self)
        
        # Info label
        info_label = QLabel(f"Preview data from column: {column_name}")
        info_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(info_label)
        
        # Table for data preview
        self.table = QTableWidget()
        self.table.setColumnCount(len(data.columns))
        self.table.setHorizontalHeaderLabels([str(col) for col in data.columns])
        self.table.setRowCount(min(20, len(data)))
        
        for row in range(min(20, len(data))):
            for col in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[row, col]))
                self.table.setItem(row, col, item)
        
        self.table.resizeColumnsToContents()
        layout.addWidget(self.table)
        
        # Column selection
        selection_layout = QHBoxLayout()
        selection_layout.addWidget(QLabel("Select column to use:"))
        
        self.column_combo = QComboBox()
        self.column_combo.addItems([str(col) for col in data.columns])
        if column_name in data.columns:
            self.column_combo.setCurrentText(str(column_name))
        selection_layout.addWidget(self.column_combo)
        
        layout.addLayout(selection_layout)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept_selection)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def accept_selection(self):
        self.selected_column = self.column_combo.currentText()
        self.confirmed = True
        self.accept()


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
        title_text = "PDF Processing Completed Successfully"
        if results.get('driver_files_created', 0) == 0:
            title_text = "PDF Processing Completed - No Matching Orders Found"
        elif results.get('failed_files'):
            title_text = "PDF Processing Completed with Some Issues"
        
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
            ("PDF Files Processed", str(results.get('processed_files', 0))),
            ("Pages Scanned", str(results.get('total_pages', 0))),
            ("Driver PDFs Created", str(results.get('driver_files_created', 0))),
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
        
        # Tab 2: Driver Details
        if results.get('driver_details'):
            driver_tab = self.create_driver_tab(results.get('driver_details', {}))
            tab_widget.addTab(driver_tab, "Driver Details")
        
        # Tab 3: Failed Files (if any)
        if results.get('failed_files'):
            failed_tab = self.create_failed_tab(results.get('failed_files', []))
            tab_widget.addTab(failed_tab, "Failed Files")
        
        # If no tabs were created, show diagnostic information
        if tab_widget.count() == 0:
            if results.get('error') == "No matching orders found in PDF files":
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
        
        message = "No matching orders were found in the PDF files.\n\n"
        message += "Troubleshooting Steps:\n"
        message += "=" * 50 + "\n\n"
        message += "1. Check that your PDF files contain order IDs that match those in your Excel file\n"
        message += "2. Ensure order IDs in PDF match those in your Excel file exactly\n"
        message += "3. Order ID matching is case-insensitive (AA061B4Y = aa061b4y)\n\n"
        message += f"Processing Summary:\n"
        message += f"- PDF files processed: {results.get('processed_files', 0)}\n"
        message += f"- Total pages scanned: {results.get('total_pages', 0)}\n"
        message += f"- No matching order IDs found\n\n"
        message += "Common Issues:\n"
        message += "- Order IDs in PDF don't match those in Excel file\n"
        message += "- PDF contains images that need OCR processing\n"
        message += "- Order IDs in Excel file don't match those in PDF\n"
        message += "- Text extraction failed from PDF pages\n\n"
        message += "Check the main application status messages for more detailed information about what was found on each page."
        
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


class TransportSorterApp(QMainWindow):
    """Main application class for Transport Sorter PySide6 version"""
    
    def __init__(self):
        super().__init__()
        
        # Application data
        self.delivery_data_values = []
        self.delivery_data_with_drivers = {}
        self.delivery_json_file = "delivery_sequence_data.json"
        self.selected_pdf_files = []
        self.selected_picking_pdf_files = []  # New variable for picking dockets
        self.selected_excel_files = []  # New variable for Excel order files
        self.selected_store_order_files = []  # New variable for store order Excel files
        self.processed_drivers = {}
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
        self.setWindowTitle("AN Dispatch Sorter")
        self.setGeometry(100, 100, 1400, 900)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        header = self.create_header()
        main_layout.addWidget(header)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("mainTabs")
        
        # Create tabs
        transport_tab = self.create_transport_tab()
        dispatch_tab = self.create_dispatch_tab()
        optimoroute_tab = OptimoRouteTab()
        
        self.tab_widget.addTab(transport_tab, "Transport Sorter")
        self.tab_widget.addTab(dispatch_tab, "Dispatch Scanning")
        self.tab_widget.addTab(optimoroute_tab, "OptimoRoute Orders")
        
        main_layout.addWidget(self.tab_widget)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status("Ready")
        
        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
    
    def create_header(self):
        """Create clean header"""
        header_frame = QFrame()
        header_frame.setFixedHeight(50)
        
        layout = QHBoxLayout(header_frame)
        layout.setContentsMargins(20, 10, 20, 10)
        
        title_label = QLabel("Transport Sorter Professional")
        title_label.setObjectName("headerTitle")
        layout.addWidget(title_label)
        
        layout.addStretch()
        
        subtitle_label = QLabel("PDF Delivery Sorting")
        subtitle_label.setObjectName("headerSubtitle")
        layout.addWidget(subtitle_label)
        
        return header_frame
    
    def create_setup_section(self):
        """Create clean setup section"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        
        # Section title
        title = QLabel("1. Setup")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
        # Output directory
        layout.addWidget(QLabel("Output Directory:"))
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("Select output directory...")
        layout.addWidget(self.output_dir_edit)
        
        output_btn = QPushButton("Browse")
        output_btn.clicked.connect(self.browse_output_directory)
        layout.addWidget(output_btn)
        
        # Spacer
        layout.addSpacing(20)
        
        # Delivery file
        layout.addWidget(QLabel("Delivery Sequence File:"))
        self.delivery_file_edit = QLineEdit()
        self.delivery_file_edit.setPlaceholderText("Select CSV/Excel file...")
        layout.addWidget(self.delivery_file_edit)
        
        delivery_btn = QPushButton("Browse")
        delivery_btn.clicked.connect(self.browse_delivery_file)
        layout.addWidget(delivery_btn)
        
        self.load_data_btn = QPushButton("Load Data")
        self.load_data_btn.setObjectName("primaryButton")
        self.load_data_btn.clicked.connect(self.load_delivery_file)
        layout.addWidget(self.load_data_btn)
        
        layout.addStretch()
        
        return section
    
    def create_data_section(self):
        """Create clean data preview section"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        
        # Section title
        title = QLabel("2. Data Preview")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
        # Data table
        self.data_table = QTableWidget()
        self.data_table.setColumnCount(4)
        self.data_table.setHorizontalHeaderLabels(["#", "Order ID", "Stop Number", "Driver"])
        self.data_table.horizontalHeader().setStretchLastSection(True)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.verticalHeader().setVisible(False)
        layout.addWidget(self.data_table)
        
        return section
    
    def create_process_section(self):
        """Create clean processing section"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        
        # Section title
        title = QLabel("3. Process Delivery PDFs")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
        # PDF files
        layout.addWidget(QLabel("Delivery PDF Files:"))
        
        # Button row
        btn_layout = QHBoxLayout()
        add_pdf_btn = QPushButton("Add PDFs")
        add_pdf_btn.clicked.connect(self.browse_pdf_files)
        
        clear_pdf_btn = QPushButton("Clear")
        clear_pdf_btn.setObjectName("secondaryButton")
        clear_pdf_btn.clicked.connect(self.clear_pdf_files)
        
        btn_layout.addWidget(add_pdf_btn)
        btn_layout.addWidget(clear_pdf_btn)
        layout.addLayout(btn_layout)
        
        # PDF list
        self.pdf_list = QListWidget()
        self.pdf_list.setMaximumHeight(120)
        layout.addWidget(self.pdf_list)
        
        # Info
        info_label = QLabel("Processes delivery PDFs and groups by driver in delivery sequence order.")
        info_label.setObjectName("infoText")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        layout.addStretch()
        
        return section
    
    def create_picking_section(self):
        """Create picking dockets section"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        
        # Section title
        title = QLabel("4. Process Picking Dockets")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
        # PDF files subsection
        pdf_label = QLabel("Picking Docket PDF Files:")
        pdf_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
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
        self.picking_pdf_list.setMaximumHeight(120)
        layout.addWidget(self.picking_pdf_list)
        
        # Info
        info_label = QLabel("Processes picking dockets with REVERSED page order so first delivery stops are at the top of the pallet.")
        info_label.setObjectName("infoText")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        layout.addStretch()
        
        return section
    
    def create_store_order_section(self):
        """Create store order upload section"""
        section = QFrame()
        section.setObjectName("section")

        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        
        # Section title
        title = QLabel("5. Store Order Upload")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
        # Description
        desc_label = QLabel("Upload store order Excel files to pick_lists table")
        desc_label.setObjectName("workflowInfo")
        layout.addWidget(desc_label)
        
        # Excel files subsection
        excel_label = QLabel("Store Order Excel Files:")
        excel_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(excel_label)
        
        # Button row
        btn_layout = QHBoxLayout()
        self.add_store_order_btn = QPushButton("Add Excel Files")
        self.add_store_order_btn.clicked.connect(self.browse_store_order_files)
        
        self.clear_store_order_btn = QPushButton("Clear")
        self.clear_store_order_btn.setObjectName("secondaryButton")
        self.clear_store_order_btn.clicked.connect(self.clear_store_order_files)
        
        btn_layout.addWidget(self.add_store_order_btn)
        btn_layout.addWidget(self.clear_store_order_btn)
        layout.addLayout(btn_layout)
        
        # Excel file list
        self.store_order_file_list = QListWidget()
        self.store_order_file_list.setMaximumHeight(120)
        layout.addWidget(self.store_order_file_list)
        
        # Upload button
        self.upload_store_order_btn = QPushButton("Upload to Supabase")
        self.upload_store_order_btn.setObjectName("primaryButton")
        self.upload_store_order_btn.clicked.connect(self.upload_store_orders_to_supabase)
        layout.addWidget(self.upload_store_order_btn)
        
        # Column mapping info
        mapping_info = QLabel("Expected Excel columns:\n• Column A: Order Number (→ order_id)\n• Column B: Item Code (→ item_code)\n• Column C: Quantity (→ quantity_required)")
        mapping_info.setObjectName("infoText")
        mapping_info.setWordWrap(True)
        layout.addWidget(mapping_info)
        
        # Status message
        self.store_order_status_label = QLabel("Ready to upload store order files")
        self.store_order_status_label.setObjectName("infoText")
        self.store_order_status_label.setWordWrap(True)
        layout.addWidget(self.store_order_status_label)
        
        layout.addStretch()
        
        return section
    
    def create_transport_tab(self):
        """Create the Transport Sorter tab with sections 1, 2, 3"""
        tab_widget = QWidget()
        tab_layout = QVBoxLayout(tab_widget)
        tab_layout.setSpacing(20)
        tab_layout.setContentsMargins(20, 20, 20, 20)
        
        # Content area - 3 column grid
        content_widget = QWidget()
        content_layout = QGridLayout(content_widget)
        content_layout.setSpacing(15)
        
        # Create sections
        setup_section = self.create_setup_section()
        data_section = self.create_data_section()
        process_section = self.create_process_section()
        
        content_layout.addWidget(setup_section, 0, 0)
        content_layout.addWidget(data_section, 0, 1)
        content_layout.addWidget(process_section, 0, 2)
        
        # Set equal column widths
        for i in range(3):
            content_layout.setColumnStretch(i, 1)
        
        tab_layout.addWidget(content_widget)
        
        # Process button
        self.process_all_btn = QPushButton("Process Delivery PDFs")
        self.process_all_btn.setObjectName("primaryButton")
        self.process_all_btn.clicked.connect(self.process_all_pdfs_and_packing)
        self.process_all_btn.setFixedHeight(50)
        tab_layout.addWidget(self.process_all_btn)
        
        return tab_widget
    
    def create_dispatch_tab(self):
        """Create the Dispatch Scanning tab with sections 4, 5 and output"""
        tab_widget = QWidget()
        tab_layout = QVBoxLayout(tab_widget)
        tab_layout.setSpacing(20)
        tab_layout.setContentsMargins(20, 20, 20, 20)
        
        # Content area - 2 column grid
        content_widget = QWidget()
        content_layout = QGridLayout(content_widget)
        content_layout.setSpacing(15)
        
        # Create sections
        picking_section = self.create_picking_section()
        store_order_section = self.create_store_order_section()
        
        content_layout.addWidget(picking_section, 0, 0)
        content_layout.addWidget(store_order_section, 0, 1)
        
        # Set equal column widths
        for i in range(2):
            content_layout.setColumnStretch(i, 1)
        
        tab_layout.addWidget(content_widget)
        
        # Process button
        self.process_picking_btn = QPushButton("Process Picking Dockets (Reversed)")
        self.process_picking_btn.setObjectName("primaryButton")
        self.process_picking_btn.clicked.connect(self.process_picking_dockets)
        self.process_picking_btn.setFixedHeight(50)
        tab_layout.addWidget(self.process_picking_btn)
        
        # Output section
        output_section = self.create_output_section()
        tab_layout.addWidget(output_section)
        
        return tab_widget
    
    def create_output_section(self):
        """Create output section for the dispatch tab"""
        section = QFrame()
        section.setObjectName("section")
        
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        
        # Section title
        title = QLabel("Output & Results")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        
        # Output directory info
        output_info = QLabel("Processed files will be saved to the selected output directory.")
        output_info.setObjectName("infoText")
        output_info.setWordWrap(True)
        layout.addWidget(output_info)
        
        # Open output directory button
        self.open_output_btn = QPushButton("Open Output Directory")
        self.open_output_btn.setObjectName("secondaryButton")
        self.open_output_btn.clicked.connect(lambda: self.open_output_directory(self.output_dir_edit.text()))
        layout.addWidget(self.open_output_btn)
        
        layout.addStretch()
        
        return section
    
    def apply_clean_styling(self):
        """Apply clean, minimal styling"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8fafc;
            }
            
            QTabWidget#mainTabs {
                background-color: #f8fafc;
            }
            
            QTabWidget#mainTabs::pane {
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                background-color: #f8fafc;
                margin-top: 5px;
            }
            
            QTabWidget#mainTabs::tab-bar {
                alignment: center;
            }
            
            QTabBar::tab {
                background-color: #e2e8f0;
                color: #374151;
                padding: 12px 24px;
                margin-right: 2px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: 500;
                font-size: 14px;
            }
            
            QTabBar::tab:selected {
                background-color: #2563eb;
                color: white;
            }
            
            QTabBar::tab:hover:!selected {
                background-color: #cbd5e1;
            }
            
            QFrame#section {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
            }
            
            QLabel {
                color: #374151;
                font-size: 13px;
            }
            
            QLabel#headerTitle {
                color: #1e293b;
                font-size: 22px;
                font-weight: bold;
            }
            
            QLabel#headerSubtitle {
                color: #64748b;
                font-size: 14px;
            }
            
            QLabel#sectionTitle {
                color: #1e293b;
                font-size: 16px;
                font-weight: bold;
                margin-bottom: 5px;
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
            
            QPushButton {
                background-color: #e2e8f0;
                color: #374151;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                min-height: 20px;
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
                border-radius: 6px;
                padding: 8px 12px;
                background-color: white;
                color: #374151;
                font-size: 13px;
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
            
            QPlainTextEdit {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: white;
                color: #374151;
                padding: 8px;
                font-family: 'Consolas', monospace;
                font-size: 12px;
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
            
            QComboBox {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px 12px;
                background-color: white;
                color: #374151;
                font-size: 13px;
            }
            
            QComboBox:hover {
                border-color: #2563eb;
            }
            
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #d1d5db;
                background-color: #f8fafc;
            }
            
            QComboBox::down-arrow {
                image: none;
                border: none;
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #374151;
            }
        """)
    
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
    
    # File handling methods
    def browse_output_directory(self):
        """Browse for output directory"""
        directory = QFileDialog.getExistingDirectory(
            self, 
            "Select Output Directory",
            str(Path.home())
        )
        if directory:
            self.output_dir_edit.setText(directory)
    
    def browse_delivery_file(self):
        """Browse for delivery sequence file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Delivery Sequence File",
            str(Path.home()),
            "Data files (*.xlsx *.xls *.csv);;Excel files (*.xlsx *.xls);;CSV files (*.csv);;All files (*.*)"
        )
        if file_path:
            self.delivery_file_edit.setText(file_path)
    
    def browse_pdf_files(self):
        """Browse for PDF files to process"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF Files",
            str(Path.home()),
            "PDF files (*.pdf);;All files (*.*)"
        )
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.selected_pdf_files:
                    self.selected_pdf_files.append(file_path)
                    self.pdf_list.addItem(Path(file_path).name)
    
    def clear_pdf_files(self):
        """Clear selected PDF files"""
        self.selected_pdf_files.clear()
        self.pdf_list.clear()
    
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
        self.disable_excel_upload()
    
    def browse_excel_files(self):
        """Browse for Excel order files to upload"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Excel Order Files",
            str(Path.home()),
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.selected_excel_files:
                    self.selected_excel_files.append(file_path)
                    self.excel_file_list.addItem(Path(file_path).name)
    
    def clear_excel_files(self):
        """Clear selected Excel files"""
        self.selected_excel_files.clear()
        self.excel_file_list.clear()
    
    def enable_excel_upload(self):
        """Enable Excel upload functionality after picking dockets are processed"""
        self.add_excel_btn.setEnabled(True)
        self.clear_excel_btn.setEnabled(True)
        self.upload_excel_btn.setEnabled(True)
        
        # Update status message
        self.excel_status_label.setText("✅ Picking dockets processed! You can now upload Excel files to match with generated barcodes.")
        self.excel_status_label.setObjectName("successText")
        self.excel_status_label.setStyleSheet("""
            QLabel {
                color: #059669;
                font-size: 12px;
                padding: 8px;
                background-color: #d1fae5;
                border-radius: 4px;
                font-weight: 500;
            }
        """)
    
    def disable_excel_upload(self):
        """Disable Excel upload functionality when picking dockets are not processed"""
        self.add_excel_btn.setEnabled(False)
        self.clear_excel_btn.setEnabled(False)
        self.upload_excel_btn.setEnabled(False)
        
        # Reset status message
        self.excel_status_label.setText("⚠️ Process picking dockets first to generate barcodes before uploading Excel files.")
        self.excel_status_label.setObjectName("warningText")
        self.excel_status_label.setStyleSheet("""
            QLabel {
                color: #d97706;
                font-size: 12px;
                padding: 8px;
                background-color: #fef3c7;
                border-radius: 4px;
                font-weight: 500;
            }
        """)
    
    def upload_excel_to_supabase(self):
        """Upload Excel files to Supabase"""
        if not self.picking_dockets_processed:
            QMessageBox.warning(
                self, 
                "Process Picking Dockets First", 
                "Please process picking dockets first to generate barcodes before uploading Excel files.\n\n"
                "The Excel upload matches your order data with the generated barcodes."
            )
            return
        
        if not self.selected_excel_files:
            QMessageBox.warning(self, "No Files Selected", "Please select Excel files first.")
            return
        
        try:
            from excel_upload_example import upload_excel_pick_list
            
            self.update_status("Uploading Excel files to Supabase...")
            
            success_count = 0
            total_files = len(self.selected_excel_files)
            
            for file_path in self.selected_excel_files:
                try:
                    success = upload_excel_pick_list(file_path)
                    if success:
                        success_count += 1
                        self.update_status(f"Uploaded {Path(file_path).name} successfully...")
                    else:
                        self.update_status(f"Failed to upload {Path(file_path).name}")
                except Exception as e:
                    self.update_status(f"Error uploading {Path(file_path).name}: {str(e)}")
            
            if success_count == total_files:
                QMessageBox.information(
                    self, 
                    "Upload Complete", 
                    f"Successfully uploaded all {success_count} Excel files to Supabase!\n\n"
                    f"Pick lists are now linked to your generated barcodes."
                )
                self.update_status(f"✅ Successfully uploaded {success_count} Excel files to Supabase")
            else:
                QMessageBox.warning(
                    self, 
                    "Upload Partially Complete", 
                    f"Uploaded {success_count} out of {total_files} files successfully.\n\n"
                    f"Check the status messages for details about failed uploads."
                )
                self.update_status(f"⚠️ Uploaded {success_count}/{total_files} Excel files")
                
        except ImportError:
            QMessageBox.critical(
                self, 
                "Missing Dependencies", 
                "Required Supabase dependencies not found.\n\n"
                "Please ensure supabase_config.py and excel_upload_example.py are in the project directory."
            )
        except Exception as e:
            QMessageBox.critical(self, "Upload Error", f"Error uploading Excel files: {str(e)}")
            self.update_status(f"❌ Error uploading Excel files: {str(e)}")
    
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
    
    def browse_store_order_files(self):
        """Browse for store order Excel files"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Store Order Excel Files",
            str(Path.home()),
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.selected_store_order_files:
                    self.selected_store_order_files.append(file_path)
                    self.store_order_file_list.addItem(Path(file_path).name)
    
    def clear_store_order_files(self):
        """Clear selected store order Excel files"""
        self.selected_store_order_files.clear()
        self.store_order_file_list.clear()
        self.store_order_status_label.setText("Ready to upload store order files")
        self.store_order_status_label.setObjectName("infoText")
        self.store_order_status_label.setStyleSheet("""
            QLabel {
                color: #64748b;
                font-size: 12px;
                padding: 8px;
                background-color: #f1f5f9;
                border-radius: 4px;
            }
        """)
    
    def upload_store_orders_to_supabase(self):
        """Upload store order Excel files to Supabase"""
        if not self.selected_store_order_files:
            QMessageBox.warning(self, "No Files Selected", "Please select store order Excel files first.")
            return
        
        try:
            from supabase_config import upload_store_orders_from_excel
            
            self.update_status("Uploading store order files to Supabase...")
            
            success_count = 0
            total_files = len(self.selected_store_order_files)
            
            for file_path in self.selected_store_order_files:
                try:
                    # Read Excel file
                    df = pd.read_excel(file_path)
                    
                    # Convert DataFrame to list of dictionaries
                    store_order_data = df.to_dict('records')
                    
                    # Upload to Supabase using the new function
                    success = upload_store_orders_from_excel(store_order_data, Path(file_path).name)
                    
                    if success:
                        success_count += 1
                        self.update_status(f"Uploaded {Path(file_path).name} successfully...")
                    else:
                        self.update_status(f"Failed to upload {Path(file_path).name}")
                except Exception as e:
                    self.update_status(f"Error uploading {Path(file_path).name}: {str(e)}")
            
            if success_count == total_files:
                QMessageBox.information(
                    self, 
                    "Upload Complete", 
                    f"Successfully uploaded all {success_count} store order files to Supabase!\n\n"
                    f"Store orders are now available in the pick_lists table."
                )
                self.update_status(f"✅ Successfully uploaded {success_count} store order files to Supabase")
                
                # Update status label
                self.store_order_status_label.setText("✅ Store order files uploaded successfully!")
                self.store_order_status_label.setObjectName("successText")
                self.store_order_status_label.setStyleSheet("""
                    QLabel {
                        color: #059669;
                        font-size: 12px;
                        padding: 8px;
                        background-color: #d1fae5;
                        border-radius: 4px;
                        font-weight: 500;
                    }
                """)
            else:
                QMessageBox.warning(
                    self, 
                    "Upload Partially Complete", 
                    f"Uploaded {success_count} out of {total_files} files successfully.\n\n"
                    f"Check the status messages for details about failed uploads."
                )
                self.update_status(f"⚠️ Uploaded {success_count}/{total_files} store order files")
                
                # Update status label
                self.store_order_status_label.setText(f"⚠️ Uploaded {success_count}/{total_files} files - see status for details")
                self.store_order_status_label.setObjectName("warningText")
                self.store_order_status_label.setStyleSheet("""
                    QLabel {
                        color: #d97706;
                        font-size: 12px;
                        padding: 8px;
                        background-color: #fef3c7;
                        border-radius: 4px;
                        font-weight: 500;
                    }
                """)
                
        except ImportError:
            QMessageBox.critical(
                self, 
                "Missing Dependencies", 
                "Required Supabase dependencies not found.\n\n"
                "Please ensure supabase_config.py is in the project directory."
            )
        except Exception as e:
            QMessageBox.critical(self, "Upload Error", f"Error uploading store order files: {str(e)}")
            self.update_status(f"❌ Error uploading store order files: {str(e)}")
            
            # Update status label
            self.store_order_status_label.setText(f"❌ Error uploading files: {str(e)}")
            self.store_order_status_label.setObjectName("warningText")
            self.store_order_status_label.setStyleSheet("""
                QLabel {
                    color: #dc2626;
                    font-size: 12px;
                    padding: 8px;
                    background-color: #fee2e2;
                    border-radius: 4px;
                    font-weight: 500;
                }
            """)
    
    # Data handling methods
    def load_delivery_file(self):
        """Load delivery sequence data from file"""
        file_path = self.delivery_file_edit.text()
        if not file_path:
            QMessageBox.warning(self, "No File Selected", "Please select a delivery sequence file first.")
            return
        
        try:
            self.update_status("Loading delivery data...")
            
            # Load data based on file type
            if file_path.lower().endswith('.csv'):
                data = pd.read_csv(file_path)
            else:
                data = pd.read_excel(file_path)
            
            if data is None or data.empty:
                QMessageBox.warning(self, "No Data", "The selected file contains no data.")
                return
            
            # Ensure we have at least 3 columns (Order ID, Stop Number, Driver Number)
            if len(data.columns) < 3:
                QMessageBox.warning(self, "Invalid Data", "The file must have at least 3 columns: Order ID, Stop Number, Driver Number.")
                return
            
            # Extract Order ID and Driver Number mapping
            self.delivery_data_values = []
            self.delivery_data_with_drivers = {}
            
            for index, row in data.iterrows():
                order_id = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                stop_number = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                driver_number = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                
                if order_id and driver_number:
                    self.delivery_data_values.append(order_id)
                    self.delivery_data_with_drivers[order_id] = {
                        'stop_number': stop_number,
                        'driver_number': driver_number
                    }
            
            # Save to JSON
            self.save_delivery_data()
            
            # Update display
            self.update_delivery_display()
            self.update_status(f"Loaded {len(self.delivery_data_values)} delivery sequences with driver assignments")
            
            # Debug: Show first few loaded entries
            debug_info = f"Successfully loaded {len(self.delivery_data_values)} delivery sequences with driver assignments.\n\n"
            debug_info += "First 5 entries:\n"
            for i, (order_id, data) in enumerate(list(self.delivery_data_with_drivers.items())[:5]):
                debug_info += f"  {i+1}. Order '{order_id}' → Driver '{data['driver_number']}'\n"
            if len(self.delivery_data_with_drivers) > 5:
                debug_info += f"  ... and {len(self.delivery_data_with_drivers) - 5} more entries"
            
            QMessageBox.information(
                self, 
                "Data Loaded", 
                debug_info
            )
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Error loading file: {str(e)}")
            self.update_status(f"Error loading file: {str(e)}")
    
    def save_delivery_data(self):
        """Save delivery data to JSON file"""
        try:
            data = {
                "delivery_sequences": self.delivery_data_values,
                "delivery_data_with_drivers": self.delivery_data_with_drivers,
                "source_file": self.delivery_file_edit.text(),
                "total_records": len(self.delivery_data_values),
                "created_date": pd.Timestamp.now().isoformat()
            }
            
            with open(self.delivery_json_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving delivery data: {e}")
    
    def load_existing_delivery_data(self):
        """Load existing delivery data from JSON file"""
        try:
            if os.path.exists(self.delivery_json_file):
                with open(self.delivery_json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.delivery_data_values = data.get("delivery_sequences", [])
                self.delivery_data_with_drivers = data.get("delivery_data_with_drivers", {})
                if "source_file" in data:
                    self.delivery_file_edit.setText(data["source_file"])
                
                self.update_delivery_display()
        except Exception as e:
            print(f"Error loading existing data: {e}")
    
    def update_delivery_display(self):
        """Update the delivery data display table"""
        self.data_table.setRowCount(len(self.delivery_data_values))
        
        for i, order_id in enumerate(self.delivery_data_values):
            self.data_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.data_table.setItem(i, 1, QTableWidgetItem(str(order_id)))
            
            # Get driver data if available
            driver_data = self.delivery_data_with_drivers.get(order_id, {})
            stop_number = driver_data.get('stop_number', '')
            driver_number = driver_data.get('driver_number', '')
            
            self.data_table.setItem(i, 2, QTableWidgetItem(str(stop_number)))
            self.data_table.setItem(i, 3, QTableWidgetItem(str(driver_number)))
        
        self.data_table.resizeColumnsToContents()
    
    # Processing methods
    def process_all_pdfs_and_packing(self):
        """Process all PDFs with packing functionality"""
        if not self.delivery_data_values:
            QMessageBox.warning(self, "No Data", "Please load delivery sequence data first.")
            return
        
        output_dir = self.output_dir_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "No Output Directory", "Please select an output directory first.")
            return
        
        if not self.selected_pdf_files:
            QMessageBox.warning(self, "No PDFs", "Please select PDF files to process.")
            return
        
        self.show_progress(True)
        self.update_status("Processing PDFs...")
        self.process_all_btn.setEnabled(False)
        
        # Start background processing
        self.processing_thread = ProcessingThread(self, "process_all")
        self.processing_thread.progress_signal.connect(self.update_status)
        self.processing_thread.finished_signal.connect(self.on_processing_finished)
        self.processing_thread.start()
    
    def on_processing_finished(self, success, result):
        """Handle processing completion"""
        self.show_progress(False)
        self.process_all_btn.setEnabled(True)
        
        if success:
            self.update_status("Processing completed successfully")
            
            # Show professional results dialog
            results_dialog = ProcessingResultsDialog(result, self)
            results_dialog.exec()
        else:
            error_msg = result.get("error", "Unknown error occurred")
            self.update_status(f"Processing failed: {error_msg}")
            
            # Show professional error dialog with any partial results
            if result.get('processed_files', 0) > 0:
                # Some processing was done, show results dialog but with error status
                results_dialog = ProcessingResultsDialog(result, self)
                results_dialog.setWindowTitle("Processing Completed with Issues")
                # Update the status icon to warning
                results_dialog.findChild(QLabel).setText("⚠")
                results_dialog.findChild(QLabel).setStyleSheet("font-size: 32px; color: #f59e0b; font-weight: bold;")
                results_dialog.exec()
            else:
                QMessageBox.critical(self, "Processing Error", f"Error during processing: {error_msg}")
    
    def process_all_pdfs_and_packing_internal(self):
        """Internal method for PDF processing"""
        import re
        
        try:
            output_dir = Path(self.output_dir_edit.text())
            output_dir.mkdir(exist_ok=True)
            
            # Dictionary to store pages grouped by driver
            driver_pages = {}
            processed_files = 0
            total_pages_processed = 0
            
            self.processing_thread.progress_signal.emit("Starting PDF processing...")
            self.processing_thread.progress_signal.emit(f"Processing {len(self.selected_pdf_files)} PDF files...")
            self.processing_thread.progress_signal.emit("Looking for exact order ID matches from Excel data...")
            
            # Debug: Show loaded delivery data
            self.processing_thread.progress_signal.emit(f"Loaded delivery data: {len(self.delivery_data_with_drivers)} orders")
            for order_id, data in list(self.delivery_data_with_drivers.items())[:5]:  # Show first 5
                self.processing_thread.progress_signal.emit(f"  Order '{order_id}' → Driver '{data['driver_number']}'")
            if len(self.delivery_data_with_drivers) > 5:
                self.processing_thread.progress_signal.emit(f"  ... and {len(self.delivery_data_with_drivers) - 5} more orders")
            
            # Process PDF files
            for pdf_file in self.selected_pdf_files:
                self.processing_thread.progress_signal.emit(f"Processing: {Path(pdf_file).name}")
                
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
                        
                        # New approach: Search for exact order ID matches from Excel data
                        # This is much simpler and more reliable than parsing "Our Order No" patterns
                        
                        order_id = None
                        matched_order_id = None
                        
                        # Search for each order ID from Excel directly in the PDF text
                        for excel_order_id in self.delivery_data_with_drivers.keys():
                            # Case-insensitive search for the exact order ID
                            if excel_order_id.upper() in page_text.upper():
                                order_id = excel_order_id  # Use the exact case from Excel
                                matched_order_id = excel_order_id
                                self.processing_thread.progress_signal.emit(
                                    f"✅ Found exact match: '{excel_order_id}' on page {page_num + 1}"
                                )
                                break
                        
                        # If no exact match found, try word boundary search for more precision
                        if not order_id:
                            for excel_order_id in self.delivery_data_with_drivers.keys():
                                # Use word boundaries to avoid partial matches
                                pattern = r'\b' + re.escape(excel_order_id) + r'\b'
                                if re.search(pattern, page_text, re.IGNORECASE):
                                    order_id = excel_order_id
                                    matched_order_id = excel_order_id
                                    self.processing_thread.progress_signal.emit(
                                        f"✅ Found word boundary match: '{excel_order_id}' on page {page_num + 1}"
                                    )
                                    break
                        
                        # Debug: Show what we found on this page
                        if order_id:
                            self.processing_thread.progress_signal.emit(
                                f"Found Order ID '{order_id}' on page {page_num + 1} of {Path(pdf_file).name}"
                            )
                        else:
                            # Debug: Show first 400 characters of page text to help troubleshoot
                            if page_text.strip():
                                preview_text = page_text.replace('\n', ' ').strip()[:400]
                                self.processing_thread.progress_signal.emit(
                                    f"Page {page_num + 1} text preview: {preview_text}..."
                                )
                                # Check if any Order patterns exist in the text
                                has_order_text = any(phrase in page_text.lower() for phrase in ['order no', 'order number', 'our order'])
                                if has_order_text:
                                    self.processing_thread.progress_signal.emit(
                                        f"⚠ Page {page_num + 1} contains 'order' text but no pattern matched"
                                    )
                                    # Show exact text around "order" for debugging
                                    lines = page_text.split('\n')
                                    for i, line in enumerate(lines):
                                        if 'order' in line.lower():
                                            self.processing_thread.progress_signal.emit(
                                                f"  Line {i+1}: '{line.strip()}'"
                                            )
                            else:
                                self.processing_thread.progress_signal.emit(
                                    f"Page {page_num + 1}: No text extracted (may need OCR)"
                                )
                        
                        if order_id:
                            # Find driver for this order (case-insensitive)
                            driver_data = None
                            for stored_order_id, data in self.delivery_data_with_drivers.items():
                                if stored_order_id.upper() == order_id.upper():
                                    driver_data = data
                                    break
                            
                            if driver_data:
                                driver_number = driver_data['driver_number']
                                stop_number = driver_data['stop_number']
                                
                                # Initialize driver group if not exists
                                if driver_number not in driver_pages:
                                    driver_pages[driver_number] = []
                                
                                # Store page info for this driver with stop number for sorting
                                driver_pages[driver_number].append({
                                    'source_pdf_path': pdf_file,
                                    'page_num': page_num,
                                    'order_id': order_id,
                                    'source_file': pdf_file,
                                    'stop_number': stop_number  # Include stop number for sorting
                                })
                                
                                self.processing_thread.progress_signal.emit(
                                    f"✓ Matched Order {order_id} → Driver {driver_number} (Stop {stop_number}, page {page_num + 1}) - INCLUDED"
                                )
                            else:
                                self.processing_thread.progress_signal.emit(
                                    f"⚠ Order {order_id} not found in Excel data (page {page_num + 1}) - SKIPPED"
                                )
                        
                        total_pages_processed += 1
                    
                    processed_files += 1
                    pdf_document.close()
                    
                except Exception as e:
                    self.processing_thread.progress_signal.emit(f"Error processing {pdf_file}: {str(e)}")
                    if 'pdf_document' in locals():
                        pdf_document.close()
                    continue
            
            # Create separate PDF files for each driver
            self.processing_thread.progress_signal.emit("Creating driver-specific PDF files...")
            
            # Show summary of what was found
            total_matched_pages = sum(len(pages) for pages in driver_pages.values())
            self.processing_thread.progress_signal.emit(f"Found {total_matched_pages} pages with matching Order IDs across {len(driver_pages)} drivers")
            
            created_files = []
            failed_files = []
            
            if not driver_pages:
                self.processing_thread.progress_signal.emit("No matching orders found in PDF files!")
                self.processing_thread.progress_signal.emit("Check that your PDF files contain order IDs that match those in your Excel file")
                return {
                    "processed_files": processed_files,
                    "total_pages": total_pages_processed,
                    "driver_files_created": 0,
                    "created_files": [],
                    "failed_files": [],
                    "driver_details": {},
                    "output_dir": str(output_dir),
                    "error": "No matching orders found in PDF files"
                }
            
            for driver_number, pages in driver_pages.items():
                if not pages:
                    continue
                
                try:
                    # Create new PDF for this driver
                    output_filename = f"Driver_{driver_number}_Orders.pdf"
                    output_path = output_dir / output_filename
                    
                    self.processing_thread.progress_signal.emit(
                        f"Creating {output_filename} with {len(pages)} pages..."
                    )
                    
                    new_pdf = fitz.open()
                    pages_added = 0
                    
                    # Add all pages for this driver
                    # Group pages by source file to minimize file opening
                    pages_by_file = {}
                    for page_info in pages:
                        source_file = page_info['source_pdf_path']
                        if source_file not in pages_by_file:
                            pages_by_file[source_file] = []
                        pages_by_file[source_file].append(page_info['page_num'])
                    
                    # Process each source file
                    for source_file, page_numbers in pages_by_file.items():
                        try:
                            source_pdf = fitz.open(source_file)
                            for page_num in page_numbers:
                                # Insert page into new PDF
                                new_pdf.insert_pdf(source_pdf, from_page=page_num, to_page=page_num)
                                pages_added += 1
                            source_pdf.close()
                        except Exception as e:
                            self.processing_thread.progress_signal.emit(
                                f"Error adding pages from {source_file}: {str(e)}"
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
                                f"✓ Successfully created {output_filename} with {pages_added} pages"
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
                    failed_files.append(f"Driver_{driver_number}_Orders.pdf")
                    self.processing_thread.progress_signal.emit(
                        f"✗ Error creating PDF for Driver {driver_number}: {str(e)}"
                    )
                    continue
            
            # All PDFs are already closed during processing
            
            # Final summary message
            self.processing_thread.progress_signal.emit("Processing complete!")
            self.processing_thread.progress_signal.emit(f"Created {len(created_files)} PDF files in {output_dir}")
            
            # Generate summary report
            summary_path = output_dir / "processing_summary.txt"
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write("PDF Processing Summary\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Total PDF files processed: {processed_files}\n")
                f.write(f"Total pages scanned: {total_pages_processed}\n")
                f.write(f"Driver PDF files created: {len(created_files)}\n")
                if failed_files:
                    f.write(f"Failed PDF files: {len(failed_files)}\n")
                f.write("\n")
                
                if created_files:
                    f.write("✓ Successfully Created PDF Files:\n")
                    for filename in created_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if failed_files:
                    f.write("✗ Failed PDF Files:\n")
                    for filename in failed_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                f.write("Driver Page Counts:\n")
                for driver_number, pages in driver_pages.items():
                    f.write(f"  Driver {driver_number}: {len(pages)} pages\n")
                
                f.write("Order ID Matches Found:\n")
                for driver_number, pages in driver_pages.items():
                    f.write(f"  Driver {driver_number}:\n")
                    for page_info in pages:
                        f.write(f"    - Order {page_info['order_id']} (Page {page_info['page_num'] + 1} from {Path(page_info['source_file']).name})\n")
                    f.write("\n")
            
            # Collect driver details for results dialog
            driver_details = {}
            for driver_number, pages in driver_pages.items():
                orders = list(set(page_info['order_id'] for page_info in pages))
                
                driver_details[driver_number] = {
                    'page_count': len(pages),
                    'orders': orders
                }
            
            return {
                "processed_files": processed_files,
                "total_pages": total_pages_processed,
                "driver_files_created": len(created_files),
                "created_files": created_files,
                "failed_files": failed_files,
                "driver_details": driver_details,
                "output_dir": str(output_dir)
            }
            
        except Exception as e:
            self.processing_thread.progress_signal.emit(f"Error: {str(e)}")
            raise e

    def process_picking_dockets(self):
        """Process picking dockets with reversed page order"""
        if not self.delivery_data_values:
            QMessageBox.warning(self, "No Data", "Please load delivery sequence data first.")
            return
        
        output_dir = self.output_dir_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "No Output Directory", "Please select an output directory first.")
            return
        
        if not self.selected_picking_pdf_files:
            QMessageBox.warning(self, "No Picking PDFs", "Please select picking docket PDF files to process.")
            return
        
        self.show_progress(True)
        self.update_status("Processing picking dockets...")
        self.process_picking_btn.setEnabled(False)
        
        # Start background processing
        self.processing_thread = ProcessingThread(self, "process_picking")
        self.processing_thread.progress_signal.connect(self.update_status)
        self.processing_thread.finished_signal.connect(self.on_picking_processing_finished)
        self.processing_thread.start()
    
    def on_picking_processing_finished(self, success, result):
        """Handle picking processing completion"""
        self.show_progress(False)
        self.process_picking_btn.setEnabled(True)
        
        if success:
            self.update_status("Picking dockets processing completed successfully")
            
            # Enable Excel upload functionality
            self.picking_dockets_processed = True
            self.enable_excel_upload()
            
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
        """Internal method for picking dockets processing with reversed page order"""
        import re
        from barcode import Code128
        from barcode.writer import ImageWriter
        import tempfile
        
        try:
            output_dir = Path(self.output_dir_edit.text())
            picking_output_dir = output_dir / "picking_dockets"
            picking_output_dir.mkdir(exist_ok=True)
            
            # Dictionary to store pages grouped by driver
            driver_pages = {}
            processed_files = 0
            total_pages_processed = 0
            
            # Dictionary to store generated barcodes for each order ID
            order_barcodes = {}
            
            self.processing_thread.progress_signal.emit("Starting picking dockets processing...")
            self.processing_thread.progress_signal.emit(f"Processing {len(self.selected_picking_pdf_files)} picking docket PDF files...")
            self.processing_thread.progress_signal.emit("Looking for exact order ID matches from Excel data...")
            
            # Debug: Show loaded delivery data
            self.processing_thread.progress_signal.emit(f"Loaded delivery data: {len(self.delivery_data_with_drivers)} orders")
            for order_id, data in list(self.delivery_data_with_drivers.items())[:5]:  # Show first 5
                self.processing_thread.progress_signal.emit(f"  Order '{order_id}' → Driver '{data['driver_number']}'")
            if len(self.delivery_data_with_drivers) > 5:
                self.processing_thread.progress_signal.emit(f"  ... and {len(self.delivery_data_with_drivers) - 5} more orders")
            
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
                        order_id = None
                        matched_order_id = None
                        
                        # Search for each order ID from Excel directly in the PDF text
                        for excel_order_id in self.delivery_data_with_drivers.keys():
                            # Case-insensitive search for the exact order ID
                            if excel_order_id.upper() in page_text.upper():
                                order_id = excel_order_id  # Use the exact case from Excel
                                matched_order_id = excel_order_id
                                self.processing_thread.progress_signal.emit(
                                    f"✅ Found exact match: '{excel_order_id}' on page {page_num + 1}"
                                )
                                break
                        
                        # If no exact match found, try word boundary search for more precision
                        if not order_id:
                            for excel_order_id in self.delivery_data_with_drivers.keys():
                                # Use word boundaries to avoid partial matches
                                pattern = r'\b' + re.escape(excel_order_id) + r'\b'
                                if re.search(pattern, page_text, re.IGNORECASE):
                                    order_id = excel_order_id
                                    matched_order_id = excel_order_id
                                    self.processing_thread.progress_signal.emit(
                                        f"✅ Found word boundary match: '{excel_order_id}' on page {page_num + 1}"
                                    )
                                    break
                        
                        # Debug: Show what we found on this page
                        if order_id:
                            self.processing_thread.progress_signal.emit(
                                f"Found Order ID '{order_id}' on page {page_num + 1} of {Path(pdf_file).name}"
                            )
                        else:
                            # Debug: Show first 400 characters of page text to help troubleshoot
                            if page_text.strip():
                                preview_text = page_text.replace('\n', ' ').strip()[:400]
                                self.processing_thread.progress_signal.emit(
                                    f"Page {page_num + 1} text preview: {preview_text}..."
                                )
                        
                        if order_id:
                            # Find driver for this order (case-insensitive)
                            driver_data = None
                            for stored_order_id, data in self.delivery_data_with_drivers.items():
                                if stored_order_id.upper() == order_id.upper():
                                    driver_data = data
                                    break
                            
                            if driver_data:
                                driver_number = driver_data['driver_number']
                                stop_number = driver_data['stop_number']
                                
                                # Initialize driver group if not exists
                                if driver_number not in driver_pages:
                                    driver_pages[driver_number] = []
                                
                                # Store page info for this driver with stop number for sorting
                                driver_pages[driver_number].append({
                                    'source_pdf_path': pdf_file,
                                    'page_num': page_num,
                                    'order_id': order_id,
                                    'source_file': pdf_file,
                                    'stop_number': stop_number  # Include stop number for sorting
                                })
                                
                                self.processing_thread.progress_signal.emit(
                                    f"✓ Matched Order {order_id} → Driver {driver_number} (Stop {stop_number}, page {page_num + 1}) - INCLUDED"
                                )
                            else:
                                self.processing_thread.progress_signal.emit(
                                    f"⚠ Order {order_id} not found in Excel data (page {page_num + 1}) - SKIPPED"
                                )
                        
                        total_pages_processed += 1
                    
                    processed_files += 1
                    pdf_document.close()
                    
                except Exception as e:
                    self.processing_thread.progress_signal.emit(f"Error processing {pdf_file}: {str(e)}")
                    if 'pdf_document' in locals():
                        pdf_document.close()
                    continue
            
            # Generate barcodes for all unique order IDs found
            self.processing_thread.progress_signal.emit("Generating barcodes for unique order IDs...")
            unique_order_ids = set()
            for driver_number, pages in driver_pages.items():
                for page_info in pages:
                    unique_order_ids.add(page_info['order_id'])
            
            # Create barcodes for each unique order ID
            for order_id in unique_order_ids:
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
            
            # Save generated barcodes to Supabase
            try:
                self.processing_thread.progress_signal.emit("Saving barcodes to Supabase database...")
                
                # Prepare barcode data for Supabase
                barcode_data_for_db = []
                for driver_number, pages in driver_pages.items():
                    for page_info in pages:
                        order_id = page_info['order_id']
                        if order_id in order_barcodes:
                            barcode_record = {
                                'order_id': order_id,
                                'driver_number': driver_number,
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
            
            # Create separate PDF files for each driver with REVERSED page order
            self.processing_thread.progress_signal.emit("Creating driver-specific picking docket PDF files with reversed page order...")
            
            # Show summary of what was found
            total_matched_pages = sum(len(pages) for pages in driver_pages.values())
            self.processing_thread.progress_signal.emit(f"Found {total_matched_pages} picking docket pages with matching Order IDs across {len(driver_pages)} drivers")
            self.processing_thread.progress_signal.emit("📋 Only including pages with order IDs present in Excel file - other pages are filtered out")
            
            created_files = []
            failed_files = []
            
            if not driver_pages:
                self.processing_thread.progress_signal.emit("No matching orders found in picking docket PDF files!")
                self.processing_thread.progress_signal.emit("Check that your picking docket PDF files contain order IDs that match those in your Excel file")
                return {
                    "processed_files": processed_files,
                    "total_pages": total_pages_processed,
                    "driver_files_created": 0,
                    "created_files": [],
                    "failed_files": [],
                    "driver_details": {},
                    "output_dir": str(picking_output_dir),
                    "error": "No matching orders found in picking docket PDF files"
                }
            
            for driver_number, pages in driver_pages.items():
                if not pages:
                    continue
                
                try:
                    # Create new PDF for this driver
                    output_filename = f"Driver_{driver_number}_Picking_Dockets.pdf"
                    output_path = picking_output_dir / output_filename
                    
                    # Sort pages by stop number, then REVERSE the order for picking
                    # This ensures that the first stops in delivery sequence are at the top of the pallet
                    pages_with_stop_numbers = []
                    for page_info in pages:
                        stop_number = page_info.get('stop_number', '0')
                        try:
                            # Try to convert to int for proper numeric sorting
                            sort_key = int(stop_number) if stop_number.isdigit() else 999999
                        except:
                            sort_key = 999999
                        pages_with_stop_numbers.append((sort_key, page_info))
                    
                    # Sort by stop number (ascending), then reverse for picking
                    pages_with_stop_numbers.sort(key=lambda x: x[0])
                    sorted_pages = [page_info for sort_key, page_info in pages_with_stop_numbers]
                    
                    # REVERSE the order for picking (first delivery stops at top of pallet)
                    reversed_pages = sorted_pages[::-1]
                    
                    self.processing_thread.progress_signal.emit(
                        f"Creating {output_filename} with {len(reversed_pages)} pages in REVERSED order..."
                    )
                    self.processing_thread.progress_signal.emit(
                        f"  First page will be: Order {reversed_pages[0]['order_id']} (Stop {reversed_pages[0]['stop_number']})"
                    )
                    self.processing_thread.progress_signal.emit(
                        f"  Last page will be: Order {reversed_pages[-1]['order_id']} (Stop {reversed_pages[-1]['stop_number']})"
                    )
                    
                    new_pdf = fitz.open()
                    pages_added = 0
                    
                    # Add all pages for this driver in reversed order with barcodes
                    for page_info in reversed_pages:
                        try:
                            # Open source PDF and get the page
                            source_pdf = fitz.open(page_info['source_pdf_path'])
                            source_page = source_pdf[page_info['page_num']]
                            
                            # Create a new page in the output PDF
                            new_page = new_pdf.new_page(width=source_page.rect.width, height=source_page.rect.height)
                            
                            # Copy the original page content
                            new_page.show_pdf_page(new_page.rect, source_pdf, page_info['page_num'])
                            
                            # Add barcode at the top center of the page
                            order_id = page_info['order_id']
                            if order_id in order_barcodes:
                                try:
                                    # Insert barcode image at the top center
                                    barcode_data = order_barcodes[order_id]
                                    
                                    # Calculate position for top center
                                    page_width = new_page.rect.width
                                    barcode_width = 700  # Even longer barcode
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
                                f"✓ Successfully created {output_filename} with {pages_added} pages in REVERSED order with barcodes"
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
                    failed_files.append(f"Driver_{driver_number}_Picking_Dockets.pdf")
                    self.processing_thread.progress_signal.emit(
                        f"✗ Error creating picking docket PDF for Driver {driver_number}: {str(e)}"
                    )
                    continue
            
            # Final summary message
            self.processing_thread.progress_signal.emit("Picking dockets processing complete!")
            self.processing_thread.progress_signal.emit(f"Created {len(created_files)} picking docket PDF files in {picking_output_dir}")
            self.processing_thread.progress_signal.emit("📝 Pages are in REVERSED order - first delivery stops are at the top!")
            self.processing_thread.progress_signal.emit(f"🏷️  Generated barcodes for {len(unique_order_ids)} unique order IDs")
            self.processing_thread.progress_signal.emit("📋 Only pages with order IDs matching Excel data were included - others were filtered out")
            
            # Generate summary report
            summary_path = picking_output_dir / "picking_dockets_summary.txt"
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write("Picking Dockets Processing Summary\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Total picking docket PDF files processed: {processed_files}\n")
                f.write(f"Total pages scanned: {total_pages_processed}\n")
                f.write(f"Driver picking docket PDF files created: {len(created_files)}\n")
                f.write(f"Unique order IDs with barcodes: {len(unique_order_ids)}\n")
                if failed_files:
                    f.write(f"Failed PDF files: {len(failed_files)}\n")
                f.write("\n")
                f.write("IMPORTANT: Pages are in REVERSED order for picking!\n")
                f.write("First delivery stops are at the top of each pallet.\n")
                f.write("Each page has a barcode at the top center for the order ID.\n\n")
                
                if created_files:
                    f.write("✓ Successfully Created Picking Docket PDF Files:\n")
                    for filename in created_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if failed_files:
                    f.write("✗ Failed PDF Files:\n")
                    for filename in failed_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                f.write("Driver Page Counts (in reversed order):\n")
                for driver_number, pages in driver_pages.items():
                    f.write(f"  Driver {driver_number}: {len(pages)} pages\n")
                
                f.write("Order ID Matches Found (in reversed order):\n")
                for driver_number, pages in driver_pages.items():
                    f.write(f"  Driver {driver_number}:\n")
                    # Sort and reverse for the summary too
                    pages_with_stop_numbers = []
                    for page_info in pages:
                        stop_number = page_info.get('stop_number', '0')
                        try:
                            sort_key = int(stop_number) if stop_number.isdigit() else 999999
                        except:
                            sort_key = 999999
                        pages_with_stop_numbers.append((sort_key, page_info))
                    
                    pages_with_stop_numbers.sort(key=lambda x: x[0])
                    sorted_pages = [page_info for sort_key, page_info in pages_with_stop_numbers]
                    reversed_pages = sorted_pages[::-1]
                    
                    for page_info in reversed_pages:
                        f.write(f"    - Order {page_info['order_id']} (Stop {page_info['stop_number']}) - Page {page_info['page_num'] + 1} from {Path(page_info['source_file']).name}\n")
                    f.write("\n")
                
                f.write("Barcodes Generated:\n")
                for order_id in sorted(unique_order_ids):
                    f.write(f"  - {order_id}\n")
            
            # Collect driver details for results dialog
            driver_details = {}
            for driver_number, pages in driver_pages.items():
                orders = list(set(page_info['order_id'] for page_info in pages))
                
                driver_details[driver_number] = {
                    'page_count': len(pages),
                    'orders': orders
                }
            
            return {
                "processed_files": processed_files,
                "total_pages": total_pages_processed,
                "driver_files_created": len(created_files),
                "created_files": created_files,
                "failed_files": failed_files,
                "driver_details": driver_details,
                "output_dir": str(picking_output_dir),
                "barcodes_generated": len(unique_order_ids)
            }
            
        except Exception as e:
            self.processing_thread.progress_signal.emit(f"Error: {str(e)}")
            raise e


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Around Noon Dispatch")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("Transport Solutions")
    
    # Apply modern style
    app.setStyle('Fusion')
    
    # Create and show the main window
    window = TransportSorterApp()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()