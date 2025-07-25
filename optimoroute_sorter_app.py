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
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize, QDate, QPropertyAnimation, QEasingCurve, QRect
from PySide6.QtGui import QFont, QPalette, QColor, QIcon, QPixmap, QPainter, QRegion


class ProcessingThread(QThread):
    """Background thread for PDF processing operations"""
    progress_signal = Signal(str)
    finished_signal = Signal(bool, dict)
    
    def __init__(self, app_instance):
        super().__init__()
        self.app = app_instance
    
    def run(self):
        try:
            result = self.app.process_all_pdfs_and_packing_internal()
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
        elif results.get('missing_order_ids'):
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
        elif results.get('missing_order_ids'):
            title_text = "PDF Processing Completed - Some Orders Missing"
        
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
        
        # Add order matching statistics if available
        if results.get('total_order_ids'):
            total_orders = results.get('total_order_ids', 0)
            found_orders = len(results.get('found_order_ids', []))
            missing_orders = len(results.get('missing_order_ids', []))
            
            stats_data.extend([
                ("Total Orders in Data", str(total_orders)),
                ("Orders Found in PDFs", str(found_orders)),
                ("Orders Missing from PDFs", str(missing_orders))
            ])
        
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
        
        # Tab 3: Missing Orders (if any)
        if results.get('missing_order_ids'):
            missing_tab = self.create_missing_tab(results.get('missing_order_ids', []), results.get('delivery_data_with_drivers', {}))
            tab_widget.addTab(missing_tab, "Missing Orders")
        
        # Tab 4: Failed Files (if any)
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
    
    def create_missing_tab(self, missing_order_ids, delivery_data_with_drivers):
        """Create tab showing missing order IDs"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Table
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Order ID", "Driver", "Stop Number", "Status"])
        table.setRowCount(len(missing_order_ids))
        
        for i, order_id in enumerate(sorted(missing_order_ids)):
            # Order ID
            order_item = QTableWidgetItem(str(order_id))
            order_item.setFlags(order_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 0, order_item)
            
            # Driver
            driver_data = delivery_data_with_drivers.get(order_id, {})
            driver_name = driver_data.get('driver_number', 'Unknown')
            driver_item = QTableWidgetItem(str(driver_name))
            driver_item.setFlags(driver_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 1, driver_item)
            
            # Stop Number
            stop_number = driver_data.get('stop_number', 'Unknown')
            stop_item = QTableWidgetItem(str(stop_number))
            stop_item.setFlags(stop_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 2, stop_item)
            
            # Status
            status_item = QTableWidgetItem("❌ Missing")
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, 3, status_item)
        
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
        message += "1. Check that your PDF files contain order IDs that match those in your delivery data\n"
        message += "2. Ensure order IDs in PDF match those in your delivery data exactly\n"
        message += "3. Order ID matching is case-insensitive (AA061B4Y = aa061b4y)\n\n"
        message += f"Processing Summary:\n"
        message += f"- PDF files processed: {results.get('processed_files', 0)}\n"
        message += f"- Total pages scanned: {results.get('total_pages', 0)}\n"
        message += f"- No matching order IDs found\n\n"
        message += "Common Issues:\n"
        message += "- Order IDs in PDF don't match those in delivery data\n"
        message += "- PDF contains images that need OCR processing\n"
        message += "- Order IDs in delivery data don't match those in PDF\n"
        message += "- Text extraction failed from PDF pages\n\n"
        message += "Check the main application status messages for more detailed information about what was found on each page."
        
        diagnostic_text.setPlainText(message)
        layout.addWidget(diagnostic_text)
        
        return widget
    
    def open_output_folder(self):
        """Open output folder"""
        if hasattr(self.parent(), 'open_output_directory'):
            self.parent().open_output_directory(self.results.get('output_dir', ''))
    
    def export_missing_orders(self, missing_order_ids, delivery_data_with_drivers):
        """Export missing orders to a text file"""
        try:
            from PySide6.QtWidgets import QFileDialog
            from pathlib import Path
            
            # Get save location
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Missing Orders",
                str(Path.home() / "missing_orders.txt"),
                "Text files (*.txt);;All files (*.*)"
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("Missing Orders Report\n")
                    f.write("=" * 50 + "\n\n")
                    f.write(f"Total missing orders: {len(missing_order_ids)}\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                    f.write("Missing Order Details:\n")
                    f.write("-" * 30 + "\n")
                    
                    for order_id in sorted(missing_order_ids):
                        driver_data = delivery_data_with_drivers.get(order_id, {})
                        driver_name = driver_data.get('driver_number', 'Unknown')
                        stop_number = driver_data.get('stop_number', 'Unknown')
                        f.write(f"Order ID: {order_id}\n")
                        f.write(f"  Driver: {driver_name}\n")
                        f.write(f"  Stop Number: {stop_number}\n")
                        f.write("\n")
                
                QMessageBox.information(
                    self,
                    "Export Successful",
                    f"Missing orders list exported to:\n{file_path}"
                )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Export Error",
                f"Failed to export missing orders: {str(e)}"
            )
    
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


class OptimoRouteApiThread(QThread):
    """Background thread for OptimoRoute API operations"""
    progress_signal = Signal(str)
    finished_signal = Signal(bool, list)
    
    def __init__(self, api_key, from_date=None, to_date=None, driver_filter=None):
        super().__init__()
        self.api_key = api_key
        self.base_url = "https://api.optimoroute.com/v1"
        self.from_date = from_date
        self.to_date = to_date
        self.driver_filter = driver_filter
    
    def run(self):
        try:
            self.progress_signal.emit("Connecting to OptimoRoute API...")
            
            headers = {
                'Content-Type': 'application/json'
            }
            
            orders = []
            
            # Use custom dates if provided, otherwise default to last 7 days
            if self.from_date and self.to_date:
                from_date = self.from_date
                to_date = self.to_date
                self.progress_signal.emit(f"Searching for orders from {from_date} to {to_date}...")
            else:
                from_date = (datetime.now() - timedelta(days=6)).strftime('%Y-%m-%d')
                to_date = datetime.now().strftime('%Y-%m-%d')
                self.progress_signal.emit("Searching for orders in the last 7 days...")
            
            request_body = {
                "dateRange": {
                    "from": from_date,
                    "to": to_date
                },
                "includeOrderData": True,
                "includeScheduleInformation": True
            }
            
            # Add driver filter if specified
            if self.driver_filter and self.driver_filter.strip() and self.driver_filter != "All Drivers":
                request_body["driverName"] = self.driver_filter.strip()
                self.progress_signal.emit(f"Filtering by driver: {self.driver_filter.strip()}")
            else:
                self.progress_signal.emit("")
            
            url = f"{self.base_url}/search_orders"
            params = {'key': self.api_key}
            
            # Handle pagination with after_tag
            after_tag = None
            page_count = 0
            max_pages = 10  # Safety limit to prevent infinite loops
            
            while page_count < max_pages:
                if after_tag:
                    request_body["after_tag"] = after_tag
                    self.progress_signal.emit(f"Fetching page {page_count + 1} of orders...")
                
                try:
                    response = requests.post(
                        url, 
                        headers=headers, 
                        params=params,
                        json=request_body,
                        timeout=15
                    )
                    
                    if response.status_code == 200:
                        data = response.json()
                        
                        if data.get('success') and data.get('orders'):
                            self.progress_signal.emit(f"Processing {len(data['orders'])} orders from page {page_count + 1}...")
                            
                            # Process each order
                            for order_item in data['orders']:
                                order_data = order_item.get('data', {})
                                schedule_info = order_item.get('scheduleInformation', {})
                                
                                if order_data:  # Only process if we have order data
                                    # Build comprehensive order data
                                    processed_order = {
                                        'id': order_data.get('id', ''),
                                        'orderNo': order_data.get('orderNo', ''),
                                        'date': order_data.get('date', ''),
                                        'address': order_data.get('location', {}).get('address', ''),
                                        'locationName': order_data.get('location', {}).get('locationName', ''),
                                        'latitude': order_data.get('location', {}).get('latitude', ''),
                                        'longitude': order_data.get('location', {}).get('longitude', ''),
                                        'duration': order_data.get('duration', 0),
                                        'priority': order_data.get('priority', ''),
                                        'type': order_data.get('type', ''),
                                        'load1': order_data.get('load1', 0),
                                        'load2': order_data.get('load2', 0),
                                        'load3': order_data.get('load3', 0),
                                        'load4': order_data.get('load4', 0),
                                        'timeWindows': order_data.get('timeWindows', []),
                                        'skills': order_data.get('skills', []),
                                        'vehicleFeatures': order_data.get('vehicleFeatures', []),
                                        'notes': order_data.get('notes', ''),
                                        'phone': order_data.get('phone', ''),
                                        'email': order_data.get('email', ''),
                                        'customField1': order_data.get('customField1', ''),
                                        'customField2': order_data.get('customField2', ''),
                                        'customField3': order_data.get('customField3', ''),
                                        'customField4': order_data.get('customField4', ''),
                                        'customField5': order_data.get('customField5', ''),
                                        'allowedWeekdays': order_data.get('allowedWeekdays', []),
                                        'notificationPreference': order_data.get('notificationPreference', ''),
                                        'assignedTo': order_data.get('assignedTo'),
                                        # Schedule information from includeScheduleInformation
                                        'driverName': schedule_info.get('driverName', '') if schedule_info else '',
                                        'driverExternalId': schedule_info.get('driverExternalId', '') if schedule_info else '',
                                        'vehicleLabel': schedule_info.get('vehicleLabel', '') if schedule_info else '',
                                        'vehicleRegistration': schedule_info.get('vehicleRegistration', '') if schedule_info else '',
                                        'scheduledAt': schedule_info.get('scheduledAt', '') if schedule_info else '',
                                        'scheduledAtDt': schedule_info.get('scheduledAtDt', '') if schedule_info else '',
                                        'arrivalTimeDt': schedule_info.get('arrivalTimeDt', '') if schedule_info else '',
                                        'stopNumber': schedule_info.get('stopNumber', '') if schedule_info else '',
                                        'travelTime': schedule_info.get('travelTime', 0) if schedule_info else 0,
                                        'distance': schedule_info.get('distance', 0) if schedule_info else 0,
                                        'status': 'scheduled' if schedule_info else 'unscheduled'
                                    }
                                    orders.append(processed_order)
                            
                            # Check for pagination
                            after_tag = data.get('after_tag')
                            if not after_tag:
                                break  # No more pages
                            
                            page_count += 1
                        else:
                            self.progress_signal.emit("No orders found in response")
                            break
                            
                    elif response.status_code == 401:
                        self.progress_signal.emit("Authentication failed - please check your API key")
                        self.finished_signal.emit(False, [])
                        return
                    else:
                        self.progress_signal.emit(f"API returned status {response.status_code}: {response.text}")
                        break
                        
                except requests.exceptions.RequestException as e:
                    self.progress_signal.emit(f"Network error: {str(e)}")
                    break
            
            if not orders:
                self.progress_signal.emit("No scheduled orders found for the selected date")
                self.finished_signal.emit(True, [])  # Return empty list instead of sample data
                return
            
            self.progress_signal.emit(f"")
            self.finished_signal.emit(True, orders)
            
        except Exception as e:
            self.progress_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(False, [])


class SettingsDialog(QDialog):
    """Settings dialog for configuring API key"""
    
    def __init__(self, current_api_key="", parent=None):
        super().__init__(parent)
        self.current_api_key = current_api_key
        self.setWindowTitle("Settings")
        self.setModal(True)
        self.resize(500, 300)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Title
        title_label = QLabel("OptimoRoute API Configuration")
        title_label.setObjectName("settingsTitle")
        layout.addWidget(title_label)
        
        # API Key section
        api_group = QGroupBox("API Key")
        api_layout = QVBoxLayout(api_group)
        
        api_info = QLabel(
            "Enter your OptimoRoute API key. You can find this in your OptimoRoute account settings."
        )
        api_info.setWordWrap(True)
        api_info.setObjectName("infoText")
        api_layout.addWidget(api_info)
        
        # API Key input
        api_layout.addWidget(QLabel("API Key:"))
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setPlaceholderText("Enter your OptimoRoute API key...")
        self.api_key_edit.setText(current_api_key)
        self.api_key_edit.setEchoMode(QLineEdit.Password)  # Hide the API key
        api_layout.addWidget(self.api_key_edit)
        
        # Show/Hide API key toggle
        show_key_layout = QHBoxLayout()
        self.show_key_checkbox = QCheckBox("Show API Key")
        self.show_key_checkbox.toggled.connect(self.toggle_api_key_visibility)
        show_key_layout.addWidget(self.show_key_checkbox)
        show_key_layout.addStretch()
        api_layout.addLayout(show_key_layout)
        
        layout.addWidget(api_group)
        
        # Help section
        help_group = QGroupBox("How to get your API Key")
        help_layout = QVBoxLayout(help_group)
        
        help_text = QLabel(
            "1. Log in to your OptimoRoute account\n"
            "2. Go to Settings → API\n"
            "3. Generate a new API key or copy an existing one\n"
            "4. Paste the key in the field above\n\n"
            "Note: Keep your API key secure and don't share it with others."
        )
        help_text.setWordWrap(True)
        help_text.setObjectName("helpText")
        help_layout.addWidget(help_text)
        
        layout.addWidget(help_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        clear_btn = QPushButton("Clear API Key")
        clear_btn.setObjectName("dangerButton")
        clear_btn.clicked.connect(self.clear_api_key)
        
        save_btn = QPushButton("Save")
        save_btn.setObjectName("primaryButton")
        save_btn.clicked.connect(self.accept)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(clear_btn)
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(save_btn)
        
        layout.addLayout(button_layout)
        
        # Apply styling
        self.apply_settings_styling()
    
    def toggle_api_key_visibility(self, show):
        """Toggle API key visibility"""
        if show:
            self.api_key_edit.setEchoMode(QLineEdit.Normal)
        else:
            self.api_key_edit.setEchoMode(QLineEdit.Password)
    
    def test_api_connection(self):
        """Test the API connection with the provided key"""
        api_key = self.api_key_edit.text().strip()
        if not api_key:
            QMessageBox.warning(self, "No API Key", "Please enter an API key to test.")
            return
        
        # Show testing message
        QMessageBox.information(self, "Testing Connection", 
                               "Testing API connection... This may take a few seconds.")
        
        # Test the connection in a background thread
        test_thread = ApiTestThread(api_key)
        test_thread.finished_signal.connect(self.on_test_finished)
        test_thread.start()
    
    def on_test_finished(self, success, message):
        """Handle API test completion"""
        if success:
            QMessageBox.information(self, "Connection Successful", 
                                   f"API connection test successful!\n\n{message}")
        else:
            QMessageBox.critical(self, "Connection Failed", 
                                f"API connection test failed:\n\n{message}")
    
    def clear_api_key(self):
        """Clear the API key field"""
        reply = QMessageBox.question(
            self,
            "Clear API Key",
            "Are you sure you want to clear the API key?\n\n"
            "This will require you to enter a new API key before using the application.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.api_key_edit.clear()
    
    def get_api_key(self):
        """Get the entered API key"""
        return self.api_key_edit.text().strip()
    
    def apply_settings_styling(self):
        """Apply styling to settings dialog"""
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
            }
            
            QLabel#settingsTitle {
                font-size: 18px;
                font-weight: bold;
                color: #1e293b;
                margin-bottom: 10px;
            }
            
            QGroupBox {
                font-weight: bold;
                color: #374151;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                background-color: #f8fafc;
            }
            
            QLabel#infoText, QLabel#helpText {
                color: #64748b;
                font-size: 12px;
                padding: 8px;
                background-color: #f1f5f9;
                border-radius: 4px;
                font-weight: normal;
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
            
            QCheckBox {
                color: #374151;
                font-size: 13px;
            }
            
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #d1d5db;
                border-radius: 3px;
                background-color: white;
            }
            
            QCheckBox::indicator:checked {
                background-color: #2563eb;
                border-color: #2563eb;
            }
            
            QCheckBox::indicator:checked::after {
                content: "✓";
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            
            QPushButton#dangerButton {
                background-color: #dc2626;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                min-height: 20px;
            }
            
            QPushButton#dangerButton:hover {
                background-color: #b91c1c;
            }
        """)


class ApiTestThread(QThread):
    """Background thread for testing API connection"""
    finished_signal = Signal(bool, str)
    
    def __init__(self, api_key):
        super().__init__()
        self.api_key = api_key
    
    def run(self):
        try:
            headers = {'Content-Type': 'application/json'}
            url = "https://api.optimoroute.com/v1/search_orders"
            params = {'key': self.api_key}
            
            # Simple test request
            request_body = {
                "dateRange": {
                    "from": "2024-01-01",
                    "to": "2024-01-01"
                },
                "includeOrderData": False,
                "includeScheduleInformation": False
            }
            
            response = requests.post(
                url, 
                headers=headers, 
                params=params,
                json=request_body,
                timeout=10
            )
            
            if response.status_code == 200:
                self.finished_signal.emit(True, "API key is valid and connection successful.")
            elif response.status_code == 401:
                self.finished_signal.emit(False, "Invalid API key. Please check your key and try again.")
            else:
                self.finished_signal.emit(False, f"API returned status {response.status_code}: {response.text}")
                
        except requests.exceptions.RequestException as e:
            self.finished_signal.emit(False, f"Network error: {str(e)}")
        except Exception as e:
            self.finished_signal.emit(False, f"Error: {str(e)}")


class OptimoRouteSorterApp(QMainWindow):
    """OptimoRoute Sorter Application for Delivery Processing"""
    
    def __init__(self):
        super().__init__()
        
        # Set application icon
        try:
            icon_path = "application_icon.ico" 
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except Exception as e:
            print(f"Could not load application icon: {e}")
        
        # Application data
        self.delivery_data_values = []
        self.delivery_data_with_drivers = {}
        self.delivery_json_file = "delivery_sequence_data.json"
        self.selected_pdf_files = []
        self.processed_drivers = {}
        self.processing_thread = None
        
        # OptimoRoute API setup
        self.api_key = self.load_api_key()
        self.optimoroute_thread = None
        self.scheduled_orders_data = []
        
        # Check if API key is configured before proceeding
        if not self.api_key:
            # Initialize UI first for the API key screen
            self.init_ui()
            self.apply_clean_styling()
            self.show_api_key_screen()
            return  # Exit initialization if no API key provided
        
        # Auto-refresh timer setup
        self.auto_refresh_timer = QTimer()
        self.auto_refresh_timer.timeout.connect(self.auto_refresh_data)
        self.auto_refresh_enabled = True
        self.auto_refresh_timer.start(2000)  # 5 seconds
        
        # Data change tracking
        self.last_data_hash = None
        self.last_order_count = 0
        
        # Initialize UI
        self.init_ui()
        self.apply_clean_styling()
        
        # Load saved output directory
        saved_output_dir = self.load_output_directory()
        if saved_output_dir:
            self.output_dir_edit.setText(saved_output_dir)
            self.update_output_button_text()
        
        # Load existing data
        self.load_existing_delivery_data()
        
        # Initialize date range display
        self.on_date_changed()
        
        # Set initial status if no existing data
        if not self.delivery_data_values:
            self.update_status("Ready - Auto-refresh enabled, checking for changes")
        
        # Setup window animation
        self.setup_window_animation()
    
    def setup_window_animation(self):
        """Setup the window opening animation - reveal content from top to bottom"""
        # Set initial opacity to 0 (invisible)
        self.setWindowOpacity(0.0)
        
        # Get the current window geometry
        current_geometry = self.geometry()
        final_width = current_geometry.width()
        final_height = current_geometry.height()
        
        # Set window to final position and size
        self.resize(final_width, final_height)
        
        # Initialize reveal progress
        self.reveal_progress = 0.0
        
        # Create fade-in animation
        self.fade_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_animation.setDuration(800)  # 800ms duration
        self.fade_animation.setStartValue(0.0)
        self.fade_animation.setEndValue(1.0)
        self.fade_animation.setEasingCurve(QEasingCurve.OutCubic)  # Smooth easing
        
        # Create reveal timer for progressive reveal
        self.reveal_timer = QTimer()
        self.reveal_timer.timeout.connect(self.update_reveal_progress)
        self.reveal_steps = 40  # Number of steps for smooth reveal
        self.reveal_step = 0
        
        # Start animations after a short delay
        QTimer.singleShot(100, lambda: self.fade_animation.start())
        QTimer.singleShot(100, lambda: self.reveal_timer.start(20))  # 20ms intervals
    
    def update_reveal_progress(self):
        """Update the reveal progress using timer-based animation"""
        self.reveal_step += 1
        
        # Calculate progress using easing curve (OutCubic)
        progress = self.reveal_step / self.reveal_steps
        if progress > 1.0:
            progress = 1.0
        
        # Apply easing curve (OutCubic: t^3)
        eased_progress = 1 - (1 - progress) ** 3
        
        self.reveal_progress = eased_progress
        self.update_reveal_mask()
        
        # Stop timer when complete
        if self.reveal_step >= self.reveal_steps:
            self.reveal_timer.stop()
            # Remove mask when animation is complete
            self.setMask(QRegion())
    
    def update_reveal_mask(self):
        """Update the window mask to reveal content from top to bottom"""
        if not hasattr(self, 'reveal_progress'):
            return
            
        progress = self.reveal_progress
        window_height = self.height()
        window_width = self.width()
        
        # Calculate how much of the window should be visible
        visible_height = int(window_height * progress)
        
        if visible_height > 0:
            # Create a region that shows only the top portion of the window
            reveal_region = QRegion(0, 0, window_width, visible_height)
            self.setMask(reveal_region)
        else:
            # Hide the entire window
            self.setMask(QRegion())
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("OptimoRoute Sorter - Delivery Processing")
        self.setGeometry(100, 100, 1400, 900)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(0, 0, 0, 20)
        
        # Header
        header_frame = self.create_header()
        main_layout.addWidget(header_frame)
        
        # Content area - 3 column grid
        content_widget = QWidget()
        content_layout = QGridLayout(content_widget)
        content_layout.setSpacing(10)
        
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
        
        main_layout.addWidget(content_widget)
        
        # Process button
        self.process_all_btn = QPushButton("Process Delivery PDFs")
        self.process_all_btn.setObjectName("primaryButton")
        self.process_all_btn.clicked.connect(self.process_all_pdfs_and_packing)
        self.process_all_btn.setFixedHeight(50)
        main_layout.addWidget(self.process_all_btn)
        
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
        header_widget = QWidget()
        header_widget.setObjectName("headerWidget")
        header_widget.setFixedHeight(80)
        
        layout = QHBoxLayout(header_widget)
        layout.setContentsMargins(0, 15, 0, 15)
        
        # Add left margin for refresh button
        layout.addSpacing(20)
        
        # Refresh button
        self.refresh_btn = QPushButton("🔄")
        self.refresh_btn.setObjectName("refreshButton")
        self.refresh_btn.clicked.connect(self.refresh_data)
        self.refresh_btn.setFixedSize(40, 40)
        self.refresh_btn.setToolTip("Manual refresh - fetch latest data from OptimoRoute")
        layout.addWidget(self.refresh_btn)
        
        # Settings button
        settings_btn = QPushButton("⚙")
        settings_btn.setObjectName("settingsButton")
        settings_btn.setToolTip("Settings - Configure API Key")
        settings_btn.clicked.connect(self.open_settings)
        settings_btn.setFixedSize(40, 40)
        layout.addWidget(settings_btn)
        
        # Spacer between buttons and title
        layout.addSpacing(15)
        
        title_label = QLabel("OptimoRoute Sorter")
        title_label.setObjectName("headerTitle")
        layout.addWidget(title_label)
        
        layout.addStretch()
        
        subtitle_label = QLabel("Delivery Processing & PDF Sorting")
        subtitle_label.setObjectName("headerSubtitle")
        layout.addWidget(subtitle_label)
        
        # Add right margin for subtitle
        layout.addSpacing(20)
        
        return header_widget
    
    def create_setup_section(self):
        """Create setup section"""
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
        self.output_dir_edit.textChanged.connect(self.update_output_button_text)
        layout.addWidget(self.output_dir_edit)
        
        self.output_btn = QPushButton("Browse")
        self.output_btn.clicked.connect(self.browse_output_directory)
        layout.addWidget(self.output_btn)
        
        # Spacer
        layout.addSpacing(15)
        
        # Date selection for scheduled orders
        layout.addWidget(QLabel("Select Date for Scheduled Orders:"))
        
        # Single date picker
        self.fetch_date = QDateEdit()
        self.fetch_date.setDate(QDate.currentDate())  # Default to today
        self.fetch_date.setCalendarPopup(True)
        self.fetch_date.setDisplayFormat("yyyy-MM-dd")
        self.fetch_date.dateChanged.connect(self.on_date_changed)
        layout.addWidget(self.fetch_date)
        
        # Spacer
        layout.addSpacing(10)
        
        # Auto-refresh status (no manual fetch button needed)
        layout.addStretch()
        
        layout.addStretch()
        
        return section
    
    def create_data_section(self):
        """Create data preview section"""
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
        
        # Set column stretch factors to ensure proper distribution
        header = self.data_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # # column - fixed width
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Order ID - stretch
        header.setSectionResizeMode(2, QHeaderView.Fixed)  # Stop Number - fixed width
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Driver - stretch
        
        # Set initial column widths
        self.data_table.setColumnWidth(0, 50)   # # column
        self.data_table.setColumnWidth(2, 100)  # Stop Number column
        
        layout.addWidget(self.data_table)
        
        return section
    
    def create_process_section(self):
        """Create processing section"""
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
            self.save_output_directory(directory)
            self.update_output_button_text()
    
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
            # Silently fail without showing error message
            pass
    
    def save_output_directory(self, directory):
        """Save output directory to configuration file"""
        try:
            config_file = "api_config.json"
            config = {}
            
            # Load existing config if it exists
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            
            # Update with new output directory
            config['output_directory'] = directory
            
            # Save back to file
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Error saving output directory: {e}")
    
    def load_output_directory(self):
        """Load output directory from configuration file"""
        try:
            config_file = "api_config.json"
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('output_directory', '')
            return ''
        except Exception as e:
            print(f"Error loading output directory: {e}")
            return ''
    
    def update_output_button_text(self):
        """Update the output directory button text based on whether a directory is set"""
        if self.output_dir_edit.text().strip():
            self.output_btn.setText("Change Output Location")
        else:
            self.output_btn.setText("Browse")
    
    # Data handling methods
    def load_from_scheduled_deliveries_internal(self):
        """Internal method to load delivery sequence data from OptimoRoute scheduled deliveries"""
        try:
            # Check if we have scheduled deliveries data
            if not hasattr(self, 'scheduled_orders_data') or not self.scheduled_orders_data:
                self.update_status("No scheduled deliveries data available")
                return
            
            # Filter for scheduled orders only (orders that have scheduling information)
            scheduled_orders = [order for order in self.scheduled_orders_data 
                              if order.get('scheduledAt') or order.get('scheduledAtDt')]
            
            if not scheduled_orders:
                self.update_status("No scheduled orders found in the OptimoRoute data")
                return
            
            self.update_status("Loading delivery data from scheduled deliveries...")
            
            # Convert OptimoRoute data to delivery sequence format
            self.delivery_data_values = []
            self.delivery_data_with_drivers = {}
            
            for order in scheduled_orders:
                order_id = str(order.get('orderNo', '')).strip()
                stop_number = str(order.get('stopNumber', '')).strip()
                driver_name = str(order.get('driverName', '')).strip()
                driver_external_id = str(order.get('driverExternalId', '')).strip()
                
                # Use driver_external_id if available, otherwise use driver_name
                driver_number = driver_external_id if driver_external_id else driver_name
                
                if order_id and driver_number:
                    self.delivery_data_values.append(order_id)
                    self.delivery_data_with_drivers[order_id] = {
                        'stop_number': stop_number,
                        'driver_number': driver_number
                    }
            
            if not self.delivery_data_values:
                self.update_status("No valid orders with driver assignments found in scheduled deliveries")
                return
            
            # Save to JSON
            self.save_delivery_data("scheduled_deliveries")
            
            # Update display
            self.update_delivery_display()
            self.update_status(f"✅ Successfully loaded {len(self.delivery_data_values)} delivery sequences from scheduled deliveries")
            
        except Exception as e:
            self.update_status(f"Error loading from scheduled deliveries: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to load from scheduled deliveries: {str(e)}")
    
    def save_delivery_data(self, source_type="scheduled_deliveries"):
        """Save delivery data to JSON file"""
        try:
            source_info = "OptimoRoute Scheduled Deliveries"
                
            data = {
                "delivery_sequences": self.delivery_data_values,
                "delivery_data_with_drivers": self.delivery_data_with_drivers,
                "source_file": source_info,
                "source_type": source_type,
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
        
        # Ensure the table fills the available width properly
        # The column stretch modes set in create_data_section will handle the distribution
    
    # OptimoRoute API methods
    def on_date_changed(self):
        """Handle date changes and provide feedback"""
        selected_date = self.fetch_date.date()
        today = QDate.currentDate()
        
        if selected_date == today:
            self.update_status(f"Selected date: {selected_date.toString('yyyy-MM-dd')} (Today)")
        elif selected_date == today.addDays(-1):
            self.update_status(f"Selected date: {selected_date.toString('yyyy-MM-dd')} (Yesterday)")
        elif selected_date > today:
            self.update_status(f"Selected date: {selected_date.toString('yyyy-MM-dd')} (Future date)")
        else:
            days_ago = today.daysTo(selected_date)
            self.update_status(f"Selected date: {selected_date.toString('yyyy-MM-dd')} ({abs(days_ago)} days ago)")
    
    def set_quick_date(self, days_offset):
        """Set quick date (today, yesterday, etc.)"""
        target_date = QDate.currentDate().addDays(days_offset)
        self.fetch_date.setDate(target_date)
        
        if days_offset == 0:
            self.update_status("Date set to today")
        elif days_offset == -1:
            self.update_status("Date set to yesterday")
        elif days_offset == -7:
            self.update_status("Date set to one week ago")
        else:
            self.update_status(f"Date set to {abs(days_offset)} days ago")
    
    def validate_date_selection(self):
        """Validate that the selected date is reasonable"""
        selected_date = self.fetch_date.date()
        today = QDate.currentDate()
        
        # Check if the date is too far in the future
        if selected_date > today.addDays(30):
            reply = QMessageBox.question(
                self,
                "Future Date",
                f"You've selected a date more than 30 days in the future ({selected_date.toString('yyyy-MM-dd')}). "
                f"There may not be any scheduled orders for future dates.\n\nDo you want to continue?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return False
        
        # Check if the date is very old
        days_ago = selected_date.daysTo(today)
        if days_ago > 90:
            reply = QMessageBox.question(
                self,
                "Old Date",
                f"You've selected a date from {days_ago} days ago ({selected_date.toString('yyyy-MM-dd')}). "
                f"Older data may not be available.\n\nDo you want to continue?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return False
        
        return True
    
    def fetch_and_load_scheduled_deliveries(self):
        """Fetch scheduled orders from OptimoRoute API and load them into the data preview"""
        if self.optimoroute_thread and self.optimoroute_thread.isRunning():
            QMessageBox.information(self, "In Progress", "Already fetching orders. Please wait...")
            return
        
        # Check if API key is configured
        if not self.api_key:
            QMessageBox.warning(self, "No API Key", 
                               "Please configure your OptimoRoute API key in Settings first.")
            self.open_settings()
            return
        
        # Validate date selection
        if not self.validate_date_selection():
            return
        
        # Get date from UI (always use "All Drivers" as default)
        selected_date = self.fetch_date.date().toString("yyyy-MM-dd")
        # Use the same date for both from and to to get orders for that specific day
        from_date = selected_date
        to_date = selected_date
        driver_filter = None  # Always use "All Drivers" (no filter)
        
        # Start background thread with the selected date
        self.optimoroute_thread = OptimoRouteApiThread(self.api_key, from_date, to_date, driver_filter)
        self.optimoroute_thread.progress_signal.connect(self.update_api_progress)
        self.optimoroute_thread.finished_signal.connect(self.on_fetch_and_load_finished)
        self.optimoroute_thread.start()
    
    def on_fetch_and_load_finished(self, success, orders):
        """Handle fetch completion and automatically load data"""
        if success:
            # Check if data has changed
            data_changed = self.has_data_changed(orders)
            
            self.scheduled_orders_data = orders
            
            # Check if any orders were found
            if not orders:
                # Show no orders found message
                selected_date = self.fetch_date.date().toString('yyyy-MM-dd')
                
                QMessageBox.information(
                    self,
                    "No Orders Found",
                    f"No scheduled orders found for {selected_date}.\n\n"
                    f"This could mean:\n"
                    f"• No orders were scheduled for this date\n"
                    f"• Orders exist but are not yet scheduled\n\n"
                    f"Try selecting a different date."
                )
                
                # Clear existing data display
                self.delivery_data_values = []
                self.delivery_data_with_drivers = {}
                self.update_delivery_display()
                self.update_status(f"No scheduled orders found for {selected_date}")
                return
            
            # Automatically load the fetched data into delivery sequence
            self.load_from_scheduled_deliveries_internal()
            
            # Show change notification if data changed
            if data_changed:
                QMessageBox.information(
                    self,
                    "Data Updated",
                    f"Successfully fetched and loaded {len(orders)} delivery sequences from OptimoRoute scheduled deliveries.\n\n"
                    f"Date: {self.fetch_date.date().toString('yyyy-MM-dd')}\n"
                    f"Driver filter: All Drivers\n"
                    f"Data mapping:\n"
                    f"• Order No → Column A\n"
                    f"• Stop# → Column B\n" 
                    f"• Driver → Column C\n\n"
                    f"You can now process delivery PDFs using this data."
                )
            else:
                QMessageBox.information(
                    self,
                    "No Changes",
                    f"Data fetched successfully but no changes detected.\n\n"
                    f"Current data: {len(orders)} orders for {self.fetch_date.date().toString('yyyy-MM-dd')}"
                )
            
        else:
            QMessageBox.warning(self, "API Error", "Failed to fetch orders from OptimoRoute API.")
    
    def update_api_progress(self, message):
        """Update API progress message"""
        self.update_status(message)
    

    
    def refresh_data(self):
        """Manual refresh functionality"""
        # Trigger immediate refresh
        self.auto_refresh_data()
        self.update_status("Manual refresh triggered")
    
    def auto_refresh_data(self):
        """Automatically refresh data from OptimoRoute API"""
        if not self.auto_refresh_enabled:
            return
            
        if self.optimoroute_thread and self.optimoroute_thread.isRunning():
            # Skip this refresh cycle if already processing
            return
        
        # Check if we have a date selected
        if not self.fetch_date.date().isValid():
            return
        
        # Check if API key is configured
        if not self.api_key:
            return
        
        # Silently trigger the fetch and load process
        self.silent_fetch_and_load_scheduled_deliveries()
    
    def silent_fetch_and_load_scheduled_deliveries(self):
        """Silent version of fetch_and_load_scheduled_deliveries for auto-refresh"""
        if self.optimoroute_thread and self.optimoroute_thread.isRunning():
            return
        
        # Get date from UI (always use "All Drivers" as default)
        selected_date = self.fetch_date.date().toString("yyyy-MM-dd")
        # Use the same date for both from and to to get orders for that specific day
        from_date = selected_date
        to_date = selected_date
        driver_filter = None  # Always use "All Drivers" (no filter)
        
        # Update status without disabling button
        self.update_status("Auto-refreshing data...")
        
        # Start background thread with the selected date
        self.optimoroute_thread = OptimoRouteApiThread(self.api_key, from_date, to_date, driver_filter)
        self.optimoroute_thread.progress_signal.connect(self.update_api_progress)
        self.optimoroute_thread.finished_signal.connect(self.on_silent_fetch_finished)
        self.optimoroute_thread.start()
    
    def on_silent_fetch_finished(self, success, orders):
        """Handle silent fetch completion for auto-refresh"""
        if success:
            # Check if data has changed before updating UI
            data_changed = self.has_data_changed(orders)
            
            if data_changed:
                self.scheduled_orders_data = orders
                
                # Check if any orders were found
                if orders:
                    # Automatically load the fetched data into delivery sequence
                    self.load_from_scheduled_deliveries_internal()
                else:
                    # Clear existing data display
                    self.delivery_data_values = []
                    self.delivery_data_with_drivers = {}
                    self.update_delivery_display()
                    selected_date = self.fetch_date.date().toString('yyyy-MM-dd')
                    self.update_status(f"Auto-refresh: No orders found for {selected_date}")
            else:
                # No status update for no changes detected
                pass
        else:
            self.update_status("Auto-refresh: API connection failed")
    
    def update_refresh_button_tooltip(self):
        """Update the refresh button tooltip based on auto-refresh state"""
        # This method is kept for compatibility but no longer used for the refresh button
        # The refresh button now has a fixed tooltip set in create_header()
        pass
    
    def calculate_data_hash(self, orders_data):
        """Calculate a hash of the orders data to detect changes"""
        if not orders_data:
            return hashlib.md5("empty".encode()).hexdigest()
        
        # Create a string representation of the data for hashing
        data_string = ""
        for order in orders_data:
            # Include key fields that would indicate a change
            data_string += f"{order.get('id', '')}{order.get('orderNo', '')}{order.get('scheduledAt', '')}{order.get('driverName', '')}{order.get('stopNumber', '')}"
        
        return hashlib.md5(data_string.encode()).hexdigest()
    
    def has_data_changed(self, new_orders_data):
        """Check if the new data is different from the last known data"""
        if not new_orders_data:
            new_hash = hashlib.md5("empty".encode()).hexdigest()
            new_count = 0
        else:
            new_hash = self.calculate_data_hash(new_orders_data)
            new_count = len(new_orders_data)
        
        # Check if hash or count has changed
        hash_changed = new_hash != self.last_data_hash
        count_changed = new_count != self.last_order_count
        
        # Update stored values
        self.last_data_hash = new_hash
        self.last_order_count = new_count
        
        return hash_changed or count_changed
    
    def open_settings(self):
        """Open settings dialog to configure API key"""
        dialog = SettingsDialog(self.api_key, self)
        if dialog.exec() == QDialog.Accepted:
            new_api_key = dialog.get_api_key()
            if new_api_key != self.api_key:
                # If API key was cleared, show API key screen
                if not new_api_key:
                    self.api_key = ""
                    self.save_api_key("")
                    self.show_api_key_screen()
                    return
                
                self.api_key = new_api_key
                self.save_api_key(new_api_key)
                self.update_status("API key updated successfully")
    
    def show_api_key_screen(self):
        """Show API key input screen as a blue overlay in the main application"""
        # Create a central widget for the API key screen
        api_key_widget = QWidget()
        api_key_widget.setObjectName("apiKeyScreen")
        self.setCentralWidget(api_key_widget)
        
        # Create layout for the API key screen
        layout = QVBoxLayout(api_key_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(50, 50, 50, 50)
        
        # Add top spacer for centering
        layout.addStretch(1)
        
        # Title
        title_label = QLabel("OptimoRoute Sorter")
        title_label.setObjectName("apiKeyTitle")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # Spacer
        layout.addSpacing(40)
        
        # API Key section (centered)
        api_section = QWidget()
        api_section.setObjectName("apiKeySection")
        api_layout = QVBoxLayout(api_section)
        api_layout.setSpacing(15)
        api_layout.setContentsMargins(30, 25, 30, 25)
        
        # API Key label
        api_label = QLabel("Enter Your OptimoRoute API Key")
        api_label.setObjectName("apiKeyLabel")
        api_label.setAlignment(Qt.AlignCenter)
        api_layout.addWidget(api_label)
        
        # API Key input
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setPlaceholderText("Enter your OptimoRoute API key...")
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.api_key_edit.setObjectName("apiKeyInput")
        self.api_key_edit.returnPressed.connect(self.validate_api_key)
        api_layout.addWidget(self.api_key_edit)
        
        # Show/Hide API key toggle
        show_key_layout = QHBoxLayout()
        self.show_key_checkbox = QCheckBox("Show API Key")
        self.show_key_checkbox.setObjectName("apiKeyCheckbox")
        self.show_key_checkbox.toggled.connect(self.toggle_api_key_visibility)
        show_key_layout.addWidget(self.show_key_checkbox)
        show_key_layout.addStretch()
        api_layout.addLayout(show_key_layout)
        
        # Help text
        help_label = QLabel(
            "Don't have an API key? Get one from your OptimoRoute account:\n"
            "1. Log in to your OptimoRoute account\n"
            "2. Go to Settings → API\n"
            "3. Generate a new API key"
        )
        help_label.setObjectName("apiKeyHelp")
        help_label.setWordWrap(True)
        help_label.setAlignment(Qt.AlignCenter)
        api_layout.addWidget(help_label)
        
        # Continue button (centered)
        continue_btn = QPushButton("Continue")
        continue_btn.setObjectName("apiKeyContinueButton")
        continue_btn.clicked.connect(self.validate_api_key)
        
        # Center the button
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(continue_btn)
        button_layout.addStretch()
        
        api_layout.addLayout(button_layout)
        
        # Center the API section horizontally
        center_layout = QHBoxLayout()
        center_layout.addStretch()
        center_layout.addWidget(api_section)
        center_layout.addStretch()
        
        # Add the centered API section to main layout
        layout.addLayout(center_layout)
        
        # Add bottom spacer for centering
        layout.addStretch(1)
        
        # Set focus to API key input
        self.api_key_edit.setFocus()
        
        # Apply API key screen styling
        self.apply_api_key_screen_styling()
    
    def toggle_api_key_visibility(self, show):
        """Toggle API key visibility"""
        if show:
            self.api_key_edit.setEchoMode(QLineEdit.Normal)
        else:
            self.api_key_edit.setEchoMode(QLineEdit.Password)
    
    def validate_api_key(self):
        """Validate API key and continue with application"""
        api_key = self.api_key_edit.text().strip()
        if not api_key:
            QMessageBox.warning(self, "No API Key", "Please enter an API key to continue.")
            return
        
        # Save the API key
        self.api_key = api_key
        self.save_api_key(api_key)
        
        # Continue with application initialization
        self.continue_initialization()
    
    def apply_api_key_screen_styling(self):
        """Apply styling for the API key screen"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2563eb;
            }
            
            QWidget#apiKeyScreen {
                background-color: #2563eb;
            }
            
            QLabel#apiKeyTitle {
                color: white;
                font-size: 36px;
                font-weight: bold;
                margin-bottom: 10px;
            }
            
            QWidget#apiKeySection {
                background-color: white;
                border-radius: 10px;
                border: none;
                max-width: 400px;
            }
            
            QLabel#apiKeyLabel {
                color: #1e293b;
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 8px;
            }
            
            QLineEdit#apiKeyInput {
                border: 2px solid #d1d5db;
                border-radius: 6px;
                padding: 12px;
                background-color: white;
                color: #374151;
                font-size: 14px;
                font-weight: 500;
            }
            
            QLineEdit#apiKeyInput:focus {
                border-color: #2563eb;
                outline: none;
            }
            
            QCheckBox#apiKeyCheckbox {
                color: #64748b;
                font-size: 14px;
                font-weight: 500;
            }
            
            QCheckBox#apiKeyCheckbox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #d1d5db;
                border-radius: 4px;
                background-color: white;
            }
            
            QCheckBox#apiKeyCheckbox::indicator:checked {
                background-color: #2563eb;
                border-color: #2563eb;
            }
            
            QCheckBox#apiKeyCheckbox::indicator:checked::after {
                content: "✓";
                color: white;
                font-weight: bold;
                font-size: 14px;
            }
            
            QLabel#apiKeyHelp {
                color: #6b7280;
                font-size: 12px;
                padding: 12px;
                background-color: #f8fafc;
                border-radius: 6px;
                border: 1px solid #e2e8f0;
            }
            
            QPushButton#apiKeyContinueButton {
                background-color: #2563eb;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: 600;
                font-size: 14px;
                min-width: 100px;
            }
            
            QPushButton#apiKeyContinueButton:hover {
                background-color: #1d4ed8;
            }
            
            QPushButton#apiKeyContinueButton:pressed {
                background-color: #1e40af;
            }
        """)
    

    
    def continue_initialization(self):
        """Continue with application initialization after API key is provided"""
        # Auto-refresh timer setup
        self.auto_refresh_timer = QTimer()
        self.auto_refresh_timer.timeout.connect(self.auto_refresh_data)
        self.auto_refresh_enabled = True
        self.auto_refresh_timer.start(2000)  # 5 seconds
        
        # Data change tracking
        self.last_data_hash = None
        self.last_order_count = 0
        
        # Reinitialize UI with the main application layout
        self.init_ui()
        self.apply_clean_styling()
        
        # Load saved output directory
        saved_output_dir = self.load_output_directory()
        if saved_output_dir:
            self.output_dir_edit.setText(saved_output_dir)
            self.update_output_button_text()
        
        # Load existing data
        self.load_existing_delivery_data()
        
        # Initialize date range display
        self.on_date_changed()
        
        # Set initial status if no existing data
        if not self.delivery_data_values:
            self.update_status("Ready - Auto-refresh enabled, checking for changes")
        
        # Setup window animation
        self.setup_window_animation()
    
    def close_application(self):
        """Close the application"""
        QApplication.quit()
    
    def load_api_key(self):
        """Load API key from configuration file"""
        config_file = "api_config.json"
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('api_key', '')
            return ''
        except Exception as e:
            print(f"Error loading API key: {e}")
            return ''
    
    def save_api_key(self, api_key):
        """Save API key to configuration file"""
        config_file = "api_config.json"
        try:
            config = {}
            
            # Load existing config if it exists
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            
            # Update with new API key
            config['api_key'] = api_key
            
            # Save back to file
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving API key: {e}")
            QMessageBox.warning(self, "Save Error", f"Could not save API key: {str(e)}")
    
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
        self.processing_thread = ProcessingThread(self)
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
                results_dialog.exec()
            else:
                QMessageBox.critical(self, "Processing Error", f"Error during processing: {error_msg}")
    
    def process_all_pdfs_and_packing_internal(self):
        """Internal method for PDF processing"""
        import re
        
        # Get delivery data for summary generation
        delivery_data_with_drivers = self.delivery_data_with_drivers
        
        try:
            output_dir = Path(self.output_dir_edit.text())
            output_dir.mkdir(exist_ok=True)
            
            # Create date-based subfolder using the selected date
            selected_date = self.fetch_date.date().toString("yyyy-MM-dd")
            date_folder = output_dir / selected_date
            date_folder.mkdir(exist_ok=True)
            
            self.processing_thread.progress_signal.emit(f"Creating output folder: {date_folder}")
            
            # Dictionary to store pages grouped by driver
            driver_pages = {}
            processed_files = 0
            total_pages_processed = 0
            
            # Track found and missing order IDs
            found_order_ids = set()
            missing_order_ids = set()
            
            self.processing_thread.progress_signal.emit("Starting PDF processing...")
            self.processing_thread.progress_signal.emit(f"Processing {len(self.selected_pdf_files)} PDF files...")
            self.processing_thread.progress_signal.emit("Looking for exact order ID matches from delivery data...")
            
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
                        
                        # Search for exact order ID matches from delivery data
                        order_id = None
                        matched_order_id = None
                        
                        # Search for each order ID from delivery data directly in the PDF text
                        for delivery_order_id in self.delivery_data_with_drivers.keys():
                            # Case-insensitive search for the exact order ID
                            if delivery_order_id.upper() in page_text.upper():
                                order_id = delivery_order_id  # Use the exact case from delivery data
                                matched_order_id = delivery_order_id
                                found_order_ids.add(delivery_order_id)  # Track found order
                                self.processing_thread.progress_signal.emit(
                                    f"✅ Found exact match: '{delivery_order_id}' on page {page_num + 1}"
                                )
                                break
                        
                        # If no exact match found, try word boundary search for more precision
                        if not order_id:
                            for delivery_order_id in self.delivery_data_with_drivers.keys():
                                # Use word boundaries to avoid partial matches
                                pattern = r'\b' + re.escape(delivery_order_id) + r'\b'
                                if re.search(pattern, page_text, re.IGNORECASE):
                                    order_id = delivery_order_id
                                    matched_order_id = delivery_order_id
                                    found_order_ids.add(delivery_order_id)  # Track found order
                                    self.processing_thread.progress_signal.emit(
                                        f"✅ Found word boundary match: '{delivery_order_id}' on page {page_num + 1}"
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
                                    f"⚠ Order {order_id} not found in delivery data (page {page_num + 1}) - SKIPPED"
                                )
                        
                        total_pages_processed += 1
                    
                    processed_files += 1
                    pdf_document.close()
                    
                except Exception as e:
                    self.processing_thread.progress_signal.emit(f"Error processing {pdf_file}: {str(e)}")
                    if 'pdf_document' in locals():
                        pdf_document.close()
                    continue
            
            # Calculate missing order IDs
            all_order_ids = set(self.delivery_data_with_drivers.keys())
            missing_order_ids = all_order_ids - found_order_ids
            
            # Report missing orders
            if missing_order_ids:
                self.processing_thread.progress_signal.emit(f"⚠ MISSING ORDERS: {len(missing_order_ids)} order IDs not found in PDF files")
                self.processing_thread.progress_signal.emit("Missing Order IDs:")
                for missing_id in sorted(missing_order_ids):
                    driver_info = self.delivery_data_with_drivers.get(missing_id, {})
                    driver_name = driver_info.get('driver_number', 'Unknown')
                    stop_number = driver_info.get('stop_number', 'Unknown')
                    self.processing_thread.progress_signal.emit(f"  - {missing_id} (Driver: {driver_name}, Stop: {stop_number})")
            else:
                self.processing_thread.progress_signal.emit("✅ All order IDs from delivery data were found in PDF files!")
            
            self.processing_thread.progress_signal.emit(f"📊 Summary: Found {len(found_order_ids)}/{len(all_order_ids)} order IDs")
            
            # Create separate PDF files for each driver
            self.processing_thread.progress_signal.emit("Creating driver-specific PDF files...")
            
            # Show summary of what was found
            total_matched_pages = sum(len(pages) for pages in driver_pages.values())
            self.processing_thread.progress_signal.emit(f"Found {total_matched_pages} pages with matching Order IDs across {len(driver_pages)} drivers")
            
            created_files = []
            failed_files = []
            
            if not driver_pages:
                self.processing_thread.progress_signal.emit("No matching orders found in PDF files!")
                self.processing_thread.progress_signal.emit("Check that your PDF files contain order IDs that match those in your delivery data")
                return {
                    "processed_files": processed_files,
                    "total_pages": total_pages_processed,
                    "driver_files_created": 0,
                    "created_files": [],
                    "failed_files": [],
                    "driver_details": {},
                    "output_dir": str(date_folder),
                    "error": "No matching orders found in PDF files"
                }
            
            for driver_number, pages in driver_pages.items():
                if not pages:
                    continue

                try:
                    # Create new PDF for this driver
                    # Count unique orders for this driver
                    unique_orders = len(set(page_info['order_id'] for page_info in pages))
                    output_filename = f"Driver_{driver_number}_{unique_orders}_Orders.pdf"
                    output_path = date_folder / output_filename

                    self.processing_thread.progress_signal.emit(
                        f"Creating {output_filename} with {len(pages)} pages ({unique_orders} unique orders)..."
                    )

                    # Sort pages by stop number first (delivery sequence order)
                    try:
                        pages.sort(key=lambda x: int(x.get('stop_number', 0)))
                    except (ValueError, TypeError):
                        # If stop numbers aren't numeric, sort as strings
                        pages.sort(key=lambda x: str(x.get('stop_number', '')))
                    
                    # Reverse pages so they print in correct order (last page prints first)
                    pages.reverse()
                    
                    self.processing_thread.progress_signal.emit(
                        f"Pages sorted by delivery sequence and reversed for correct printing order"
                    )

                    new_pdf = fitz.open()
                    pages_added = 0

                    # Add all pages for this driver in reversed delivery sequence order
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
                    # Count unique orders for error message
                    unique_orders = len(set(page_info['order_id'] for page_info in pages))
                    failed_files.append(f"Driver_{driver_number}_{unique_orders}_Orders.pdf")
                    self.processing_thread.progress_signal.emit(
                        f"✗ Error creating PDF for Driver {driver_number}: {str(e)}"
                    )
                    continue
            
            # Create Reversed Picking folder with opposite order
            self.processing_thread.progress_signal.emit("Creating Reversed Picking folder...")
            
            reversed_picking_folder = date_folder / "Reversed Picking Orders"
            reversed_picking_folder.mkdir(exist_ok=True)
            
            reversed_created_files = []
            reversed_failed_files = []
            
            for driver_number, pages in driver_pages.items():
                if not pages:
                    continue

                try:
                    # Create new PDF for this driver in reversed picking order
                    unique_orders = len(set(page_info['order_id'] for page_info in pages))
                    output_filename = f"Driver_{driver_number}_{unique_orders}_Orders.pdf"
                    reversed_output_path = reversed_picking_folder / output_filename

                    self.processing_thread.progress_signal.emit(
                        f"Creating reversed picking {output_filename} with {len(pages)} pages ({unique_orders} unique orders)..."
                    )

                    # Sort pages by stop number (delivery sequence order) - NO REVERSE
                    # This means first deliveries will be picked last
                    reversed_pages = pages.copy()  # Make a copy to avoid modifying original
                    try:
                        reversed_pages.sort(key=lambda x: int(x.get('stop_number', 0)))
                    except (ValueError, TypeError):
                        # If stop numbers aren't numeric, sort as strings
                        reversed_pages.sort(key=lambda x: str(x.get('stop_number', '')))
                    
                    # DO NOT reverse - keep delivery sequence order for reversed picking
                    self.processing_thread.progress_signal.emit(
                        f"Pages sorted by delivery sequence for reversed picking (first deliveries picked last)"
                    )

                    new_pdf = fitz.open()
                    pages_added = 0

                    # Add all pages for this driver in delivery sequence order (first delivery picked last)
                    # Group pages by source file to minimize file opening
                    pages_by_file = {}
                    for page_info in reversed_pages:
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
                                f"Error adding pages from {source_file} to reversed picking PDF: {str(e)}"
                            )
                            continue
                    
                    # Only save if we successfully added pages
                    if pages_added > 0:
                        new_pdf.save(str(reversed_output_path))
                        new_pdf.close()
                        
                        # Verify the file was created
                        if reversed_output_path.exists():
                            reversed_created_files.append(output_filename)
                            self.processing_thread.progress_signal.emit(
                                f"✓ Successfully created reversed picking {output_filename} with {pages_added} pages"
                            )
                        else:
                            reversed_failed_files.append(output_filename)
                            self.processing_thread.progress_signal.emit(
                                f"✗ Failed to create reversed picking {output_filename} - file not found after save"
                            )
                    else:
                        new_pdf.close()
                        reversed_failed_files.append(output_filename)
                        self.processing_thread.progress_signal.emit(
                            f"✗ No pages added to reversed picking {output_filename}"
                        )
                        
                except Exception as e:
                    # Count unique orders for error message
                    unique_orders = len(set(page_info['order_id'] for page_info in pages))
                    reversed_failed_files.append(f"Driver_{driver_number}_{unique_orders}_Orders.pdf")
                    self.processing_thread.progress_signal.emit(
                        f"✗ Error creating reversed picking PDF for Driver {driver_number}: {str(e)}"
                    )
                    continue
            
            # Final summary message
            self.processing_thread.progress_signal.emit("Processing complete!")
            self.processing_thread.progress_signal.emit(f"Created {len(created_files)} PDF files in {date_folder}")
            self.processing_thread.progress_signal.emit(f"Created {len(reversed_created_files)} reversed picking PDF files in {reversed_picking_folder}")
            
            # Generate summary report
            summary_path = date_folder / "processing_summary.txt"
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write("PDF Processing Summary\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Processing Date: {selected_date}\n")
                f.write(f"Output Folder: {date_folder}\n")
                f.write(f"Total PDF files processed: {processed_files}\n")
                f.write(f"Total pages scanned: {total_pages_processed}\n")
                f.write(f"Driver PDF files created: {len(created_files)}\n")
                f.write(f"Reversed picking PDF files created: {len(reversed_created_files)}\n")
                if failed_files:
                    f.write(f"Failed PDF files: {len(failed_files)}\n")
                if reversed_failed_files:
                    f.write(f"Failed reversed picking PDF files: {len(reversed_failed_files)}\n")
                f.write("\n")
                
                # Order matching summary
                f.write("Order Matching Summary:\n")
                f.write(f"  Total orders in delivery data: {len(all_order_ids)}\n")
                f.write(f"  Orders found in PDF files: {len(found_order_ids)}\n")
                f.write(f"  Orders missing from PDF files: {len(missing_order_ids)}\n")
                f.write("\n")
                
                if created_files:
                    f.write("✓ Successfully Created PDF Files:\n")
                    for filename in created_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if reversed_created_files:
                    f.write("✓ Successfully Created Reversed Picking PDF Files:\n")
                    for filename in reversed_created_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if failed_files:
                    f.write("✗ Failed PDF Files:\n")
                    for filename in failed_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if reversed_failed_files:
                    f.write("✗ Failed Reversed Picking PDF Files:\n")
                    for filename in reversed_failed_files:
                        f.write(f"  - {filename}\n")
                    f.write("\n")
                
                if missing_order_ids:
                    f.write("❌ Missing Orders (not found in PDF files):\n")
                    for order_id in sorted(missing_order_ids):
                        # Get driver info from the delivery data
                        driver_info = delivery_data_with_drivers.get(order_id, {})
                        driver_name = driver_info.get('driver_number', 'Unknown')
                        stop_number = driver_info.get('stop_number', 'Unknown')
                        f.write(f"  - {order_id} (Driver: {driver_name}, Stop: {stop_number})\n")
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
                "output_dir": str(date_folder),
                "found_order_ids": list(found_order_ids),
                "missing_order_ids": list(missing_order_ids),
                "total_order_ids": len(all_order_ids),
                "delivery_data_with_drivers": delivery_data_with_drivers
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
            
            QWidget#headerWidget {
                background-color: #2563eb;
                margin-bottom: 10px;
            }
            
            QLabel#headerTitle {
                color: white;
                font-size: 28px;
                font-weight: bold;
            }
            
            QLabel#headerSubtitle {
                color: #e2e8f0;
                font-size: 16px;
            }
            
            QPushButton#refreshButton {
                background-color: white;
                color: #2563eb;
                border: 2px solid rgba(255, 255, 255, 0.8);
                border-radius: 20px;
                font-size: 18px;
                font-weight: bold;
                padding: 0px;
            }
            
            QPushButton#refreshButton:hover {
                background-color: #f8fafc;
                border-color: white;
            }
            
            QPushButton#refreshButton:pressed {
                background-color: #e2e8f0;
                border-color: white;
            }
            
            QPushButton#settingsButton {
                background-color: white;
                color: #2563eb;
                border: 2px solid rgba(255, 255, 255, 0.8);
                border-radius: 20px;
                font-size: 18px;
                font-weight: bold;
                padding: 0px;
            }
            
            QPushButton#settingsButton:hover {
                background-color: #f8fafc;
                border-color: white;
            }
            
            QPushButton#settingsButton:pressed {
                background-color: #e2e8f0;
                border-color: white;
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
            
            QLabel#sectionTitle {
                color: #1e293b;
                font-size: 16px;
                font-weight: bold;
                margin-bottom: 5px;
            }
            
            QLabel#infoText {
                color: #64748b;
                font-size: 12px;
                padding: 8px;
                background-color: #f1f5f9;
                border-radius: 4px;
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
                border-radius: 0px;
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
            
            QPushButton#dangerButton {
                background-color: #dc2626;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                min-height: 20px;
            }
            
            QPushButton#dangerButton:hover {
                background-color: #b91c1c;
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
            
            QComboBox, QDateEdit {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px 12px;
                background-color: white;
                color: #374151;
                font-size: 13px;
                min-height: 20px;
            }
            
            QComboBox:hover, QDateEdit:hover {
                border-color: #2563eb;
            }
            
            QDateEdit:focus {
                border-color: #2563eb;
                outline: none;
            }
            
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 25px;
                border-left: 1px solid #d1d5db;
                background-color: #f8fafc;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            
            QDateEdit::down-arrow {
                image: none;
                border: none;
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #374151;
            }
            
            QDateEdit::down-arrow:hover {
                border-top-color: #2563eb;
            }
            
            QCalendarWidget {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 8px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                gridline-color: #f1f5f9;
                min-width: 350px;
                min-height: 250px;
            }
            
            QCalendarWidget QWidget {
                background-color: white;
                color: #374151;
            }
            
            QCalendarWidget QAbstractItemView:enabled {
                background-color: white;
                color: #374151;
                border: none;
                font-size: 13px;
                selection-background-color: #2563eb;
                selection-color: white;
                outline: none;
            }
            
            QCalendarWidget QAbstractItemView::item {
                padding: 5px;
                border: none;
                background-color: transparent;
            }
            
            QCalendarWidget QAbstractItemView::item:hover {
                background-color: #eff6ff;
                color: #1e40af;
                border-radius: 4px;
            }
            
            QCalendarWidget QAbstractItemView::item:selected {
                background-color: #2563eb;
                color: white;
                border-radius: 4px;
                font-weight: bold;
            }
            
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #f8fafc;
                border-bottom: 1px solid #e2e8f0;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 8px;
            }
            
            QCalendarWidget QToolButton {
                background-color: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 4px;
                color: #374151;
                font-weight: 500;
                padding: 4px 8px;
                margin: 2px;
                min-width: 30px;
                min-height: 25px;
            }
            
            QCalendarWidget QToolButton:hover {
                background-color: #e2e8f0;
                color: #1e293b;
                border-color: #cbd5e1;
            }
            
            QCalendarWidget QToolButton:pressed {
                background-color: #cbd5e1;
                color: #1e293b;
            }
            
            QCalendarWidget QToolButton::menu-indicator {
                image: none;
                width: 0px;
            }
            
            QCalendarWidget QSpinBox {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 4px 8px;
                color: #374151;
                font-weight: 500;
                min-width: 60px;
            }
            
            QCalendarWidget QSpinBox:hover {
                border-color: #2563eb;
            }
            
            QCalendarWidget QSpinBox::up-button, QCalendarWidget QSpinBox::down-button {
                background-color: #f8fafc;
                border: 1px solid #d1d5db;
                width: 16px;
                height: 12px;
            }
            
            QCalendarWidget QSpinBox::up-button:hover, QCalendarWidget QSpinBox::down-button:hover {
                background-color: #e2e8f0;
            }
            
            QCalendarWidget QHeaderView::section {
                background-color: #f8fafc;
                color: #6b7280;
                border: none;
                border-bottom: 1px solid #e2e8f0;
                padding: 8px 4px;
                font-weight: 600;
                font-size: 12px;
                text-align: center;
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
    window = OptimoRouteSorterApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 