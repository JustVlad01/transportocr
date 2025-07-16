# optimoroute_tab.py
import sys
import requests
import json
from datetime import datetime, timedelta
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QProgressBar, QComboBox, QDateEdit, QGroupBox, QFrame,
    QTextEdit, QSplitter, QLineEdit
)
from PySide6.QtCore import Qt, QThread, Signal, QDate
from PySide6.QtGui import QFont


class OptimoRouteApiThread(QThread):
    """Background thread for OptimoRoute API operations"""
    progress_signal = Signal(str)
    finished_signal = Signal(bool, list)
    
    def __init__(self, api_key):
        super().__init__()
        self.api_key = api_key
        self.base_url = "https://api.optimoroute.com/v1"
    
    def run(self):
        try:
            self.progress_signal.emit("Connecting to OptimoRoute API...")
            
            # Fetch all orders - we'll use a different approach since bulk get requires specific order IDs
            # Let's try to get orders for today and recent dates
            headers = {
                'Content-Type': 'application/json'
            }
            
            orders = []
            
            # Try to get orders for the last 7 days
            for i in range(7):
                date = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
                self.progress_signal.emit(f"Fetching orders for {date}...")
                
                url = f"{self.base_url}/get_routes"
                params = {
                    'key': self.api_key,
                    'date': date
                }
                
                try:
                    response = requests.get(url, headers=headers, params=params, timeout=10)
                    
                    if response.status_code == 200:
                        data = response.json()
                        if data.get('success') and data.get('routes'):
                            # Extract orders from routes
                            for route in data['routes']:
                                if 'stops' in route:
                                    for stop in route['stops']:
                                        if stop.get('orderNo') and stop['orderNo'] != '-':
                                            order_data = {
                                                'id': stop.get('id', ''),
                                                'orderNo': stop.get('orderNo', ''),
                                                'date': date,
                                                'address': stop.get('address', ''),
                                                'locationName': stop.get('locationName', ''),
                                                'latitude': stop.get('latitude', ''),
                                                'longitude': stop.get('longitude', ''),
                                                'scheduledAt': stop.get('scheduledAt', ''),
                                                'driverName': route.get('driverName', ''),
                                                'vehicleLabel': route.get('vehicleLabel', ''),
                                                'status': 'active'  # Default status
                                            }
                                            orders.append(order_data)
                    elif response.status_code == 401:
                        self.progress_signal.emit("Authentication failed - please check your API key")
                        self.finished_signal.emit(False, [])
                        return
                    else:
                        self.progress_signal.emit(f"API returned status {response.status_code} for {date}")
                        
                except requests.exceptions.RequestException as e:
                    self.progress_signal.emit(f"Network error for {date}: {str(e)}")
                    continue
            
            if not orders:
                # If no orders found in routes, try a simpler API test
                self.progress_signal.emit("No orders found in routes, testing API connection...")
                
                url = f"{self.base_url}/get_routes"
                params = {
                    'key': self.api_key,
                    'date': datetime.now().strftime('%Y-%m-%d')
                }
                
                response = requests.get(url, headers=headers, params=params, timeout=10)
                
                if response.status_code == 200:
                    self.progress_signal.emit("API connection successful, but no orders found")
                    # Create a sample entry to show the connection works
                    orders = [{
                        'id': 'sample',
                        'orderNo': 'No orders found',
                        'date': datetime.now().strftime('%Y-%m-%d'),
                        'address': 'API connection successful',
                        'locationName': 'No active orders in the last 7 days',
                        'latitude': '',
                        'longitude': '',
                        'scheduledAt': '',
                        'driverName': '',
                        'vehicleLabel': '',
                        'status': 'info'
                    }]
                else:
                    self.progress_signal.emit(f"API connection failed: {response.status_code}")
                    self.finished_signal.emit(False, [])
                    return
            
            self.progress_signal.emit(f"Successfully fetched {len(orders)} orders")
            self.finished_signal.emit(True, orders)
            
        except Exception as e:
            self.progress_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(False, [])


class OptimoRouteTab(QWidget):
    """OptimoRoute orders tab widget"""
    
    def __init__(self, api_key="3ac9317b7972340ccf529ef24f9374fbfYhFnF5FyX4"):
        super().__init__()
        self.api_key = api_key
        self.orders_data = []
        self.processing_thread = None
        
        self.init_ui()
        self.apply_styling()
    
    def init_ui(self):
        """Initialize the user interface"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Header section
        header_frame = self.create_header()
        layout.addWidget(header_frame)
        
        # Controls section
        controls_frame = self.create_controls()
        layout.addWidget(controls_frame)
        
        # Main content area with splitter
        splitter = QSplitter(Qt.Vertical)
        
        # Orders table
        table_frame = self.create_table_section()
        splitter.addWidget(table_frame)
        
        # Details section
        details_frame = self.create_details_section()
        splitter.addWidget(details_frame)
        
        # Set initial splitter sizes (70% table, 30% details)
        splitter.setSizes([700, 300])
        
        layout.addWidget(splitter)
        
        # Status bar
        status_frame = self.create_status_section()
        layout.addWidget(status_frame)
    
    def create_header(self):
        """Create header section"""
        frame = QFrame()
        frame.setObjectName("headerFrame")
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Title and subtitle
        title_layout = QVBoxLayout()
        
        title_label = QLabel("OptimoRoute Orders")
        title_label.setObjectName("headerTitle")
        title_layout.addWidget(title_label)
        
        subtitle_label = QLabel("View and manage your OptimoRoute orders")
        subtitle_label.setObjectName("headerSubtitle")
        title_layout.addWidget(subtitle_label)
        
        layout.addLayout(title_layout)
        layout.addStretch()
        
        # API Status indicator
        self.api_status_label = QLabel("● Disconnected")
        self.api_status_label.setObjectName("apiStatusDisconnected")
        layout.addWidget(self.api_status_label)
        
        return frame
    
    def create_controls(self):
        """Create controls section"""
        frame = QFrame()
        frame.setObjectName("controlsFrame")
        
        layout = QHBoxLayout(frame)
        
        # Refresh button
        self.refresh_btn = QPushButton("🔄 Fetch Orders")
        self.refresh_btn.setObjectName("primaryButton")
        self.refresh_btn.clicked.connect(self.fetch_orders)
        layout.addWidget(self.refresh_btn)
        
        # Date filter
        date_label = QLabel("From Date:")
        layout.addWidget(date_label)
        
        self.date_filter = QDateEdit()
        self.date_filter.setDate(QDate.currentDate().addDays(-7))
        self.date_filter.setCalendarPopup(True)
        layout.addWidget(self.date_filter)
        
        # Driver filter
        driver_label = QLabel("Driver:")
        layout.addWidget(driver_label)
        
        self.driver_filter = QComboBox()
        self.driver_filter.addItem("All Drivers")
        layout.addWidget(self.driver_filter)
        
        # Status filter
        status_label = QLabel("Status:")
        layout.addWidget(status_label)
        
        self.status_filter = QComboBox()
        self.status_filter.addItems(["All Status", "Active", "Completed", "Pending"])
        layout.addWidget(self.status_filter)
        
        layout.addStretch()
        
        # Orders count
        self.orders_count_label = QLabel("Orders: 0")
        self.orders_count_label.setObjectName("ordersCount")
        layout.addWidget(self.orders_count_label)
        
        return frame
    
    def create_table_section(self):
        """Create orders table section"""
        frame = QFrame()
        frame.setObjectName("tableFrame")
        
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Table title
        table_title = QLabel("Orders List")
        table_title.setObjectName("sectionTitle")
        layout.addWidget(table_title)
        
        # Orders table
        self.orders_table = QTableWidget()
        self.orders_table.setColumnCount(8)
        self.orders_table.setHorizontalHeaderLabels([
            "Order No", "Date", "Address", "Location", "Driver", "Vehicle", "Scheduled At", "Status"
        ])
        
        # Configure table
        header = self.orders_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Order No
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Date
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Address
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Location
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Driver
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Vehicle
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Scheduled At
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Status
        
        self.orders_table.setAlternatingRowColors(True)
        self.orders_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.orders_table.itemSelectionChanged.connect(self.on_order_selected)
        
        layout.addWidget(self.orders_table)
        
        return frame
    
    def create_details_section(self):
        """Create order details section"""
        frame = QFrame()
        frame.setObjectName("detailsFrame")
        
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Details title
        details_title = QLabel("Order Details")
        details_title.setObjectName("sectionTitle")
        layout.addWidget(details_title)
        
        # Details text area
        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setPlaceholderText("Select an order to view details...")
        layout.addWidget(self.details_text)
        
        return frame
    
    def create_status_section(self):
        """Create status section"""
        frame = QFrame()
        frame.setObjectName("statusFrame")
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Status label
        self.status_label = QLabel("Ready to fetch orders")
        self.status_label.setObjectName("statusLabel")
        layout.addWidget(self.status_label)
        
        layout.addStretch()
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumWidth(200)
        layout.addWidget(self.progress_bar)
        
        return frame
    
    def fetch_orders(self):
        """Fetch orders from OptimoRoute API"""
        if self.processing_thread and self.processing_thread.isRunning():
            QMessageBox.information(self, "In Progress", "Already fetching orders. Please wait...")
            return
        
        # Show progress
        self.show_progress(True)
        self.refresh_btn.setEnabled(False)
        self.update_status("Initializing API connection...")
        self.update_api_status(False)
        
        # Start background thread
        self.processing_thread = OptimoRouteApiThread(self.api_key)
        self.processing_thread.progress_signal.connect(self.update_status)
        self.processing_thread.finished_signal.connect(self.on_fetch_finished)
        self.processing_thread.start()
    
    def on_fetch_finished(self, success, orders):
        """Handle fetch completion"""
        self.show_progress(False)
        self.refresh_btn.setEnabled(True)
        
        if success:
            self.orders_data = orders
            self.populate_table()
            self.update_status(f"Successfully fetched {len(orders)} orders")
            self.update_api_status(True)
            self.update_filters()
        else:
            self.update_status("Failed to fetch orders")
            self.update_api_status(False)
            QMessageBox.warning(
                self, 
                "API Error", 
                "Failed to fetch orders from OptimoRoute.\n\n"
                "Please check:\n"
                "• Your API key is correct\n"
                "• Your internet connection\n"
                "• OptimoRoute service status"
            )
    
    def populate_table(self):
        """Populate the orders table with data"""
        self.orders_table.setRowCount(len(self.orders_data))
        
        for row, order in enumerate(self.orders_data):
            # Order No
            self.orders_table.setItem(row, 0, QTableWidgetItem(str(order.get('orderNo', ''))))
            
            # Date
            self.orders_table.setItem(row, 1, QTableWidgetItem(str(order.get('date', ''))))
            
            # Address
            self.orders_table.setItem(row, 2, QTableWidgetItem(str(order.get('address', ''))))
            
            # Location Name
            self.orders_table.setItem(row, 3, QTableWidgetItem(str(order.get('locationName', ''))))
            
            # Driver
            self.orders_table.setItem(row, 4, QTableWidgetItem(str(order.get('driverName', ''))))
            
            # Vehicle
            self.orders_table.setItem(row, 5, QTableWidgetItem(str(order.get('vehicleLabel', ''))))
            
            # Scheduled At
            self.orders_table.setItem(row, 6, QTableWidgetItem(str(order.get('scheduledAt', ''))))
            
            # Status
            status_item = QTableWidgetItem(str(order.get('status', 'Unknown')))
            if order.get('status') == 'active':
                status_item.setBackground(Qt.green)
            elif order.get('status') == 'completed':
                status_item.setBackground(Qt.blue)
            elif order.get('status') == 'info':
                status_item.setBackground(Qt.yellow)
            self.orders_table.setItem(row, 7, status_item)
        
        # Update count
        self.orders_count_label.setText(f"Orders: {len(self.orders_data)}")
    
    def on_order_selected(self):
        """Handle order selection"""
        current_row = self.orders_table.currentRow()
        if current_row >= 0 and current_row < len(self.orders_data):
            order = self.orders_data[current_row]
            self.show_order_details(order)
    
    def show_order_details(self, order):
        """Show detailed information for selected order"""
        details = f"""
Order Details:
═══════════════════════════════════════

Order Number: {order.get('orderNo', 'N/A')}
Order ID: {order.get('id', 'N/A')}
Date: {order.get('date', 'N/A')}

Location Information:
─────────────────────────────────────
Address: {order.get('address', 'N/A')}
Location Name: {order.get('locationName', 'N/A')}
Coordinates: {order.get('latitude', 'N/A')}, {order.get('longitude', 'N/A')}

Assignment Information:
─────────────────────────────────────
Driver: {order.get('driverName', 'N/A')}
Vehicle: {order.get('vehicleLabel', 'N/A')}
Scheduled Time: {order.get('scheduledAt', 'N/A')}
Status: {order.get('status', 'N/A')}

Raw Data:
─────────────────────────────────────
{json.dumps(order, indent=2)}
        """
        self.details_text.setPlainText(details.strip())
    
    def update_filters(self):
        """Update filter dropdowns based on current data"""
        # Update driver filter
        drivers = set()
        for order in self.orders_data:
            driver = order.get('driverName', '')
            if driver:
                drivers.add(driver)
        
        current_driver = self.driver_filter.currentText()
        self.driver_filter.clear()
        self.driver_filter.addItem("All Drivers")
        for driver in sorted(drivers):
            self.driver_filter.addItem(driver)
        
        # Restore selection if possible
        index = self.driver_filter.findText(current_driver)
        if index >= 0:
            self.driver_filter.setCurrentIndex(index)
    
    def update_status(self, message):
        """Update status message"""
        self.status_label.setText(message)
    
    def update_api_status(self, connected):
        """Update API connection status"""
        if connected:
            self.api_status_label.setText("● Connected")
            self.api_status_label.setObjectName("apiStatusConnected")
        else:
            self.api_status_label.setText("● Disconnected")
            self.api_status_label.setObjectName("apiStatusDisconnected")
        
        # Refresh styling
        self.api_status_label.style().unpolish(self.api_status_label)
        self.api_status_label.style().polish(self.api_status_label)
    
    def show_progress(self, show=True):
        """Show or hide progress bar"""
        self.progress_bar.setVisible(show)
        if show:
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
        else:
            self.progress_bar.setRange(0, 1)
            self.progress_bar.setValue(1)
    
    def apply_styling(self):
        """Apply custom styling to the widget"""
        self.setStyleSheet("""
            QFrame#headerFrame {
                background-color: #f8fafc;
                border-bottom: 2px solid #e2e8f0;
                padding: 15px 0px;
            }
            
            QLabel#headerTitle {
                color: #1e293b;
                font-size: 24px;
                font-weight: bold;
                margin: 0px;
            }
            
            QLabel#headerSubtitle {
                color: #64748b;
                font-size: 14px;
                margin: 0px;
            }
            
            QLabel#apiStatusConnected {
                color: #059669;
                font-weight: bold;
                font-size: 14px;
            }
            
            QLabel#apiStatusDisconnected {
                color: #dc2626;
                font-weight: bold;
                font-size: 14px;
            }
            
            QFrame#controlsFrame {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
            }
            
            QFrame#tableFrame {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
            }
            
            QFrame#detailsFrame {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
            }
            
            QFrame#statusFrame {
                background-color: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 8px 15px;
            }
            
            QLabel#sectionTitle {
                color: #1e293b;
                font-size: 16px;
                font-weight: bold;
                margin-bottom: 10px;
            }
            
            QLabel#ordersCount {
                color: #1e293b;
                font-weight: bold;
                font-size: 14px;
            }
            
            QLabel#statusLabel {
                color: #374151;
                font-size: 13px;
            }
            
            QPushButton#primaryButton {
                background-color: #2563eb;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 14px;
            }
            
            QPushButton#primaryButton:hover {
                background-color: #1d4ed8;
            }
            
            QPushButton#primaryButton:disabled {
                background-color: #9ca3af;
            }
            
            QTableWidget {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: white;
                gridline-color: #f1f5f9;
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
                border-right: 1px solid #e2e8f0;
                padding: 10px 8px;
                font-weight: 600;
                color: #374151;
                font-size: 13px;
            }
            
            QTextEdit {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: #f8fafc;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 12px;
                color: #374151;
                padding: 10px;
            }
            
            QComboBox, QDateEdit {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px 10px;
                background-color: white;
                color: #374151;
                font-size: 13px;
                min-width: 120px;
            }
            
            QComboBox:hover, QDateEdit:hover {
                border-color: #2563eb;
            }
            
            QProgressBar {
                border: 1px solid #d1d5db;
                border-radius: 4px;
                text-align: center;
                background-color: white;
                color: #374151;
                font-size: 12px;
            }
            
            QProgressBar::chunk {
                background-color: #2563eb;
                border-radius: 3px;
            }
        """) 