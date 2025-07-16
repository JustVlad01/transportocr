import os
from supabase import create_client, Client
from datetime import datetime
from typing import List, Dict, Optional

# Hardcoded Supabase credentials (for development only)
SUPABASE_URL = "https://doftypeumwgvirppcuim.supabase.co"  # Replace with your actual URL
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRvZnR5cGV1bXdndmlycHBjdWltIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDYwNzQxMjcsImV4cCI6MjA2MTY1MDEyN30.v8XG5wOU50Jy9qca6MG_mVqtvXf96lKjagiwPh5DsqA"  # Replace with your actual anon key

# Note: Replace the above with your actual Supabase credentials
# You can find these in your Supabase project settings

# Create Supabase client
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def get_supabase_client():
    """Get the configured Supabase client"""
    return supabase

# ================================
# BARCODE MANAGEMENT FUNCTIONS
# ================================

def save_generated_barcodes(barcodes_data: List[Dict]) -> bool:
    """
    Save generated barcodes to Supabase
    
    Args:
        barcodes_data: List of dictionaries containing:
            - order_id: The order ID (what's encoded in the barcode)
            - driver_number: Driver number
            - pdf_file_name: Name of the PDF file
            - page_number: Page number in the PDF
            - barcode_type: Type of barcode (default: 'Code128')
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Prepare data for insertion
        records = []
        for barcode_data in barcodes_data:
            record = {
                'order_id': barcode_data['order_id'],
                'driver_number': str(barcode_data.get('driver_number', '')),
                'pdf_file_name': barcode_data.get('pdf_file_name', ''),
                'page_number': barcode_data.get('page_number', 0),
                'barcode_type': barcode_data.get('barcode_type', 'Code128'),
                'status': 'generated'
            }
            records.append(record)
        
        # Insert all records at once
        result = supabase.table('dispatch.generated_barcodes').upsert(records, on_conflict='order_id').execute()
        
        print(f"✅ Successfully saved {len(records)} barcodes to Supabase")
        return True
        
    except Exception as e:
        print(f"❌ Error saving barcodes to Supabase: {e}")
        return False

def get_barcode_info(order_id: str) -> Optional[Dict]:
    """Get barcode information for a specific order ID"""
    try:
        result = supabase.table('dispatch.generated_barcodes').select("*").eq('order_id', order_id).execute()
        
        if result.data:
            return result.data[0]
        return None
        
    except Exception as e:
        print(f"❌ Error getting barcode info: {e}")
        return None

def update_barcode_status(order_id: str, status: str) -> bool:
    """Update the status of a barcode (generated, scanned, picked, completed)"""
    try:
        result = supabase.table('dispatch.generated_barcodes').update({
            'status': status,
            'updated_at': datetime.now().isoformat()
        }).eq('order_id', order_id).execute()
        
        return True
        
    except Exception as e:
        print(f"❌ Error updating barcode status: {e}")
        return False

def record_barcode_scan(order_id: str, scanned_by: str = None, scanner_device: str = None, location: str = None) -> bool:
    """Record when a barcode is scanned"""
    try:
        scan_record = {
            'order_id': order_id,
            'scanned_by': scanned_by,
            'scanner_device': scanner_device,
            'location': location
        }
        
        # Insert scan record
        result = supabase.table('dispatch.scan_history').insert(scan_record).execute()
        
        # Update barcode status to 'scanned'
        update_barcode_status(order_id, 'scanned')
        
        print(f"✅ Recorded scan for order {order_id}")
        return True
        
    except Exception as e:
        print(f"❌ Error recording scan: {e}")
        return False

# ================================
# PICK LIST MANAGEMENT FUNCTIONS
# ================================

def upload_pick_list_from_excel(excel_data: List[Dict], excel_file_name: str) -> bool:
    """
    Upload pick list data from Excel file
    
    Expected Excel columns:
    - order_number (Column A): The order ID (e.g., A060JR7)
    - items (Column B): Description of the item (e.g., "item 1")
    - quantity (Column C): How many to pick (e.g., 3)
    
    Args:
        excel_data: List of dictionaries containing pick list items
        excel_file_name: Name of the Excel file
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        records = []
        
        # Group items by order to assign sequence numbers
        order_sequences = {}
        
        for i, item in enumerate(excel_data):
            # Handle different possible column names
            order_id = item.get('order_number') or item.get('order_id') or item.get('Order Number') or ''
            item_description = item.get('items') or item.get('item_description') or item.get('Items') or ''
            quantity = item.get('quantity') or item.get('quantity_required') or item.get('Quantity') or 0
            
            # Skip empty rows
            if not order_id or not item_description:
                continue
            
            # Auto-assign sequence number for each order
            if order_id not in order_sequences:
                order_sequences[order_id] = 0
            order_sequences[order_id] += 1
            
            # Auto-generate item code if not provided
            item_code = item.get('item_code') or f"{order_id}_ITEM_{order_sequences[order_id]:03d}"
            
            record = {
                'order_id': str(order_id).strip(),
                'item_code': item_code,
                'item_description': str(item_description).strip(),
                'quantity_required': int(quantity) if str(quantity).isdigit() else 0,
                'pick_sequence': order_sequences[order_id],
                'pick_location': item.get('pick_location') or item.get('location') or None,
                'status': 'pending',
                'excel_file_name': excel_file_name
            }
            records.append(record)
        
        if not records:
            print("❌ No valid records found in Excel file")
            return False
        
        # Insert all records
        result = supabase.table('pick_lists').insert(records).execute()
        
        print(f"✅ Successfully uploaded {len(records)} pick list items from {excel_file_name}")
        
        # Show summary by order
        order_counts = {}
        for record in records:
            order_id = record['order_id']
            order_counts[order_id] = order_counts.get(order_id, 0) + 1
        
        print("📋 Items uploaded by order:")
        for order_id, count in order_counts.items():
            print(f"  - Order {order_id}: {count} items")
        
        return True
        
    except Exception as e:
        print(f"❌ Error uploading pick list: {e}")
        print(f"Excel data sample: {excel_data[:2] if excel_data else 'No data'}")
        return False

def get_pick_list_for_order(order_id: str) -> List[Dict]:
    """Get all items that need to be picked for a specific order"""
    try:
        result = supabase.table('pick_lists').select("*").eq('order_id', order_id).order('pick_sequence').execute()
        
        return result.data if result.data else []
        
    except Exception as e:
        print(f"❌ Error getting pick list: {e}")
        return []

def update_pick_item_status(order_id: str, item_code: str, quantity_picked: int, picked_by: str = None) -> bool:
    """Update the picked quantity and status for a specific item"""
    try:
        update_data = {
            'quantity_picked': quantity_picked,
            'picked_at': datetime.now().isoformat(),
            'picked_by': picked_by,
            'status': 'picked'
        }
        
        result = supabase.table('pick_lists').update(update_data).eq('order_id', order_id).eq('item_code', item_code).execute()
        
        return True
        
    except Exception as e:
        print(f"❌ Error updating pick item: {e}")
        return False

def get_barcode_scan_history(order_id: str) -> List[Dict]:
    """Get scan history for a specific order"""
    try:
        result = supabase.table('dispatch.scan_history').select("*").eq('order_id', order_id).order('scanned_at', desc=True).execute()
        
        return result.data if result.data else []
        
    except Exception as e:
        print(f"❌ Error getting scan history: {e}")
        return []

def upload_store_orders_from_excel(excel_data: List[Dict], excel_file_name: str) -> bool:
    """
    Upload store orders from Excel file to dispatch_orders table
    
    Expected Excel columns:
    - Column A: ordernumber
    - Column B: itemcode  
    - Column C: product_description
    - Column D: barcode
    - Column E: customer_type
    - Column F: quantity
    
    Args:
        excel_data: List of dictionaries containing store order items
        excel_file_name: Name of the Excel file
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        records = []
        
        for i, item in enumerate(excel_data):
            # Handle the specific column mapping for dispatch_orders table
            # Column A: ordernumber
            ordernumber_raw = (item.get('ordernumber') or 
                              item.get('order number') or 
                              item.get('Order Number') or 
                              item.get('order_number') or 
                              item.get(list(item.keys())[0] if item.keys() else None) or '')
            ordernumber = str(ordernumber_raw).strip()
            
            # Column B: itemcode
            itemcode_raw = (item.get('itemcode') or 
                           item.get('item code') or 
                           item.get('Item Code') or 
                           item.get('item_code') or 
                           item.get(list(item.keys())[1] if len(item.keys()) > 1 else None) or '')
            itemcode = str(itemcode_raw).strip()
            
            # Column C: product_description
            product_description_raw = (item.get('product_description') or 
                                     item.get('product description') or 
                                     item.get('Product Description') or 
                                     item.get('description') or 
                                     item.get('Description') or 
                                     item.get(list(item.keys())[2] if len(item.keys()) > 2 else None) or '')
            product_description = str(product_description_raw).strip()
            
            # Column D: barcode
            barcode_raw = (item.get('barcode') or 
                          item.get('Barcode') or 
                          item.get('bar_code') or 
                          item.get('Bar Code') or 
                          item.get(list(item.keys())[3] if len(item.keys()) > 3 else None) or '')
            barcode = str(barcode_raw).strip()
            
            # Column E: customer_type
            customer_type_raw = (item.get('customer_type') or 
                               item.get('customer type') or 
                               item.get('Customer Type') or 
                               item.get('customer') or 
                               item.get('Customer') or 
                               item.get(list(item.keys())[4] if len(item.keys()) > 4 else None) or '')
            customer_type = str(customer_type_raw).strip()
            
            # Column F: quantity
            quantity_value = (item.get('quantity') or 
                            item.get('Quantity') or 
                            item.get('qty') or 
                            item.get('Qty') or 
                            item.get(list(item.keys())[5] if len(item.keys()) > 5 else None) or 0)
            
            # Skip empty rows - at minimum we need ordernumber and itemcode
            if not ordernumber or not itemcode:
                continue
            
            # Convert quantity to integer
            try:
                quantity = int(float(str(quantity_value))) if quantity_value else 0
            except (ValueError, TypeError):
                quantity = 0
            
            # Create record for dispatch_orders table
            record = {
                'ordernumber': str(ordernumber),
                'itemcode': str(itemcode),
                'product_description': str(product_description) if product_description else None,
                'barcode': str(barcode) if barcode else None,
                'customer_type': str(customer_type) if customer_type else None,
                'quantity': quantity
            }
            records.append(record)
        
        if not records:
            print("❌ No valid records found in Excel file")
            print(f"Sample data: {excel_data[:3] if excel_data else 'No data'}")
            return False
        
        # Insert all records into dispatch_orders table
        result = supabase.table('dispatch_orders').insert(records).execute()
        
        print(f"✅ Successfully uploaded {len(records)} dispatch order items from {excel_file_name}")
        
        # Show summary by order
        order_counts = {}
        for record in records:
            ordernumber = record['ordernumber']
            order_counts[ordernumber] = order_counts.get(ordernumber, 0) + 1
        
        print("📋 Dispatch orders uploaded by order:")
        for ordernumber, count in order_counts.items():
            print(f"  - Order {ordernumber}: {count} items")
        
        return True
        
    except Exception as e:
        print(f"❌ Error uploading dispatch orders: {e}")
        print(f"Excel data sample: {excel_data[:2] if excel_data else 'No data'}")
        return False

# ================================
# LEGACY FUNCTIONS (for backward compatibility)
# ================================

def insert_delivery_data(table_name: str, data: dict):
    """Insert data into a Supabase table"""
    try:
        result = supabase.table(table_name).insert(data).execute()
        return result
    except Exception as e:
        print(f"Error inserting data: {e}")
        return None

def get_delivery_data(table_name: str, filters: dict = None):
    """Get data from a Supabase table with optional filters"""
    try:
        query = supabase.table(table_name).select("*")
        
        if filters:
            for key, value in filters.items():
                query = query.eq(key, value)
        
        result = query.execute()
        return result.data
    except Exception as e:
        print(f"Error getting data: {e}")
        return None

def update_delivery_data(table_name: str, data: dict, filters: dict):
    """Update data in a Supabase table"""
    try:
        query = supabase.table(table_name).update(data)
        
        for key, value in filters.items():
            query = query.eq(key, value)
        
        result = query.execute()
        return result
    except Exception as e:
        print(f"Error updating data: {e}")
        return None

def delete_delivery_data(table_name: str, filters: dict):
    """Delete data from a Supabase table"""
    try:
        query = supabase.table(table_name).delete()
        
        for key, value in filters.items():
            query = query.eq(key, value)
        
        result = query.execute()
        return result
    except Exception as e:
        print(f"Error deleting data: {e}")
        return None 