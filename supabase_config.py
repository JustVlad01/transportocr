import os
import uuid
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
        if not barcodes_data:
            print("⚠️ No barcode data provided")
            return True
        
        print(f"📋 Processing {len(barcodes_data)} barcode records...")
        
        # Prepare data for insertion
        records = []
        for i, barcode_data in enumerate(barcodes_data):
            try:
                # Clean and validate the data
                order_id = str(barcode_data.get('order_id', '')).strip()
                if not order_id:
                    print(f"⚠️ Skipping barcode record {i + 1}: missing order_id")
                    continue
                
                # Clean string values to prevent JSON issues
                def clean_string(value) -> str:
                    if value is None:
                        return ""
                    # Handle Path objects and other non-string types
                    if hasattr(value, '__fspath__'):  # Path objects
                        s = str(value)
                    else:
                        s = str(value).strip()
                    # Remove problematic characters that can cause JSON issues
                    s = s.replace('\x00', '').replace('\r', ' ').replace('\n', ' ')
                    s = s.replace('\t', ' ').replace('\b', '').replace('\f', ' ')
                    # Remove any remaining control characters
                    s = ''.join(char for char in s if ord(char) >= 32 or char in '\n\r\t')
                    return s
                
                # Ensure all values are JSON-serializable
                try:
                    page_number = barcode_data.get('page_number', 0)
                    if page_number is None:
                        page_number = 0
                    page_number = int(page_number)
                except (ValueError, TypeError):
                    page_number = 0
                
                record = {
                    'order_id': clean_string(order_id),
                    'driver_number': clean_string(barcode_data.get('driver_number', '')),
                    'pdf_file_name': clean_string(barcode_data.get('pdf_file_name', '')),
                    'page_number': page_number,
                    'barcode_type': clean_string(barcode_data.get('barcode_type', 'Code128')),
                    'status': 'generated'
                }
                
                # Validate record before adding
                if len(record['order_id']) > 50:
                    record['order_id'] = record['order_id'][:50]
                if len(record['driver_number']) > 20:
                    record['driver_number'] = record['driver_number'][:20]
                if len(record['pdf_file_name']) > 255:
                    record['pdf_file_name'] = record['pdf_file_name'][:255]
                if len(record['barcode_type']) > 20:
                    record['barcode_type'] = record['barcode_type'][:20]
                
                records.append(record)
                
            except Exception as record_error:
                print(f"⚠️ Error processing barcode record {i + 1}: {str(record_error)}")
                print(f"   Record data: {barcode_data}")
                continue
        
        if not records:
            print("⚠️ No valid barcode records to save")
            return True
        
        print(f"📋 Saving {len(records)} valid barcode records to Supabase...")
        
        # Test JSON serialization before sending to Supabase
        try:
            import json
            json.dumps(records)
        except Exception as json_error:
            print(f"❌ JSON serialization error: {str(json_error)}")
            print(f"❌ Error type: {type(json_error).__name__}")
            # Try to identify the problematic record
            for i, record in enumerate(records):
                try:
                    json.dumps(record)
                except Exception as record_json_error:
                    print(f"❌ JSON error in record {i + 1}: {str(record_json_error)}")
                    print(f"   Problematic record: {record}")
                    # Try to identify the problematic field
                    for key, value in record.items():
                        try:
                            json.dumps({key: value})
                        except Exception as field_error:
                            print(f"   ❌ Problematic field '{key}': {value} (type: {type(value)})")
                            print(f"   ❌ Field error: {str(field_error)}")
                    return False
        
        # Insert all records at once
        result = supabase.table('dispatch.generated_barcodes').upsert(records, on_conflict='order_id').execute()
        
        print(f"✅ Successfully saved {len(records)} barcodes to Supabase")
        return True
        
    except Exception as e:
        print(f"❌ Error saving barcodes to Supabase: {e}")
        print(f"❌ Error type: {type(e).__name__}")
        
        # Check if it's a JSON error
        if "JSON could not be generated" in str(e):
            print("🔍 JSON generation error detected in barcode data")
            print("🔍 This usually indicates invalid characters or data types in the records")
        
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

def upload_store_orders_from_excel(excel_data: List[Dict], excel_file_name: str, created_at_override: Optional[str] = None, pdf_file_name: Optional[str] = None) -> bool:
    """
    Upload store orders from Excel file to dispatch_orders table
    *** ORDER PRESERVATION: Records are uploaded in EXACT Excel file order ***
    
    Expected Excel columns:
    - Column A: ordernumber
    - Column B: itemcode  
    - Column C: product_description
    - Column D: barcode
    - Column E: customer_type
    - Column F: quantity
    - Column G: sitename (optional - will be uploaded if present)
    - Column H: accountcode (optional - will be uploaded if present)
    - Column I: dispatchcode (optional - will be uploaded if present)
    - Column J: route (optional - will be uploaded if present)
    
    Args:
        excel_data: List of dictionaries containing store order items (in Excel row order)
        excel_file_name: Name of the Excel file
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # First, validate that we have data
        if not excel_data:
            print("❌ No data provided in excel_data")
            return False
        
        print(f"📋 Processing {len(excel_data)} rows from Excel file: {excel_file_name}")
        print(f"📋 Excel columns detected: {list(excel_data[0].keys()) if excel_data else 'No columns'}")
        
        records = []

        def normalize_key(key: str) -> str:
            # Lowercase and keep only alphanumeric characters
            return ''.join(ch.lower() for ch in str(key) if ch.isalnum())

        def get_value(item_row: Dict, candidate_keys: List[str], default: str = '') -> str:
            # Build a normalized lookup map once per row
            norm_map = {normalize_key(k): k for k in item_row.keys()}
            for cand in candidate_keys:
                norm = normalize_key(cand)
                if norm in norm_map:
                    return item_row.get(norm_map[norm], default)
            return default

        def truncate(value: str, max_len: int) -> str:
            if value is None:
                return value
            s = str(value)
            return s if len(s) <= max_len else s[:max_len]

        def clean_string(value) -> str:
            """Clean string values to prevent JSON encoding issues"""
            if value is None:
                return ""
            # Convert to string and remove any problematic characters
            s = str(value).strip()
            # Remove null bytes and other problematic characters
            s = s.replace('\x00', '').replace('\r', ' ').replace('\n', ' ')
            return s

        # Process records in exact Excel order (enumerate gives us the sequence)
        for excel_row_index, item in enumerate(excel_data):
            try:
                # Column A: ordernumber (support variants like 'OrderNumber', 'Order Number', 'order_number')
                ordernumber_raw = get_value(item, ['ordernumber', 'OrderNumber', 'Order Number', 'order_number', 'order id', 'orderid', 'Order ID'], '')
                ordernumber = clean_string(ordernumber_raw)

                # Column B: itemcode (support 'ItemCode', 'Item Code', 'item_code')
                itemcode_raw = get_value(item, ['itemcode', 'ItemCode', 'Item Code', 'item_code', 'Code'], '')
                itemcode = clean_string(itemcode_raw)

                # Column C: product_description
                product_description_raw = get_value(item, ['product_description', 'Product Description', 'product description', 'description', 'Description'], '')
                product_description = clean_string(product_description_raw)

                # Column D: barcode
                barcode_raw = get_value(item, ['barcode', 'Barcode', 'bar_code', 'Bar Code'], '')
                barcode = clean_string(barcode_raw)

                # Column E: customer_type (often missing; avoid mapping unrelated fields like 'Source.Name')
                customer_type_raw = get_value(item, ['customer_type', 'Customer Type', 'customer type'], '')
                customer_type = clean_string(customer_type_raw)

                # Column F: quantity
                quantity_value = get_value(item, ['quantity', 'Quantity', 'qty', 'Qty'], 0)

                # NEW: Column G: sitename (exact match with database column)
                sitename_raw = get_value(item, ['sitename', 'SiteName', 'Site Name', 'site name'], '')
                sitename = clean_string(sitename_raw)

                # NEW: Column H: accountcode (exact match with database column)
                accountcode_raw = get_value(item, ['accountcode', 'AccountCode', 'Account Code', 'account code'], '')
                accountcode = clean_string(accountcode_raw)

                # NEW: Column I: dispatchcode (exact match with database column)
                dispatchcode_raw = get_value(item, ['dispatchcode', 'DispatchCode', 'Dispatch Code', 'dispatch code'], '')
                dispatchcode = clean_string(dispatchcode_raw)

                # NEW: Column J: route (exact match with database column)
                route_raw = get_value(item, ['route', 'Route'], '')
                route = clean_string(route_raw)

                # Skip empty rows - at minimum we need ordernumber and itemcode
                if not ordernumber or not itemcode:
                    print(f"⚠️ Skipping row {excel_row_index + 1}: missing ordernumber or itemcode")
                    continue

                # Convert quantity to integer
                try:
                    if quantity_value in (None, '', 'nan', 'NaN'):
                        quantity = 0
                    else:
                        quantity = int(float(str(quantity_value)))
                except (ValueError, TypeError):
                    quantity = 0

                # Enforce DB length limits to avoid 22001 errors
                ordernumber_db = truncate(ordernumber, 50)
                itemcode_db = truncate(itemcode, 50)
                barcode_db = truncate(barcode, 100) if barcode else None
                customer_type_db = truncate(customer_type, 50) if customer_type else None
                sitename_db = truncate(sitename, 100) if sitename else None
                accountcode_db = truncate(accountcode, 100) if accountcode else None
                dispatchcode_db = truncate(dispatchcode, 100) if dispatchcode else None
                route_db = truncate(route, 100) if route else None

                # Create record for dispatch_orders table with Excel row sequence preservation
                record = {
                    'ordernumber': ordernumber_db,
                    'itemcode': itemcode_db,
                    'product_description': product_description if product_description else None,
                    'barcode': barcode_db,
                    'customer_type': customer_type_db,
                    'quantity': quantity,
                    'excel_row_sequence': excel_row_index + 1,  # CRITICAL: Preserves Excel file row order (1-based)
                    'order_start_time': None  # Explicitly set to NULL to prevent automatic timestamp
                }
                
                # Add pdf_file_name from the row data if available
                if 'pdf_file_name' in item and item['pdf_file_name']:
                    record['pdf_file_name'] = item['pdf_file_name']
                elif pdf_file_name:
                    record['pdf_file_name'] = pdf_file_name
                
                # Add the four new columns if they have values
                if sitename_db:
                    record['sitename'] = sitename_db
                if accountcode_db:
                    record['accountcode'] = accountcode_db
                if dispatchcode_db:
                    record['dispatchcode'] = dispatchcode_db
                if route_db:
                    record['route'] = route_db
                
                if created_at_override:
                    record['created_at'] = created_at_override
                
                records.append(record)
                
            except Exception as row_error:
                print(f"⚠️ Error processing row {excel_row_index + 1}: {str(row_error)}")
                print(f"   Row data: {item}")
                continue
        
        if not records:
            print("❌ No valid records found in Excel file")
            print(f"Sample data: {excel_data[:3] if excel_data else 'No data'}")
            return False
        
        print(f"📋 Successfully processed {len(records)} valid records out of {len(excel_data)} total rows")
        
        # Insert records in batches to maintain order
        # Note: Supabase should preserve the insertion order when we provide explicit sequence numbers
        print(f"📋 Uploading {len(records)} records to dispatch_orders table...")
        
        try:
            # Try to insert into dispatch_orders table
            result = supabase.table('dispatch_orders').insert(records).execute()
            print(f"✅ Successfully uploaded {len(records)} dispatch order items from {excel_file_name}")
            print(f"🔢 Excel row order preserved using sequence numbers 1-{len(records)}")
            
            # Upload to crate_verification table
            print(f"📋 Uploading crate verification data...")
            upload_crate_verification_data(records, excel_file_name, created_at_override)
            
        except Exception as db_error:
            print(f"❌ Database upload error: {str(db_error)}")
            print(f"❌ Error details: {type(db_error).__name__}")
            print(f"❌ Error message: {getattr(db_error, 'message', 'No message')}")
            print(f"❌ Error code: {getattr(db_error, 'code', 'No code')}")
            
            # Try to identify the problematic record
            if "JSON could not be generated" in str(db_error):
                print("🔍 JSON generation error detected - checking for problematic data...")
                for i, record in enumerate(records):
                    try:
                        # Test JSON serialization for each record
                        import json
                        json.dumps(record)
                    except Exception as json_error:
                        print(f"❌ JSON error in record {i + 1}: {str(json_error)}")
                        print(f"   Problematic record: {record}")
                        return False
            
            # Check if it's a table not found error
            if "relation" in str(db_error).lower() and "does not exist" in str(db_error).lower():
                print("🔍 Table not found error - checking if dispatch_orders table exists")
                print("🔍 This might be a schema or table name issue")
            
            return False
        
        # Show summary by order with sequence info
        order_counts = {}
        for record in records:
            ordernumber = record['ordernumber']
            if ordernumber not in order_counts:
                order_counts[ordernumber] = {'count': 0, 'first_sequence': record['excel_row_sequence']}
            order_counts[ordernumber]['count'] += 1
        
        print("📋 Dispatch orders uploaded by order (with picking sequence):")
        for ordernumber, info in order_counts.items():
            print(f"  - Order {ordernumber}: {info['count']} items (starting at row {info['first_sequence']})")
        
        return True
        
    except Exception as e:
        print(f"❌ Error uploading dispatch orders: {e}")
        print(f"❌ Error type: {type(e).__name__}")
        print(f"Excel data sample: {excel_data[:2] if excel_data else 'No data'}")
        return False


def upload_crate_verification_data(dispatch_records: List[Dict], excel_file_name: str, created_at_override: Optional[str] = None) -> bool:
    """
    Upload crate verification data to crate_verification table
    Groups dispatch records by ordernumber and creates summary records
    Note: total_items and total_crates are left blank (NULL) as requested
    
    Args:
        dispatch_records: List of dispatch order records
        excel_file_name: Name of the Excel file
        created_at_override: Optional timestamp override
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Group records by ordernumber to create summary data
        order_summaries = {}
        
        for record in dispatch_records:
            ordernumber = record.get('ordernumber')
            if not ordernumber:
                continue
                
            if ordernumber not in order_summaries:
                order_summaries[ordernumber] = {
                    'ordernumber': ordernumber,
                    'sitename': record.get('sitename'),
                    'routenumber': record.get('route'),  # route field maps to routenumber
                    'total_items': None,  # Left blank as requested
                    'total_crates': None  # Left blank as requested
                }
        
        if not order_summaries:
            print("⚠️ No valid order summaries found for crate verification")
            return False
        
        # Convert to list of records for upload
        crate_verification_records = list(order_summaries.values())
        
        # Add created_at override if provided
        if created_at_override:
            for record in crate_verification_records:
                record['created_at'] = created_at_override
        
        print(f"📋 Uploading {len(crate_verification_records)} crate verification records...")
        
        # Upload to crate_verification table
        result = supabase.table('crate_verification').insert(crate_verification_records).execute()
        print(f"✅ Successfully uploaded {len(crate_verification_records)} crate verification records")
        
        # Show summary
        print("📋 Crate verification records uploaded:")
        for record in crate_verification_records:
            print(f"  - Order {record['ordernumber']}: Site: {record['sitename'] or 'N/A'}, Route: {record['routenumber'] or 'N/A'} (total_items and total_crates left blank)")
        
        return True
        
    except Exception as e:
        print(f"❌ Error uploading crate verification data: {e}")
        print(f"❌ Error type: {type(e).__name__}")
        return False


def upload_order_updates_from_excel(excel_data: List[Dict], excel_file_name: str, created_at_override: Optional[str] = None) -> bool:
    """
    Upload order updates from Excel file to dispatch_orders_update table
    *** ORDER PRESERVATION: Records are uploaded in EXACT Excel file order ***
    
    Expected Excel columns for dispatch_orders_update table:
    - Column A: ordernumber (required, unique)
    - Column B: itemcode  
    - Column C: product_description
    - Column D: barcode
    - Column E: customer_type
    - Column F: quantity
    - Column G: quantity_picked (optional)
    - Column H: error_counter (optional)
    - Column I: picker_name (optional)
    - Column J: scanned_by (optional)
    - Column K: full_or_partial_picking (optional)
    - Column L: bakery_items (optional)
    - Column M: sitename (optional)
    - Column N: accountcode (optional)
    - Column O: dispatchcode (optional)
    - Column P: route (optional)
    
    Args:
        excel_data: List of dictionaries containing order update items (in Excel row order)
        excel_file_name: Name of the Excel file
        created_at_override: Optional timestamp override for created_at field
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # First, validate that we have data
        if not excel_data:
            print("❌ No data provided in excel_data")
            return False
        
        print(f"📋 Processing {len(excel_data)} rows from update Excel file: {excel_file_name}")
        print(f"📋 Excel columns detected: {list(excel_data[0].keys()) if excel_data else 'No columns'}")
        
        records = []

        def normalize_key(key: str) -> str:
            # Lowercase and keep only alphanumeric characters
            return ''.join(ch.lower() for ch in str(key) if ch.isalnum())

        def get_value(item_row: Dict, candidate_keys: List[str], default: str = '') -> str:
            # Build a normalized lookup map once per row
            norm_map = {normalize_key(k): k for k in item_row.keys()}
            for cand in candidate_keys:
                norm = normalize_key(cand)
                if norm in norm_map:
                    return item_row.get(norm_map[norm], default)
            return default

        def truncate(value: str, max_len: int) -> str:
            if value is None:
                return value
            s = str(value)
            return s if len(s) <= max_len else s[:max_len]

        def clean_string(value) -> str:
            """Clean string values to prevent JSON encoding issues"""
            if value is None:
                return ""
            # Convert to string and remove any problematic characters
            s = str(value).strip()
            # Remove null bytes and other problematic characters
            s = s.replace('\x00', '').replace('\r', ' ').replace('\n', ' ')
            return s

        # Process records in exact Excel order (enumerate gives us the sequence)
        for excel_row_index, item in enumerate(excel_data):
            try:
                # Column A: ordernumber (required, unique)
                ordernumber_raw = get_value(item, ['ordernumber', 'OrderNumber', 'Order Number', 'order_number', 'order id', 'orderid', 'Order ID'], '')
                ordernumber = clean_string(ordernumber_raw)

                # Column B: itemcode
                itemcode_raw = get_value(item, ['itemcode', 'ItemCode', 'Item Code', 'item_code', 'Code'], '')
                itemcode = clean_string(itemcode_raw)

                # Column C: product_description
                product_description_raw = get_value(item, ['product_description', 'Product Description', 'product description', 'description', 'Description'], '')
                product_description = clean_string(product_description_raw)

                # Column D: barcode
                barcode_raw = get_value(item, ['barcode', 'Barcode', 'bar_code', 'Bar Code'], '')
                barcode = clean_string(barcode_raw)

                # Column E: customer_type
                customer_type_raw = get_value(item, ['customer_type', 'Customer Type', 'customer type'], '')
                customer_type = clean_string(customer_type_raw)

                # Column F: quantity
                quantity_value = get_value(item, ['quantity', 'Quantity', 'qty', 'Qty'], 0)

                # Column G: quantity_picked (optional)
                quantity_picked_value = get_value(item, ['quantity_picked', 'Quantity Picked', 'quantity picked', 'picked'], 0)

                # Column H: error_counter (optional)
                error_counter_value = get_value(item, ['error_counter', 'Error Counter', 'error counter', 'errors'], 0)

                # Column I: picker_name (optional)
                picker_name_raw = get_value(item, ['picker_name', 'Picker Name', 'picker name', 'picker'], '')
                picker_name = clean_string(picker_name_raw)

                # Column J: scanned_by (optional)
                scanned_by_raw = get_value(item, ['scanned_by', 'Scanned By', 'scanned by', 'scanner'], '')
                scanned_by = clean_string(scanned_by_raw)

                # Column K: full_or_partial_picking (optional)
                full_or_partial_raw = get_value(item, ['full_or_partial_picking', 'Full Or Partial Picking', 'full or partial picking', 'partial'], '')
                full_or_partial_picking = None
                if full_or_partial_raw:
                    # Convert to boolean
                    full_or_partial_picking = str(full_or_partial_raw).lower() in ['true', '1', 'yes', 'full', 'complete']

                # Column L: bakery_items (optional)
                bakery_items_raw = get_value(item, ['bakery_items', 'Bakery Items', 'bakery items', 'bakery'], '')
                bakery_items = None
                if bakery_items_raw:
                    # Convert to boolean
                    bakery_items = str(bakery_items_raw).lower() in ['true', '1', 'yes']

                # Column M: sitename (optional)
                sitename_raw = get_value(item, ['sitename', 'SiteName', 'Site Name', 'site name'], '')
                sitename = clean_string(sitename_raw)

                # Column N: accountcode (optional)
                accountcode_raw = get_value(item, ['accountcode', 'AccountCode', 'Account Code', 'account code'], '')
                accountcode = clean_string(accountcode_raw)

                # Column O: dispatchcode (optional)
                dispatchcode_raw = get_value(item, ['dispatchcode', 'DispatchCode', 'Dispatch Code', 'dispatch code'], '')
                dispatchcode = clean_string(dispatchcode_raw)

                # Column P: route (optional)
                route_raw = get_value(item, ['route', 'Route'], '')
                route = clean_string(route_raw)

                # Skip empty rows - at minimum we need ordernumber
                if not ordernumber:
                    print(f"⚠️ Skipping row {excel_row_index + 1}: missing ordernumber")
                    continue

                # Convert numeric values
                try:
                    if quantity_value in (None, '', 'nan', 'NaN'):
                        quantity = 0
                    else:
                        quantity = int(float(str(quantity_value)))
                except (ValueError, TypeError):
                    quantity = 0

                try:
                    if quantity_picked_value in (None, '', 'nan', 'NaN'):
                        quantity_picked = 0
                    else:
                        quantity_picked = int(float(str(quantity_picked_value)))
                except (ValueError, TypeError):
                    quantity_picked = 0

                try:
                    if error_counter_value in (None, '', 'nan', 'NaN'):
                        error_counter = 0
                    else:
                        error_counter = int(float(str(error_counter_value)))
                except (ValueError, TypeError):
                    error_counter = 0

                # Enforce DB length limits to avoid 22001 errors
                ordernumber_db = truncate(ordernumber, 50)
                itemcode_db = truncate(itemcode, 50)
                barcode_db = truncate(barcode, 100) if barcode else None
                customer_type_db = truncate(customer_type, 50) if customer_type else None
                picker_name_db = truncate(picker_name, 100) if picker_name else None
                scanned_by_db = truncate(scanned_by, 100) if scanned_by else None
                sitename_db = truncate(sitename, 100) if sitename else None
                accountcode_db = truncate(accountcode, 100) if accountcode else None
                dispatchcode_db = truncate(dispatchcode, 100) if dispatchcode else None
                route_db = truncate(route, 100) if route else None

                # Create record for dispatch_orders_update table with Excel row sequence preservation
                record = {
                    'id': str(uuid.uuid4()),  # Generate UUID for primary key
                    'ordernumber': ordernumber_db,
                    'itemcode': itemcode_db,
                    'product_description': product_description if product_description else None,
                    'barcode': barcode_db,
                    'customer_type': customer_type_db,
                    'quantity': quantity,
                    'quantity_picked': quantity_picked,
                    'error_counter': error_counter,
                    'picker_name': picker_name_db,
                    'scanned_by': scanned_by_db,
                    'full_or_partial_picking': full_or_partial_picking,
                    'bakery_items': bakery_items,
                    'excel_row_sequence': excel_row_index + 1  # CRITICAL: Preserves Excel file row order (1-based)
                }
                
                # Add optional columns if they have values
                if sitename_db:
                    record['sitename'] = sitename_db
                if accountcode_db:
                    record['accountcode'] = accountcode_db
                if dispatchcode_db:
                    record['dispatchcode'] = dispatchcode_db
                if route_db:
                    record['route'] = route_db
                
                if created_at_override:
                    record['created_at'] = created_at_override
                
                records.append(record)
                
            except Exception as row_error:
                print(f"⚠️ Error processing row {excel_row_index + 1}: {str(row_error)}")
                print(f"   Row data: {item}")
                continue
        
        if not records:
            print("❌ No valid records found in update Excel file")
            print(f"Sample data: {excel_data[:3] if excel_data else 'No data'}")
            return False
        
        print(f"📋 Successfully processed {len(records)} valid records out of {len(excel_data)} total rows")
        
        # Insert records in batches to maintain order
        print(f"📋 Uploading {len(records)} records to dispatch_orders_update table...")
        
        try:
            result = supabase.table('dispatch_orders_update').insert(records).execute()
            print(f"✅ Successfully uploaded {len(records)} order update items from {excel_file_name}")
            print(f"🔢 Excel row order preserved using sequence numbers 1-{len(records)}")
        except Exception as db_error:
            print(f"❌ Database upload error: {str(db_error)}")
            print(f"❌ Error details: {type(db_error).__name__}")
            
            # Try to identify the problematic record
            if "JSON could not be generated" in str(db_error):
                print("🔍 JSON generation error detected - checking for problematic data...")
                for i, record in enumerate(records):
                    try:
                        # Test JSON serialization for each record
                        import json
                        json.dumps(record)
                    except Exception as json_error:
                        print(f"❌ JSON error in record {i + 1}: {str(json_error)}")
                        print(f"   Problematic record: {record}")
                        return False
            return False
        
        # Show summary by order with sequence info
        order_counts = {}
        for record in records:
            ordernumber = record['ordernumber']
            if ordernumber not in order_counts:
                order_counts[ordernumber] = {'count': 0, 'first_sequence': record['excel_row_sequence']}
            order_counts[ordernumber]['count'] += 1
        
        print(f"📊 Update Summary:")
        print(f"   - Total orders updated: {len(order_counts)}")
        print(f"   - Total items uploaded: {len(records)}")
        print(f"   - Excel row sequence range: 1-{len(records)}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error in upload_order_updates_from_excel: {str(e)}")
        print(f"❌ Error type: {type(e).__name__}")
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