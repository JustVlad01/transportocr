#!/usr/bin/env python3
"""
Example script showing how to upload pick list data from Excel to Supabase
"""

import pandas as pd
from supabase_config import upload_pick_list_from_excel, get_pick_list_for_order, record_barcode_scan

def upload_excel_pick_list(excel_file_path):
    """
    Upload pick list data from Excel file to Supabase
    
    Expected Excel format (matching your current format):
    | order_number | items   | quantity |
    |--------------|---------|----------|
    | A060JR7      | item 1  | 3        |
    | A060JR7      | item 2  | 2        |
    | A060JR7      | item 3  | 6        |
    | A060JR7      | item 4  | 1        |
    | A060JR7      | item 5  | 1        |
    """
    try:
        # Read Excel file
        df = pd.read_excel(excel_file_path)
        
        print(f"📊 Excel file columns: {list(df.columns)}")
        print(f"📊 First few rows:")
        print(df.head())
        
        # Convert DataFrame to list of dictionaries
        pick_list_data = df.to_dict('records')
        
        # Upload to Supabase
        success = upload_pick_list_from_excel(pick_list_data, excel_file_path)
        
        if success:
            print(f"✅ Successfully uploaded pick list from {excel_file_path}")
            return True
        else:
            print(f"❌ Failed to upload pick list from {excel_file_path}")
            return False
            
    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        return False

def demo_barcode_scanning():
    """
    Demonstrate the barcode scanning workflow
    """
    print("\n=== Barcode Scanning Demo ===")
    
    # Simulate scanning a barcode
    scanned_order_id = "A060JR7"  # This would come from your barcode scanner
    
    print(f"📱 Scanned barcode for Order ID: {scanned_order_id}")
    
    # Record the scan
    scan_success = record_barcode_scan(
        order_id=scanned_order_id,
        scanned_by="John Doe",
        scanner_device="Handheld Scanner 001",
        location="Warehouse A"
    )
    
    if scan_success:
        print("✅ Scan recorded successfully")
        
        # Get the pick list for this order
        pick_list = get_pick_list_for_order(scanned_order_id)
        
        if pick_list:
            print(f"\n📋 Pick List for Order {scanned_order_id}:")
            print("-" * 50)
            for item in pick_list:
                print(f"• {item['item_code']}: {item['item_description']}")
                print(f"  Quantity: {item['quantity_required']}")
                print(f"  Location: {item['pick_location']}")
                print(f"  Sequence: {item['pick_sequence']}")
                print()
        else:
            print(f"❌ No pick list found for Order {scanned_order_id}")
    else:
        print("❌ Failed to record scan")

if __name__ == "__main__":
    print("=== Excel Upload and Barcode Scanning Demo ===")
    
    # Example: Upload pick list from Excel
    # excel_file_path = "pick_list_example.xlsx"
    # upload_excel_pick_list(excel_file_path)
    
    # Example: Demonstrate barcode scanning
    demo_barcode_scanning()
    
    print("\n=== Instructions ===")
    print("1. First, run your transport sorter to generate barcodes")
    print("2. Upload an Excel file with pick list data using upload_excel_pick_list()")
    print("3. When someone scans a barcode, it will:")
    print("   - Record the scan in the database")
    print("   - Return the pick list for that order")
    print("   - Show what items need to be picked")
    print("\n📋 Excel file should have columns:")
    print("   - order_id: The order ID that matches your barcodes")
    print("   - item_code: Unique code for each item")
    print("   - item_description: Description of the item")
    print("   - quantity_required: How many to pick")
    print("   - pick_location: Where to find the item")
    print("   - pick_sequence: Order to pick items") 