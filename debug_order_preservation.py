#!/usr/bin/env python3
"""
Debug script to test Excel order preservation in Supabase upload
"""

import pandas as pd
from supabase_config import get_supabase_client, upload_store_orders_from_excel

def test_order_preservation(excel_file_path):
    """Test if Excel file order is preserved when uploading to Supabase"""
    
    print("=== DEBUGGING EXCEL ORDER PRESERVATION ===\n")
    
    try:
        # Step 1: Read Excel file and show original order
        print("📊 Step 1: Reading Excel file...")
        df = pd.read_excel(excel_file_path)
        
        print(f"📋 Excel file has {len(df)} rows")
        print(f"📋 Columns: {list(df.columns)}")
        print("\n🔍 First 10 rows in ORIGINAL Excel order:")
        print("-" * 80)
        for i, row in df.head(10).iterrows():
            print(f"Row {i+1}: {row.iloc[0]} | {row.iloc[1]} | {row.iloc[2] if len(row) > 2 else 'N/A'}")
        
        # Step 2: Convert to records and show order
        print(f"\n📊 Step 2: Converting to records (preserves order)...")
        store_order_data = df.to_dict('records')
        
        print(f"📋 Converted to {len(store_order_data)} records")
        print("\n🔍 First 10 records with sequence numbers:")
        print("-" * 80)
        for i, record in enumerate(store_order_data[:10]):
            order_num = record.get(list(record.keys())[0]) if record.keys() else 'N/A'
            item_code = record.get(list(record.keys())[1]) if len(record.keys()) > 1 else 'N/A'
            print(f"Sequence {i+1}: {order_num} | {item_code}")
        
        # Step 3: Test upload to Supabase
        print(f"\n📊 Step 3: Testing upload to Supabase...")
        success = upload_store_orders_from_excel(store_order_data, excel_file_path)
        
        if success:
            print("✅ Upload successful! Now checking database order...")
            
            # Step 4: Query database to verify order
            supabase = get_supabase_client()
            
            # Get the uploaded data ordered by excel_row_sequence
            result = supabase.table('dispatch_orders').select(
                'ordernumber, itemcode, excel_row_sequence'
            ).order('excel_row_sequence').limit(10).execute()
            
            print(f"\n🔍 First 10 rows from database (ordered by excel_row_sequence):")
            print("-" * 80)
            for row in result.data:
                print(f"DB Sequence {row['excel_row_sequence']}: {row['ordernumber']} | {row['itemcode']}")
            
            # Step 5: Compare Excel vs Database order
            print(f"\n🔍 ORDER COMPARISON:")
            print("-" * 80)
            print("Excel Order → Database Order")
            
            for i, (excel_record, db_row) in enumerate(zip(store_order_data[:10], result.data)):
                excel_order = excel_record.get(list(excel_record.keys())[0])
                excel_item = excel_record.get(list(excel_record.keys())[1])
                
                db_order = db_row['ordernumber']
                db_item = db_row['itemcode']
                db_seq = db_row['excel_row_sequence']
                
                match_symbol = "✅" if (excel_order == db_order and excel_item == db_item) else "❌"
                
                print(f"{match_symbol} Row {i+1}: {excel_order}|{excel_item} → Seq{db_seq}: {db_order}|{db_item}")
            
            # Final verdict
            excel_first_5 = [(r.get(list(r.keys())[0]), r.get(list(r.keys())[1])) for r in store_order_data[:5]]
            db_first_5 = [(r['ordernumber'], r['itemcode']) for r in result.data[:5]]
            
            if excel_first_5 == db_first_5:
                print(f"\n✅ SUCCESS: Excel order is preserved in database!")
            else:
                print(f"\n❌ ISSUE: Excel order is NOT preserved in database!")
                print(f"Excel first 5: {excel_first_5}")
                print(f"DB first 5: {db_first_5}")
        else:
            print("❌ Upload failed!")
            
    except Exception as e:
        print(f"❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()

def check_database_schema():
    """Check if the dispatch_orders table has the excel_row_sequence column"""
    
    print("\n=== CHECKING DATABASE SCHEMA ===\n")
    
    try:
        supabase = get_supabase_client()
        
        # Try to query the table with excel_row_sequence column
        result = supabase.table('dispatch_orders').select('id, excel_row_sequence').limit(1).execute()
        
        if result.data is not None:
            print("✅ dispatch_orders table exists with excel_row_sequence column")
            return True
        else:
            print("❌ Table exists but might not have excel_row_sequence column")
            return False
            
    except Exception as e:
        print(f"❌ Database schema issue: {e}")
        print("\n🔧 SOLUTION: Run this SQL in your Supabase SQL editor:")
        print("-" * 60)
        print("ALTER TABLE dispatch_orders ADD COLUMN excel_row_sequence INT4;")
        print("CREATE INDEX idx_dispatch_orders_excel_sequence ON dispatch_orders(excel_row_sequence);")
        print("-" * 60)
        return False

def clear_test_data():
    """Clear test data from dispatch_orders table"""
    
    print("\n=== CLEARING TEST DATA ===\n")
    
    try:
        supabase = get_supabase_client()
        
        # Delete all records from dispatch_orders
        result = supabase.table('dispatch_orders').delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
        
        print("✅ Test data cleared from dispatch_orders table")
        
    except Exception as e:
        print(f"❌ Error clearing test data: {e}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python debug_order_preservation.py <excel_file_path>")
        print("Example: python debug_order_preservation.py your_order_format.xlsx")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    
    # Check database schema first
    schema_ok = check_database_schema()
    
    if schema_ok:
        # Clear any existing test data
        clear_test_data()
        
        # Test order preservation
        test_order_preservation(excel_file)
    else:
        print("\n⚠️ Please fix the database schema first, then run this script again.") 