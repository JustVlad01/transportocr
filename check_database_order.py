#!/usr/bin/env python3
"""
Script to check what's actually in your dispatch_orders table
"""

from supabase_config import get_supabase_client

def check_current_database_order():
    """Check what's currently in the dispatch_orders table"""
    
    print("=== CHECKING CURRENT DATABASE CONTENT ===\n")
    
    try:
        supabase = get_supabase_client()
        
        # Get all data from dispatch_orders table
        print("📊 Fetching ALL data from dispatch_orders table...")
        
        # First, check without ordering (default database order)
        result_default = supabase.table('dispatch_orders').select(
            'ordernumber, itemcode, product_description, excel_row_sequence, created_at'
        ).execute()
        
        print(f"📋 Found {len(result_default.data)} total records in dispatch_orders table")
        
        if not result_default.data:
            print("❌ No data found in dispatch_orders table!")
            print("💡 Make sure you've uploaded your Excel file using the dispatch scanning app")
            return
        
        # Show default order (what you might be seeing in Supabase dashboard)
        print(f"\n🔍 DEFAULT ORDER (what Supabase dashboard shows):")
        print("-" * 100)
        print("Row | Order Number | Item Code | Product Description | Sequence | Created")
        print("-" * 100)
        
        for i, row in enumerate(result_default.data[:15]):  # Show first 15
            seq = row.get('excel_row_sequence', 'NULL')
            created = row.get('created_at', '')[:19] if row.get('created_at') else 'NULL'
            product = row.get('product_description', '')[:30] + '...' if len(row.get('product_description', '')) > 30 else row.get('product_description', '')
            
            print(f"{i+1:3} | {row['ordernumber']:12} | {row['itemcode']:9} | {product:30} | {seq:8} | {created}")
        
        if len(result_default.data) > 15:
            print(f"... and {len(result_default.data) - 15} more rows")
        
        # Now check with proper Excel order (ordered by excel_row_sequence)
        print(f"\n🔍 CORRECT EXCEL ORDER (ordered by excel_row_sequence):")
        print("-" * 100)
        print("Row | Order Number | Item Code | Product Description | Sequence | Created")
        print("-" * 100)
        
        result_ordered = supabase.table('dispatch_orders').select(
            'ordernumber, itemcode, product_description, excel_row_sequence, created_at'
        ).order('excel_row_sequence').execute()
        
        for i, row in enumerate(result_ordered.data[:15]):  # Show first 15
            seq = row.get('excel_row_sequence', 'NULL')
            created = row.get('created_at', '')[:19] if row.get('created_at') else 'NULL'
            product = row.get('product_description', '')[:30] + '...' if len(row.get('product_description', '')) > 30 else row.get('product_description', '')
            
            print(f"{i+1:3} | {row['ordernumber']:12} | {row['itemcode']:9} | {product:30} | {seq:8} | {created}")
        
        # Compare the two orders
        print(f"\n🔍 ORDER COMPARISON:")
        print("-" * 60)
        
        default_first_10 = [(r['ordernumber'], r['itemcode']) for r in result_default.data[:10]]
        ordered_first_10 = [(r['ordernumber'], r['itemcode']) for r in result_ordered.data[:10]]
        
        if default_first_10 == ordered_first_10:
            print("✅ Default order matches Excel order - your data is already in correct order!")
        else:
            print("❌ Default order does NOT match Excel order")
            print("🔧 SOLUTION: Always query with ORDER BY excel_row_sequence")
            print("\nCorrect SQL query:")
            print("SELECT * FROM dispatch_orders ORDER BY excel_row_sequence;")
        
        # Check for missing sequence numbers
        print(f"\n🔍 SEQUENCE NUMBER CHECK:")
        print("-" * 40)
        
        missing_sequences = [r for r in result_default.data if r.get('excel_row_sequence') is None]
        if missing_sequences:
            print(f"⚠️  Found {len(missing_sequences)} records WITHOUT excel_row_sequence!")
            print("These are probably old uploads before order preservation was implemented.")
            print("Consider re-uploading your Excel file to get proper sequence numbers.")
        else:
            print("✅ All records have excel_row_sequence numbers")
        
        # Show sequence number range
        sequences = [r.get('excel_row_sequence') for r in result_default.data if r.get('excel_row_sequence') is not None]
        if sequences:
            print(f"📊 Sequence numbers range: {min(sequences)} to {max(sequences)}")
        
    except Exception as e:
        print(f"❌ Error checking database: {e}")
        import traceback
        traceback.print_exc()

def show_correct_queries():
    """Show the correct SQL queries to view data in Excel order"""
    
    print("\n=== CORRECT SQL QUERIES FOR EXCEL ORDER ===\n")
    
    print("🔧 To view data in EXACT Excel order, use these queries:")
    print("-" * 60)
    
    print("1. Get all data in Excel order:")
    print("   SELECT * FROM dispatch_orders ORDER BY excel_row_sequence;")
    print()
    
    print("2. Get data for specific order in Excel order:")
    print("   SELECT * FROM dispatch_orders")
    print("   WHERE ordernumber = 'A062M57'")
    print("   ORDER BY excel_row_sequence;")
    print()
    
    print("3. Get picking sequence for all orders:")
    print("   SELECT ordernumber, itemcode, product_description, excel_row_sequence")
    print("   FROM dispatch_orders")
    print("   ORDER BY excel_row_sequence;")
    print()
    
    print("4. Get items to pick in order:")
    print("   SELECT ordernumber, itemcode, quantity, excel_row_sequence as pick_order")
    print("   FROM dispatch_orders")
    print("   WHERE ordernumber = 'YOUR_ORDER_ID'")
    print("   ORDER BY excel_row_sequence;")
    print()
    
    print("⚠️  IMPORTANT: Always use 'ORDER BY excel_row_sequence' to maintain Excel order!")

if __name__ == "__main__":
    check_current_database_order()
    show_correct_queries() 