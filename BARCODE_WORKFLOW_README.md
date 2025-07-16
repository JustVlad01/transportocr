# 🏷️ Barcode & Pick List Workflow

This system integrates barcode generation with Supabase database storage and pick list management. Here's how it all works together:

## 📋 **Complete Workflow Overview**

### 1. **Generate Barcodes** 
- Run your transport sorter app to process picking dockets
- Barcodes are generated for each unique order ID
- **Barcodes are saved to Supabase as TEXT/NUMBERS, not images**
- Each barcode contains the order ID (e.g., "A060JR7")

### 2. **Upload Pick Lists**
- Create an Excel file with picking instructions
- Upload it to Supabase using the provided functions
- This links order IDs with specific items to pick

### 3. **Scan & Pick**
- Workers scan barcodes with handheld scanners
- System retrieves pick list for that order
- Items are marked as picked when completed

---

## 🗄️ **Database Schema**

Your Supabase database will have these tables in the `dispatch` schema:

### `dispatch.generated_barcodes`
- `order_id` - The order number (what's IN the barcode)
- `driver_number` - Which driver this order belongs to
- `pdf_file_name` - Source PDF file name
- `page_number` - Page number in PDF
- `barcode_type` - Type of barcode (Code128)
- `status` - generated, scanned, picked, completed

### `dispatch.pick_lists`
- `order_id` - Links to the barcode (from your order_number column)
- `item_code` - Auto-generated unique identifier (e.g., A060JR7_ITEM_001)
- `item_description` - What to pick (from your items column)
- `quantity_required` - How many to pick (from your quantity column)
- `quantity_picked` - How many have been picked (starts at 0)
- `pick_location` - Where to find it (optional, can be added later)
- `pick_sequence` - Order to pick items (auto-assigned: 1, 2, 3, 4, 5)
- `status` - pending, picked, completed

### `dispatch.scan_history`
- `order_id` - Which order was scanned
- `scanned_at` - When it was scanned
- `scanned_by` - Who scanned it
- `scanner_device` - Which device was used

---

## 🚀 **Setup Instructions**

### 1. **Database Setup**
```sql
-- Run this SQL in your Supabase database
-- (Copy from supabase_schema.sql)
CREATE SCHEMA IF NOT EXISTS dispatch;

CREATE TABLE dispatch.generated_barcodes (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    order_id VARCHAR(50) NOT NULL UNIQUE,
    barcode_type VARCHAR(20) DEFAULT 'Code128',
    driver_number VARCHAR(20),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    -- ... (see supabase_schema.sql for complete schema)
);
```

### 2. **Environment Setup**
Make sure you have:
- ✅ `.env` file with your Supabase credentials
- ✅ Required Python packages installed
- ✅ Supabase tables created

### 3. **Generate Barcodes**
```python
# Run your transport sorter app
python main.py

# This will:
# - Generate barcodes for each order
# - Save them to Supabase automatically
# - Create PDF files with barcodes
```

### 4. **Upload Pick Lists**
```python
# Use the example script to upload your Excel file
from excel_upload_example import upload_excel_pick_list

# Upload your Excel file (with order_number, items, quantity columns)
upload_excel_pick_list("your_order_file.xlsx")

# This will automatically:
# - Read your order_number, items, quantity columns
# - Generate item codes: A060JR7_ITEM_001, A060JR7_ITEM_002, etc.
# - Assign pick sequence: 1, 2, 3, 4, 5
# - Save to dispatch.pick_lists table
```

---

## 📊 **Excel File Format**

Your Excel file should match this exact format (as shown in your image):

| order_number | items   | quantity |
|--------------|---------|----------|
| A060JR7      | item 1  | 3        |
| A060JR7      | item 2  | 2        |
| A060JR7      | item 3  | 6        |
| A060JR7      | item 4  | 1        |
| A060JR7      | item 5  | 1        |

**Column Requirements:**
- **Column A: `order_number`** - Must match the order IDs in your delivery data (e.g., A060JR7)
- **Column B: `items`** - Description of what needs to be picked (e.g., "item 1", "item 2")  
- **Column C: `quantity`** - How many of each item to pick (e.g., 3, 2, 6)

**Auto-Generated Fields:**
- `item_code` - Automatically created (A060JR7_ITEM_001, A060JR7_ITEM_002, etc.)
- `pick_sequence` - Automatically assigned (1, 2, 3, 4, 5)
- `pick_location` - Optional field (can be added later)

---

## 📱 **Barcode Scanning Workflow**

### What happens when someone scans a barcode:

1. **Scanner reads the barcode** → Gets order ID (e.g., "A060JR7")
2. **System records the scan** → Saves who, when, where
3. **System retrieves pick list** → Shows all items for that order
4. **Worker picks items** → Marks them as completed
5. **Order status updates** → Tracks progress

### Example Code:
```python
from supabase_config import record_barcode_scan, get_pick_list_for_order

# When barcode is scanned
order_id = "A060JR7"  # From barcode scanner

# Record the scan
record_barcode_scan(
    order_id=order_id,
    scanned_by="Worker Name",
    scanner_device="Scanner001",
    location="Warehouse A"
)

# Get what needs to be picked
pick_list = get_pick_list_for_order(order_id)

# Show pick list to worker - using your exact data format
for item in pick_list:
    print(f"Pick {item['quantity_required']} of {item['item_description']}")
    print(f"Item Code: {item['item_code']}")
    print(f"Sequence: {item['pick_sequence']}")
    print("---")

# Example output for A060JR7:
# Pick 3 of item 1
# Item Code: A060JR7_ITEM_001
# Sequence: 1
# ---
# Pick 2 of item 2  
# Item Code: A060JR7_ITEM_002
# Sequence: 2
# ---
# Pick 6 of item 3
# Item Code: A060JR7_ITEM_003
# Sequence: 3
# ---
```

---

## 🔑 **Key Points**

### **About Barcodes:**
- ✅ Barcodes are saved as **TEXT/NUMBERS** in the database
- ✅ The barcode contains the order ID (e.g., "A060JR7")
- ✅ When scanned, it returns the order ID as text
- ❌ We don't save barcode images - just the data inside them

### **About Pick Lists:**
- ✅ Pick lists are uploaded separately from Excel
- ✅ They're linked to barcodes by `order_id`
- ✅ Each order can have multiple items to pick
- ✅ Items are picked in sequence order

### **About Scanning:**
- ✅ Scanning records who, when, where
- ✅ Immediately shows pick list for that order
- ✅ Tracks picking progress
- ✅ Updates order status automatically

---

## 🛠️ **Available Functions**

### Barcode Functions:
- `save_generated_barcodes()` - Save barcodes to database
- `get_barcode_info()` - Get info about a barcode
- `update_barcode_status()` - Update barcode status
- `record_barcode_scan()` - Record when scanned

### Pick List Functions:
- `upload_pick_list_from_excel()` - Upload from Excel
- `get_pick_list_for_order()` - Get items to pick
- `update_pick_item_status()` - Mark items as picked
- `get_barcode_scan_history()` - See scan history

---

## 🔍 **Testing Your Setup**

### 1. **Test Barcode Generation**
```python
python main.py
# Check if barcodes appear in Supabase
```

### 2. **Test Excel Upload**
```python
python excel_upload_example.py
# Check if pick lists appear in Supabase
```

### 3. **Test Scanning Simulation**
```python
from supabase_config import record_barcode_scan, get_pick_list_for_order

# Simulate scanning
record_barcode_scan("A060JR7", "Test User")
pick_list = get_pick_list_for_order("A060JR7")
print(pick_list)
```

---

## 📞 **Support**

If you have issues:
1. Check your `.env` file has correct Supabase credentials
2. Verify all database tables are created
3. Make sure Excel file has correct column names
4. Check that order IDs match between delivery data and pick lists

---

## 🎯 **Summary**

This system creates a complete picking workflow:
1. **Transport sorter** generates barcodes → saves to Supabase
2. **Excel upload** provides pick lists → links to barcodes
3. **Barcode scanning** triggers picking → shows what to pick
4. **Progress tracking** monitors completion → updates status

The key insight: **Barcodes store order IDs as text, not images!** When scanned, they return the order number, which is used to look up the pick list in your database. 