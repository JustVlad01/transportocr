-- Supabase Database Schema for Transport Sorter
-- All tables will be created in the 'dispatch' schema

-- Create the dispatch schema if it doesn't exist
CREATE SCHEMA IF NOT EXISTS dispatch;

-- Table 1: Generated Barcodes
CREATE TABLE dispatch.generated_barcodes (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    order_id VARCHAR(50) NOT NULL UNIQUE,
    barcode_type VARCHAR(20) DEFAULT 'Code128',
    driver_number VARCHAR(20),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    pdf_file_name VARCHAR(255),
    page_number INTEGER,
    status VARCHAR(20) DEFAULT 'generated' -- generated, scanned, picked, completed
);

-- Table 2: Order Details (from Excel uploads)
CREATE TABLE dispatch.order_details (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    order_id VARCHAR(50) NOT NULL,
    item_description TEXT,
    quantity INTEGER,
    location VARCHAR(100),
    pick_sequence INTEGER,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    excel_file_name VARCHAR(255),
    
    -- Foreign key to barcodes table
    FOREIGN KEY (order_id) REFERENCES dispatch.generated_barcodes(order_id)
);

-- Table 3: Scan History
CREATE TABLE dispatch.scan_history (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    order_id VARCHAR(50) NOT NULL,
    scanned_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    scanned_by VARCHAR(100),
    scanner_device VARCHAR(100),
    location VARCHAR(100),
    
    -- Foreign key to barcodes table
    FOREIGN KEY (order_id) REFERENCES dispatch.generated_barcodes(order_id)
);

-- Table 4: Pick Lists (what needs to be picked for each order)
CREATE TABLE dispatch.pick_lists (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    order_id VARCHAR(50) NOT NULL,
    item_code VARCHAR(50) DEFAULT NULL, -- Optional, auto-generated if not provided
    item_description TEXT NOT NULL,
    quantity_required INTEGER NOT NULL,
    quantity_picked INTEGER DEFAULT 0,
    pick_location VARCHAR(100) DEFAULT NULL, -- Optional
    pick_sequence INTEGER DEFAULT NULL, -- Auto-assigned if not provided
    picked_at TIMESTAMP WITH TIME ZONE,
    picked_by VARCHAR(100),
    status VARCHAR(20) DEFAULT 'pending', -- pending, picked, completed
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    excel_file_name VARCHAR(255),
    
    -- Foreign key to barcodes table
    FOREIGN KEY (order_id) REFERENCES dispatch.generated_barcodes(order_id)
);

-- Create indexes for performance
CREATE INDEX idx_dispatch_generated_barcodes_order_id ON dispatch.generated_barcodes(order_id);
CREATE INDEX idx_dispatch_order_details_order_id ON dispatch.order_details(order_id);
CREATE INDEX idx_dispatch_scan_history_order_id ON dispatch.scan_history(order_id);
CREATE INDEX idx_dispatch_pick_lists_order_id ON dispatch.pick_lists(order_id);
CREATE INDEX idx_dispatch_pick_lists_status ON dispatch.pick_lists(status);

-- Table 5: Dispatch Orders (Store order uploads with Excel row order preserved)
CREATE TABLE dispatch_orders (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    ordernumber VARCHAR(50) NOT NULL,
    itemcode VARCHAR(50) NOT NULL,
    product_description TEXT,
    barcode VARCHAR(100),
    customer_type VARCHAR(50),
    quantity INT2,
    quantity_picked INT2,
    error_counter INT2,
    excel_row_sequence INT4 NOT NULL,  -- Preserves Excel file row order for picking sequence
    order_start_time TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    order_finish_time TIMESTAMP WITH TIME ZONE,
    order_duration INTERVAL,
    picker_name VARCHAR(50),
    scanned_by VARCHAR(50),
    full_or_partial_picking BOOLEAN,
    item_skipped BOOLEAN,
    delivery_date DATE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create indexes for performance
CREATE INDEX idx_dispatch_orders_ordernumber ON dispatch_orders(ordernumber);
CREATE INDEX idx_dispatch_orders_itemcode ON dispatch_orders(itemcode);
CREATE INDEX idx_dispatch_orders_barcode ON dispatch_orders(barcode);
CREATE INDEX idx_dispatch_orders_excel_sequence ON dispatch_orders(excel_row_sequence);

-- Table 6: Crate Verification (for tracking order verification data)
CREATE TABLE crate_verification (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    ordernumber VARCHAR(50) NOT NULL,
    sitename VARCHAR(100),
    routenumber VARCHAR(50),
    total_items INTEGER DEFAULT 0,
    total_crates INTEGER DEFAULT 0,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create indexes for performance
CREATE INDEX idx_crate_verification_ordernumber ON crate_verification(ordernumber);
CREATE INDEX idx_crate_verification_sitename ON crate_verification(sitename);
CREATE INDEX idx_crate_verification_routenumber ON crate_verification(routenumber);

-- Row Level Security (RLS) policies can be added here if needed
ALTER TABLE dispatch.generated_barcodes ENABLE ROW LEVEL SECURITY;
ALTER TABLE dispatch.order_details ENABLE ROW LEVEL SECURITY;
ALTER TABLE dispatch.scan_history ENABLE ROW LEVEL SECURITY;
ALTER TABLE dispatch.pick_lists ENABLE ROW LEVEL SECURITY;
ALTER TABLE dispatch_orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE crate_verification ENABLE ROW LEVEL SECURITY;
