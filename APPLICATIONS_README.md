# Transport Sorter - Separate Applications

This project has been split into **two standalone applications** for different departments:

## 📋 Applications Overview

### 1. **OptimoRoute Sorter** (`optimoroute_sorter_app.py`)
**Purpose**: Delivery processing and PDF sorting
**For**: Delivery scheduling and route management teams

**Features**:
- ✅ Fetch scheduled deliveries from OptimoRoute API
- ✅ Filter by date range and driver
- ✅ Load delivery sequence data
- ✅ Process delivery PDFs and group by driver
- ✅ Generate driver-specific PDF files in delivery sequence order

### 2. **Dispatch Scanning** (`dispatch_scanning_app.py`)
**Purpose**: Picking dockets and store order processing
**For**: Warehouse/dispatch teams

**Features**:
- ✅ Process picking docket PDFs with **REVERSED page order**
- ✅ Generate barcodes for each order
- ✅ Upload store order Excel files to Supabase
- ✅ Create picking PDFs optimized for pallet loading

---

## 🚀 How to Run

### Method 1: Double-click batch files (Windows)
```
Double-click: run_optimoroute_sorter.bat
Double-click: run_dispatch_scanning.bat
```

### Method 2: Command line
```bash
# OptimoRoute Sorter
python optimoroute_sorter_app.py

# Dispatch Scanning
python dispatch_scanning_app.py
```

---

## 📖 Usage Instructions

### **OptimoRoute Sorter Workflow**:
1. **Setup**: Select output directory
2. **Date & Driver**: Choose date range and optional driver filter
3. **Fetch & Load**: Click "🔄 Fetch & Load Scheduled Deliveries"
4. **Add PDFs**: Select delivery PDF files to process
5. **Process**: Click "Process Delivery PDFs"

### **Dispatch Scanning Workflow**:
1. **Important**: Run OptimoRoute Sorter first to load delivery data
2. **Picking PDFs**: Add picking docket PDF files
3. **Process**: Click "Process Picking Dockets (Reversed)"
4. **Store Orders**: Upload Excel files to Supabase (optional)

---

## 🔗 Data Flow Between Applications

```
OptimoRoute Sorter → delivery_sequence_data.json → Dispatch Scanning
```

**Important**: The Dispatch Scanning app requires delivery sequence data from OptimoRoute Sorter to match orders with drivers.

---

## 📁 File Structure

```
transport-sorter/
├── optimoroute_sorter_app.py      # App 1: Delivery processing
├── dispatch_scanning_app.py       # App 2: Picking & store orders
├── run_optimoroute_sorter.bat     # Launcher for App 1
├── run_dispatch_scanning.bat      # Launcher for App 2
├── delivery_sequence_data.json    # Shared data file
├── requirements.txt               # Dependencies
├── supabase_config.py            # Database config
└── APPLICATIONS_README.md         # This file
```

---

## 📋 Dependencies

Both applications require the same dependencies:
```bash
pip install -r requirements.txt
```

**Required packages**:
- PySide6 (GUI)
- pandas (Data processing)
- requests (API calls)
- PyMuPDF (PDF processing)
- pytesseract (OCR)
- reportlab (PDF generation)
- python-barcode (Barcode generation)

---

## 🎯 Department Usage

### **Delivery Team** → Use `OptimoRoute Sorter`
- Fetch delivery schedules
- Process delivery PDFs
- Generate driver route files

### **Warehouse Team** → Use `Dispatch Scanning`
- Process picking dockets (reversed order)
- Upload store orders
- Generate barcoded picking lists

---

## 🔧 Configuration

### API Key (OptimoRoute Sorter)
The OptimoRoute API key is embedded in the code. To change it:
```python
# In optimoroute_sorter_app.py, line ~XXX
self.api_key = "your-new-api-key-here"
```

### Supabase (Dispatch Scanning)
Ensure `supabase_config.py` is properly configured for database uploads.

---

## 🆘 Troubleshooting

### Common Issues:

**1. "No delivery sequence data" in Dispatch Scanning**
- **Solution**: Run OptimoRoute Sorter first and fetch delivery data

**2. "Supabase not available" error**
- **Solution**: Check that `supabase_config.py` exists and is properly configured

**3. "No matching orders found"**
- **Solution**: Ensure order IDs in PDFs match those in delivery data

**4. API connection failed**
- **Solution**: Check internet connection and API key validity

---

## 📞 Support

Each application has built-in status messages and error dialogs to help troubleshoot issues. Check the status bar at the bottom of each application for detailed information.

---

**Last Updated**: December 2024
**Version**: 2.0 (Separated Applications) 