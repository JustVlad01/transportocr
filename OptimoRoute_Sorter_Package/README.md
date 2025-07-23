# OptimoRoute Sorter - Installation Guide

## Quick Start

### Option 1: Using the Batch File (Windows - Recommended)
1. Double-click `install_and_run.bat`
2. The script will automatically install Python dependencies and launch the application

### Option 2: Manual Installation

#### Prerequisites
- Python 3.8 or higher
- Tesseract OCR (for PDF text extraction)

#### Step 1: Install Python Dependencies
```bash
pip install -r requirements.txt
```

#### Step 2: Install Tesseract OCR

**Windows:**
1. Download Tesseract installer from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install to default location (usually C:\Program Files\Tesseract-OCR)
3. Add to PATH or update the pytesseract path in the code if needed

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

#### Step 3: Run the Application
```bash
python optimoroute_sorter_app.py
```

## Application Features

### 1. OptimoRoute API Integration
- Fetches scheduled deliveries from your OptimoRoute account
- Automatically loads delivery sequence data
- Filter by driver and date range

### 2. PDF Processing
- Processes delivery PDF files
- Extracts order IDs using text recognition and OCR
- Groups pages by driver according to delivery sequence
- Creates separate PDF files for each driver

### 3. Smart Order Matching
- Matches PDF content with delivery data
- Case-insensitive order ID matching
- Handles both text-based and image-based PDFs

## Usage Instructions

### Setup Phase
1. **Set Output Directory**: Choose where processed PDF files will be saved
2. **Configure Date Range**: Set the date range for fetching scheduled deliveries
3. **Driver Filter** (Optional): Filter deliveries for specific drivers

### Data Loading
1. Click "Fetch & Load Scheduled Deliveries"
2. The app will connect to OptimoRoute API and load your scheduled orders
3. Review the data in the preview table

### PDF Processing
1. Click "Add PDFs" to select delivery PDF files
2. Click "Process Delivery PDFs" to start processing
3. The app will:
   - Scan each PDF page for order IDs
   - Match order IDs with delivery data
   - Group pages by assigned driver
   - Create separate PDF files for each driver

### Results
- Driver-specific PDF files are created in your output directory
- A processing summary is generated
- Detailed results dialog shows what was processed

## Troubleshooting

### Common Issues

**"No matching orders found"**
- Ensure PDF files contain order IDs that match your delivery data
- Check that order IDs in PDFs match the format in OptimoRoute
- Try different PDF files or check OCR quality

**API Connection Issues**
- Verify your OptimoRoute API key is correct
- Check your internet connection
- Ensure your OptimoRoute account has API access

**OCR Issues**
- Install Tesseract OCR properly
- Ensure PDF quality is good enough for text recognition
- Check that Tesseract is in your system PATH

### Getting Help
1. Check the application status messages for detailed information
2. Review the processing summary file created in your output directory
3. Ensure all dependencies are properly installed

## File Structure
```
OptimoRoute_Sorter_Package/
├── optimoroute_sorter_app.py     # Main application
├── requirements.txt              # Python dependencies
├── install_and_run.bat          # Windows quick-start script
├── run_app.bat                  # Windows run script
├── README.md                    # This file
└── delivery_sequence_data.json  # Sample/existing delivery data
```

## API Configuration
The application uses OptimoRoute API to fetch scheduled deliveries. The API key is embedded in the application. If you need to change it, edit the `api_key` variable in `optimoroute_sorter_app.py`.

## Support
For technical support or questions about this application, contact your system administrator.
