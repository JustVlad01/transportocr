# Transport Sorter - PDF Scanner

A Python application for scanning PDF documents using OCR (Optical Character Recognition) with a visual interface. The application can read Excel and CSV files, extract data from the first column, and process PDF files using Tesseract OCR and PyMuPDF.

## Features

- **Visual Interface**: User-friendly GUI built with tkinter
- **Multiple File Formats**: Read and extract values from Excel (.xlsx, .xls) and CSV (.csv) files
- **JSON Storage**: Store extracted data as JSON files
- **PDF Processing**: Scan single or multiple PDF files
- **OCR Capabilities**: Extract text from PDFs using Tesseract OCR
- **Data Management**: Automatically save and load processed data

## Requirements

- Python 3.7+
- Tesseract OCR installed on your system

## Installation

### 1. Install Tesseract OCR

#### Ubuntu/Debian:
```bash
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-eng
```

#### Windows:
Download and install from: https://github.com/UB-Mannheim/tesseract/wiki

#### macOS:
```bash
brew install tesseract
```

### 2. Install Python Dependencies

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

### 1. Run the Application

```bash
# Using the convenience script
./run.sh

# Or manually
source venv/bin/activate
python3 main.py
```

### 2. Load Data File

1. Click **"Browse File"** to select your Excel (.xlsx, .xls) or CSV (.csv) file
2. Click **"Load & Save Data"** to extract first column values and save as JSON
3. The data will be displayed in the preview table

### 3. Process PDF Files

1. After loading data, click **"Browse PDFs"** to select one or more PDF files
2. Click **"Process PDFs with OCR"** to scan the documents
3. Results will be saved in the `ocr_results/` directory

## File Structure

```
transport-sorter/
├── main.py                 # Main application file
├── requirements.txt        # Python dependencies
├── setup.py               # Setup and system check script
├── run.sh                 # Convenience script to run the app
├── README.md              # This file
├── venv/                  # Virtual environment (created after setup)
├── transport_data.json    # Stored data (created after first use)
└── ocr_results/           # OCR processing results (created after first use)
    └── ocr_results_*.json # Timestamped result files
```

## Output Format

### Data Storage (transport_data.json)
```json
{
  "source_file": "/path/to/data/file.xlsx",
  "file_type": "Excel",
  "column_a_values": ["value1", "value2", "value3"],
  "total_records": 3,
  "created_date": "2024-01-01T12:00:00"
}
```

### OCR Results (ocr_results/ocr_results_*.json)
```json
{
  "reference_data": ["value1", "value2", "value3"],
  "processed_files": [
    {
      "file_name": "document.pdf",
      "file_path": "/path/to/document.pdf",
      "extracted_text": [
        {
          "page": 1,
          "text": "Extracted text from page 1"
        }
      ],
      "page_count": 1,
      "processing_date": "2024-01-01T12:00:00"
    }
  ],
  "total_files": 1
}
```

## Supported File Formats

### Input Data Files
- **Excel**: .xlsx, .xls
- **CSV**: .csv (comma-separated values)

The application will automatically detect the file format and process accordingly. It extracts values from the first column of the file.

### PDF Files
- All standard PDF formats
- Both text-based and image-based PDFs (OCR will be applied to image-based content)

## Troubleshooting

### Common Issues

1. **Tesseract not found**: Make sure Tesseract OCR is installed and in your PATH
2. **Import errors**: Install all dependencies using `pip install -r requirements.txt`
3. **Data file errors**: Ensure your Excel/CSV file has data in the first column
4. **PDF processing fails**: Check that PDF files are not corrupted and accessible

### Getting Help

If you encounter issues:
1. Check that all dependencies are installed: `python3 setup.py`
2. Verify Tesseract OCR is working: `tesseract --version`
3. Ensure data files are readable and not password-protected
4. Check that the virtual environment is activated: `source venv/bin/activate`

## Dependencies

- `pandas>=1.5.0` - Excel and CSV file processing
- `PyMuPDF>=1.23.0` - PDF manipulation
- `pytesseract>=0.3.10` - Tesseract OCR interface
- `Pillow>=9.0.0` - Image processing
- `openpyxl>=3.1.0` - Excel file reading

## License

This project is open source and available under the MIT License. 