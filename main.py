import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from pathlib import Path
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

class TransportSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Transport Sorter - PDF Scanner")
        self.root.geometry("800x600")
        
        # Application data
        self.excel_data = []
        self.json_file = "transport_data.json"
        
        # Setup the UI
        self.setup_ui()
        
        # Load existing data if available
        self.load_existing_data()
    
    def setup_ui(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Transport Sorter - PDF Scanner", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Excel File Section
        excel_frame = ttk.LabelFrame(main_frame, text="Step 1: Load Excel Data", padding="10")
        excel_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        excel_frame.columnconfigure(1, weight=1)
        
        ttk.Label(excel_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.excel_file_var = tk.StringVar()
        self.excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_var, state="readonly")
        self.excel_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        self.browse_excel_btn = ttk.Button(excel_frame, text="Browse Excel", 
                                          command=self.browse_excel_file,
                                          cursor="hand2")
        self.browse_excel_btn.grid(row=0, column=2)
        
        self.load_excel_btn = ttk.Button(excel_frame, text="Load & Save Data", 
                                        command=self.load_excel_data,
                                        cursor="hand2")
        self.load_excel_btn.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        
        # Data Display Section
        data_frame = ttk.LabelFrame(main_frame, text="Loaded Data Preview", padding="10")
        data_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
        # Treeview for data display
        self.tree = ttk.Treeview(data_frame, columns=("Value",), show="tree headings", height=8)
        self.tree.heading("#0", text="Index")
        self.tree.heading("Value", text="Column A Values")
        self.tree.column("#0", width=80)
        self.tree.column("Value", width=200)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # PDF Processing Section
        pdf_frame = ttk.LabelFrame(main_frame, text="Step 2: Process PDF Files", padding="10")
        pdf_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        pdf_frame.columnconfigure(1, weight=1)
        
        ttk.Label(pdf_frame, text="PDF Files:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.pdf_files_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_files_var, state="readonly")
        self.pdf_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        self.browse_pdf_btn = ttk.Button(pdf_frame, text="Browse PDFs", 
                                        command=self.browse_pdf_files,
                                        cursor="hand2")
        self.browse_pdf_btn.grid(row=0, column=2)
        
        self.process_pdf_btn = ttk.Button(pdf_frame, text="Process PDFs with OCR", 
                                         command=self.process_pdf_files,
                                         cursor="hand2")
        self.process_pdf_btn.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        
        # Status Section
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        status_frame.columnconfigure(0, weight=1)
        
        ttk.Label(status_frame, text="Status:").grid(row=0, column=0, sticky=tk.W)
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                     foreground="green", font=('Arial', 10, 'bold'))
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Progress bar
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        
        # Configure grid weights for main_frame
        main_frame.rowconfigure(2, weight=1)
    
    def browse_excel_file(self):
        """Browse and select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_var.set(file_path)
    
    def load_excel_data(self):
        """Load data from Excel file column A and save as JSON"""
        excel_file = self.excel_file_var.get()
        if not excel_file:
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        
        try:
            self.status_var.set("Loading Excel data...")
            self.progress.start()
            
            # Read Excel file
            df = pd.read_excel(excel_file)
            
            # Get column A values (first column)
            if df.empty:
                messagebox.showerror("Error", "Excel file is empty!")
                return
            
            # Extract column A values, skip empty cells
            column_a_values = df.iloc[:, 0].dropna().astype(str).tolist()
            
            if not column_a_values:
                messagebox.showerror("Error", "No data found in column A!")
                return
            
            self.excel_data = column_a_values
            
            # Save to JSON
            data_to_save = {
                "source_file": excel_file,
                "column_a_values": self.excel_data,
                "total_records": len(self.excel_data),
                "created_date": pd.Timestamp.now().isoformat()
            }
            
            with open(self.json_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, indent=2, ensure_ascii=False)
            
            # Update display
            self.update_data_display()
            
            self.status_var.set(f"Successfully loaded {len(self.excel_data)} records and saved to JSON")
            messagebox.showinfo("Success", f"Data loaded successfully!\n{len(self.excel_data)} records saved to {self.json_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
            self.status_var.set("Error loading Excel file")
        finally:
            self.progress.stop()
    
    def update_data_display(self):
        """Update the treeview with loaded data"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add new items
        for i, value in enumerate(self.excel_data, 1):
            self.tree.insert("", "end", text=str(i), values=(value,))
    
    def load_existing_data(self):
        """Load existing JSON data if available"""
        if os.path.exists(self.json_file):
            try:
                with open(self.json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.excel_data = data.get('column_a_values', [])
                    self.update_data_display()
                    self.status_var.set(f"Loaded existing data: {len(self.excel_data)} records")
            except Exception as e:
                print(f"Error loading existing data: {e}")
    
    def browse_pdf_files(self):
        """Browse and select PDF files"""
        file_paths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_paths:
            self.pdf_files = file_paths
            file_names = [os.path.basename(path) for path in file_paths]
            display_text = f"{len(file_names)} file(s): " + ", ".join(file_names[:3])
            if len(file_names) > 3:
                display_text += f"... and {len(file_names) - 3} more"
            self.pdf_files_var.set(display_text)
    
    def process_pdf_files(self):
        """Process PDF files with OCR"""
        if not hasattr(self, 'pdf_files') or not self.pdf_files:
            messagebox.showerror("Error", "Please select PDF files first!")
            return
        
        if not self.excel_data:
            messagebox.showerror("Error", "Please load Excel data first!")
            return
        
        try:
            self.status_var.set("Processing PDF files with OCR...")
            self.progress.start()
            
            # Create results directory
            results_dir = Path("ocr_results")
            results_dir.mkdir(exist_ok=True)
            
            results = []
            
            for pdf_path in self.pdf_files:
                pdf_name = os.path.basename(pdf_path)
                self.status_var.set(f"Processing {pdf_name}...")
                
                # Process PDF with OCR
                pdf_result = self.extract_text_from_pdf(pdf_path)
                results.append({
                    "file_name": pdf_name,
                    "file_path": pdf_path,
                    "extracted_text": pdf_result["text"],
                    "page_count": pdf_result["page_count"],
                    "processing_date": pd.Timestamp.now().isoformat()
                })
            
            # Save results
            output_file = results_dir / f"ocr_results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump({
                    "reference_data": self.excel_data,
                    "processed_files": results,
                    "total_files": len(results)
                }, f, indent=2, ensure_ascii=False)
            
            self.status_var.set(f"Successfully processed {len(results)} PDF files")
            messagebox.showinfo("Success", 
                              f"PDF processing completed!\n"
                              f"Processed {len(results)} files\n"
                              f"Results saved to: {output_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process PDF files:\n{str(e)}")
            self.status_var.set("Error processing PDF files")
        finally:
            self.progress.stop()
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF using PyMuPDF and OCR"""
        doc = fitz.open(pdf_path)
        extracted_text = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # First try to extract text directly
            text = page.get_text()
            
            # If no text found, use OCR
            if not text.strip():
                # Convert page to image
                mat = fitz.Matrix(2, 2)  # 2x zoom for better OCR quality
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                
                # Use PIL to create image
                image = Image.open(io.BytesIO(img_data))
                
                # Use Tesseract OCR
                try:
                    text = pytesseract.image_to_string(image)
                except Exception as ocr_error:
                    text = f"OCR Error on page {page_num + 1}: {str(ocr_error)}"
            
            extracted_text.append({
                "page": page_num + 1,
                "text": text.strip()
            })
        
        doc.close()
        
        return {
            "text": extracted_text,
            "page_count": len(doc)
        }

def main():
    root = tk.Tk()
    app = TransportSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()