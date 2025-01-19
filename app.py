import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter

class ExpenseData:
    def __init__(self):
        self.file_name = None
        self.legal_entity = None
        self.currency = None
        self.amount_due_employee = None
        self.amount_due_company_card = None
        self.total_paid_by_company = None

def clean_amount(amount_str):
    """Convert string amount to float, handling different formats"""
    if not amount_str:
        return None
        
    # Remove any currency codes and symbols
    amount_str = re.sub(r'(?:USD|EUR|GBP|[€$£])\s*', '', amount_str.strip())
    
    # Handle Spanish format (1.234,56) and English format (1,234.56)
    if ',' in amount_str and '.' in amount_str:
        if amount_str.find('.') < amount_str.find(','):
            # Spanish format
            amount_str = amount_str.replace('.', '').replace(',', '.')
        else:
            # English format
            amount_str = amount_str.replace(',', '')
    elif ',' in amount_str:
        amount_str = amount_str.replace(',', '')
            
    try:
        return float(amount_str)
    except ValueError:
        return None

def extract_expense_data(text):
    """Extract specified expense report fields from text"""
    data = ExpenseData()
    
    # Extract legal entity with precise pattern
    legal_entity_match = re.search(r'Custom 10-Legal Entity\s*:\s*(\S+)', text)
    if legal_entity_match:
        data.legal_entity = legal_entity_match.group(1).strip()
        # Clean any remaining artifacts or extra spaces
        data.legal_entity = re.sub(r'[^\w$]', '', data.legal_entity)
    
    # Extract currency
    currency_match = re.search(r'Currency\s*:\s*([^\n]+)', text)
    if currency_match:
        data.currency = currency_match.group(1).strip()
    
    # Extract amounts with exact patterns
    amount_patterns = {
        'amount_due_employee': r'Amount Due Employee\s*:\s*(?:USD|EUR|€)?\s*([\d,.]+)',
        'amount_due_company_card': r'Amount Due Company Card\s*:\s*(?:USD|EUR|€)?\s*([\d,.]+)',
        'total_paid_by_company': r'Total Paid By Company\s*:\s*(?:USD|EUR|€)?\s*([\d,.]+)'
    }
    
    for field, pattern in amount_patterns.items():
        match = re.search(pattern, text)
        if match:
            amount_str = match.group(1)
            if amount_str:
                setattr(data, field, clean_amount(amount_str))
    
    return data

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Expense Report Processor")
        self.root.geometry("800x600")
        
        # Main container with padding
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(self.main_frame, 
                              text="PDF Expense Report Processor",
                              font=('Helvetica', 16, 'bold'))
        title_label.pack(pady=20)
        
        # Selected files label
        self.files_label = ttk.Label(self.main_frame, text="No files selected")
        self.files_label.pack(pady=10)
        
        # Buttons
        ttk.Button(self.main_frame, text="Select PDFs", 
                  command=self.select_files).pack(pady=5)
        ttk.Button(self.main_frame, text="Process PDFs", 
                  command=self.process_files).pack(pady=5)
        
        # Progress bar and status
        self.progress = ttk.Progressbar(self.main_frame, length=400, mode='determinate')
        self.progress.pack(pady=20)
        
        self.status_label = ttk.Label(self.main_frame, text="Ready")
        self.status_label.pack(pady=5)
        
        # Results text widget
        self.results_frame = ttk.Frame(self.main_frame)
        self.results_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.results_text = tk.Text(self.results_frame, height=10, wrap=tk.WORD)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, 
                                command=self.results_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        # Store selected files
        self.selected_files = []

    def select_files(self):
        """Open file dialog to select PDFs"""
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf")]
        )
        if files:
            self.selected_files = files
            num_files = len(files)
            self.files_label.config(text=f"{num_files} files selected")

    def process_files(self):
        """Process the selected PDF files"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select PDF files first")
            return
        
        # Create output directory
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = os.path.join(os.path.expanduser("~"), "Desktop", 
                                f"Processed_Expenses_{timestamp}")
        os.makedirs(output_dir, exist_ok=True)
        
        total_files = len(self.selected_files)
        self.progress['value'] = 0
        processed_data = []
        
        try:
            for index, pdf_file in enumerate(self.selected_files):
                # Update status
                filename = os.path.basename(pdf_file)
                self.status_label.config(text=f"Processing: {filename}")
                self.root.update()
                
                try:
                    # Create directory for this PDF's pages
                    pdf_dir = os.path.join(output_dir, os.path.splitext(filename)[0])
                    os.makedirs(pdf_dir, exist_ok=True)
                    
                    # Read the PDF
                    pdf_reader = PdfReader(pdf_file)
                    num_pages = len(pdf_reader.pages)
                    combined_text = ""
                    
                    # Extract text and split pages
                    for page_num in range(num_pages):
                        # Extract text
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()
                        combined_text += page_text + "\n"
                        
                        # Save individual page
                        pdf_writer = PdfWriter()
                        pdf_writer.add_page(page)
                        output_path = os.path.join(pdf_dir, f'page_{page_num + 1}.pdf')
                        with open(output_path, 'wb') as output_file:
                            pdf_writer.write(output_file)
                        
                        self.results_text.insert(tk.END, 
                                              f"Extracted page {page_num + 1} from {filename}\n")
                        self.results_text.see(tk.END)
                        self.root.update()
                    
                    # Extract data
                    expense_data = extract_expense_data(combined_text)
                    expense_data.file_name = filename
                    
                    # Add to processed data
                    processed_data.append({
                        'File Name': expense_data.file_name,
                        'Legal Entity': expense_data.legal_entity,
                        'Currency': expense_data.currency,
                        'Amount Due Employee': expense_data.amount_due_employee,
                        'Amount Due Company Card': expense_data.amount_due_company_card,
                        'Total Paid By Company': expense_data.total_paid_by_company
                    })
                    
                    # Update progress
                    self.progress['value'] = (index + 1) / total_files * 100
                    self.root.update()
                    
                    # Log success
                    self.results_text.insert(tk.END, 
                                          f"Successfully processed: {filename} - {num_pages} pages\n")
                    
                except Exception as e:
                    error_msg = f"Error processing {filename}: {str(e)}\n"
                    self.results_text.insert(tk.END, error_msg)
                    print(error_msg)
            
            # Create CSV with extracted data
            if processed_data:
                import pandas as pd
                df = pd.DataFrame(processed_data)
                csv_path = os.path.join(output_dir, 'expense_data.csv')
                df.to_csv(csv_path, index=False)
                
                self.results_text.insert(tk.END, f"\nProcessing complete!\n")
                self.results_text.insert(tk.END, f"Data exported to: {csv_path}\n")

            self.status_label.config(text="Processing complete!")
            messagebox.showinfo("Complete", f"Files processed and saved to:\n{output_dir}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        self.progress['value'] = 0

def main():
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()