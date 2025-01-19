import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import re
import io
import zipfile
from datetime import datetime

class ExpenseData:
    def __init__(self):
        self.file_name = None
        self.legal_entity = None
        self.currency = None
        self.amount_due_employee = None
        self.amount_due_company_card = None
        self.total_paid_by_company = None

def clean_amount(amount_str):
    """Convert string amount to float, handling different formats including negative values"""
    if not amount_str:
        return None
        
    # Store if the amount is negative (enclosed in parentheses)
    is_negative = amount_str.startswith('(') and amount_str.endswith(')')
    
    # Remove parentheses if present
    if is_negative:
        amount_str = amount_str[1:-1]
    
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
        value = float(amount_str)
        return -abs(value) if (is_negative or amount_str.startswith('-')) else value
    except ValueError:
        return None

def extract_expense_data(text, max_pages_to_search=2):
    """Extract specified expense report fields from text"""
    data = ExpenseData()
    
    # Extract legal entity with precise pattern
    legal_entity_match = re.search(r'Custom 10-Legal Entity\s*:\s*(\S+)', text)
    if legal_entity_match:
        data.legal_entity = legal_entity_match.group(1).strip()
        data.legal_entity = re.sub(r'[^\w$]', '', data.legal_entity)
    
    # Extract currency
    currency_match = re.search(r'Currency\s*:\s*([^\n]+)', text)
    if currency_match:
        data.currency = currency_match.group(1).strip()
    
    # Extract amounts with exact patterns
    amount_patterns = {
        'amount_due_employee': r'Amount Due Employee\s*:\s*(?:USD|EUR|€)?\s*([\d,.()-]+)',
        'amount_due_company_card': r'Amount Due Company Card\s*:\s*(?:USD|EUR|€)?\s*([\d,.()-]+)',
        'total_paid_by_company': r'Total Paid By Company\s*:\s*(?:USD|EUR|€)?\s*([\d,.()-]+)'
    }
    
    for field, pattern in amount_patterns.items():
        match = re.search(pattern, text)
        if match:
            amount_str = match.group(1)
            if amount_str:
                setattr(data, field, clean_amount(amount_str))
    
    return data

def split_pdf_to_pages(pdf_reader):
    """Split PDF into individual pages and return a dict of BytesIO objects"""
    pages = {}
    for page_num in range(len(pdf_reader.pages)):
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[page_num])
        page_buffer = io.BytesIO()
        pdf_writer.write(page_buffer)
        page_buffer.seek(0)
        pages[f'page_{page_num + 1}'] = page_buffer
    return pages

def create_zip_of_pdfs(pdf_pages_dict, original_filename):
    """Create a zip file containing individual PDF pages"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        base_filename = original_filename.rsplit('.', 1)[0]
        for page_num, page_buffer in pdf_pages_dict.items():
            zf.writestr(
                f"{base_filename}/{page_num}.pdf",
                page_buffer.getvalue()
            )
    zip_buffer.seek(0)
    return zip_buffer

def process_pdfs(uploaded_files):
    """Process uploaded PDF files and extract expense data"""
    processed_data = []
    all_pdf_pages = {}
    
    for uploaded_file in uploaded_files:
        try:
            # Read PDF content
            pdf_reader = PdfReader(uploaded_file)
            
            # Get text from first two pages for expense data
            summary_text = ""
            for page_num in range(min(2, len(pdf_reader.pages))):
                summary_text += pdf_reader.pages[page_num].extract_text() + "\n"
            
            # Extract data
            expense_data = extract_expense_data(summary_text)
            expense_data.file_name = uploaded_file.name
            
            # Split PDF into pages
            pdf_pages = split_pdf_to_pages(pdf_reader)
            all_pdf_pages[uploaded_file.name] = pdf_pages
            
            # Add to processed data
            processed_data.append({
                'File Name': expense_data.file_name,
                'Legal Entity': expense_data.legal_entity,
                'Currency': expense_data.currency,
                'Amount Due Employee': expense_data.amount_due_employee,
                'Amount Due Company Card': expense_data.amount_due_company_card,
                'Total Paid By Company': expense_data.total_paid_by_company
            })
            
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")
    
    return processed_data, all_pdf_pages

def main():
    st.set_page_config(page_title="PDF Expense Report Processor", layout="wide")
    
    st.title("PDF Expense Report Processor")
    
    st.write("""
    ### Instructions
    1. Upload one or more PDF expense reports
    2. The app will extract key information and split the PDFs into pages
    3. Download the Excel summary and/or individual processed PDFs
    """)
    
    uploaded_files = st.file_uploader(
        "Choose PDF files", 
        type="pdf", 
        accept_multiple_files=True,
        help="You can select multiple PDF files at once"
    )
    
    if uploaded_files:
        with st.spinner('Processing PDFs...'):
            processed_data, all_pdf_pages = process_pdfs(uploaded_files)
            
            if processed_data:
                df = pd.DataFrame(processed_data)
                
                # Display results
                st.subheader("Extracted Data")
                st.dataframe(df)
                
                # Create columns for download buttons
                col1, col2 = st.columns(2)
                
                # Excel download button
                with col1:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df.to_excel(writer, sheet_name='Expense Data', index=False)
                        workbook = writer.book
                        worksheet = writer.sheets['Expense Data']
                        
                        header_format = workbook.add_format({
                            'bold': True,
                            'text_wrap': True,
                            'valign': 'top',
                            'bg_color': '#D9D9D9',
                            'border': 1
                        })
                        
                        for col_num, value in enumerate(df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                        
                        for idx, col in enumerate(df.columns):
                            series = df[col]
                            max_len = max(
                                series.astype(str).map(len).max(),
                                len(str(series.name))
                            ) + 2
                            worksheet.set_column(idx, idx, max_len)
                    
                    buffer.seek(0)
                    st.download_button(
                        label="📊 Download Excel Summary",
                        data=buffer,
                        file_name=f"expense_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                
                # PDF downloads
                with col2:
                    # Create a zip file containing all PDFs
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    
                    for original_filename, pdf_pages in all_pdf_pages.items():
                        zip_buffer = create_zip_of_pdfs(pdf_pages, original_filename)
                        
                        st.download_button(
                            label=f"📄 Download {original_filename} (Split Pages)",
                            data=zip_buffer,
                            file_name=f"{original_filename.rsplit('.', 1)[0]}_{timestamp}_pages.zip",
                            mime="application/zip"
                        )
                
                st.success(f"Successfully processed {len(processed_data)} files!")
            else:
                st.warning("No data could be extracted from the uploaded files.")

if __name__ == "__main__":
    main()
