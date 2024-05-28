import os
import pandas as pd
import pdfplumber  # New library for better PDF parsing
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)

def parse_files_in_folder(folder_path, output_excel_path):
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_excel_path)
    os.makedirs(output_dir, exist_ok=True)
    
    # Create a Pandas Excel writer using openpyxl as the engine.
    writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')

    try:
        # Loop through all files in the folder
        for filename in os.listdir(folder_path):
            if filename.endswith('-read'):  # Skip files ending with '-read'
                continue
            file_path = os.path.join(folder_path, filename)
            
            # Check if the file is a PDF, TXT, or CSV
            if filename.endswith('.pdf'):
                # Parse PDF
                data = parse_pdf(file_path)
            elif filename.endswith('.txt'):
                # Parse TXT
                data = parse_txt(file_path)
            elif filename.endswith('.csv'):
                # Read CSV
                data = pd.read_csv(file_path)
            else:
                # Skip if not a supported file type
                continue

            if data.empty:
                logging.warning(f"No data extracted from file: {filename}")
                continue

            data['Category'] = 'Category_Value'

            # Write to Excel with sheet name as filename
            sheet_name = os.path.splitext(filename)[0][:31]  # Use filename without extension as sheet name, limit to 31 chars
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            
            os.rename(file_path, os.path.join(folder_path, filename + '-read'))

        # Save the Excel file
        writer._save()
    except Exception as e:
        logging.error(f"An error occurred while processing files, All the files seems to have been read before: {e}")


def parse_pdf(file_path):
    rows = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                rows.extend(parse_pdf_text(text))
    
    if rows:
        data = pd.DataFrame(rows, columns=['Date', 'Description', 'Currency', 'Amount', 'Fees', 'Total'])
    else:
        data = pd.DataFrame(columns=['Date', 'Description', 'Currency', 'Amount', 'Fees', 'Total'])

    return data

def parse_pdf_text(text):
    rows = []
    account_activity_pattern = re.compile(
        r'(\d{2}/\d{2}/\d{4})'  # Date
        r'\s+(.*?)'  # Description
        r'\s+([A-Z]{3})'  # Currency
        r'(-?\d+\.\d{2})'  # Amount
        r'\s+USDID:\s+.*?'  # Skip the ID part
        r'USD(-?\d+\.\d{2})'  # Fees
        r'(-?\d+\.\d{2})'  # Total
    )

    lines = text.split('\n')
    for line in lines:
        account_activity_match = account_activity_pattern.search(line)
        if account_activity_match:
            rows.append(list(account_activity_match.groups()))

    return rows

def parse_txt(file_path):
    # Read the TXT file
    with open(file_path, 'r') as file:
        lines = file.readlines()
    
    rows = []
    for line in lines:
        if line.strip():
            # Using regex to extract date, description, and amount
            match = re.match(r'(\d{4}-\d{2}-\d{2})\s+(.*?)\s+(-?\d+\.\d{2})$', line.strip())
            if match:
                rows.append(list(match.groups()))

    data = pd.DataFrame(rows, columns=['Date', 'Description', 'Amount'])
    return data

# Example usage
folder_path = 'client_docs'  # Use relative path
output_excel_path = 'processed_files/csv_extracted.xlsx'  # Use relative path
parse_files_in_folder(folder_path, output_excel_path)
