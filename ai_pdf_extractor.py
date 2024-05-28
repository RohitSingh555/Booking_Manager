import pdfplumber
import openpyxl
import os
import logging
import ollama
import json
import re

# Setup basic configuration for logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to extract text from PDF using pdfplumber
def extract_text_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page_number, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text()
                logging.info("Extracted text from page %d of PDF: %s", page_number, page_text)
                text += page_text + "\n"
            return text
    except Exception as e:
        logging.error("Failed to process PDF %s: %s", pdf_path, str(e))
        return ""

# Function to preprocess extracted text
def preprocess_text(text):
    lines = text.split('\n')
    transactions = []
    for line in lines:
        # Filter out lines that are not transactions
        if re.search(r'\d{2}/\d{2}/\d{4}', line) and re.search(r'USD', line):
            transactions.append(line)
    return "\n".join(transactions)

# Function to process PayPal data using Ollama AI
def process_paypal_data_with_ollama(text):
    data = []
    try:
        prompt = (
            "Extract and parse PayPal transactions from the following text and provide a JSON with Date, Description, Amount, "
            "and Category fields:\nJSON should look like this:\n"
            "{\n"
            '    "Date": "MM-dd-yyyy",\n'
            '    "Description": "Product or service bought, fetch the description of it",\n'
            '    "Amount": "Amount took to buy that resource.",\n'
            '    "Category": "Categorize if it is PayPal, for food, for what kind of product it is, categorize it"\n'
            "}\n"
            f"{text}"
        )
        response = ollama.generate(
            model='gemma:2b',
            prompt=prompt
        )
        parsed_data = response['response']
        if not parsed_data:
            raise ValueError("Empty response from Ollama AI")
        transactions = json.loads(parsed_data)
        for transaction in transactions:
            data.append((transaction['Date'], transaction['Description'], transaction['Amount'], transaction['Category']))
    except Exception as e:
        logging.error("Failed to process PayPal data with Ollama AI: %s", str(e))
    return data

# Function to process eBay data using Ollama AI
def process_ebay_data_with_ollama(text):
    data = []
    try:
        prompt = (
            "Extract and parse eBay transactions from the following text and provide a JSON with Date, Description, Amount, "
            "and Category fields:\nJSON should look like this:\n"
            "{\n"
            '    "Date": "MM-dd-yyyy",\n'
            '    "Description": "Product or service bought, fetch the description of it",\n'
            '    "Amount": "Amount took to buy that resource.",\n'
            '    "Category": "Categorize if it is eBay, for food, for what kind of product it is, categorize it"\n'
            "}\n"
            f"{text}"
        )
        response = ollama.generate(
            model='gemma:2b',
            prompt=prompt
        )
        parsed_data = response['response']
        if not parsed_data:
            raise ValueError("Empty response from Ollama AI")
        transactions = json.loads(parsed_data)
        for transaction in transactions:
            data.append((transaction['Date'], transaction['Description'], transaction['Amount'], transaction['Category']))
    except Exception as e:
        logging.error("Failed to process eBay data with Ollama AI: %s", str(e))
    return data

# Function to save data to Excel
def save_to_excel(data, file_name, sheet_name):
    try:
        if os.path.exists(file_name):
            wb = openpyxl.load_workbook(file_name)
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Remove the default sheet

        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(title=sheet_name)
            ws.append(("Date", "Description", "Amount", "Category"))  # Only add headers if sheet is newly created
        else:
            ws = wb[sheet_name]

        for row in data:
            ws.append(row)

        # Remove duplicate rows
        remove_duplicate_rows(ws)

        wb.save(file_name)
        logging.info("Data written to %s in %s", sheet_name, file_name)
    except Exception as e:
        logging.error("Failed to save data to Excel %s: %s", file_name, str(e))

def remove_duplicate_rows(sheet):
    rows = [tuple(cell.value for cell in row) for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=sheet.max_column)]
    unique_rows = set(rows)
    sheet.delete_rows(2, sheet.max_row)
    for row in unique_rows:
        sheet.append(row)

def identify_and_process_pdfs(directory, output_excel):
    pdf_files = [f for f in os.listdir(directory) if f.endswith('.pdf')]
    logging.info("Found %d PDF files in the directory.", len(pdf_files))
    
    for filename in pdf_files:
        pdf_path = os.path.join(directory, filename)
        logging.info("Processing PDF: %s", pdf_path)
        
        text = extract_text_from_pdf(pdf_path)
        text = preprocess_text(text)
        
        if "ebay" in filename.lower():
            logging.info("Processing eBay PDF: %s", filename)
            data = process_ebay_data_with_ollama(text)
            save_to_excel(data, output_excel, "eBay")
        elif filename.endswith("PDF"):
            logging.info("Skipping PDF: %s", filename)
        else:
            logging.info("Processing PayPal PDF: %s", filename)
            data = process_paypal_data_with_ollama(text)
            if filename.lower() == "creed.pdf":
                print("Extracted data from creed.pdf:")
            if not data:
                logging.warning("No data extracted from PayPal PDF: %s", filename)
            else:
                save_to_excel(data, output_excel, "PayPal")

def main():
    directory_path = "client_docs"
    output_excel = "processed_files/AI_pdf_output_data.xlsx"
    if not os.path.exists(directory_path):
        logging.warning("Directory %s does not exist. Please check the path.", directory_path)
        return
    identify_and_process_pdfs(directory_path, output_excel)

if __name__ == "__main__":
    main()
