import os
import pandas as pd
import fitz  

def parse_files_in_folder(folder_path, output_excel_path):
    output_dir = os.path.dirname(output_excel_path)
    os.makedirs(output_dir, exist_ok=True)
    
    writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
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
            continue

        data['Category'] = 'Category_Value'

        sheet_name = os.path.splitext(filename)[0][:31]  
        data.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save the Excel file
    writer._save()

def parse_pdf(file_path):
    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text()
    
    rows = [line.split(maxsplit=2) for line in text.split('\n') if line.strip()]
    data = pd.DataFrame(rows, columns=['Date', 'Description', 'Amount'])
    return data

def parse_txt(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()
    
    rows = [line.split(maxsplit=2) for line in lines if line.strip()]
    data = pd.DataFrame(rows, columns=['Date', 'Description', 'Amount'])
    return data

# Example usage
folder_path = 'client_docs' 
output_excel_path = 'processed_files/test.xlsx' 
parse_files_in_folder(folder_path, output_excel_path)
