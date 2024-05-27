import openpyxl
import json
from groq import Groq
import os
from dotenv import load_dotenv
import requests
import re

# Load environment variables
load_dotenv()
api_key = os.getenv("GROQ_API_KEY")

# Categories
categories = {
    "Income",
    "Expenses",
    "Business Expenses",
    "Tax Deductible Expenses",
    "Subscriptions",
    "Uncertain Expenses"
}

def read_excel_file(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    max_row = sheet.max_row
    max_column = sheet.max_column

    date_column = None
    description_column = None
    amount_column = None

    for col_idx in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=col_idx).value
        if cell_value and any(keyword in cell_value.lower() for keyword in ["date", "total"]):
            date_column = col_idx
        elif cell_value and "description" in cell_value.lower():
            description_column = col_idx
        elif cell_value and "amount" in cell_value.lower():
            amount_column = col_idx

    # Check if all required columns are found
    if date_column is None or description_column is None or amount_column is None:
        raise ValueError("Required columns not found in the Excel file.")

    # Extract data from the identified columns
    data = []
    for row_idx in range(2, max_row + 1):
        entry = {
            "date": sheet.cell(row=row_idx, column=date_column).value,
            "description": sheet.cell(row=row_idx, column=description_column).value,
            "amount": sheet.cell(row=row_idx, column=amount_column).value
        }
        data.append(entry)

    return data

# Function to create or update an Excel file from response data
def create_excel_file(response_data, categories, file_name):
    try:
        wb = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  

    sheets = {}
    for category in categories:
        if category in wb.sheetnames:
            sheets[category] = wb[category]
        else:
            sheets[category] = wb.create_sheet(title=category)
            sheets[category].append(["Date", "Amount", "Description", "Transaction Type"])

    for entry in response_data:
        category = entry["category"]
        if category in sheets:
            sheets[category].append([entry["date"], entry["amount"], entry["description"], entry["transaction_type"]])

    wb.save(file_name)

def get_groq_response(categorized_data):
    client = Groq(
        api_key=os.environ.get("GROQ_API_KEY")
    )

    chat_completion = client.chat.completions.create(
        messages=[
            {
                "role": "user",
                "content": f"Please note: I don't want code! {categorized_data} \n Take this data and give me a json which has Date , Amount, Description, transaction type (credit or debit) and category(give the category from these 4 'Income, Expenses, Business Expenses and Uncertain Expenses') analyze the data and description to give me transaction_type and category don't provide null, and always return json for the whole data don't skip anything..",
            }
        ],
        model="llama3-70b-8192",
    )

    response = chat_completion.choices[0].message.content
    print("Raw response from Groq:", response)  
    return response

def extract_json_from_string(string):
    json_objects = []
    stack = []
    start_index = -1

    for i, char in enumerate(string):
        if char == '{':
            if not stack:
                start_index = i
            stack.append('{')
        elif char == '}':
            if stack:
                stack.pop()
                if not stack:
                    json_str = string[start_index:i + 1]
                    try:
                        json_obj = json.loads(json_str)
                        json_objects.append(json_obj)
                    except json.JSONDecodeError:
                        continue

    return json_objects

# Main function
def main():
    file_path = "processed_files/test.xlsx"
    data = read_excel_file(file_path)

    response = get_groq_response(data)

    try:
        json_objects = extract_json_from_string(response)
        if not json_objects:
            raise ValueError("No valid JSON objects found in the response.")
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON from response: {e}")
        return

    print("Extracted JSON objects:", json_objects)

    create_excel_file(json_objects, categories, "categorized_data.xlsx")

if __name__ == "__main__":
    main()
