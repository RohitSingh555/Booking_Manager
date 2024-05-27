import openpyxl
from datetime import datetime, timedelta

def read_excel_file(file_path, sheets_to_read):
    wb = openpyxl.load_workbook(file_path)
    data = {sheet: [] for sheet in sheets_to_read}

    for sheet_name in sheets_to_read:
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                entry = {
                    "date": row[0],
                    "amount": row[1],
                    "description": row[2],
                    "source": row[3]
                }
                data[sheet_name].append(entry)
        else:
            print(f"Sheet {sheet_name} not found in the file.")

    return data

def parse_date(date_str):
    for fmt in ('%m/%d/%y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%b %d, %Y'):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Date format for {date_str} is not recognized.")

import re

def preprocess_amount(amount_str):
    # Convert to string if amount_str is not already a string
    if not isinstance(amount_str, str):
        amount_str = str(amount_str)
    
    # Remove any non-numeric characters using regex
    cleaned_amount = re.sub(r'[^0-9.-]', '', amount_str)
    # If there is a '-' at the end, treat it as a negative sign
    cleaned_amount= cleaned_amount.lstrip('-')
    # Convert the cleaned amount to a float
    float_amount = float(cleaned_amount)
    return float_amount

def calculate_weekly_balances(data):
    combined_data = data["Income"] + data["Expenses"]
    combined_data.sort(key=lambda x: parse_date(x["date"]))
    
    start_date = parse_date(combined_data[0]["date"])
    end_date = parse_date(combined_data[-1]["date"])

    weekly_data = []
    current_date = start_date
    while current_date <= end_date:
        week_start = current_date
        week_end = current_date + timedelta(days=6)
        weekly_income = sum(float(preprocess_amount(entry["amount"])) for entry in data["Income"] if week_start <= parse_date(entry["date"]) <= week_end)
        weekly_expenses = sum(float(preprocess_amount(entry["amount"])) for entry in data["Expenses"] if week_start <= parse_date(entry["date"]) <= week_end)

        weekly_balance = weekly_income - weekly_expenses
        weekly_data.append({
            "week_start": week_start.strftime('%m/%d/%y'),
            "income": weekly_income,
            "expenses": weekly_expenses,
            "balance": weekly_balance
        })

        current_date += timedelta(days=7)

    return weekly_data

def calculate_balance_summary(data):
    total_income = sum(float(entry["amount"]) for entry in data["Income"])
    total_expenses = sum(float(entry["amount"]) for entry in data["Expenses"])
    net_income = total_income - total_expenses
    available_budget = net_income
    daily_budget = net_income / 365
    weekly_budget = net_income / 52
    yearly_budget = net_income

    return {
        "Total Income": total_income,
        "Total Expenses": total_expenses,
        "Net Income": net_income,
        "Available Budget": available_budget,
        "Daily Budget": daily_budget,
        "Weekly Budget": weekly_budget,
        "Yearly Budget": yearly_budget
    }

def write_to_excel(file_path, weekly_balances, balance_summary, account_balances):
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    # Overwrite the existing Weekly Budget sheet or create a new one
    if "Weekly Budget" not in wb.sheetnames:
        weekly_budget_sheet = wb.create_sheet(title="Weekly Budget")
        weekly_budget_sheet.append(["Week Start", "Income", "Expenses", "Balance"])
    else:
        weekly_budget_sheet = wb["Weekly Budget"]
        # Clear existing data
        weekly_budget_sheet.delete_rows(2, weekly_budget_sheet.max_row)

    # Write weekly balances data to the Weekly Budget sheet
    for weekly_balance in weekly_balances:
        weekly_budget_sheet.append([
            weekly_balance["week_start"],
            weekly_balance["income"],
            weekly_balance["expenses"],
            weekly_balance["balance"]
        ])

    # Overwrite or create Balance Summary sheet
    if "Balance Summary" not in wb.sheetnames:
        balance_summary_sheet = wb.create_sheet(title="Balance Summary")
    else:
        balance_summary_sheet = wb["Balance Summary"]
        # Clear existing data
        balance_summary_sheet.delete_rows(2, balance_summary_sheet.max_row)

    # Write balance summary data to the Balance Summary sheet
    # balance_summary_sheet.append(["Description", "Amount"])
    for description, amount in balance_summary.items():
        balance_summary_sheet.append([description, amount])

    # Overwrite or create Balances sheet
    if "Balances" not in wb.sheetnames:
        balances_sheet = wb.create_sheet(title="Balances")
    else:
        balances_sheet = wb["Balances"]
        # Clear existing data
        balances_sheet.delete_rows(2, balances_sheet.max_row)

    # Write account balances data to the Balances sheet
    balances_sheet.append(["Account Type", "Amount"])
    for account_type, balance in account_balances.items():
        balances_sheet.append([account_type, balance])

    wb.save(file_path)

def main():
    file_path = "categorized_data.xlsx"
    sheets_to_read = ["Income", "Expenses", "Business Expenses", "Tax Deductible Expenses", "Subscriptions", "Uncertain Expenses"]
    
    data = read_excel_file(file_path, sheets_to_read)

    if data["Income"] and data["Expenses"]:
        weekly_balances = calculate_weekly_balances(data)
        balance_summary = calculate_balance_summary(data)

        # Get account balances from the user
        account_balances = {}
        num_accounts = int(input("Enter the number of accounts: "))
        for _ in range(num_accounts):
            account_type = input("Enter Account Type: ")
            balance = float(input("Enter Balance: "))
            account_balances[account_type] = balance

        write_to_excel(file_path, weekly_balances, balance_summary, account_balances)
        print("Weekly Budget, Balance Summary, and Account Balances have been successfully written to the file.")
    else:
        print("No data found in Income or Expenses sheets.")

if __name__ == "__main__":
    main()
