import datetime
from openpyxl import Workbook, load_workbook
import json
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def clear_screen():
    os.system("clear")

def record_transaction(amount, description, transaction_type, bank, transactions_sheet):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    transactions_sheet.append([timestamp, transaction_type, amount, description, bank])

def total_bank(bank_accounts):
    return sum(bank_accounts.values())

def generate_financial_statements(workbook, transactions_sheet):
    total_income = 0.0
    total_expenses = 0.0
    
    for row in transactions_sheet.iter_rows(min_row=2, values_only=True):
        if row[1] == "Income":
            total_income += row[2]
        elif row[1] == "Expense":
            total_expenses += row[2]
    
    ws_statements = workbook.create_sheet(title="Financial Statements 2024")
    ws_statements.append(["Date", "Total Income", "Total Expenses", "Net Income"])
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    net_income = total_income - total_expenses
    ws_statements.append([current_date, total_income, total_expenses, net_income])
    logging.info("Financial statements added to the workbook.")

def generate_daily_expenses_sheet(workbook, transactions_sheet):
    ws_daily_expenses = workbook.create_sheet(title="Daily Expenses")
    ws_daily_expenses.append(["Date", "Description", "Amount", "Bank"])
    
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    
    for row in transactions_sheet.iter_rows(min_row=2, values_only=True):
        if row[0].split()[0] == current_date:
            if row[1] == "Expense":
                ws_daily_expenses.append([row[0], row[3], row[2], row[4]])
    
    logging.info("Daily expenses sheet added to the workbook.")

def save_bank_accounts(bank_accounts):
    with open("bank_accounts.json", "w") as file:
        json.dump(bank_accounts, file)

def load_bank_accounts():
    if os.path.exists("bank_accounts.json"):
        with open("bank_accounts.json", "r") as file:
            return json.load(file)
    else:
        return {}

def initialize_bank_accounts():
    bank_accounts = {
        "RHB": float(input("Enter your latest amount in RHB: ")),
        "Bank Islam": float(input("Enter your latest amount in Bank Islam: ")),
        "Cimb Bank": float(input("Enter your latest amount in Cimb Bank: "))
    }
    
    if input("Do you have any other bank? [y/n]: ").lower() == "y":
        bank_name = input("What is your bank? : ")
        bank_accounts[bank_name] = float(input(f"Enter your latest amount in {bank_name}: "))
    
    return bank_accounts

def get_transaction_type():
    while True:
        transaction_type = input("\nIncome [I] or Expense [E] and SAVE [S] to save file: ").upper()
        if transaction_type in ["I", "E", "S", "SAVE"]:
            return transaction_type
        else:
            print("Invalid transaction type. Please enter I, E, or S.")

def get_bank_selection(bank_accounts):
    while True:
        print("\n[1] RHB Bank")
        print("[2] Bank Islam")
        print("[3] CIMB Bank")
        print("[4] OTHERS\n")
        loc_bank = input("Which bank did you use? [Enter number]: ")
        if loc_bank.isdigit():
            loc_bank = int(loc_bank)
            if loc_bank == 1:
                return "RHB"
            elif loc_bank == 2:
                return "Bank Islam"
            elif loc_bank == 3:
                return "Cimb Bank"
            elif loc_bank == 4:
                return input("Enter the bank name: ")
            else:
                print("Invalid selection. Please try again.")
        else:
            if loc_bank in bank_accounts:
                return loc_bank
            else:
                print("Bank not found. Please try again.")

def handle_transaction(bank_accounts, ws, transaction_type):
    description = input("Enter the transaction description: ")
    amount_input = input("Enter the transaction amount: ")

    try:
        amount = float(amount_input)
    except ValueError:
        print("Invalid input. Please enter a valid number.")
        return

    selected_bank = get_bank_selection(bank_accounts)
    
    if transaction_type in ["E", "EXPENSE"]:
        if bank_accounts[selected_bank] < amount:
            print("Insufficient funds in the selected bank account. Please try again.")
            return

        bank_accounts[selected_bank] -= amount
        record_transaction(amount, description, "Expense", selected_bank, ws)
        print("Expense recorded.")
        
    elif transaction_type in ["I", "INCOME"]:
        bank_accounts[selected_bank] += amount
        record_transaction(amount, description, "Income", selected_bank, ws)
        print("Income recorded.")

    print("\nUpdated Account Balances:")
    for bank, amount in bank_accounts.items():
        print(f"{bank}: {amount}")

    print(f"Total Amount: {total_bank(bank_accounts)}")

def main():
    timestamp2 = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    clear_screen()
    print("Welcome to Finance Management\n")
    print(f"{timestamp2}\n")
    
    bank_accounts = load_bank_accounts()
    if not bank_accounts:
        bank_accounts = initialize_bank_accounts()

    print("\nTotal Amount: ", total_bank(bank_accounts))
    for bank, amount in bank_accounts.items():
        print(f"{bank}: {amount}")

    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    filename = f"finance_{current_date}_recorded.xlsx"
    
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Transactions"
        ws.append(["Date", "Transaction Type", "Amount", "Description", "Bank"])

    try:
        while True:
            transaction_type = get_transaction_type()
            if transaction_type in ["S", "SAVE"]:
                wb.save(filename)
                print(f"File saved as {filename}")
                save_bank_accounts(bank_accounts)
                break

            handle_transaction(bank_accounts, ws, transaction_type)

    except KeyboardInterrupt:
        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        print("\n\nTransaction recording interrupted. Saving data to file...\n")
        interrupted_filename = f"finance_{current_time}_interrupted.xlsx"
        wb.save(interrupted_filename)
        print(f"\nData saved successfully as {interrupted_filename}")

    # Display total income and total expenses
    print("\nTotal Income:", total_income)
    print("Total Expenses:", total_expenses)

    # Generate financial statements and daily expenses sheet
    generate_financial_statements(wb, ws)
    generate_daily_expenses_sheet(wb, ws)
    wb.save(filename)
    print(f"All data saved in {filename}")

if __name__ == "__main__":
    main()
