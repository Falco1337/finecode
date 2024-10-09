import datetime
from openpyxl import Workbook, load_workbook
import json
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")

def record_transaction(amount, description, transaction_type, bank, transactions_sheet):
    """Record a financial transaction to the Excel sheet."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    transactions_sheet.append([timestamp, transaction_type, amount, description, bank])

def total_bank(bank_accounts):
    """Return the total balance across all bank accounts."""
    return sum(bank_accounts.values())

def generate_financial_statements(workbook, transactions_sheet):
    """Generate financial statements (income, expenses, and net income)."""
    total_income, total_expenses = 0.0, 0.0
    
    for row in transactions_sheet.iter_rows(min_row=2, values_only=True):
        if row[1] == "Income":
            total_income += row[2]
        elif row[1] == "Expense":
            total_expenses += row[2]
    
    ws_statements = workbook.create_sheet(title="Financial Statements Latest")
    ws_statements.append(["Date", "Total Income", "Total Expenses", "Net Income"])
    
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    net_income = total_income - total_expenses
    ws_statements.append([current_date, total_income, total_expenses, net_income])
    
    logging.info("Financial statements added to the workbook.")

def generate_daily_expenses_sheet(workbook, transactions_sheet):
    """Create a sheet for tracking daily expenses."""
    ws_daily_expenses = workbook.create_sheet(title="Daily Expenses")
    ws_daily_expenses.append(["Date", "Description", "Amount", "Bank"])
    
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    
    for row in transactions_sheet.iter_rows(min_row=2, values_only=True):
        if row[0].split()[0] == current_date and row[1] == "Expense":
            ws_daily_expenses.append([row[0], row[3], row[2], row[4]])
    
    logging.info("Daily expenses sheet added to the workbook.")

def save_bank_accounts(bank_accounts):
    """Save bank account balances to a JSON file."""
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    with open(f"bank_accounts_{current_date}.json", "w") as file:
        json.dump(bank_accounts, file)
    logging.info(f"Bank accounts saved to bank_accounts_{current_date}.json.")

def load_bank_accounts():
    """Load bank account balances from a JSON file if available."""
    if os.path.exists("bank_accounts.json"):
        with open("bank_accounts.json", "r") as file:
            return json.load(file)
    return {}

def initialize_bank_accounts():
    """Initialize bank accounts with user input."""
    bank_accounts = {
        "RHB": float(input("Latest amount in RHB: ")),
        "Bank Islam": float(input("Latest amount in Bank Islam: ")),
        "CIMB Bank": float(input("Latest amount in CIMB Bank: "))
    }
    
    if input("Do you have any other bank? [y/n]: ").strip().lower() == "y":
        bank_name = input("Enter the bank name: ").strip()
        bank_accounts[bank_name] = float(input(f"Enter the latest amount in {bank_name}: "))
    
    return bank_accounts

def get_transaction_type():
    """Prompt user to input a valid transaction type."""
    while True:
        transaction_type = input("\nIncome [I], Expense [E], or SAVE [S] to save: ").strip().upper()
        if transaction_type in ["I", "E", "S"]:
            return transaction_type
        logging.warning("Invalid transaction type. Please enter I, E, or S.")

def get_bank_selection(bank_accounts):
    """Let user select a bank or input a new bank name."""
    bank_choices = list(bank_accounts.keys()) + ["OTHERS"]
    
    print("\nSelect your bank:")
    for idx, bank in enumerate(bank_choices, 1):
        print(f"[{idx}] {bank}")
    
    while True:
        choice = input("Enter the bank number or name: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(bank_choices):
            selected_bank = bank_choices[int(choice) - 1]
            if selected_bank == "OTHERS":
                return input("Enter the bank name: ").strip()
            return selected_bank
        elif choice in bank_accounts:
            return choice
        else:
            logging.warning("Invalid selection. Try again.")

def handle_transaction(bank_accounts, transactions_sheet, transaction_type, total_income, total_expenses):
    """Process a financial transaction based on the user's input."""
    description = input("Enter the transaction description: ").strip()
    amount_input = input("Enter the transaction amount: ").strip()

    try:
        amount = float(amount_input)
    except ValueError:
        logging.error("Invalid input. Please enter a valid number.")
        return total_income, total_expenses

    selected_bank = get_bank_selection(bank_accounts)

    if transaction_type == "E":
        if bank_accounts[selected_bank] < amount:
            logging.warning("Insufficient funds. Please try again.")
            return total_income, total_expenses
        bank_accounts[selected_bank] -= amount
        total_expenses += amount
        record_transaction(amount, description, "Expense", selected_bank, transactions_sheet)
        logging.info("Expense recorded.")
    elif transaction_type == "I":
        bank_accounts[selected_bank] += amount
        total_income += amount
        record_transaction(amount, description, "Income", selected_bank, transactions_sheet)
        logging.info("Income recorded.")

    logging.info("Updated Account Balances:")
    for bank, amount in bank_accounts.items():
        logging.info(f"{bank}: {amount}")
    
    logging.info(f"Total Balance: {total_bank(bank_accounts)}")
    return total_income, total_expenses

def main():
    """Main entry point for the Finance Management system."""
    clear_screen()
    logging.info("/nWelcome to Finance Management")
    
    bank_accounts = load_bank_accounts() or initialize_bank_accounts()
    
    logging.info(f"\nTotal Amount: {total_bank(bank_accounts)}")
    for bank, amount in bank_accounts.items():
        logging.info(f"{bank}: {amount}")

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

    total_income, total_expenses = 0.0, 0.0

    try:
        while True:
            transaction_type = get_transaction_type()
            if transaction_type == "S":
                wb.save(filename)
                logging.info(f"File saved as {filename}")
                save_bank_accounts(bank_accounts)
                break

            total_income, total_expenses = handle_transaction(bank_accounts, ws, transaction_type, total_income, total_expenses)

    except KeyboardInterrupt:
        interrupted_filename = f"finance_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_interrupted.xlsx"
        logging.warning("\nTransaction recording interrupted. Saving data...")
        wb.save(interrupted_filename)
        logging.info(f"Data saved as {interrupted_filename}")

    logging.info(f"\nTotal Income: {total_income}")
    logging.info(f"Total Expenses: {total_expenses}")

    generate_financial_statements(wb, ws)
    generate_daily_expenses_sheet(wb, ws)
    wb.save(filename)
    logging.info(f"All data saved in {filename}")

if __name__ == "__main__":
    main()
