import datetime
import json
import os
import logging
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt, Confirm
from openpyxl import Workbook, load_workbook

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

console = Console()

def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")

def initialize_workbook():
    """Initialize the workbook and load the transaction sheet."""
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    filename = f"finance_{current_date}_recorded.xlsx"
    
    if os.path.exists(filename):
        workbook = load_workbook(filename)
        transactions_sheet = workbook.active
    else:
        workbook = Workbook()
        transactions_sheet = workbook.active
        transactions_sheet.title = "Transactions"
        transactions_sheet.append(["Date", "Transaction Type", "Amount", "Description", "Bank"])
    return filename, workbook, transactions_sheet

def load_bank_accounts():
    """Load bank account balances from a JSON file if available."""
    if os.path.exists("bank_accounts.json"):
        with open("bank_accounts.json", "r") as file:
            return json.load(file)
    return {}

def save_bank_accounts(bank_accounts):
    """Save bank account balances to a JSON file."""
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    with open(f"bank_accounts_{current_date}.json", "w") as file:
        json.dump(bank_accounts, file)
    logging.info(f"Bank accounts saved to bank_accounts_{current_date}.json.")

def record_transaction(transactions_sheet, amount, description, transaction_type, bank):
    """Record a financial transaction to the Excel sheet."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    transactions_sheet.append([timestamp, transaction_type, amount, description, bank])

def update_balance_display(bank_accounts, descriptions):
    """Display the total balance and account details."""
    total_balance = sum(bank_accounts.values())
    console.print(f"\n[bold cyan]Total Balance: {total_balance:.2f}[/bold cyan]")

    table = Table(title="Account Balances")
    table.add_column("Bank", justify="left")
    table.add_column("Amount", justify="right")
    table.add_column("Description", justify="left")

    for bank, amount in bank_accounts.items():
        table.add_row(bank, f"{amount:.2f}", descriptions.get(bank, "N/A"))
    
    console.print(table)

def handle_transaction(transaction_type, bank_accounts, descriptions, transactions_sheet):
    """Handle the transaction and update balances."""
    description = Prompt.ask("Enter the transaction description")
    bank = Prompt.ask("Enter the bank name")
    
    amount_input = Prompt.ask("Enter the transaction amount")
    
    try:
        amount = float(amount_input)
    except ValueError:
        console.print("[bold red]Invalid amount. Please enter a valid number.[/bold red]")
        return
    
    if transaction_type == "Expense":
        if bank_accounts.get(bank, 0) < amount:
            console.print(f"[bold red]Insufficient funds in {bank}. Please try again.[/bold red]")
            return
        bank_accounts[bank] -= amount
    elif transaction_type == "Income":
        bank_accounts[bank] = bank_accounts.get(bank, 0) + amount

    # Record transaction
    record_transaction(transactions_sheet, amount, description, transaction_type, bank)

    # Update the latest description for the bank
    descriptions[bank] = description
    
    # Update display
    update_balance_display(bank_accounts, descriptions)

    # Log transaction
    logging.info(f"{transaction_type} of {amount} recorded for {bank}.")

def generate_financial_statements(workbook, transactions_sheet):
    """Generate financial statements (income, expenses, and net income)."""
    total_income, total_expenses = 0.0, 0.0
    
    for row in transactions_sheet.iter_rows(min_row=2, values_only=True):
        if row[1] == "Income":
            total_income += row[2]
        elif row[1] == "Expense":
            total_expenses += row[2]
    
    ws_statements = workbook.create_sheet(title="Financial Statements")
    ws_statements.append(["Date", "Total Income", "Total Expenses", "Net Income"])
    
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    net_income = total_income - total_expenses
    ws_statements.append([current_date, total_income, total_expenses, net_income])
    
    logging.info("Financial statements added to the workbook.")

def save_data(workbook, bank_accounts, descriptions, transactions_sheet):
    """Save the workbook and bank accounts."""
    filename, workbook, transactions_sheet = initialize_workbook()
    if Confirm.ask("Are you sure you want to save your data?"):
        generate_financial_statements(workbook, transactions_sheet)
        workbook.save(filename)
        save_bank_accounts(bank_accounts)
        console.print(f"[bold green]Data has been saved in {filename}[/bold green]")

def initialize_bank_accounts():
    """Initialize bank accounts with user input."""
    bank_accounts = {}
    while True:
        bank_name = Prompt.ask("Enter the bank name")
        latest_amount = Prompt.ask(f"Enter the latest amount in {bank_name}")
        try:
            bank_accounts[bank_name] = float(latest_amount)
        except ValueError:
            console.print("[bold red]Invalid amount. Please enter a valid number.[/bold red]")
            continue

        if not Confirm.ask("Do you want to add another bank?"):
            break

    return bank_accounts

def show_help():
    """Display help information to the user."""
    console.print("[bold blue]Available Commands:[/bold blue]")
    console.print("[bold green]I:[/bold green] Income transaction")
    console.print("[bold green]E:[/bold green] Expense transaction")
    console.print("[bold green]S:[/bold green] Save data")
    console.print("[bold green]H:[/bold green] Show help")
    console.print("[bold green]Q:[/bold green] Quit the application")

def main():
    """Main function to run the Finance Management system."""
    clear_screen()    
    console.print("[bold green]Welcome to Finance Management[/bold green]\n")
    bank_accounts = initialize_bank_accounts()
    clear_screen()
    descriptions = {}  # Initialize descriptions dictionary
    update_balance_display(bank_accounts, descriptions)

    filename, workbook, transactions_sheet = initialize_workbook()

    while True:
        transaction_type = Prompt.ask("\nChoose transaction type: [I] Income, [E] Expense, [S] Save, [H] Help, [Q] Quit", default="Q").strip().upper()
        if transaction_type in ['S', 'SAVE']:
            save_data(workbook, bank_accounts, descriptions, transactions_sheet)
        elif transaction_type in ['I', 'E']:
            handle_transaction("Income" if transaction_type == 'I' else "Expense", bank_accounts, descriptions, transactions_sheet)
        elif transaction_type in ['H', 'HELP']:
            show_help()
        elif transaction_type in ['Q', 'QUIT']:
            console.print("[bold yellow]Exiting the Finance Management system.[/bold yellow]")
            break
            
if __name__ == "__main__":
    main()
