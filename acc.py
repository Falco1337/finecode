import datetime
import json
import os
import logging
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt, Confirm
from rich.progress import Progress
from openpyxl import Workbook, load_workbook

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
console = Console()

def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")

def initialize_workbook():
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
    try:
        if os.path.exists("bank_accounts.json"):
            with open("bank_accounts.json", "r") as file:
                return json.load(file)
    except json.JSONDecodeError:
        console.print("[bold red]Error loading bank accounts. The file may be corrupted.[/bold red]")
    return {}

def save_bank_accounts(bank_accounts):
    with open("bank_accounts.json", "w") as file:
        json.dump(bank_accounts, file)
    logging.info("Bank accounts saved to bank_accounts.json.")

def record_transaction(transactions_sheet, amount, description, transaction_type, bank):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    transactions_sheet.append([timestamp, transaction_type, amount, description, bank])

def validate_amount(amount_input):
    try:
        amount = float(amount_input)
        if amount <= 0:
            raise ValueError("Amount must be a positive number.")
        return amount
    except ValueError as e:
        console.print(f"[bold red]{e}[/bold red]")
        return None

def display_balance(bank_accounts, descriptions):
    total_balance = sum(bank_accounts.values())
    console.print(f"\n[bold cyan]Total Balance: RM {total_balance:.2f}[/bold cyan]\n")

    table = Table(title="💰 Aidil Finance Management 2024 💸")
    table.add_column("Bank", justify="left", style="magenta")
    table.add_column("Amount", justify="right", style="green")
    table.add_column("Description", justify="left", style="yellow")

    max_balance = max((abs(amount) for amount in bank_accounts.values()), default=1)
    
    for bank, amount in bank_accounts.items():
        table.add_row(bank, f"RM {amount:.2f}", descriptions.get(bank, "Balance"))
    console.print(table)

    console.print("\n[bold blue]Account Balances Monitoring:[/bold blue]\n")
    for bank, amount in bank_accounts.items():
        bar_length = int((abs(amount) / max_balance) * 30)  
        bar_color = "red" if amount < 0 else "green"  
        bar = f"[{bar_color}]{'█' * bar_length}[/{bar_color}]"
        console.print(f"{bank.ljust(15)} | {bar} RM {amount:.2f}")

def show_transaction_history(transactions_sheet):
    console.print("\n[bold blue]Transaction History:[/bold blue]\n")
    for row in transactions_sheet.iter_rows(min_row=2, values_only=True):
        console.print(f"{row[0]} | {row[1]} | RM {row[2]:.2f} | {row[3]} | {row[4]}")

def process_transaction(transaction_type, bank_accounts, descriptions, transactions_sheet):
    description = Prompt.ask("Enter a brief description of the transaction")
    bank = Prompt.ask("Which bank used ?")
    amount_input = Prompt.ask("Enter the transaction amount")
    
    amount = validate_amount(amount_input)
    if amount is None:
        return

    if transaction_type == "Expense" and bank_accounts.get(bank, 0) < amount:
        console.print(f"[bold red]Insufficient funds in {bank} for this expense.[/bold red]")
        return

    bank_accounts[bank] = bank_accounts.get(bank, 0) + (amount if transaction_type == "Income" else -amount)

    with Progress(console=console, transient=True) as progress:
        task = progress.add_task(f"[cyan]Recording {transaction_type.lower()}...[/cyan]", total=100)
        record_transaction(transactions_sheet, amount, description, transaction_type, bank)
        descriptions[bank] = description
        progress.update(task, advance=100)
    
    display_balance(bank_accounts, descriptions)
    logging.info(f"{transaction_type} of RM {amount} recorded for {bank}.")

def generate_financial_statements(workbook, transactions_sheet):
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

def save_data(workbook, bank_accounts, descriptions):
    if Confirm.ask("Are you sure you want to save your data ?"):
        filename, _, transactions_sheet = initialize_workbook()
        generate_financial_statements(workbook, transactions_sheet)
        workbook.save(filename)
        save_bank_accounts(bank_accounts)
        console.print(f"[bold green]Data saved successfully in {filename}[/bold green]")

def initialize_bank_accounts():
    bank_accounts = load_bank_accounts()
    if bank_accounts:
        console.print("[bold green]Loaded bank account data from previous session.[/bold green]\n")
    else:
        console.print("[bold yellow]No previous bank account data found. Please enter initial bank account details.[/bold yellow]\n")
        try:
            while True:
                bank_name = input("Enter the bank name: ")
                if bank_name.lower() == 'q':
                    break
                latest_amount = Prompt.ask(f"What is the current balance in {bank_name}?")
                amount = validate_amount(latest_amount)
                if amount is None:
                    continue
                bank_accounts[bank_name] = amount
                if not Confirm.ask("Would you like to add another bank ?"):
                    break
        except KeyboardInterrupt:
            console.print("\n[bold yellow]Exiting bank account entry...[/bold yellow]")
    
    return bank_accounts

def show_help():
    console.print("[bold blue]Available Commands:[/bold blue]")
    console.print("[bold green]I:[/bold green] Income transaction\n[bold green]E:[/bold green] Expense transaction\n[bold green]E:[/bold green] Expense transaction\n[bold green]S:[/bold green] Save data\n[bold green]T:[/bold green] Show transaction history\n[bold green]H:[/bold green] Show help\n[bold green]Q:[/bold green] Quit the application")

def main():
    clear_screen()    
    console.print("[bold green]\n🌟 Welcome to Aidil Finance Management System ! 🌟\n[/bold green]")
    bank_accounts = initialize_bank_accounts()
    descriptions = {}
    clear_screen()
    display_balance(bank_accounts, descriptions)

    filename, workbook, transactions_sheet = initialize_workbook()
    
    try:
        while True:
            command = Prompt.ask("\nChoose an option: \n[I] Income, \n[E] Expense, \n[S] Save, \n[T] Transactions, \n[H] Help, \n[Q] Quit", default="Q").strip().upper()
            if command in ['S', 'SAVE']:
                save_data(workbook, bank_accounts, descriptions)
            elif command in ['I', 'E']:
                transaction_type = "Income" if command == 'I' else "Expense"
                process_transaction(transaction_type, bank_accounts, descriptions, transactions_sheet)
            elif command in ['T', 'TRANSACTIONS']:
                show_transaction_history(transactions_sheet)
            elif command in ['H', 'HELP']:
                show_help()
            elif command in ['Q', 'QUIT']:
                console.print("[bold yellow]👋 Exiting the Finance Management System. Goodbye![/bold yellow]")
                break
            else:
                console.print("[bold red]Invalid command. Please try again.[/bold red]")
    except KeyboardInterrupt:
        console.print("[bold yellow]👋 Exiting.... Goodbye ![/bold yellow]")

if __name__ == "__main__":
    main()
