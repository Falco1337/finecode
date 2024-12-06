import datetime
import json
import os
import logging
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt, Confirm
from rich.progress import Progress
from openpyxl import Workbook, load_workbook

BANK_ACCOUNTS_FILE = "bank_accounts.json"
DATE_FORMAT = "%Y-%m-%d"
LOGGING_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

logging.basicConfig(level=logging.INFO, format=LOGGING_FORMAT)
console = Console()

def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")

def get_current_month_filename():
    current_month = datetime.datetime.now().strftime("%B")
    return f"Finance_statements_{current_month}.xlsx"

def initialize_workbook():
    filename = get_current_month_filename()
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
    if os.path.exists(BANK_ACCOUNTS_FILE):
        try:
            with open(BANK_ACCOUNTS_FILE, "r") as file:
                return json.load(file)
        except json.JSONDecodeError:
            console.print("[bold red]Error: Bank accounts file is corrupted.[/bold red]")
    return {}

def save_bank_accounts(bank_accounts):
    try:
        with open(BANK_ACCOUNTS_FILE, "w") as file:
            json.dump(bank_accounts, file)
        logging.info("Bank accounts saved successfully.")
    except IOError:
        console.print("[bold red]Error: Unable to save bank accounts.[/bold red]")

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

def auto_save(workbook, bank_accounts):
    filename = get_current_month_filename()
    try:
        workbook.save(filename)
        save_bank_accounts(bank_accounts)
        logging.info(f"Data auto-saved to {filename}")
    except IOError as e:
        console.print(f"[bold red]Error auto-saving data: {e}[/bold red]")

def process_transaction(transaction_type, bank_accounts, descriptions, transactions_sheet, workbook):
    description = Prompt.ask("Enter a brief description of the transaction")
    bank = Prompt.ask("Which bank will be used?")
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
    auto_save(workbook, bank_accounts)

def show_help():
    console.print("[bold blue]Available Commands:[/bold blue]")
    console.print("[bold green]I:[/bold green] Income transaction\n[bold green]E:[/bold green] Expense transaction\n[bold green]S:[/bold green] Save data\n[bold green]T:[/bold green] Show transaction history\n[bold green]H:[/bold green] Show help\n[bold green]Q:[/bold green] Quit the application")

def save_data(workbook, bank_accounts, descriptions):
    if Confirm.ask("Are you sure you want to save your data?"):
        filename = get_current_month_filename()
        workbook.save(filename)
        save_bank_accounts(bank_accounts)
        console.print(f"[bold green]Data saved successfully in {filename}[/bold green]")

def main():
    clear_screen()
    console.print("[bold green]\n🌟 Welcome to Aidil Finance Management System! 🌟\n[/bold green]")
    bank_accounts = load_bank_accounts()
    descriptions = {}
    if not bank_accounts:
        console.print("[bold yellow]No bank accounts found. Please set up your accounts.[/bold yellow]")
        while True:
            try:
                bank_name = input("Enter the bank name (or 'q' to quit): ")
                if bank_name.lower() == 'q':
                    break
                balance_input = Prompt.ask(f"Enter the balance for {bank_name}")
                balance = validate_amount(balance_input)
                if balance is not None:
                    bank_accounts[bank_name] = balance
            except KeyboardInterrupt:
                console.print("\n[bold yellow]Setup cancelled.[/bold yellow]")
                break
    filename, workbook, transactions_sheet = initialize_workbook()
    while True:
        try:
            command = Prompt.ask("\nChoose an option: [I] Income, [E] Expense, [S] Save, [T] Transactions, [H] Help, [Q] Quit", default="Q").strip().upper()
            if command == "I":
                process_transaction("Income", bank_accounts, descriptions, transactions_sheet, workbook)
            elif command == "E":
                process_transaction("Expense", bank_accounts, descriptions, transactions_sheet, workbook)
            elif command == "S":
                save_data(workbook, bank_accounts, descriptions)
            elif command == "T":
                show_transaction_history(transactions_sheet)
            elif command == "H":
                show_help()
            elif command == "Q":
                console.print("[bold yellow]👋 Goodbye![/bold yellow]")
                break
            else:
                console.print("[bold red]Invalid option. Try again.[/bold red]")
        except KeyboardInterrupt:
            console.print("\n[bold yellow]Exiting...[/bold yellow]")
            break

if __name__ == "__main__":
    main()