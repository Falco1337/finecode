import atexit
import datetime
import json
import os
import logging
import time
import requests
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt
from openpyxl import Workbook

BANK_ACCOUNTS_FILE = "bank_total.json"
ASSETS_FILE = "assets.json"
TRANSACTIONS_FILE = "transactions.json"
DATE_FORMAT = "%Y-%m-%d"
LOGGING_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

ALPHA_VANTAGE_API_KEY = os.getenv("ALPHA_VANTAGE_API_KEY", "Y3PQLKE4QF2U56GV")

logging.basicConfig(level=logging.INFO, format=LOGGING_FORMAT)
console = Console()

class BankAccountManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.accounts = self._load_data()

    def _load_data(self):
        return load_json_data(self.file_path, {})

    def update_balances(self):
        console.print("[bold cyan]Please enter the latest balances for your bank accounts[/bold cyan]:")
        while True:
            bank_name = Prompt.ask("Enter the name of the bank (or type 'no' or 'n' to stop)")
            if bank_name.lower() in ["no", "n"]:
                break
            try:
                balance = float(Prompt.ask(f"Enter the latest balance for {bank_name}"))
                self.accounts[bank_name] = balance
            except ValueError:
                console.print("[bold red]Error: Please enter a valid number.[/bold red]")
        save_json_data(self.file_path, self.accounts)
        console.print("[bold green]Bank balances updated successfully![/bold green]")

    def get_total_balance(self):
        return sum(self.accounts.values())

class TransactionManager:
    def __init__(self, file_path, bank_manager):
        self.file_path = file_path
        self.transactions = self._load_data()
        self.bank_manager = bank_manager

    def _load_data(self):
        return load_json_data(self.file_path, [])

    def process_transaction(self, transaction_type):
        bank = Prompt.ask(f"Enter bank for {transaction_type}:")
        try:
            amount = float(Prompt.ask(f"Enter amount for {transaction_type}:"))
        except ValueError:
            console.print("[bold red]Error: Please enter a valid number.[/bold red]")
            return

        if transaction_type == "Expense" and self.bank_manager.accounts.get(bank, 0) < amount:
            console.print("[bold red]Error: Insufficient funds![/bold red]")
            return

        description = Prompt.ask(f"Enter a description for the {transaction_type.lower()}: ")

        self.bank_manager.accounts[bank] = self.bank_manager.accounts.get(bank, 0) + amount if transaction_type == "Income" else self.bank_manager.accounts.get(bank, 0) - amount
        self.transactions.append({
            "bank": bank,
            "amount": amount,
            "type": transaction_type,
            "description": description,
            "date": datetime.datetime.now().strftime(DATE_FORMAT)
        })

        save_json_data(self.file_path, self.transactions)
        save_json_data(BANK_ACCOUNTS_FILE, self.bank_manager.accounts)
        console.print(f"[bold green]{transaction_type} recorded successfully![/bold green]")

class AssetManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.assets = self._load_data()

    def _load_data(self):
        return load_json_data(self.file_path, {"ASNB": 981.47, "Gold": 435.60, "Properties": 300000.00})

    def add_asset(self):
        asset_name = Prompt.ask("Enter asset name:")
        try:
            asset_value = float(Prompt.ask("Enter asset value:"))
        except ValueError:
            console.print("[bold red]Error: Please enter a valid number.[/bold red]")
            return

        self.assets[asset_name] = self.assets.get(asset_name, 0) + asset_value
        save_json_data(self.file_path, self.assets)
        console.print("[bold green]Asset added successfully![/bold green]")

def clear_screen():
    try:
        os.system("cls" if os.name == "nt" else "clear")
    except:
        print("\n" * 100)

def load_json_data(file_path, default_data={}):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r") as file:
                return json.load(file)
        except (json.JSONDecodeError, IOError):
            console.print(f"[bold red]Error: {file_path} is corrupted. Initializing new data.[/bold red]")
    return default_data

def save_json_data(file_path, data):
    try:
        with open(file_path, "w") as file:
            json.dump(data, file, indent=4)
        logging.info(f"Data saved successfully to {file_path}.")
    except IOError:
        console.print(f"[bold red]Error: Unable to save data to {file_path}.[/bold red]")

def auto_save(bank_manager, asset_manager, transaction_manager):
    save_json_data(BANK_ACCOUNTS_FILE, bank_manager.accounts)
    save_json_data(ASSETS_FILE, asset_manager.assets)
    save_json_data(TRANSACTIONS_FILE, transaction_manager.transactions)
    console.print("[bold green]Data auto-saved successfully![/bold green]")

def display_balance(bank_manager, asset_manager):
    total_bank_balance = bank_manager.get_total_balance()
    total_assets = sum(asset_manager.assets.values())
    total_balance = total_bank_balance + total_assets

    transactions = load_json_data(TRANSACTIONS_FILE, [])
    total_income = sum(transaction["amount"] for transaction in transactions if transaction["type"] == "Income")
    total_expense = sum(transaction["amount"] for transaction in transactions if transaction["type"] == "Expense")

    clear_screen()
    console.print("[bold green]\n\U0001F31F Aydiel Accountant Assistant System 2025 \U0001F31F\n[/bold green]")
    console.print(f"[bold cyan]💰 Total Balance: RM {total_balance:.2f}[/bold cyan]\n")

    bank_table = Table(title="\n\U0001F4B3 Bank Accounts", show_header=True, header_style="bold magenta")
    bank_table.add_column("Bank", justify="left", style="cyan", no_wrap=True)
    bank_table.add_column("Balance", justify="right", style="green")

    for bank, amount in bank_manager.accounts.items():
        bank_table.add_row(bank, f"RM {amount:.2f}")

    console.print(bank_table)

    income_expense_table = Table(title="\n\U0001F4C8 Income & Expense Summary", show_header=True, header_style="bold blue")
    income_expense_table.add_column("Type", justify="left", style="cyan", no_wrap=True)
    income_expense_table.add_column("Amount", justify="right", style="green")

    income_expense_table.add_row("Total Income", f"RM {total_income:.2f}")
    income_expense_table.add_row("Total Expense", f"RM {total_expense:.2f}")
    income_expense_table.add_row("Net Savings", f"RM {total_income - total_expense:.2f}", style="bold yellow")

    console.print(income_expense_table)

    asset_table = Table(title="\n\U0001F4B8 Assets", show_header=True, header_style="bold magenta")
    asset_table.add_column("Asset", justify="left", style="cyan", no_wrap=True)
    asset_table.add_column("Value", justify="right", style="green")

    for asset, value in asset_manager.assets.items():
        asset_table.add_row(asset, f"RM {value:.2f}")

    console.print(asset_table)

def generate_reports(transaction_manager):
    transactions = transaction_manager.transactions
    if not transactions:
        console.print("[bold yellow]No transactions found.[/bold yellow]")
        return

    bank_usage = {}
    for transaction in transactions:
        bank = transaction.get("bank")
        amount = transaction.get("amount")
        description = transaction.get("description", "N/A")
        if bank and amount:
            bank_usage[bank] = bank_usage.get(bank, {"amount": 0, "description": description})
            bank_usage[bank]["amount"] += amount

    console.print("\n[bold cyan]Bank Usage Summary:[/bold cyan]")
    for bank, data in bank_usage.items():
        console.print(f"{bank}: RM {data['amount']:.2f}: {data['description']}")

    report_filename = f"account_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Bank", "Amount", "Type", "Description", "Date"])
    for transaction in transactions:
        ws.append([
            transaction.get("bank"),
            transaction.get("amount"),
            transaction.get("type"),
            transaction.get("description", "N/A"),
            transaction.get("date")
        ])
    try:
        wb.save(report_filename)
        console.print(f"[bold green]Report saved as {report_filename}[/bold green]")
    except Exception as e:
        console.print(f"[bold red]Error saving report: {str(e)}[/bold red]")

def fetch_stock_price(symbol):
    if symbol.upper() in ["USD/MYR", "MYR/USD"]:
        url = f"https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency={symbol[:3]}&to_currency={symbol[4:]}&apikey={ALPHA_VANTAGE_API_KEY}"
    else:
        url = f"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={symbol}&interval=5min&apikey={ALPHA_VANTAGE_API_KEY}"

    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        if symbol.upper() in ["USD/MYR", "MYR/USD"]:
            if "Realtime Currency Exchange Rate" in data:
                exchange_rate = data["Realtime Currency Exchange Rate"]["5. Exchange Rate"]
                console.print(f"[bold green]{symbol} Exchange Rate: {exchange_rate}[/bold green]")
            else:
                console.print("[bold red]Error: No exchange rate data available.[/bold red]")
        elif "Time Series (5min)" in data:
            latest_time = sorted(data["Time Series (5min)"].keys())[-1]
            stock_price = data["Time Series (5min)"][latest_time]["1. open"]
            console.print(f"[bold green]{symbol} Latest Price ({latest_time}): {stock_price}[/bold green]")
        else:
            console.print("[bold red]Error: No data available. Check API limits or try again later.[/bold red]")
    except (KeyError, ValueError, requests.RequestException) as e:
        console.print(f"[bold red]Error fetching data: {str(e)}[/bold red]")

def fetch_world_trends():
    console.print("\n[bold cyan]Fetching World Market Prices...[/bold cyan]")
    fetch_stock_price("USD/MYR")
    fetch_stock_price("XAU/USD")

def run_application():
    clear_screen()
    console.print("[bold green]\n\U0001F31F Aydiel Accountant Assistant System 2025 \U0001F31F\n[/bold green]")

    bank_manager = BankAccountManager(BANK_ACCOUNTS_FILE)
    asset_manager = AssetManager(ASSETS_FILE)
    transaction_manager = TransactionManager(TRANSACTIONS_FILE, bank_manager)

    atexit.register(auto_save, bank_manager, asset_manager, transaction_manager)

    if not bank_manager.accounts:
        bank_manager.update_balances()

    try:
        while True:
            command = Prompt.ask("\nChoose an option:\n[B] Balance\n[I] Income\n[E] Expense\n[A] Add Asset\n[R] Reports\n[T] World Trends\n[Q] Quit", default="Q").strip().upper()
            if command == "B":
                display_balance(bank_manager, asset_manager)
            elif command == "I":
                transaction_manager.process_transaction("Income")
            elif command == "E":
                transaction_manager.process_transaction("Expense")
            elif command == "A":
                asset_manager.add_asset()
            elif command == "R":
                generate_reports(transaction_manager)
            elif command == "T":
                fetch_world_trends()
            elif command == "Q":
                console.print("[bold green]Good Bye![/bold green]")
                clear_screen()
                time.sleep(1)
                break
    except KeyboardInterrupt:
        clear_screen()
        auto_save(bank_manager, asset_manager, transaction_manager)
        console.print("[bold green]Your data has been saved. You can refer to the latest files. Thank you![/bold green]")
        time.sleep(2)
