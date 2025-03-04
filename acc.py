import datetime
import json
import os
import logging
import time
import pandas as pd
import matplotlib.pyplot as plt
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

ALPHA_VANTAGE_API_KEY = "Y3PQLKE4QF2U56GV."
ALPHA_VANTAGE_URL = 'https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=IBM&interval=5min&apikey=ALPHA_VANTAGE_API_KEY'

logging.basicConfig(level=logging.INFO, format=LOGGING_FORMAT)
console = Console()

def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")

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

def display_balance(bank_accounts, assets):
    total_bank_balance = sum(bank_accounts.values())
    total_assets = sum(assets.values())
    total_balance = total_bank_balance + total_assets
    
    clear_screen()
    console.print(f"\n[bold cyan]Total Balance: RM {total_balance:.2f}[/bold cyan]\n")
    
    bank_table = Table(title="\U0001F4B3 Bank Accounts")
    bank_table.add_column("Bank", justify="left", style="cyan")
    bank_table.add_column("Balance", justify="right", style="green")
    
    for bank, amount in bank_accounts.items():
        bank_table.add_row(bank, f"RM {amount:.2f}")
    
    asset_table = Table(title="\U0001F4B8 Assets")
    asset_table.add_column("Asset", justify="left", style="yellow")
    asset_table.add_column("Value", justify="right", style="green")
    
    for asset, value in assets.items():
        asset_table.add_row(asset, f"RM {value:.2f}")
    
    console.print(bank_table)
    console.print(asset_table)

def generate_reports():
    transactions = load_json_data(TRANSACTIONS_FILE, [])
    if not transactions:
        console.print("[bold yellow]No transactions found.[/bold yellow]")
        return
    
    df = pd.DataFrame(transactions)
    if "bank" not in df.columns or "amount" not in df.columns:
        console.print("[bold red]Error: Invalid transaction data format.[/bold red]")
        return
    
    bank_usage = df.groupby("bank")["amount"].sum()
    
    console.print("\n[bold cyan]Bank Usage Summary:[/bold cyan]")
    for bank, amount in bank_usage.items():
        console.print(f"{bank}: RM {amount:.2f}")
    
    # Save to Excel
    df.to_excel("account_report.xlsx", index=False)
    console.print("[bold green]Report saved as account_report.xlsx[/bold green]")
    
    # Bar Chart for Bank Usage
    plt.bar(bank_usage.index, bank_usage.values, color='blue')
    plt.xlabel("Banks")
    plt.ylabel("Total Usage (RM)")
    plt.title("Bank Usage Report")
    plt.xticks(rotation=45)
    plt.show()

def fetch_stock_price(symbol):
    if symbol.upper() == "USD/MYR":
        url = "https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=USD&to_currency=MYR&apikey=Y3PQLKE4QF2U56GV"
    elif symbol.upper() == "MYR/USD":
        url = "https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=MYR&to_currency=USD&apikey=Y3PQLKE4QF2U56GV"
    else:
        url = f"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={symbol}&interval=5min&apikey=Y3PQLKE4QF2U56GV"
    
    response = requests.get(url)
    
    try:
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

def process_transaction(transaction_type, bank_accounts, transactions):
    bank = Prompt.ask(f"Enter bank for {transaction_type}:")
    amount = float(Prompt.ask(f"Enter amount for {transaction_type}:"))
    
    if transaction_type == "Expense" and bank_accounts.get(bank, 0) < amount:
        console.print("[bold red]Error: Insufficient funds![/bold red]")
        return
    
    bank_accounts[bank] = bank_accounts.get(bank, 0) + amount if transaction_type == "Income" else bank_accounts.get(bank, 0) - amount
    transactions.append({"bank": bank, "amount": amount, "type": transaction_type, "date": datetime.datetime.now().strftime(DATE_FORMAT)})
    
    save_json_data(BANK_ACCOUNTS_FILE, bank_accounts)
    save_json_data(TRANSACTIONS_FILE, transactions)
    console.print(f"[bold green]{transaction_type} recorded successfully![/bold green]")

def add_asset(assets):
    asset_name = Prompt.ask("Enter asset name:")
    asset_value = float(Prompt.ask("Enter asset value:"))
    
    assets[asset_name] = assets.get(asset_name, 0) + asset_value
    save_json_data(ASSETS_FILE, assets)
    console.print("[bold green]Asset added successfully![/bold green]")

def main():
    clear_screen()
    console.print("[bold green]\n\U0001F31F Accountant Assistant System 2025 \U0001F31F\n[/bold green]")
    
    bank_accounts = load_json_data(BANK_ACCOUNTS_FILE, {})
    assets = load_json_data(ASSETS_FILE, {"ASNB": 981.47, "Gold": 435.60})
    transactions = load_json_data(TRANSACTIONS_FILE, [])
    
    while True:
        command = Prompt.ask("\nChoose an option:\n[B] Balance\n[I] Income\n[E] Expense\n[A] Add Asset\n[R] Reports\n[T] World Trends\n[Q] Quit", default="Q").strip().upper()
        if command == "B":
            display_balance(bank_accounts, assets)
        elif command == "I":
            process_transaction("Income", bank_accounts, transactions)
        elif command == "E":
            process_transaction("Expense", bank_accounts, transactions)
        elif command == "A":
            add_asset(assets)
        elif command == "R":
            generate_reports()
        elif command == "T":
            fetch_world_trends()
        elif command == "Q":
            print("Good Bye!")
            clear_screen()
            time.sleep(1)
            break

if __name__ == "__main__":
    main()
