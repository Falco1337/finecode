import os
import datetime
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

os.system("clear")
def record_transaction(amount, description, transaction_type, transaction_sheet):
    date = datetime.datetime.now().strftime("%Y-%m-%d")
    time = datetime.datetime.now().strftime("%H:%M:%S")
    transaction_sheet.append([date, time, transaction_type, amount, description])

def total_bank(bank_accounts):
    return sum(bank_accounts.values())

def generate_financial_statements():
    pass

def save_bank_accounts(bank_accounts):
    with open("bank_accounts.json", "w") as file:
        json.dump(bank_accounts, file)

def load_bank_accounts():
    if os.path.exists("bank_accounts.json"):
        with open("bank_accounts.json", "r") as file:
            return json.load(file)
    else:
        return None

def main():
    print("\nWelcome to Finance Management")

    # Load bank accounts from file or initialize if file does not exist
    bank_accounts = load_bank_accounts()
    if bank_accounts is None:
        bank_accounts = {
            "RHB": float(input("RHB latest amount: ")),
            "Bank Islam": float(input("Bank Islam latest amount: ")),
            "Cimb Bank": float(input("Cimb Bank latest amount: "))
        }

        # Add other banks if any
        other_bank = input("Do you have any other bank? [y/n]: ")
        if other_bank.lower() == "y":
            bank_name = input("Bank Name: ")
            bank_accounts[bank_name] = float(input(f"Enter {bank_name} latest amount: "))

    total_income = 0
    total_expenses = 0

    print("\nTotal Amount:", total_bank(bank_accounts))
    for bank, amount in bank_accounts.items():
        print(f"{bank}: {amount}")

    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Time", "Income/Expenses", "Amount", "Description"])

    bank_options = ["BI (Bank Islam)", "RHB", "CB (CIMB)"]

    while True:
        new = datetime.datetime.now().strftime("[%Y-%m-%d]")
        transaction_type = input("\nIncome [I] / Expense [E] / 'save' to save file: ").upper()
        if transaction_type == "SAVE":
            # Add total formula to Excel sheet
            last_row = ws.max_row
            ws[f"D{last_row+1}"] = f"=SUM(D2:D{last_row})"
            ws[f"F{last_row+1}"] = f"=SUM(F2:F{last_row})"
            
            wb.save(f"{new}.xlsx")
            print(f"File saved as {new}.xlsx")
            save_bank_accounts(bank_accounts)
            break
        desc = input("Description of activities: ")
        amount_input = float(input(f"Amount for {transaction_type}: "))

        try:
            amount = float(amount_input)
        except ValueError:
            print("Invalid input! Please enter a valid number")
            continue

        if transaction_type.upper() == "E" or transaction_type.lower() == "expense":
            for account in bank_accounts:
                bank_accounts[account] -= amount
            record_transaction(amount, desc, "Expense", ws, total_expenses)
            total_expenses += amount
            print("Expense successfully recorded!")
            if total_expenses > 150:
                print("You have spent more than 150! Save your money!")
        elif transaction_type.upper() == "I" or transaction_type.lower() == "income":
            print("Select the bank to deposit the income:")
            for i, option in enumerate(bank_options, 1):
                print(f"{i}. {option}")

            while True:
                choice = input("Enter the number corresponding to your choice: ")
                if choice.isdigit():
                    choice = int(choice)
                    if 1 <= choice <= len(bank_options):
                        selected_bank = bank_options[choice - 1]
                        break
                    else:
                        print("Invalid choice! Please enter a number within the range.")
                else:
                    print("Invalid input! Please enter a number.")

            if selected_bank.startswith("BI"):
                bank_name = "Bank Islam"
            elif selected_bank.startswith("RHB"):
                bank_name = "RHB"
            elif selected_bank.startswith("CB"):
                bank_name = "Cimb Bank"
            
            if bank_name in bank_accounts:
                bank_accounts[bank_name] += amount
                record_transaction(amount, desc, "Income", ws, total_income)
                total_income += amount
                print("Income successfully recorded!")
            else:
                print("Invalid bank name!")
        else:
            print("Invalid transaction type!")

        total_amount = total_bank(bank_accounts)
        print(f"\n{bank_accounts}")
        print(f"Total amount: {total_amount}")
        print(f"Total Income: {total_income}")
        print(f"Total Expenses: {total_expenses}")

    generate_financial_statements()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nYou have to learn how to save money")
