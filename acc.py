import datetime
from openpyxl import Workbook, load_workbook
import json
import os

def clear_screen():
    if os.name == 'posix':
        os.system("clear")
    else:
        os.system("cls")

def record_transaction(amount, description, transaction_type, bank, transactions_sheet):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    transactions_sheet.append([timestamp, transaction_type, amount, description, bank])

def total_bank(bank_accounts):
    return sum(bank_accounts.values())

def generate_financial_statements():
    # Logic to read from transactions file and generate statements
    pass

def late_bank(bank_accounts):
    with open("bank_accounts.json", "w") as file:
        json.dump(bank_accounts, file)

def load_bank():
    if os.path.exists("bank_accounts.json"):
        with open("bank_accounts.json", "r") as file:
            return json.load(file)
    else:
        return None

def main():
    clear_screen()
    print("Welcome to Finance Management\n")
    
    # Load bank account from latest file
    bank_accounts = load_bank()
    if not bank_accounts:
        bank_accounts = {
            "RHB": float(input("Enter your latest amount in RHB: ")),
            "Bank Islam": float(input("Enter your latest amount in Bank Islam: ")),
            "Cimb Bank": float(input("Enter your latest amount in Cimb Bank: "))
        }
        
        other_bank = input("Do you have any other bank? [y/n]: ")
        if other_bank.lower() == "y":
            bank_name = input("What is your bank? : ")
            bank_accounts[bank_name] = float(input(f"Enter your latest amount in {bank_name}: "))

    total_income = 0.0
    total_expenses = 0.0

    print("\nTotal Amount: ", total_bank(bank_accounts))
    for bank, amount in bank_accounts.items():
        print(f"{bank}: {amount}")

    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Transaction Type", "Amount", "Description", "Bank"])

    try:
        while True:
            new = datetime.datetime.now().strftime("%Y-%m-%d")
            transaction_type = input("\nIncome [I] or Expense [E] and 'save' to save file: ").upper()
            if transaction_type == 'SAVE':
                last_row = ws.max_row
                ws[f"C{last_row + 1}"] = f"=SUM(C2:C{last_row})"
                wb.save(f"{new}.xlsx")  # name of transaction file when saving
                print(f"File saved as {new}.xlsx")
                late_bank(bank_accounts)
                break
            description = input("Enter the transaction activities: ")
            amount_input = input("Enter the transaction amount: ")
            
            try:
                amount = float(amount_input)
            except ValueError:
                print("Invalid input. Please enter a valid number.")
                continue

            if transaction_type == "E":
                print("[1] RHB Bank")
                print("[2] Bank Islam")
                print("[3] CIMB Bank")
                print("[4] OTHERS")
                loc_bank = input("Which bank did you use? (Enter the bank name or number): ")
                if loc_bank.isdigit():
                    loc_bank = int(loc_bank)
                    if loc_bank == 1:
                        selected_bank = "RHB"
                    elif loc_bank == 2:
                        selected_bank = "Bank Islam"
                    elif loc_bank == 3:
                        selected_bank = "Cimb Bank"
                    elif loc_bank == 4:
                        selected_bank = input("Enter the bank name: ")
                    else:
                        print("Invalid selection. Please try again.")
                        continue
                else:
                    selected_bank = loc_bank
                
                if selected_bank not in bank_accounts:
                    print("Bank not found. Please try again.")
                    continue

                if bank_accounts[selected_bank] < amount:
                    print("Insufficient funds in the selected bank account.")
                    continue

                bank_accounts[selected_bank] -= amount
                record_transaction(amount, description, "Expense", selected_bank, ws)
                total_expenses += amount
                print("Expense recorded.")
                
                if total_expenses > 200:
                    print("You have spent more than RM 200. Consider saving money.")
            elif transaction_type == "I":
                record_transaction(amount, description, "Income", "N/A", ws)
                total_income += amount
                print("Income recorded.")
            else:
                print("Invalid transaction type.")
            
            print("\nUpdated Account Balances:")
            for bank, amount in bank_accounts.items():
                print(f"{bank}: {amount}")

            total_amount = total_bank(bank_accounts)
            print(f"Total Amount: {total_amount}")

    except KeyboardInterrupt:
        print("\n\nTransaction recording interrupted. Saving data to file...")
        wb.save("transactions.xlsx")
        print("Data saved successfully.")

    # Display total income and total expenses
    print("\nTotal Income:", total_income)
    print("Total Expenses:", total_expenses)

    # Generate financial statements
    generate_financial_statements()

if __name__ == "__main__":
    main()
