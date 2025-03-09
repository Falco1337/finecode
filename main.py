# main.py
import os
import subprocess
import sys
import time
from rich.console import Console
from acca import run_application

console = Console()

REQUIRED_LIBRARIES = [
    "atexit",
    "datetime",
    "json",
    "os",
    "logging",
    "requests",
    "rich",
    "openpyxl"
]

BANK_ACCOUNTS_FILE = "bank_total.json"

def clear_screen():
    try:
        os.system("cls" if os.name == "nt" else "clear")
    except:
        print("\n" * 100)

def check_libraries():
    console.print("[bold cyan]Checking libraries...[/bold cyan]\n")
    time.sleep(5)
    missing_libraries = []
    for lib in REQUIRED_LIBRARIES:
        try:
            __import__(lib)
            console.print(f"{lib}... [bold green]Done[/bold green]")
            time.sleep(5)
        except ImportError:
            console.print(f"{lib}... [bold red]Not detected[/bold red]")
            missing_libraries.append(lib)
    return missing_libraries

def install_libraries(libraries):
    if not libraries:
        return
    console.print("[bold cyan]Installing missing libraries...[/bold cyan]")
    for lib in libraries:
        console.print(f"Installing {lib} please wait...")
        start_time = time.time()
        try:
            subprocess.check_call([sys.executable, "pip3", "-m", "install", lib])
            console.print(f"{lib}... [bold green]Done[/bold green]")
            time.sleep(5)
        except subprocess.CalledProcessError:
            console.print(f"{lib}... [bold red]Failed to install[/bold red]")
        elapsed_time = time.time() - start_time
        console.print(f"ETA: {time.strftime('%H:%M:%S', time.gmtime(elapsed_time))} minutes")

def check_bank_file():
    console.print("[bold cyan]Checking bank_total.json...[/bold cyan]")
    time.sleep(5)
    if os.path.exists(BANK_ACCOUNTS_FILE):
        console.print("[bold green]Found![/bold green]")
        return True
    else:
        console.print("[bold red]Not found![/bold red]")
        time.sleep(2)
        console.print("[bold green]Please Wait...[/bold green]")
        time.sleep(5)
        return False

def main():
    clear_screen()
    missing_libraries = check_libraries()
    if missing_libraries:
        install_libraries(missing_libraries)
    clear_screen()
    if check_bank_file():
        console.print("[bold green]Bank data detected. Proceeding to main application...[/bold green]")
    else:
        console.print("[bold yellow]No bank data found. Please update bank balances.[/bold yellow]")
    time.sleep(2)
    clear_screen()
    run_application()

if __name__ == "__main__":
    main()
