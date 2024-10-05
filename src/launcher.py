import subprocess
import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
from rich.panel import Panel
from rich.table import Table
import os
import winreg
from time import sleep
from pathlib import Path
import pandas as pd
import ctypes
import json
from configparser import ConfigParser
import sys

app = typer.Typer()
console = Console()

def display_ascii_logo():
    ascii_art = """
    ██████╗  █████╗ ████████╗██╗  ██╗    ███████╗██╗███╗   ██╗██████╗ ███████╗██████╗
    ██╔══██╗██╔══██╗╚══██╔══╝██║  ██║    ██╔════╝██║████╗  ██║██╔══██╗██╔════╝██╔══██╗
    ██████╔╝███████║   ██║   ███████║    █████╗  ██║██╔██╗ ██║██║  ██║█████╗  ██████╔╝
    ██╔═══╝ ██╔══██║   ██║   ██╔══██║    ██╔══╝  ██║██║╚██╗██║██║  ██║██╔══╝  ██╔══██╗
    ██║     ██║  ██║   ██║   ██║  ██║    ██║     ██║██║ ╚████║██████╔╝███████╗██║  ██║
    ╚═╝     ╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝    ╚═╝     ╚═╝╚═╝  ╚═══╝╚═════╝ ╚══════╝╚═╝  ╚═╝
    ~ PATHFINDER 6000 - Programmatic HOT STEAMY PATH REGISTRY INJECTOR ~
    """
    console.print(Panel(ascii_art, title="🔥 HOT STEAMY PATH Injection! 🔥", expand=False))

def load_config():
    config = ConfigParser()
    config_path = Path('config.ini')
    if config_path.exists():
        config.read(config_path)
        return config
    else:
        console.print("[yellow]Config file not found. Using default settings.[/yellow]")
        return None

def load_config_solutions():
    excel_path = Path('config_json_handling_strong_table.xlsx')
    json_path = Path('config.json')
    
    if excel_path.is_file():
        return pd.read_excel(excel_path, sheet_name='Sheet1')
    elif json_path.is_file():
        with open(json_path, 'r') as f:
            return json.load(f)
    else:
        console.print("[yellow]No configuration file found. Proceeding with default settings.[/yellow]")
        return None

def apply_config_solutions(config_data):
    if config_data is None:
        console.print("[yellow]Skipping configuration solutions as no valid file was loaded.[/yellow]")
        return
    
    if isinstance(config_data, pd.DataFrame):
        for _, row in config_data.iterrows():
            console.print(f"[yellow]Problem:[/yellow] {row['Problem']}")
            console.print(f"[green]Solution:[/green] {row['Solution with configparser & json']}")
            console.print(f"[blue]Impact:[/blue] {row['Impact on Customer Experience']}")
    elif isinstance(config_data, dict):
        for problem, solution in config_data.items():
            console.print(f"[yellow]Problem:[/yellow] {problem}")
            console.print(f"[green]Solution:[/green] {solution['solution']}")
            console.print(f"[blue]Impact:[/blue] {solution['impact']}")
    
    sleep(1)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except AttributeError:
        return False

def run_as_admin():
    if not is_admin():
        console.print("[yellow]Attempting to run with administrative privileges...[/yellow]")
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        except Exception as e:
            console.print(f"[red]Failed to elevate privileges: {e}[/red]")
            console.print("[red]Please run this script as an administrator.[/red]")
        sys.exit(0)

def registry_operations():
    if not is_admin():
        console.print("[red]Error: Administrative privileges are required to perform registry operations.[/red]")
        run_as_admin()
        return
    
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as reg_conn:
            console.print("[green]Connected to the registry.[/green]")
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, r"SOFTWARE\PathFinder6000\InstalledPaths", 0, winreg.KEY_WRITE) as reg_key:
                console.print("[green]Registry key created or opened successfully.[/green]")
                
                backup_file_path = Path.home() / "pathfinder_backup.reg"
                winreg.SaveKey(reg_key, str(backup_file_path))
                console.print(f"[green]Registry key saved to {backup_file_path}.[/green]")
                
                restore_file_path = backup_file_path
                winreg.LoadKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\PathFinder6000", str(restore_file_path))
                console.print(f"[green]Registry key loaded from {restore_file_path}.[/green]")
                
                sub_key_info = winreg.QueryInfoKey(reg_key)
                console.print(f"[green]Subkey count: {sub_key_info[0]} - Values count: {sub_key_info[1]}[/green]")
                
                winreg.FlushKey(reg_key)
                console.print("[green]Registry key changes flushed.[/green]")
                
                try:
                    winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, r"SOFTWARE\PathFinder6000\OldPath")
                    console.print("[green]Registry subkey deleted.[/green]")
                except FileNotFoundError:
                    console.print("[yellow]OldPath subkey not found. Skipping deletion.[/yellow]")
                
                env_str = "%SYSTEMROOT%"
                expanded_str = winreg.ExpandEnvironmentStrings(env_str)
                console.print(f"[green]Expanded string: {expanded_str}[/green]")
        
        console.print("[green]Registry operations completed successfully.[/green]")
    except Exception as e:
        console.print(f"[red]Registry operation failed: {e}[/red]")

def get_installed_software():
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    ]
    
    software_list = []
    
    for hive, reg_path in reg_paths:
        try:
            with winreg.OpenKey(hive, reg_path) as reg_key:
                for i in range(winreg.QueryInfoKey(reg_key)[0]):
                    try:
                        sub_key = winreg.EnumKey(reg_key, i)
                        sub_key_path = f"{reg_path}\\{sub_key}"
                        with winreg.OpenKey(hive, sub_key_path) as sub_key_open:
                            try:
                                display_name = winreg.QueryValueEx(sub_key_open, "DisplayName")[0]
                                install_location = winreg.QueryValueEx(sub_key_open, "InstallLocation")[0]
                                install_path = Path(install_location)
                                if install_path.exists() and install_path.is_dir():
                                    software_list.append((display_name, str(install_path)))
                                else:
                                    console.print(f"[yellow]Skipping {display_name}: Invalid or missing install path.[/yellow]")
                            except FileNotFoundError:
                                pass
                    except WindowsError:
                        continue
        except Exception as e:
            console.print(f"[red]Error accessing registry at {reg_path}: {e}[/red]")

    return software_list

def write_paths_to_registry(software_paths):
    reg_path = r"SOFTWARE\PathFinder6000\InstalledPaths"
    
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path) as reg_key:
            for name, path in software_paths:
                try:
                    winreg.SetValueEx(reg_key, name, 0, winreg.REG_SZ, path)
                except Exception as e:
                    console.print(f"[red]Error writing {name} to registry: {e}[/red]")
        console.print(f"[green]Successfully updated the registry with {len(software_paths)} software paths![/green]")
    except Exception as e:
        console.print(f"[red]Error updating registry: {e}[/red]")

@app.command()
def enumerate_installed_software():
    """Enumerates installed software and updates the registry."""
    
    display_ascii_logo()

    console.print("\n🔍 Enumerating installed software...")

    # Load configuration
    config = load_config()
    config_solutions = load_config_solutions()
    apply_config_solutions(config_solutions)

    # Check for admin privileges before proceeding
    if not is_admin():
        console.print("[red]Administrative privileges are required. Please run this script as an administrator.[/red]")
        run_as_admin()
        return

    # Display progress bar
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeElapsedColumn(),
    ) as progress:
        task = progress.add_task("[green]Scanning registry entries...", total=100)
        
        software_paths = get_installed_software()

        if not software_paths:
            console.print("[red]No software paths found![/red]")
            return

        for _ in range(100):
            sleep(0.05)
            progress.update(task, advance=1)

    # Writing the paths to the registry
    write_paths_to_registry(software_paths)

    # Run registry operations at the end
    registry_operations()

if __name__ == "__main__":
    app()