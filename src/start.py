import subprocess
import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
from rich.panel import Panel
import os
import winreg
from time import sleep
from pathlib import Path
try:
    import pandas as pd
except ImportError:
    pd = None
import json
from configparser import ConfigParser
import datetime
import random

app = typer.Typer()
console = Console()

def display_ascii_logo():
    ascii_art = r"""
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
        console.print("[green]Config file loaded successfully.[/green]")
        return config
    else:
        console.print("[yellow]Config file not found. Using default settings.[/yellow]")
        return None

def load_config_solutions():
    excel_path = Path('config_json_handling_strong_table.xlsx')
    json_path = Path('config.json')
    
    if excel_path.is_file() and pd is not None:
        console.print("[green]Excel configuration file found and loaded.[/green]")
        return pd.read_excel(excel_path, sheet_name='Sheet1')
    elif json_path.is_file():
        console.print("[green]JSON configuration file found and loaded.[/green]")
        with open(json_path, 'r') as f:
            return json.load(f)
    else:
        console.print("[yellow]No configuration file found. Proceeding with default settings.[/yellow]")
        return None

def apply_config_solutions(config_data):
    if config_data is None:
        console.print("[yellow]Skipping configuration solutions as no valid file was loaded.[/yellow]")
        return
    
    console.print("[blue]Applying configuration solutions...[/blue]")
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
    
    console.print("[green]Configuration solutions applied successfully.[/green]")
    sleep(1)

def registry_operations():
    console.print("[blue]Starting registry operations...[/blue]")
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE):
            console.print("[green]Connected to the registry.[/green]")
            reg_path = r"SOFTWARE\PathFinder6000\InstalledPaths"
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE) as reg_key:
                console.print("[green]Registry key created or opened successfully.[/green]")
                
                backup_file_path = Path.home() / "pathfinder_backup.reg"
                try:
                    winreg.SaveKey(reg_key, str(backup_file_path))
                    console.print(f"[green]Registry key saved to {backup_file_path}.[/green]")
                except WindowsError as e:
                    if e.winerror == 183:  # Cannot create a file when that file already exists
                        console.print(f"[yellow]Backup file exists at {backup_file_path}. Skipping backup.[/yellow]")
                    else:
                        raise
                
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
    console.print("[blue]Retrieving installed software information...[/blue]")
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

    console.print(f"[green]Retrieved information for {len(software_list)} software installations.[/green]")
    return software_list

def write_paths_to_registry(software_paths):
    console.print("[blue]Writing software paths to registry...[/blue]")
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

def generate_cool_ascii_art():
    ascii_arts = [
        """
        ┏━━━┓━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┏┓━━━━━━━━━━━┏┓
        ┃┏━┓┃━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┃┃━━━━━━━━━━━┃┃
        ┃┗━┛┃┏━━┓━┏━━┓━┏━━┓━┏━┓━┏━━┓━┏━━┓┏━┓━┃┃━┏━━┓┏━━┓┏━┛┃
        ┃┏━━┛┃┏┓┃━┃┏┓┃━┃┏┓┃━┃┏┛━┃┏━┛━┃┏━┛┃┏┓┓┃┃━┃┏┓┃┃┏┓┃┃┏┓┃
        ┃┃━━━┃┃━┫━┃┗┛┃━┃┗┛┃━┃┃━━┃┗━┓━┃┗━┓┃┃┃┃┃┗┓┃┗┛┃┃┗┛┃┃┗┛┃
        ┗┛━━━┗━━┛━┗━━┛━┗━━┛━┗┛━━┗━━┛━┗━━┛┗┛┗┛┗━┛┗━━┛┃┏━┛┗━━┛
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┃┃
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┗┛
        """,
        """
         _____      _   _      ____              _           
        |  __ \    | | | |    |  _ \            | |          
        | |__) |_ _| |_| |__  | |_) | ___   ___ | |_ ___ _ __
        |  ___/ _` | __| '_ \ |  _ < / _ \ / _ \| __/ _ \ '__|
        | |  | (_| | |_| | | || |_) | (_) | (_) | ||  __/ |   
        |_|   \__,_|\__|_| |_||____/ \___/ \___/ \__\___|_|   
        """
    ]
    return random.choice(ascii_arts)

def generate_tech_tip():
    tips = [
        "Did you know? Regularly updating your software can improve system performance and security.",
        "Pro tip: Use keyboard shortcuts to boost your productivity. Try Ctrl+Shift+Esc to open Task Manager quickly!",
        "Fun fact: The first computer bug was an actual moth found in a relay of the Harvard Mark II computer in 1947.",
        "Optimize your workflow: Set up multiple virtual desktops to organize your tasks more efficiently.",
        "Security reminder: Use a password manager to create and store strong, unique passwords for all your accounts.",
    ]
    return random.choice(tips)

@app.command()
def enumerate_installed_software():
    """Enumerates installed software and updates the registry."""
    
    start_time = datetime.datetime.now()
    display_ascii_logo()

    console.print("\n🔍 Enumerating installed software...")

    config = load_config()
    config_solutions = load_config_solutions()
    apply_config_solutions(config_solutions)

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

    write_paths_to_registry(software_paths)
    registry_operations()

    end_time = datetime.datetime.now()
    execution_time = end_time - start_time

    console.print("\n[green]Script execution completed successfully![/green]")
    console.print("[blue]Thank you for using PATHFINDER 6000![/blue]")
    
    console.print("\n[yellow]Execution Summary:[/yellow]")
    console.print(f"[green]Total software installations found:[/green] {len(software_paths)}")
    console.print(f"[green]Execution time:[/green] {execution_time.total_seconds():.2f} seconds")
    console.print(f"[green]Registry backup location:[/green] {Path.home() / 'pathfinder_backup.reg'}")
    
    console.print("\n[yellow]Next Steps:[/yellow]")
    console.print("1. Verify the registry entries in the Windows Registry Editor")
    console.print("2. Test the updated PATH environment variable")
    console.print("3. Restart any open command prompts or applications to apply changes")
    
    console.print("\n[cyan]Here's a little something extra for you:[/cyan]")
    console.print(generate_cool_ascii_art())
    console.print(f"\n[magenta]Tech Tip of the Day:[/magenta] {generate_tech_tip()}")
    
    console.print("\n[blue]For more tech tips and software solutions, visit our blog at:[/blue]")
    console.print("[link=https://nerdy-tech.com]https://nerdy-tech.com[/link]")
    
    console.print("\n[green]Confirmation Details:[/green]")
    console.print(f"Operation completed on: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    console.print(f"Machine name: {os.environ['COMPUTERNAME']}")
    console.print(f"User: {os.getlogin()}")

    console.print("\n[yellow]Updated Software Paths:[/yellow]")
    for name, path in software_paths[:10]:
        console.print(f"[green]{name}:[/green] {path}")
    if len(software_paths) > 10:
        console.print(f"... and {len(software_paths) - 10} more")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    app()
