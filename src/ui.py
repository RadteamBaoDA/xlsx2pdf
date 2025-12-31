from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn, TimeRemainingColumn, MofNCompleteColumn
from rich.table import Table
from rich.panel import Panel
import datetime
from pathlib import Path

from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn, TimeRemainingColumn, MofNCompleteColumn
from rich.table import Table
from rich.panel import Panel
from rich.layout import Layout
from rich.live import Live
from rich.console import Group
from rich.align import Align
from collections import deque
import datetime

console = Console()

class LogConsole:
    def __init__(self, max_lines=20):
        self.max_lines = max_lines
        self.logs = deque(maxlen=max_lines)
    
    def add_log(self, message):
        self.logs.append(message)
        
    def __rich__(self):
        return Panel(
            "\n".join(self.logs),
            title="Log Console",
            border_style="cyan",
            height=self.max_lines + 2
        )

def create_layout(progress, log_console):
    # Use Group to stack progress and logs tightly without filling entire screen height
    return Group(
        progress,
        "\n", # Add a small spacer
        log_console
    )

def create_progress_instance():
    # Return a Progress instance without starting it (no context manager)
    return Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        MofNCompleteColumn(),
        TextColumn("• ETA:"),
        TimeRemainingColumn(),
    )

def print_summary(total, success, error, skipped, error_files):
    table = Table(title="Conversion Summary")
    table.add_column("Metric", style="cyan")
    table.add_column("Count", style="magenta")

    table.add_row("Total Files", str(total))
    table.add_row("Successful", str(success))
    table.add_row("Errors", str(error))
    table.add_row("Skipped", str(skipped))

    console.print(table)

    if error_files:
        console.print("\n[bold red]Error Files:[/bold red]")
        for f in error_files:
            console.print(f"- {f}")

def save_summary_report(total, success, error, skipped, error_files, lang_distribution=None, output_file="summary_report.txt", logs_folder="logs"):
    """Saves the summary report to a timestamped text file in the logs folder."""
    try:
        # Create logs folder if it doesn't exist
        logs_path = Path(logs_folder)
        logs_path.mkdir(parents=True, exist_ok=True)
        
        # Create timestamped filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        base_path = Path(output_file)
        name_part = base_path.stem
        ext_part = base_path.suffix
        timestamped_filename = f"{name_part}_{timestamp}{ext_part}"
        full_path = logs_path / timestamped_filename
        
        with open(full_path, "w", encoding="utf-8") as f:
            f.write("Excel to PDF Conversion Summary Report\n")
            f.write(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("-" * 40 + "\n")
            f.write(f"Total Files : {total}\n")
            f.write(f"Successful  : {success}\n")
            f.write(f"Errors      : {error}\n")
            f.write(f"Skipped     : {skipped}\n")
            f.write("-" * 40 + "\n")
            
            if lang_distribution:
                f.write("\nLanguage Distribution:\n")
                for lang, count in sorted(lang_distribution.items()):
                    f.write(f"  {lang}: {count} files\n")
                f.write("-" * 40 + "\n")
            
            if error_files:
                f.write("\nError Files:\n")
                for line in error_files:
                    f.write(f"- {line}\n")
        
        console.print(f"\n[dim]Summary report saved to: {full_path}[/dim]")
        return str(full_path)
    except Exception as e:
        console.print(f"[red]Error saving summary report: {e}[/red]")
        return None

def print_banner():
    console.print(Panel.fit("[bold blue]Excel to PDF Converter[/bold blue]", border_style="blue"))
