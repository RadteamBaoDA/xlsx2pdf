"""
Scenario Mode Converter - Convert files using scenario configuration

This script allows you to convert multiple folders with different settings
using a scenario YAML file that defines groups of folders and their configs.

Usage:
    python scenario_converter.py --scenario scenarios/example_scenario.yaml
    python scenario_converter.py --scenario scenarios/simple_scenario.yaml --dry-run
"""

import argparse
import multiprocessing
import os
import time
import logging
from pathlib import Path
from src.core.scenario_manager import ScenarioManager
from src.interface import OfficeConverter
from src.core.logger import setup_logger, log_error, log_info, get_queue_logger
from src.ui import create_progress_instance, create_layout, LogConsole, print_banner, Live, UIHandler


def convert_worker(input_path, output_path, config, pid_queue, log_queue):
    """Worker function for file conversion."""
    try:
        get_queue_logger(log_queue)
        
        # Use OfficeConverter with the group's config
        converter = OfficeConverter(config)
        result = converter.convert(input_path, output_path, pid_queue)
        
        if not result.success:
            raise Exception(result.error or "Conversion failed")
            
    except Exception:
        raise


def kill_process_tree(pid):
    """Kill process and all children."""
    try:
        import psutil
        parent = psutil.Process(pid)
        for child in parent.children(recursive=True):
            try:
                child.kill()
            except:
                pass
        parent.kill()
    except:
        pass


def main():
    parser = argparse.ArgumentParser(
        description="Scenario-based Office to PDF Converter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert using scenario file
  python scenario_converter.py --scenario scenarios/example_scenario.yaml
  
  # Dry run to see what would be converted
  python scenario_converter.py --scenario scenarios/simple_scenario.yaml --dry-run
  
  # Filter by file type
  python scenario_converter.py --scenario scenarios/example_scenario.yaml --file-types excel
        """
    )
    parser.add_argument("--scenario", required=True, help="Path to scenario YAML file")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be converted without converting")
    parser.add_argument("--file-types", default="all", help="File types to convert: all, excel, word, powerpoint")
    args = parser.parse_args()

    # Load scenario
    try:
        scenario = ScenarioManager(args.scenario)
    except Exception as e:
        print(f"Error loading scenario: {e}")
        return

    # Print scenario summary
    print("\n" + "="*70)
    print(scenario.get_scenario_summary())
    print("="*70 + "\n")

    # Determine file extensions to scan
    file_type_map = {
        'excel': ['.xlsx', '.xls', '.xlsm', '.xlsb'],
        'word': ['.docx', '.doc', '.docm', '.dotx', '.dotm'],
        'powerpoint': ['.pptx', '.ppt', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm']
    }
    
    if args.file_types.lower() == 'all':
        scan_extensions = None  # All Office files
    else:
        requested_types = [t.strip().lower() for t in args.file_types.split(',')]
        scan_extensions = [ext for file_type in requested_types 
                          if file_type in file_type_map 
                          for ext in file_type_map[file_type]]

    # Collect files from all groups
    files_with_groups = scenario.get_all_files(scan_extensions)
    
    if not files_with_groups:
        print("No files found to convert.")
        return
    
    print(f"Found {len(files_with_groups)} file(s) to convert\n")
    
    # Group summary
    group_counts = {}
    for _, group in files_with_groups:
        group_counts[group.name] = group_counts.get(group.name, 0) + 1
    
    print("Files per group:")
    for group_name, count in group_counts.items():
        print(f"  {group_name}: {count} files")
    print()

    # Dry run mode
    if args.dry_run:
        print("DRY RUN - No files will be converted\n")
        print("Files to be processed:")
        for file_path, group in files_with_groups:
            file_type = scenario.get_file_type(file_path)
            output_path = scenario.get_output_path_for_file(file_path, group)
            print(f"  [{group.name}] [{file_type}] {file_path}")
            print(f"    -> {output_path}")
        print(f"\nTotal: {len(files_with_groups)} files")
        return

    # Setup UI Components
    log_console = LogConsole(max_lines=20)
    progress = create_progress_instance()
    layout = create_layout(progress, log_console)

    # Setup Logging
    root_logger, actual_log_file, actual_error_file = setup_logger(
        'scenario_conversion.log',
        'scenario_errors.log',
        'INFO',
        'logs'
    )
    
    root_logger.addHandler(UIHandler(log_console))

    print_banner()
    log_info(f"Starting scenario conversion: {scenario.scenario_data.get('name', 'Unnamed')}")

    success_count = 0
    error_count = 0
    error_files_list = []
    group_stats = {}  # Track stats per group

    timeout_seconds = 45 * 60  # Default timeout
    log_queue = multiprocessing.Queue()

    with Live(layout, refresh_per_second=10):
        task = progress.add_task("[cyan]Converting...", total=len(files_with_groups))
        
        for file_path, group in files_with_groups:
            # Get file type and output path
            file_type = scenario.get_file_type(file_path)
            output_path = scenario.get_output_path_for_file(file_path, group)
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Get timeout from group config
            timeout_seconds = group.config.get('timeout_minutes', 45) * 60
            
            progress.update(task, description=f"[cyan][{group.name}] Processing: {os.path.basename(file_path)}")
            
            pid_queue = multiprocessing.Queue()
            
            # Start conversion worker
            p = multiprocessing.Process(
                target=convert_worker,
                args=(file_path, output_path, group.config, pid_queue, log_queue)
            )
            p.start()
            
            start_time = time.time()
            office_pid = None
            
            while p.is_alive():
                # Drain log queue
                while not log_queue.empty():
                    try:
                        record = log_queue.get_nowait()
                        root_logger.handle(record)
                    except:
                        break
                
                # Get PID
                if office_pid is None and not pid_queue.empty():
                    try:
                        office_pid = pid_queue.get_nowait()
                    except:
                        pass
                
                # Check timeout
                if time.time() - start_time > timeout_seconds:
                    log_error(file_path, f"Timeout after {timeout_seconds/60:.0f} minutes")
                    p.terminate()
                    p.join(timeout=5)
                    if p.is_alive():
                        p.kill()
                    
                    if office_pid:
                        kill_process_tree(office_pid)
                    
                    error_count += 1
                    error_files_list.append(f"[{group.name}] {file_path} (Timeout)")
                    
                    # Update group stats
                    if group.name not in group_stats:
                        group_stats[group.name] = {'success': 0, 'error': 0}
                    group_stats[group.name]['error'] += 1
                    break
                
                time.sleep(0.05)
            
            # Drain remaining logs
            while not log_queue.empty():
                try:
                    record = log_queue.get_nowait()
                    root_logger.handle(record)
                except:
                    break
            
            # Check result
            if not p.is_alive() and time.time() - start_time <= timeout_seconds:
                p.join()
                if p.exitcode == 0:
                    success_count += 1
                    log_info(f"[{group.name}] Successfully converted: {os.path.basename(file_path)}")
                    
                    # Update group stats
                    if group.name not in group_stats:
                        group_stats[group.name] = {'success': 0, 'error': 0}
                    group_stats[group.name]['success'] += 1
                else:
                    error_count += 1
                    error_files_list.append(f"[{group.name}] {file_path}")
                    
                    # Update group stats
                    if group.name not in group_stats:
                        group_stats[group.name] = {'success': 0, 'error': 0}
                    group_stats[group.name]['error'] += 1
            
            progress.advance(task)

    # Print summary
    total_files = len(files_with_groups)
    print("\n" + "="*70)
    print("CONVERSION SUMMARY")
    print("="*70)
    print(f"Total files: {total_files}")
    print(f"Successful: {success_count}")
    print(f"Failed: {error_count}")
    print(f"Success rate: {(success_count/total_files*100):.1f}%")
    
    print("\nResults by group:")
    for group_name, stats in group_stats.items():
        total = stats['success'] + stats['error']
        success_rate = (stats['success'] / total * 100) if total > 0 else 0
        print(f"  {group_name}: {stats['success']}/{total} succeeded ({success_rate:.1f}%)")
    
    if error_files_list:
        print(f"\nFailed files ({len(error_files_list)}):")
        for error_file in error_files_list[:10]:  # Show first 10
            print(f"  - {error_file}")
        if len(error_files_list) > 10:
            print(f"  ... and {len(error_files_list) - 10} more")
    
    print("="*70)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
