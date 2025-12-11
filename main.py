import argparse
import multiprocessing
import os
import time
import psutil
import traceback
import logging
from pathlib import Path
from src.utils import load_config, get_output_path, ensure_dir, copy_to_enhanced
from src.converter import ExcelConverter
from src.logger import setup_logger, log_error, log_info, get_queue_logger
from src.ui import create_progress_instance, create_layout, LogConsole, print_summary, print_banner, save_summary_report, Live
from src.language_detector import LanguageDetector

def convert_worker(input_path, output_path, config, pid_queue, log_queue, lang_code=None):
    """
    Worker function to run conversion in a separate process.
    
    Args:
        input_path: Input Excel file path
        output_path: Output PDF file path
        config: Configuration dictionary
        pid_queue: Queue for Excel process PID
        log_queue: Queue for logging
        lang_code: Detected language code (optional, for logging)
    """
    try:
        # Setup logging for worker to send to main process
        get_queue_logger(log_queue)
        
        if lang_code:
            logging.info(f"[{Path(input_path).name}] Language: {lang_code}")
        
        converter = ExcelConverter(config)
        converter.convert(input_path, output_path, pid_queue)
    except Exception:
        # Errors are logged in converter, but we raise to signal failure to parent
        raise

def kill_process_tree(pid):
    """
    Kills a process and its children.
    """
    try:
        parent = psutil.Process(pid)
        for child in parent.children(recursive=True):
            child.kill()
        parent.kill()
    except psutil.NoSuchProcess:
        pass

class UIHandler(logging.Handler):
    """
    Logging handler that sends messages to the LogConsole UI.
    """
    def __init__(self, log_console):
        super().__init__()
        self.log_console = log_console
        self.setFormatter(logging.Formatter('%(asctime)s - %(message)s', datefmt='%H:%M:%S'))

    def emit(self, record):
        try:
            msg = self.format(record)
            style = "white"
            if record.levelno >= logging.ERROR:
                style = "red"
            elif record.levelno >= logging.WARNING:
                style = "yellow"
            elif record.levelno == logging.INFO:
                style = "green"
            
            self.log_console.add_log(f"[{style}]{msg}[/{style}]")
        except Exception:
            self.handleError(record)

def main():
    parser = argparse.ArgumentParser(description="Excel to PDF Converter")
    parser.add_argument("--input", default="./input", help="Input directory containing Excel files")
    parser.add_argument("--output", default="./output", help="Output directory for PDF files")
    parser.add_argument("--config", default="config.yaml", help="Path to configuration file")
    args = parser.parse_args()

    # Load Config
    try:
        config = load_config(args.config)
    except Exception as e:
        print(f"Error loading config: {e}")
        return

    # Prepare paths
    input_root = os.path.abspath(args.input)
    output_root = os.path.abspath(args.output)
    
    # Enhanced root (optional) - used when prepare_for_print is enabled
    prepare_for_print = config.get('excel', {}).get('prepare_for_print', True)
    enhanced_dir_name = config.get('excel', {}).get('enhanced_dir', 'enhanced_files')
    enhanced_root = os.path.abspath(enhanced_dir_name) if prepare_for_print else None

    # Scan files first
    print("Scanning files...")
    
    # Directories to exclude from scan to prevent loops or processing intermediate files
    # We primarily want to exclude the enhanced_files directory because it contains copy of xlsx files
    exclude_dirs = set()
    if enhanced_root:
        exclude_dirs.add(enhanced_root)
    
    # Also exclude hidden directories and venv
    
    excel_files = []
    for root, dirs, files in os.walk(input_root):
        # Filter directories in-place
        dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(root, d)) not in exclude_dirs and not d.startswith('.')]
        
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx', '.xlsm')):
                if not file.startswith('~$'): # Ignore temp files
                    # Check against excluded dirs for current root
                    if not any(os.path.abspath(root).startswith(ex_dir) for ex_dir in exclude_dirs):
                        excel_files.append(os.path.join(root, file))

    total_files = len(excel_files)
    
    # Setup UI Components
    log_console_lines = config.get('logging', {}).get('log_console_lines', 20)
    log_console = LogConsole(max_lines=log_console_lines)
    progress = create_progress_instance()
    layout = create_layout(progress, log_console)

    # Setup Logging
    log_file = config.get('logging', {}).get('log_file', 'conversion.log')
    error_file = config.get('logging', {}).get('error_file', 'errors.log')
    log_level = config.get('logging', {}).get('log_level', 'INFO')
    
    root_logger = setup_logger(log_file, error_file, log_level)
    
    # Attach UI Handler
    root_logger.addHandler(UIHandler(log_console))

    print_banner()
    
    # Initialize language detector
    lang_detector = LanguageDetector(config)
    if lang_detector.is_enabled():
        log_info("Language classification enabled")
        log_info(f"Mode: {lang_detector.mode}")

    success_count = 0
    error_count = 0
    skipped_count = 0
    error_files_list = []
    lang_distribution = {}  # Track files by language

    timeout_minutes = config.get('timeout_minutes', 45)
    timeout_seconds = timeout_minutes * 60

    log_queue = multiprocessing.Queue()

    with Live(layout, refresh_per_second=10):
        task = progress.add_task("[cyan]Converting...", total=total_files)
        
        for input_path in excel_files:
            # Check queue before starting (optional)
            
            file_to_convert = input_path
            
            if prepare_for_print:
                progress.update(task, description=f"[cyan]Preparing: {os.path.basename(input_path)}")
                # We can log this too
                log_info(f"[{os.path.basename(input_path)}] Preparing for print")
                
                try:
                    file_to_convert = copy_to_enhanced(input_path, input_root, enhanced_root)
                except Exception as e:
                    log_error(input_path, f"Failed to copy/prepare: {e}")
                    error_count += 1
                    error_files_list.append(f"{input_path} (Prepare Failed)")
                    progress.advance(task)
                    continue

            # Language detection and classification
            lang_code = 'other'
            if lang_detector.is_enabled():
                try:
                    # For filename mode, detect before opening workbook
                    if lang_detector.mode == 'filename':
                        lang_code = lang_detector.classify_file(input_path)
                        output_path = lang_detector.get_output_path(input_path, output_root, lang_code)
                    else:
                        # For auto mode, we need to detect from content (will be done in worker if needed)
                        # For now, use filename as fallback
                        lang_code = lang_detector.detect_language_from_filename(Path(input_path).stem)
                        output_path = lang_detector.get_output_path(input_path, output_root, lang_code)
                    
                    # Track language distribution
                    lang_distribution[lang_code] = lang_distribution.get(lang_code, 0) + 1
                    
                except Exception as e:
                    log_error(input_path, f"Language detection failed: {e}")
                    output_path = get_output_path(input_path, input_root, output_root, config.get('output_suffix', '_x'))
            else:
                output_path = get_output_path(input_path, input_root, output_root, config.get('output_suffix', '_x'))
            
            progress.update(task, description=f"[cyan]Processing: {os.path.basename(input_path)}")
            # Log separation
            # log_info(input_path, "Starting conversion")
            
            pid_queue = multiprocessing.Queue()
            
            p = multiprocessing.Process(target=convert_worker, args=(file_to_convert, output_path, config, pid_queue, log_queue, lang_code))
            p.start()
            
            start_time = time.time()
            excel_pid = None
            
            while p.is_alive():
                # Drain Log Queue
                while not log_queue.empty():
                    try:
                        record = log_queue.get_nowait()
                        # Handling record in root logger will trigger File Handlers AND UIHandler
                        root_logger.handle(record)
                    except:
                        break
                
                # Check for PID (one time)
                if excel_pid is None and not pid_queue.empty():
                    try:
                        excel_pid = pid_queue.get_nowait()
                    except:
                        pass
                
                # Check Timeout
                if time.time() - start_time > timeout_seconds:
                    log_error(input_path, f"Timeout after {timeout_minutes} minutes")
                    p.terminate()
                    p.join(timeout=5)
                    if p.is_alive():
                        p.kill() # Force kill
                    
                    if excel_pid:
                        kill_process_tree(excel_pid)
                        
                    error_count += 1
                    error_files_list.append(f"{input_path} (Timeout)")
                    break
                
                time.sleep(0.05)
            
            # Process finished or killed.
            # Drain remaining logs
            while not log_queue.empty():
                try:
                    record = log_queue.get_nowait()
                    root_logger.handle(record)
                except:
                    break

            # Check exit code if not timeout (p.is_alive() is False here)
            if not p.is_alive(): 
                # If we killed it due to timeout, it's already handled.
                # But we need to distinguish timeout from crash vs success.
                # The loop break handles timeout. If loop finished naturally:
                if time.time() - start_time <= timeout_seconds:
                    p.join() # Ensure cleanup
                    if p.exitcode == 0:
                        success_count += 1
                        log_info(f"Successfully converted: {input_path}")
                    else:
                        # If exitcode is not 0 and NOT timeout
                        # Timeout logic above terminates it, so exitcode might be set.
                        # We should set a flag if timeout occurred.
                        # Refactor: use a flag.
                        pass # Already logged error in worker or above logic needs refinement
            
            # Small Refactor for robustness:
            # If timeout occurred, we broke the loop. 
            # If natural finish, we fell through.
            # We can check exitcode.
            
            # However, I need to know if I already counted it as error.
            # Let's check error_list for this file?
            # Or use a flag 'timeout_occurred'.
            
            if f"{input_path} (Timeout)" not in error_files_list:
                # Normal finish (success or crash)
                if p.exitcode == 0:
                    # Success already handled? No.
                    pass 
                else:
                     # Check if we already logged success? No.
                     # If exitcode != 0, it is error.
                     # Worker logs exception.
                     # We Count it.
                     if p.exitcode != 0:
                         error_count += 1
                         error_files_list.append(input_path)

            progress.advance(task)

    # Print language distribution if enabled
    if lang_detector.is_enabled() and lang_distribution:
        log_info("\n=== Language Distribution ===")
        for lang, count in sorted(lang_distribution.items()):
            log_info(f"{lang}: {count} files")
    
    print_summary(total_files, success_count, error_count, skipped_count, error_files_list)
    save_summary_report(total_files, success_count, error_count, skipped_count, error_files_list, lang_distribution if lang_detector.is_enabled() else None)

if __name__ == "__main__":
    multiprocessing.freeze_support() 
    main()
