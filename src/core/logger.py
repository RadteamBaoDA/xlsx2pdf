import logging
import os
from datetime import datetime
from pathlib import Path
from logging.handlers import RotatingFileHandler, QueueHandler

def create_timestamped_filename(base_name, logs_folder="logs"):
    """
    Creates a timestamped filename in the specified logs folder.
    Format: base_name_yyyymmddhhmmss.ext
    """
    # Ensure logs folder exists
    logs_path = Path(logs_folder)
    logs_path.mkdir(parents=True, exist_ok=True)
    
    # Create timestamp
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    
    # Split filename and extension
    base_path = Path(base_name)
    name_part = base_path.stem
    ext_part = base_path.suffix
    
    # Create timestamped filename
    timestamped_name = f"{name_part}_{timestamp}{ext_part}"
    
    return logs_path / timestamped_name

def setup_logger(log_file, error_file, log_level="INFO", logs_folder="logs"):
    """
    Sets up the logger with file handlers for general logs and errors.
    Creates timestamped log files in the specified logs folder.
    """
    level_map = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR
    }
    level = level_map.get(str(log_level).upper(), logging.INFO)

    # Create timestamped log file paths
    log_file_path = create_timestamped_filename(log_file, logs_folder)
    error_file_path = create_timestamped_filename(error_file, logs_folder)

    # Create handlers
    log_handler = RotatingFileHandler(log_file_path, maxBytes=5*1024*1024, backupCount=2, encoding='utf-8')
    error_handler = RotatingFileHandler(error_file_path, maxBytes=5*1024*1024, backupCount=2, encoding='utf-8')
    
    # Create formatters and add it to handlers
    log_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    log_handler.setFormatter(log_format)
    error_handler.setFormatter(log_format)
    
    # Set levels for handlers
    # File log essentially captures everything equal or above the Configured Level
    log_handler.setLevel(level)
    error_handler.setLevel(logging.ERROR)
    
    # Get the root logger
    logger = logging.getLogger()
    logger.setLevel(level) 
    
    # Clear existing handlers to prevent duplication if called multiple times
    if logger.hasHandlers():
        logger.handlers.clear()

    # Add handlers to the logger
    logger.addHandler(log_handler)
    logger.addHandler(error_handler)
    
    # Log the file paths for user reference
    print(f"Logging to: {log_file_path}")
    print(f"Error logging to: {error_file_path}")
    
    return logger, str(log_file_path), str(error_file_path)

def get_queue_logger(queue):
    """
    Configures the root logger to send records to a multiprocessing queue.
    Used by worker processes.
    """
    logger = logging.getLogger()
    # Clear existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()
        
    logger.setLevel(logging.DEBUG) # Send everything to queue, let listener filter? 
    # Or strict level? Let's use INFO default for now, config passed later?
    # Actually, worker will assume parent set the level? No, new process.
    # We should pass config level to worker setup.
    
    handler = QueueHandler(queue)
    logger.addHandler(handler)
    return logger

def log_error(file_path, error_msg):
    logging.error(f"Failed to convert {file_path}: {error_msg}")

def log_info(msg):
    logging.info(msg)
