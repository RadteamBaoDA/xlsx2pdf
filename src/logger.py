import logging
import os
from logging.handlers import RotatingFileHandler, QueueHandler

def setup_logger(log_file, error_file, log_level="INFO"):
    """
    Sets up the logger with file handlers for general logs and errors.
    """
    level_map = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR
    }
    level = level_map.get(str(log_level).upper(), logging.INFO)

    # Create handlers
    log_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=2, encoding='utf-8')
    error_handler = RotatingFileHandler(error_file, maxBytes=5*1024*1024, backupCount=2, encoding='utf-8')
    
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
    
    return logger

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
