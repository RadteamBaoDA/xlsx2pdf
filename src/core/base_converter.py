"""
Base converter abstract class following the Strategy pattern.
All Office file converters should inherit from this base class.
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, Dict, Any
import logging
import os


class BaseConverter(ABC):
    """
    Abstract base class for all Office document converters.
    Implements the Template Method pattern for conversion workflow.
    """
    
    def __init__(self, config: Dict[str, Any]):
        """
        Initialize the converter with configuration.
        
        Args:
            config: Configuration dictionary containing converter settings
        """
        self.config = config
        self._setup_logging()
    
    def _setup_logging(self):
        """Setup logging for the converter."""
        if not logging.getLogger().handlers:
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s'
            )
    
    @abstractmethod
    def convert(self, input_path: str, output_path: str, pid_queue=None) -> bool:
        """
        Convert an Office file to PDF.
        
        Args:
            input_path: Path to the input Office file
            output_path: Path where the PDF should be saved
            pid_queue: Optional queue for sending process ID to parent
        
        Returns:
            bool: True if conversion succeeded, False otherwise
        
        Raises:
            Exception: If conversion fails
        """
        pass
    
    @abstractmethod
    def validate_input(self, input_path: str) -> bool:
        """
        Validate that the input file is appropriate for this converter.
        
        Args:
            input_path: Path to the input file
        
        Returns:
            bool: True if file is valid, False otherwise
        """
        pass
    
    def _ensure_output_directory(self, output_path: str):
        """
        Ensure the output directory exists.
        
        Args:
            output_path: Path where the output file will be saved
        """
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
    
    def _cleanup_temp_files(self, *file_paths):
        """
        Clean up temporary files.
        
        Args:
            *file_paths: Variable number of file paths to delete
        """
        for file_path in file_paths:
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    logging.debug(f"Removed temporary file: {file_path}")
                except Exception as e:
                    logging.warning(f"Failed to remove temporary file {file_path}: {e}")
    
    @property
    @abstractmethod
    def supported_extensions(self) -> list:
        """
        Return list of supported file extensions.
        
        Returns:
            list: List of supported file extensions (e.g., ['.xlsx', '.xls'])
        """
        pass
    
    def can_convert(self, file_path: str) -> bool:
        """
        Check if this converter can handle the given file.
        
        Args:
            file_path: Path to the file
        
        Returns:
            bool: True if this converter supports the file extension
        """
        ext = Path(file_path).suffix.lower()
        return ext in self.supported_extensions
