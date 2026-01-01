"""
Unified Office to PDF Converter Interface

This module provides a simple, clean interface for converting Microsoft Office files
to PDF format. It supports Excel, Word, and PowerPoint files and can handle both
single file conversions and batch operations.

Usage Examples:
    # Single file conversion
    from src.interface import convert_single
    result = convert_single('document.docx', 'output.pdf')
    
    # Batch conversion
    from src.interface import convert_batch
    results = convert_batch(['file1.xlsx', 'file2.docx'], 'output_folder')
    
    # Using the OfficeConverter class directly
    from src.interface import OfficeConverter
    converter = OfficeConverter(config)
    result = converter.convert('document.pptx', 'output.pdf')
"""

import os
import logging
from pathlib import Path
from typing import List, Dict, Optional, Any, Union
from dataclasses import dataclass, field
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing

from ..core import load_config, ensure_dir
from ..features.excel import ExcelConverter
from ..features.word import WordConverter
from ..features.powerpoint import PowerPointConverter


@dataclass
class ConversionResult:
    """
    Result of a file conversion operation.
    
    Attributes:
        input_path: Path to the input file
        output_path: Path to the output PDF file (None if conversion failed)
        success: Whether the conversion succeeded
        error: Error message if conversion failed (None if successful)
        duration: Time taken for conversion in seconds
        file_type: Type of office file (excel, word, powerpoint)
    """
    input_path: str
    output_path: Optional[str] = None
    success: bool = False
    error: Optional[str] = None
    duration: float = 0.0
    file_type: Optional[str] = None
    timestamp: datetime = field(default_factory=datetime.now)
    
    def __str__(self):
        """String representation of conversion result."""
        status = "✓ SUCCESS" if self.success else "✗ FAILED"
        return f"[{status}] {self.input_path} -> {self.output_path or 'N/A'} ({self.duration:.2f}s)"
    
    def to_dict(self):
        """Convert result to dictionary."""
        return {
            'input_path': self.input_path,
            'output_path': self.output_path,
            'success': self.success,
            'error': self.error,
            'duration': self.duration,
            'file_type': self.file_type,
            'timestamp': self.timestamp.isoformat()
        }


class OfficeConverter:
    """
    Unified interface for converting Microsoft Office files to PDF.
    
    This class automatically selects the appropriate converter based on file type
    and provides a simple, consistent interface for all Office file conversions.
    
    Attributes:
        config: Configuration dictionary for converter settings
        converters: Dictionary mapping file extensions to converter instances
    """
    
    # Mapping of file extensions to converter classes
    CONVERTER_MAP = {
        '.xlsx': ('excel', ExcelConverter),
        '.xls': ('excel', ExcelConverter),
        '.xlsm': ('excel', ExcelConverter),
        '.xlsb': ('excel', ExcelConverter),
        '.docx': ('word', WordConverter),
        '.doc': ('word', WordConverter),
        '.docm': ('word', WordConverter),
        '.dotx': ('word', WordConverter),
        '.dotm': ('word', WordConverter),
        '.pptx': ('powerpoint', PowerPointConverter),
        '.ppt': ('powerpoint', PowerPointConverter),
        '.pptm': ('powerpoint', PowerPointConverter),
        '.ppsx': ('powerpoint', PowerPointConverter),
        '.ppsm': ('powerpoint', PowerPointConverter),
        '.potx': ('powerpoint', PowerPointConverter),
        '.potm': ('powerpoint', PowerPointConverter),
    }
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the Office converter.
        
        Args:
            config: Configuration dictionary. If None, loads from default config.yaml
        """
        if config is None:
            try:
                config = load_config()
            except Exception as e:
                logging.warning(f"Could not load config, using defaults: {e}")
                config = {}
        
        self.config = config
        
        # Initialize converters (lazy initialization for better performance)
        self._converters = {}
        self._initialize_converters()
    
    def _initialize_converters(self):
        """Initialize converter instances."""
        try:
            self._converters['excel'] = ExcelConverter(self.config)
            self._converters['word'] = WordConverter(self.config)
            self._converters['powerpoint'] = PowerPointConverter(self.config)
            logging.info("Initialized all Office converters")
        except Exception as e:
            logging.error(f"Error initializing converters: {e}")
            raise
    
    def get_converter(self, file_path: str):
        """
        Get the appropriate converter for a given file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Tuple of (converter_instance, file_type) or (None, None) if unsupported
        """
        ext = Path(file_path).suffix.lower()
        
        if ext not in self.CONVERTER_MAP:
            return None, None
        
        file_type, _ = self.CONVERTER_MAP[ext]
        converter = self._converters.get(file_type)
        
        return converter, file_type
    
    def is_supported(self, file_path: str) -> bool:
        """
        Check if a file type is supported for conversion.
        
        Args:
            file_path: Path to the file
            
        Returns:
            bool: True if file type is supported
        """
        ext = Path(file_path).suffix.lower()
        return ext in self.CONVERTER_MAP
    
    def convert(self, input_path: str, output_path: str, pid_queue=None) -> ConversionResult:
        """
        Convert a single Office file to PDF.
        
        Args:
            input_path: Path to the input Office file
            output_path: Path where the PDF should be saved
            pid_queue: Optional queue for sending process ID to parent
            
        Returns:
            ConversionResult: Result of the conversion operation
        """
        start_time = datetime.now()
        
        # Validate input file
        if not os.path.exists(input_path):
            return ConversionResult(
                input_path=input_path,
                success=False,
                error=f"File not found: {input_path}",
                duration=0.0
            )
        
        # Get appropriate converter
        converter, file_type = self.get_converter(input_path)
        
        if converter is None:
            ext = Path(input_path).suffix
            return ConversionResult(
                input_path=input_path,
                success=False,
                error=f"Unsupported file type: {ext}",
                duration=0.0
            )
        
        # Ensure output directory exists
        ensure_dir(output_path)
        
        # Perform conversion
        try:
            logging.info(f"Converting {file_type} file: {input_path} -> {output_path}")
            converter.convert(input_path, output_path, pid_queue)
            
            duration = (datetime.now() - start_time).total_seconds()
            
            return ConversionResult(
                input_path=input_path,
                output_path=output_path,
                success=True,
                duration=duration,
                file_type=file_type
            )
            
        except Exception as e:
            duration = (datetime.now() - start_time).total_seconds()
            error_msg = str(e)
            logging.error(f"Conversion failed for {input_path}: {error_msg}")
            
            return ConversionResult(
                input_path=input_path,
                success=False,
                error=error_msg,
                duration=duration,
                file_type=file_type
            )
    
    def convert_batch(
        self,
        input_files: List[str],
        output_dir: str,
        preserve_structure: bool = True,
        base_dir: Optional[str] = None,
        max_workers: Optional[int] = None
    ) -> List[ConversionResult]:
        """
        Convert multiple Office files to PDF in batch.
        
        Args:
            input_files: List of input file paths
            output_dir: Directory where PDFs should be saved
            preserve_structure: If True, preserve folder structure from base_dir
            base_dir: Base directory for calculating relative paths (required if preserve_structure=True)
            max_workers: Maximum number of parallel workers (None = use CPU count)
            
        Returns:
            List[ConversionResult]: Results for each file conversion
        """
        if not input_files:
            logging.warning("No input files provided for batch conversion")
            return []
        
        if preserve_structure and base_dir is None:
            base_dir = os.path.commonpath(input_files)
            logging.info(f"Using common path as base_dir: {base_dir}")
        
        # Prepare conversion tasks
        tasks = []
        for input_file in input_files:
            if preserve_structure and base_dir:
                # Preserve folder structure
                rel_path = os.path.relpath(input_file, base_dir)
                output_path = os.path.join(output_dir, rel_path)
                output_path = os.path.splitext(output_path)[0] + '.pdf'
            else:
                # Flatten structure
                filename = Path(input_file).stem + '.pdf'
                output_path = os.path.join(output_dir, filename)
            
            tasks.append((input_file, output_path))
        
        # Execute conversions
        results = []
        
        if max_workers == 1 or len(tasks) == 1:
            # Sequential execution
            for input_file, output_path in tasks:
                result = self.convert(input_file, output_path)
                results.append(result)
                logging.info(str(result))
        else:
            # Parallel execution
            if max_workers is None:
                max_workers = min(multiprocessing.cpu_count(), len(tasks))
            
            logging.info(f"Starting batch conversion with {max_workers} workers")
            
            # Note: For Office COM objects, parallel execution may have limitations
            # Consider using sequential execution for stability
            for input_file, output_path in tasks:
                result = self.convert(input_file, output_path)
                results.append(result)
                logging.info(str(result))
        
        # Summary
        successful = sum(1 for r in results if r.success)
        failed = len(results) - successful
        logging.info(f"Batch conversion complete: {successful} succeeded, {failed} failed")
        
        return results
    
    @staticmethod
    def get_supported_extensions() -> List[str]:
        """
        Get list of all supported file extensions.
        
        Returns:
            List of supported extensions (e.g., ['.xlsx', '.docx', '.pptx'])
        """
        return list(OfficeConverter.CONVERTER_MAP.keys())
    
    def get_conversion_statistics(self, results: List[ConversionResult]) -> Dict[str, Any]:
        """
        Calculate statistics from conversion results.
        
        Args:
            results: List of conversion results
            
        Returns:
            Dictionary with statistics
        """
        if not results:
            return {
                'total': 0,
                'successful': 0,
                'failed': 0,
                'success_rate': 0.0,
                'total_duration': 0.0,
                'avg_duration': 0.0
            }
        
        total = len(results)
        successful = sum(1 for r in results if r.success)
        failed = total - successful
        total_duration = sum(r.duration for r in results)
        
        return {
            'total': total,
            'successful': successful,
            'failed': failed,
            'success_rate': (successful / total * 100) if total > 0 else 0.0,
            'total_duration': total_duration,
            'avg_duration': total_duration / total if total > 0 else 0.0,
            'by_type': self._statistics_by_type(results)
        }
    
    def _statistics_by_type(self, results: List[ConversionResult]) -> Dict[str, Dict[str, Any]]:
        """Calculate statistics grouped by file type."""
        stats_by_type = {}
        
        for result in results:
            if result.file_type:
                if result.file_type not in stats_by_type:
                    stats_by_type[result.file_type] = {
                        'total': 0,
                        'successful': 0,
                        'failed': 0
                    }
                
                stats_by_type[result.file_type]['total'] += 1
                if result.success:
                    stats_by_type[result.file_type]['successful'] += 1
                else:
                    stats_by_type[result.file_type]['failed'] += 1
        
        return stats_by_type


# Convenience functions for simple usage

def convert_single(
    input_path: str,
    output_path: str,
    config: Optional[Dict[str, Any]] = None
) -> ConversionResult:
    """
    Convert a single Office file to PDF.
    
    Convenience function for quick single-file conversions.
    
    Args:
        input_path: Path to the input Office file
        output_path: Path where the PDF should be saved
        config: Optional configuration dictionary
        
    Returns:
        ConversionResult: Result of the conversion
        
    Example:
        >>> from src.interface import convert_single
        >>> result = convert_single('document.docx', 'output.pdf')
        >>> print(result.success)
        True
    """
    converter = OfficeConverter(config)
    return converter.convert(input_path, output_path)


def convert_batch(
    input_files: List[str],
    output_dir: str,
    config: Optional[Dict[str, Any]] = None,
    preserve_structure: bool = True,
    base_dir: Optional[str] = None,
    max_workers: Optional[int] = None
) -> List[ConversionResult]:
    """
    Convert multiple Office files to PDF in batch.
    
    Convenience function for batch conversions.
    
    Args:
        input_files: List of input file paths
        output_dir: Directory where PDFs should be saved
        config: Optional configuration dictionary
        preserve_structure: If True, preserve folder structure
        base_dir: Base directory for relative paths
        max_workers: Maximum parallel workers
        
    Returns:
        List[ConversionResult]: Results for each conversion
        
    Example:
        >>> from src.interface import convert_batch
        >>> files = ['doc1.docx', 'sheet1.xlsx', 'pres1.pptx']
        >>> results = convert_batch(files, 'output_folder')
        >>> print(f"{sum(r.success for r in results)} files converted")
    """
    converter = OfficeConverter(config)
    return converter.convert_batch(
        input_files,
        output_dir,
        preserve_structure=preserve_structure,
        base_dir=base_dir,
        max_workers=max_workers
    )


def get_converter_for_file(file_path: str, config: Optional[Dict[str, Any]] = None):
    """
    Get the appropriate converter instance for a file.
    
    Args:
        file_path: Path to the file
        config: Optional configuration dictionary
        
    Returns:
        Converter instance or None if unsupported
        
    Example:
        >>> from src.interface import get_converter_for_file
        >>> converter = get_converter_for_file('document.docx')
        >>> if converter:
        >>>     converter.convert('input.docx', 'output.pdf')
    """
    office_converter = OfficeConverter(config)
    converter, _ = office_converter.get_converter(file_path)
    return converter
