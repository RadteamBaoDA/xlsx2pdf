"""
Office to PDF Conversion Package

A comprehensive solution for converting Microsoft Office files (Excel, Word, PowerPoint)
to PDF format with advanced formatting and layout options.

Architecture:
- core: Base classes, utilities, and shared functionality
- features: Feature-based converters (excel, word, powerpoint)
- interface: Clean API for external integration

Usage:
    # Quick conversion
    from src.interface import convert_single, convert_batch
    
    # Single file
    result = convert_single('document.docx', 'output.pdf')
    
    # Batch conversion
    results = convert_batch(['file1.xlsx', 'file2.docx'], 'output_dir')
    
    # Advanced usage
    from src.interface import OfficeConverter
    converter = OfficeConverter(config)
    result = converter.convert('file.pptx', 'output.pdf')
"""

# Import main interface for easy access
from .interface import (
    OfficeConverter,
    ConversionResult,
    convert_single,
    convert_batch
)

# Import converters for direct access if needed
from .features.excel import ExcelConverter
from .features.word import WordConverter
from .features.powerpoint import PowerPointConverter

# Import core utilities
from .core import (
    BaseConverter,
    load_config,
    setup_logger,
    LanguageDetector
)

__version__ = '2.0.0'
__all__ = [
    # Interface
    'OfficeConverter',
    'ConversionResult',
    'convert_single',
    'convert_batch',
    
    # Converters
    'ExcelConverter',
    'WordConverter',
    'PowerPointConverter',
    
    # Core
    'BaseConverter',
    'load_config',
    'setup_logger',
    'LanguageDetector',
]
