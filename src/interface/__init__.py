"""Interface layer for easy integration with external code."""
from .converter_interface import OfficeConverter, ConversionResult, convert_single, convert_batch

__all__ = ['OfficeConverter', 'ConversionResult', 'convert_single', 'convert_batch']
