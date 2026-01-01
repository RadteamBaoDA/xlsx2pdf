"""
Example usage of the Office to PDF Converter Interface

This file demonstrates various ways to use the refactored converter interface.
"""

from src.interface import (
    convert_single,
    convert_batch,
    OfficeConverter,
    ConversionResult
)
from src.core import load_config
import os


def example_single_conversion():
    """Example: Convert a single file"""
    print("\n=== Example 1: Single File Conversion ===")
    
    result = convert_single(
        input_path='input/sample.xlsx',
        output_path='output/sample.pdf'
    )
    
    print(f"Success: {result.success}")
    print(f"Duration: {result.duration:.2f}s")
    if result.error:
        print(f"Error: {result.error}")


def example_batch_conversion():
    """Example: Convert multiple files"""
    print("\n=== Example 2: Batch Conversion ===")
    
    files = [
        'input/document1.docx',
        'input/spreadsheet1.xlsx',
        'input/presentation1.pptx'
    ]
    
    results = convert_batch(
        input_files=files,
        output_dir='output',
        preserve_structure=False  # Flatten to output folder
    )
    
    for result in results:
        print(result)


def example_with_custom_config():
    """Example: Use custom configuration"""
    print("\n=== Example 3: Custom Configuration ===")
    
    config = load_config('config.yaml')
    
    # Modify config for specific needs
    config['word_options'] = {
        'create_bookmarks': True,
        'optimize_for_print': True
    }
    
    converter = OfficeConverter(config)
    result = converter.convert('input/document.docx', 'output/document.pdf')
    
    print(result)


def example_batch_with_statistics():
    """Example: Batch conversion with statistics"""
    print("\n=== Example 4: Batch with Statistics ===")
    
    # Scan directory for Office files
    input_dir = 'input'
    office_files = []
    
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith(('.xlsx', '.docx', '.pptx', '.xls', '.doc', '.ppt')):
                office_files.append(os.path.join(root, file))
    
    if not office_files:
        print("No Office files found in input directory")
        return
    
    # Convert all files
    converter = OfficeConverter()
    results = converter.convert_batch(
        input_files=office_files,
        output_dir='output',
        preserve_structure=True,
        base_dir=input_dir
    )
    
    # Get statistics
    stats = converter.get_conversion_statistics(results)
    
    print(f"\n📊 Conversion Statistics:")
    print(f"Total files: {stats['total']}")
    print(f"Successful: {stats['successful']}")
    print(f"Failed: {stats['failed']}")
    print(f"Success rate: {stats['success_rate']:.1f}%")
    print(f"Total duration: {stats['total_duration']:.2f}s")
    print(f"Average duration: {stats['avg_duration']:.2f}s")
    
    if stats.get('by_type'):
        print("\nBy File Type:")
        for file_type, type_stats in stats['by_type'].items():
            print(f"  {file_type}: {type_stats['successful']}/{type_stats['total']} succeeded")


def example_check_supported_files():
    """Example: Check which files are supported"""
    print("\n=== Example 5: Check Supported Files ===")
    
    files = [
        'document.docx',
        'spreadsheet.xlsx',
        'presentation.pptx',
        'image.jpg',  # Not supported
        'text.txt'    # Not supported
    ]
    
    converter = OfficeConverter()
    
    for file in files:
        supported = converter.is_supported(file)
        status = "✓ Supported" if supported else "✗ Not supported"
        print(f"{status}: {file}")
    
    print(f"\nAll supported extensions:")
    print(OfficeConverter.get_supported_extensions())


def example_type_specific_conversion():
    """Example: Using specific converters directly"""
    print("\n=== Example 6: Type-Specific Converters ===")
    
    from src.features.excel import ExcelConverter
    from src.features.word import WordConverter
    from src.features.powerpoint import PowerPointConverter
    
    config = load_config()
    
    # Use Excel converter directly for fine-grained control
    excel_converter = ExcelConverter(config)
    if excel_converter.can_convert('input/data.xlsx'):
        excel_converter.convert('input/data.xlsx', 'output/data.pdf')
        print("Excel file converted")
    
    # Use Word converter
    word_converter = WordConverter(config)
    if word_converter.can_convert('input/report.docx'):
        word_converter.convert('input/report.docx', 'output/report.pdf')
        print("Word file converted")
    
    # Use PowerPoint converter
    ppt_converter = PowerPointConverter(config)
    if ppt_converter.can_convert('input/slides.pptx'):
        ppt_converter.convert('input/slides.pptx', 'output/slides.pdf')
        print("PowerPoint file converted")


def example_error_handling():
    """Example: Proper error handling"""
    print("\n=== Example 7: Error Handling ===")
    
    converter = OfficeConverter()
    
    files = [
        'input/exists.xlsx',
        'input/missing.docx',  # File doesn't exist
        'input/corrupted.pptx'  # Might be corrupted
    ]
    
    for file in files:
        result = converter.convert(file, f'output/{os.path.basename(file)}.pdf')
        
        if result.success:
            print(f"✓ {file} converted successfully")
        else:
            print(f"✗ {file} failed: {result.error}")


if __name__ == '__main__':
    """
    Run all examples.
    Note: Make sure you have sample files in the 'input' folder before running.
    """
    
    print("=" * 60)
    print("Office to PDF Converter - Usage Examples")
    print("=" * 60)
    
    # Uncomment the examples you want to run:
    
    # example_single_conversion()
    # example_batch_conversion()
    # example_with_custom_config()
    # example_batch_with_statistics()
    example_check_supported_files()
    # example_type_specific_conversion()
    # example_error_handling()
    
    print("\n" + "=" * 60)
    print("Examples completed!")
    print("=" * 60)
