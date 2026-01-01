"""
Standalone Example - Using the converter in another project

This example shows how to use the Office to PDF converter
after copying the 'src' folder to your project.

No config file required! Everything can be configured via code.
"""

# Assuming you copied 'src' folder to your project
# You can rename it to anything, e.g., 'office_converter'
from src.interface import OfficeConverter, convert_single, convert_batch
import os


def example_1_simplest_usage():
    """Simplest possible usage - no config needed!"""
    print("\n=== Example 1: Simplest Usage ===")
    
    # Just convert - uses sensible defaults
    result = convert_single('input/sample.docx', 'output/sample.pdf')
    
    print(f"Success: {result.success}")
    if result.error:
        print(f"Error: {result.error}")


def example_2_with_custom_config():
    """Pass configuration as dictionary - no YAML file needed!"""
    print("\n=== Example 2: Custom Config via Dictionary ===")
    
    # Define your configuration in code
    my_config = {
        'word_options': {
            'create_bookmarks': True,
            'optimize_for_print': True
        },
        'excel': {
            'prepare_for_print': False  # Faster conversion
        },
        'pdf_trim': {
            'enabled': False  # Keep original margins
        }
    }
    
    # Create converter with your config
    converter = OfficeConverter(config=my_config)
    
    # Convert files
    result = converter.convert('input/document.docx', 'output/document.pdf')
    print(f"Converted: {result.success}")


def example_3_batch_conversion():
    """Convert multiple files at once"""
    print("\n=== Example 3: Batch Conversion ===")
    
    files = [
        'input/report.docx',
        'input/data.xlsx',
        'input/presentation.pptx'
    ]
    
    # Simple batch conversion with defaults
    results = convert_batch(
        input_files=files,
        output_dir='output',
        preserve_structure=False  # All PDFs in same folder
    )
    
    # Show results
    for result in results:
        status = "✓" if result.success else "✗"
        print(f"{status} {os.path.basename(result.input_path)}")


def example_4_scan_and_convert():
    """Scan a folder and convert all Office files"""
    print("\n=== Example 4: Scan and Convert All ===")
    
    # Custom configuration
    config = {
        'word_options': {'create_bookmarks': True},
        'powerpoint_options': {'output_type': 'slides'}
    }
    
    converter = OfficeConverter(config)
    
    # Scan for Office files
    input_folder = 'input'
    office_files = []
    
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            filepath = os.path.join(root, file)
            if converter.is_supported(filepath):
                office_files.append(filepath)
    
    print(f"Found {len(office_files)} Office files")
    
    # Convert all
    results = converter.convert_batch(
        input_files=office_files,
        output_dir='output',
        preserve_structure=True,
        base_dir=input_folder
    )
    
    # Statistics
    stats = converter.get_conversion_statistics(results)
    print(f"\nSuccess rate: {stats['success_rate']:.1f}%")
    print(f"Total time: {stats['total_duration']:.2f}s")


def example_5_minimal_function():
    """Minimal function you can use in your project"""
    print("\n=== Example 5: Minimal Integration Function ===")
    
    def convert_to_pdf(input_file, output_file=None, options=None):
        """
        Convert any Office file to PDF.
        
        Args:
            input_file: Path to Office file (.docx, .xlsx, .pptx, etc.)
            output_file: Path for PDF output (optional, auto-generates if None)
            options: Dict of conversion options (optional)
            
        Returns:
            True if successful, False otherwise
        """
        if output_file is None:
            # Auto-generate output path
            output_file = os.path.splitext(input_file)[0] + '.pdf'
        
        # Convert with optional config
        if options:
            converter = OfficeConverter(config=options)
            result = converter.convert(input_file, output_file)
        else:
            result = convert_single(input_file, output_file)
        
        return result.success
    
    # Usage of the minimal function
    success = convert_to_pdf('input/document.docx')
    print(f"Converted: {success}")
    
    # With custom options
    success = convert_to_pdf(
        'input/spreadsheet.xlsx',
        'output/report.pdf',
        options={'excel': {'prepare_for_print': False}}
    )
    print(f"Converted with options: {success}")


def example_6_error_handling():
    """Proper error handling"""
    print("\n=== Example 6: Error Handling ===")
    
    converter = OfficeConverter()
    
    files = [
        'input/exists.docx',
        'input/missing.xlsx',  # This file doesn't exist
        'input/valid.pptx'
    ]
    
    successful = 0
    failed = 0
    
    for file in files:
        result = converter.convert(file, f'output/{os.path.basename(file)}.pdf')
        
        if result.success:
            successful += 1
            print(f"✓ {os.path.basename(file)} - {result.duration:.2f}s")
        else:
            failed += 1
            print(f"✗ {os.path.basename(file)} - {result.error}")
    
    print(f"\nTotal: {successful} succeeded, {failed} failed")


def example_7_integration_class():
    """Create a reusable class for your project"""
    print("\n=== Example 7: Reusable Integration Class ===")
    
    class MyProjectConverter:
        """Wrapper class for easy integration in your project"""
        
        def __init__(self, **kwargs):
            """
            Initialize with optional settings.
            
            Kwargs:
                word_bookmarks: bool
                excel_optimize: bool
                pdf_trim: bool
            """
            config = {}
            
            if 'word_bookmarks' in kwargs:
                config['word_options'] = {'create_bookmarks': kwargs['word_bookmarks']}
            
            if 'excel_optimize' in kwargs:
                config['excel'] = {'prepare_for_print': kwargs['excel_optimize']}
            
            if 'pdf_trim' in kwargs:
                config['pdf_trim'] = {'enabled': kwargs['pdf_trim']}
            
            self.converter = OfficeConverter(config if config else None)
        
        def convert(self, input_path, output_dir='output'):
            """Convert any Office file to PDF"""
            output_path = os.path.join(
                output_dir,
                os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
            )
            
            result = self.converter.convert(input_path, output_path)
            return result.success, result.error
        
        def convert_many(self, files, output_dir='output'):
            """Convert multiple files"""
            results = self.converter.convert_batch(files, output_dir)
            stats = self.converter.get_conversion_statistics(results)
            return stats
    
    # Usage
    my_converter = MyProjectConverter(
        word_bookmarks=True,
        excel_optimize=False,
        pdf_trim=False
    )
    
    success, error = my_converter.convert('input/test.docx')
    print(f"Converted: {success}")


def main():
    """Run all examples"""
    print("=" * 70)
    print("Office to PDF Converter - Standalone Usage Examples")
    print("Copy the 'src' folder to your project and use like this!")
    print("=" * 70)
    
    # Uncomment the examples you want to run:
    
    # example_1_simplest_usage()
    # example_2_with_custom_config()
    # example_3_batch_conversion()
    # example_4_scan_and_convert()
    example_5_minimal_function()
    # example_6_error_handling()
    # example_7_integration_class()
    
    print("\n" + "=" * 70)
    print("That's how easy it is to use in another project!")
    print("Just copy 'src' folder and import - no config file needed!")
    print("=" * 70)


if __name__ == '__main__':
    main()
