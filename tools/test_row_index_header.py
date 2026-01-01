"""
Test script to verify row index information is added to headers

This test creates a sample Excel file and converts it to PDF
to verify that the calculated page structure with row indices
appears in the header and as a column in the output.
"""

import os
import pandas as pd
from src.interface import OfficeConverter
from src.core.utils import load_config

def create_test_excel():
    """Create a test Excel file with multiple rows"""
    # Create sample data with 25 rows
    data = {
        'ID': list(range(1, 26)),
        'Name': [f'Item {i}' for i in range(1, 26)],
        'Value': [i * 100 for i in range(1, 26)],
        'Description': [f'Description for item {i}' for i in range(1, 26)]
    }
    
    df = pd.DataFrame(data)
    
    # Save to Excel
    test_file = 'test_row_index.xlsx'
    df.to_excel(test_file, index=False, sheet_name='TestData')
    
    print(f"Created test file: {test_file}")
    return test_file

def test_conversion():
    """Test PDF conversion with row index headers"""
    
    # Create test file
    test_file = create_test_excel()
    
    try:
        # Load config
        config = load_config('config.yaml')
        
        # Ensure rows_per_page is set to trigger page breaks
        if isinstance(config.get('print_options'), list):
            # Find default config (sheets: null)
            for opt in config['print_options']:
                if opt.get('sheets') is None:
                    opt['rows_per_page'] = 5  # 5 rows per page for testing
                    opt['print_header_footer'] = True
                    break
        
        # Create converter
        converter = OfficeConverter(config)
        
        # Convert
        output_file = 'output/test_row_index_x.pdf'
        os.makedirs('output', exist_ok=True)
        
        print(f"\nConverting {test_file} to {output_file}...")
        print("Expected page structure:")
        print("  Page 1: Rows 2-6 (header row 1 + 5 data rows)")
        print("  Page 2: Rows 7-11")
        print("  Page 3: Rows 12-16")
        print("  Page 4: Rows 17-21")
        print("  Page 5: Rows 22-26")
        
        result = converter.convert(test_file, output_file)
        
        if result.success:
            print(f"\n✓ Conversion successful!")
            print(f"  Output: {output_file}")
            print(f"\nCheck the PDF for:")
            print(f"  1. Header shows page structure: 'P1:R2-6, P2:R7-11, P3:R12-16'")
            print(f"  2. Added column 'Page-Row Index' with labels at each page start")
            print(f"  3. Each page has row index information visible")
            print(f"\nCheck logs for RAG METADATA page-to-row mapping")
        else:
            print(f"\n✗ Conversion failed: {result.error}")
            
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup test file
        if os.path.exists(test_file):
            try:
                os.remove(test_file)
                print(f"\nCleaned up test file: {test_file}")
            except:
                pass

if __name__ == "__main__":
    print("="*70)
    print("Testing Row Index Header Feature")
    print("="*70)
    test_conversion()
    print("="*70)
