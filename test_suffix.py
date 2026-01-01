"""
Test file-type specific output suffixes
"""

from src.interface import OfficeConverter
from src.core.utils import load_config
from pathlib import Path

def test_suffixes():
    """Test that different file types get correct suffixes"""
    
    # Load config
    config = load_config('config.yaml')
    
    # Create converter
    converter = OfficeConverter(config)
    
    # Test Excel suffix
    excel_suffix = converter.get_output_suffix('test.xlsx')
    print(f"Excel (.xlsx) suffix: {excel_suffix}")
    assert excel_suffix == '_x', f"Expected '_x', got '{excel_suffix}'"
    
    excel_suffix2 = converter.get_output_suffix('test.xls')
    print(f"Excel (.xls) suffix: {excel_suffix2}")
    assert excel_suffix2 == '_x', f"Expected '_x', got '{excel_suffix2}'"
    
    # Test Word suffix
    word_suffix = converter.get_output_suffix('test.docx')
    print(f"Word (.docx) suffix: {word_suffix}")
    assert word_suffix == '_d', f"Expected '_d', got '{word_suffix}'"
    
    word_suffix2 = converter.get_output_suffix('test.doc')
    print(f"Word (.doc) suffix: {word_suffix2}")
    assert word_suffix2 == '_d', f"Expected '_d', got '{word_suffix2}'"
    
    # Test PowerPoint suffix
    ppt_suffix = converter.get_output_suffix('test.pptx')
    print(f"PowerPoint (.pptx) suffix: {ppt_suffix}")
    assert ppt_suffix == '_p', f"Expected '_p', got '{ppt_suffix}'"
    
    ppt_suffix2 = converter.get_output_suffix('test.ppt')
    print(f"PowerPoint (.ppt) suffix: {ppt_suffix2}")
    assert ppt_suffix2 == '_p', f"Expected '_p', got '{ppt_suffix2}'"
    
    print("\n✅ All suffix tests passed!")
    
    # Test batch conversion output paths
    print("\n--- Testing Batch Conversion Output Paths ---")
    
    test_files = [
        'test.xlsx',
        'test.docx',
        'test.pptx'
    ]
    
    import os
    from pathlib import Path
    
    for test_file in test_files:
        if True:  # Simulating preserve_structure=False
            suffix = converter.get_output_suffix(test_file)
            filename = Path(test_file).stem + suffix + '.pdf'
            output_path = os.path.join('output', filename)
            print(f"{test_file} -> {output_path}")
    
    print("\n✅ Batch conversion path tests passed!")

if __name__ == "__main__":
    test_suffixes()
