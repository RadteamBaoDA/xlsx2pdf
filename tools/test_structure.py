"""
Simple test script to verify the refactored structure works correctly.
"""

import sys
from pathlib import Path

def test_imports():
    """Test that all imports work correctly."""
    print("Testing imports...")
    
    try:
        # Test core imports
        from src.core import BaseConverter, load_config, setup_logger, LanguageDetector
        print("✓ Core imports successful")
        
        # Test feature imports
        from src.features.excel import ExcelConverter
        from src.features.word import WordConverter
        from src.features.powerpoint import PowerPointConverter
        print("✓ Feature converter imports successful")
        
        # Test interface imports
        from src.interface import (
            OfficeConverter, 
            ConversionResult, 
            convert_single, 
            convert_batch
        )
        print("✓ Interface imports successful")
        
        # Test get_converter_for_file separately
        from src.interface.converter_interface import get_converter_for_file
        print("✓ Interface utility functions successful")
        
        # Test package imports
        from src import (
            OfficeConverter as OC,
            ExcelConverter as EC,
            WordConverter as WC,
            PowerPointConverter as PC
        )
        print("✓ Package-level imports successful")
        
        return True
        
    except Exception as e:
        print(f"✗ Import failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_converter_instantiation():
    """Test that converters can be instantiated."""
    print("\nTesting converter instantiation...")
    
    try:
        from src.interface import OfficeConverter
        from src.features.excel import ExcelConverter
        from src.features.word import WordConverter
        from src.features.powerpoint import PowerPointConverter
        
        # Test with empty config
        config = {}
        
        office_conv = OfficeConverter(config)
        print("✓ OfficeConverter instantiated")
        
        excel_conv = ExcelConverter(config)
        print("✓ ExcelConverter instantiated")
        
        word_conv = WordConverter(config)
        print("✓ WordConverter instantiated")
        
        ppt_conv = PowerPointConverter(config)
        print("✓ PowerPointConverter instantiated")
        
        return True
        
    except Exception as e:
        print(f"✗ Instantiation failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_supported_extensions():
    """Test that supported extensions are correctly defined."""
    print("\nTesting supported extensions...")
    
    try:
        from src.interface import OfficeConverter
        
        supported = OfficeConverter.get_supported_extensions()
        print(f"✓ Supported extensions: {len(supported)} types")
        print(f"  Extensions: {', '.join(sorted(supported))}")
        
        # Verify key extensions
        assert '.xlsx' in supported, "Excel (.xlsx) not in supported"
        assert '.docx' in supported, "Word (.docx) not in supported"
        assert '.pptx' in supported, "PowerPoint (.pptx) not in supported"
        print("✓ All key file types supported")
        
        return True
        
    except Exception as e:
        print(f"✗ Extension test failed: {e}")
        return False


def test_converter_selection():
    """Test that correct converter is selected for each file type."""
    print("\nTesting converter selection...")
    
    try:
        from src.interface import OfficeConverter
        
        converter = OfficeConverter({})
        
        test_files = [
            ('test.xlsx', 'excel'),
            ('test.docx', 'word'),
            ('test.pptx', 'powerpoint'),
            ('test.xls', 'excel'),
            ('test.doc', 'word'),
            ('test.ppt', 'powerpoint'),
        ]
        
        for filename, expected_type in test_files:
            conv, file_type = converter.get_converter(filename)
            assert conv is not None, f"No converter for {filename}"
            assert file_type == expected_type, f"Wrong type for {filename}: {file_type} != {expected_type}"
            print(f"✓ {filename} → {file_type}")
        
        # Test unsupported file
        conv, file_type = converter.get_converter('test.txt')
        assert conv is None, "Should return None for unsupported file"
        print("✓ Unsupported files handled correctly")
        
        return True
        
    except Exception as e:
        print(f"✗ Converter selection failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_converter_interface():
    """Test the converter interface methods."""
    print("\nTesting converter interface...")
    
    try:
        from src.interface import OfficeConverter
        
        converter = OfficeConverter({})
        
        # Test is_supported
        assert converter.is_supported('file.xlsx') == True
        assert converter.is_supported('file.docx') == True
        assert converter.is_supported('file.pptx') == True
        assert converter.is_supported('file.txt') == False
        print("✓ is_supported() works correctly")
        
        # Test can_convert on converters
        excel_conv, _ = converter.get_converter('test.xlsx')
        assert excel_conv.can_convert('test.xlsx') == True
        assert excel_conv.can_convert('test.docx') == False
        print("✓ can_convert() works correctly")
        
        return True
        
    except Exception as e:
        print(f"✗ Interface test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_base_converter_inheritance():
    """Test that all converters inherit from BaseConverter."""
    print("\nTesting inheritance...")
    
    try:
        from src.core import BaseConverter
        from src.features.excel import ExcelConverter
        from src.features.word import WordConverter
        from src.features.powerpoint import PowerPointConverter
        
        assert issubclass(ExcelConverter, BaseConverter), "ExcelConverter not subclass of BaseConverter"
        assert issubclass(WordConverter, BaseConverter), "WordConverter not subclass of BaseConverter"
        assert issubclass(PowerPointConverter, BaseConverter), "PowerPointConverter not subclass of BaseConverter"
        
        print("✓ All converters inherit from BaseConverter")
        
        # Test that they have required methods
        config = {}
        converters = [
            ExcelConverter(config),
            WordConverter(config),
            PowerPointConverter(config)
        ]
        
        for conv in converters:
            assert hasattr(conv, 'convert'), f"{conv.__class__.__name__} missing convert()"
            assert hasattr(conv, 'validate_input'), f"{conv.__class__.__name__} missing validate_input()"
            assert hasattr(conv, 'supported_extensions'), f"{conv.__class__.__name__} missing supported_extensions"
            assert hasattr(conv, 'can_convert'), f"{conv.__class__.__name__} missing can_convert()"
        
        print("✓ All converters have required methods")
        
        return True
        
    except Exception as e:
        print(f"✗ Inheritance test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Run all tests."""
    print("=" * 60)
    print("Office to PDF Converter - Structure Verification")
    print("=" * 60)
    
    tests = [
        ("Imports", test_imports),
        ("Instantiation", test_converter_instantiation),
        ("Supported Extensions", test_supported_extensions),
        ("Converter Selection", test_converter_selection),
        ("Interface Methods", test_converter_interface),
        ("Inheritance", test_base_converter_inheritance),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"\n✗ {test_name} crashed: {e}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{status}: {test_name}")
    
    print(f"\nResult: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 All tests passed! Structure is working correctly.")
        return 0
    else:
        print(f"\n⚠️ {total - passed} test(s) failed. Please review the errors above.")
        return 1


if __name__ == '__main__':
    sys.exit(main())
