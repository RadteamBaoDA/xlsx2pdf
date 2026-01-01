"""
Test portability features - verify the package works without config file
"""
from src.interface import OfficeConverter, convert_single

print("=" * 60)
print("Portability Test - No Config File Required")
print("=" * 60)

# Test 1: No config file
print('\nTest 1: Creating converter without config file...')
converter = OfficeConverter()
print('✓ Success! Converter created with default config')

# Test 2: With custom config
print('\nTest 2: Creating converter with custom config...')
config = {
    'word_options': {'create_bookmarks': True},
    'excel': {'prepare_for_print': False}
}
converter2 = OfficeConverter(config)
print('✓ Success! Converter created with custom config')

# Test 3: Check supported types
print('\nTest 3: Checking supported file types...')
supported = OfficeConverter.get_supported_extensions()
print(f'✓ Supports {len(supported)} file types')

excel_types = [ext for ext in supported if ext in [".xlsx", ".xls", ".xlsm"]]
word_types = [ext for ext in supported if ext in [".docx", ".doc", ".docm"]]
ppt_types = [ext for ext in supported if ext in [".pptx", ".ppt", ".pptm"]]

print(f'  Excel: {excel_types}')
print(f'  Word: {word_types}')
print(f'  PowerPoint: {ppt_types}')

# Test 4: Verify default config
print('\nTest 4: Checking default configuration...')
assert converter.config is not None, "Config should not be None"
print(f'✓ Config loaded (from file or defaults)')
print(f'  Has {len(converter.config)} configuration sections')

# Check that converters are initialized
assert len(converter._converters) > 0, "Converters should be initialized"
print(f'✓ {len(converter._converters)} converters initialized')

# Test 5: Verify it works without any config.yaml file
print('\nTest 5: Testing truly standalone usage...')
import os
import tempfile
import sys

# Create a temp directory and test in isolation
with tempfile.TemporaryDirectory() as tmpdir:
    # Save current dir
    old_cwd = os.getcwd()
    
    try:
        # Change to empty directory (no config.yaml)
        os.chdir(tmpdir)
        
        # Should work with defaults
        test_converter = OfficeConverter()
        assert test_converter.config is not None
        assert 'word_options' in test_converter.config
        assert 'excel' in test_converter.config
        print('✓ Works in directory without config.yaml')
        
    finally:
        os.chdir(old_cwd)

print('\n✓ Default config has all required sections')

print("\n" + "=" * 60)
print("All portability tests passed!")
print("The src folder is ready to be copied to any project!")
print("=" * 60)
