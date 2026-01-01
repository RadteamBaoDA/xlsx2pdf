# How to Use This Package in Another Project

## Option 1: Copy the `src` folder to your project

### Step 1: Copy the folder
```bash
# Copy the entire src folder to your project
cp -r /path/to/xlsx2pdf/src /path/to/your-project/office_converter
```

### Step 2: Import and use with custom config

```python
# In your project
from office_converter.interface import convert_single, convert_batch, OfficeConverter

# Option A: Simple conversion with default config
result = convert_single('document.docx', 'output.pdf')
print(f"Success: {result.success}")

# Option B: Pass custom config as parameter
custom_config = {
    'word_options': {
        'create_bookmarks': True,
        'optimize_for_print': True
    },
    'excel': {
        'prepare_for_print': True
    },
    'powerpoint_options': {
        'output_type': 'slides',
        'include_hidden_slides': False
    }
}

# Create converter with custom config
converter = OfficeConverter(config=custom_config)
result = converter.convert('file.xlsx', 'output.pdf')

# Batch conversion with custom config
files = ['doc1.docx', 'sheet1.xlsx', 'pres1.pptx']
results = converter.convert_batch(
    input_files=files,
    output_dir='output',
    preserve_structure=True
)

# Get statistics
stats = converter.get_conversion_statistics(results)
print(f"Success rate: {stats['success_rate']:.1f}%")
```

## Option 2: Install as package (recommended)

### Step 1: Install from source
```bash
cd /path/to/xlsx2pdf
pip install -e .
```

### Step 2: Use in any project
```python
from src.interface import OfficeConverter, convert_single

# Works anywhere!
result = convert_single('document.docx', 'output.pdf')
```

## Configuration Options

### No Config File Required!
You can pass all configuration directly as a dictionary:

```python
from src.interface import OfficeConverter

config = {
    # Excel options
    'excel': {
        'prepare_for_print': True,
        'enhanced_dir': 'enhanced_files'
    },
    
    # Print options (for Excel)
    'print_options': {
        'mode': 'auto',  # auto, one_page, table_row_break, etc.
        'page_size': 'A4',
        'orientation': 'auto',  # auto, portrait, landscape
        'scaling': 'fit_columns',
        'margins': 'normal'
    },
    
    # PDF trimming (for Excel)
    'pdf_trim': {
        'enabled': True,
        'margin_threshold': 10,
        'min_margin': 5
    },
    
    # Word options
    'word_options': {
        'create_bookmarks': True,
        'optimize_for_print': True,
        'include_doc_properties': True,
        'keep_form_fields': True
    },
    
    # PowerPoint options
    'powerpoint_options': {
        'output_type': 'slides',  # slides, notes, handouts
        'handout_order': 'vertical',
        'slides_per_page': 1,
        'include_hidden_slides': False,
        'frame_slides': False,
        'print_comments': False
    },
    
    # Logging options (optional)
    'logging': {
        'log_file': 'conversion.log',
        'error_file': 'errors.log',
        'log_level': 'INFO',
        'logs_folder': 'logs'
    }
}

converter = OfficeConverter(config)
```

### Or use YAML config file
```python
from src.core import load_config
from src.interface import OfficeConverter

# Load from YAML file
config = load_config('my_config.yaml')
converter = OfficeConverter(config)
```

## Minimal Example (Zero Configuration)

```python
# Just copy the src folder and import!
from src.interface import convert_single

# Works with default settings
result = convert_single('any_office_file.xlsx', 'output.pdf')

if result.success:
    print(f"✓ Converted in {result.duration:.2f}s")
else:
    print(f"✗ Failed: {result.error}")
```

## Complete Integration Example

```python
import os
from src.interface import OfficeConverter

def convert_office_files_in_folder(input_folder, output_folder):
    """
    Convert all Office files in a folder to PDF.
    
    Args:
        input_folder: Folder containing Office files
        output_folder: Folder to save PDFs
    """
    # Custom configuration
    config = {
        'word_options': {'create_bookmarks': True},
        'excel': {'prepare_for_print': False},  # Faster conversion
        'pdf_trim': {'enabled': False}  # Keep original margins
    }
    
    # Initialize converter
    converter = OfficeConverter(config)
    
    # Scan for Office files
    office_files = []
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            filepath = os.path.join(root, file)
            if converter.is_supported(filepath):
                office_files.append(filepath)
    
    print(f"Found {len(office_files)} Office files")
    
    # Convert all files
    results = converter.convert_batch(
        input_files=office_files,
        output_dir=output_folder,
        preserve_structure=True,
        base_dir=input_folder
    )
    
    # Report results
    stats = converter.get_conversion_statistics(results)
    print(f"\nConversion Complete:")
    print(f"  Total: {stats['total']}")
    print(f"  Successful: {stats['successful']}")
    print(f"  Failed: {stats['failed']}")
    print(f"  Success rate: {stats['success_rate']:.1f}%")
    
    # Show failed files
    for result in results:
        if not result.success:
            print(f"  ✗ {result.input_path}: {result.error}")
    
    return results

# Usage
if __name__ == '__main__':
    results = convert_office_files_in_folder('./input', './output')
```

## Requirements

Only these packages needed (no config file required):
```
pywin32>=300
rich>=10.0.0
pyyaml>=5.4.0  # Only if using YAML config files
psutil>=5.8.0
pypdf>=3.0.0
```

Install:
```bash
pip install pywin32 rich pyyaml psutil pypdf
```

## Troubleshooting

### Import Error
If you get import errors after copying the folder:
```python
# Make sure the folder name matches your imports
# If you renamed 'src' to 'office_converter', use:
from office_converter.interface import convert_single

# Or add to Python path
import sys
sys.path.insert(0, '/path/to/copied/folder')
from src.interface import convert_single
```

### No Config File
The package works without any config file! Just pass config as dict:
```python
converter = OfficeConverter({'word_options': {'create_bookmarks': True}})
```

### Missing Dependencies
```bash
pip install pywin32 rich pypdf psutil pyyaml
```

## API Quick Reference

```python
from src.interface import OfficeConverter, convert_single, convert_batch

# Single file
result = convert_single(input_path, output_path, config=None)

# Batch
results = convert_batch(files, output_dir, config=None)

# With converter class
converter = OfficeConverter(config=None)  # config is optional
result = converter.convert(input_path, output_path)
results = converter.convert_batch(files, output_dir)
stats = converter.get_conversion_statistics(results)
```

## That's It!

Just copy the `src` folder and start converting Office files to PDF! 🎉
