# Using Office to PDF Converter in Your Project

## 🚀 Quick Start (Copy & Use)

### Step 1: Copy the `src` folder to your project

```bash
cp -r /path/to/xlsx2pdf/src /path/to/your-project/office2pdf
```

### Step 2: Import and use (no config file needed!)

```python
from office2pdf.interface import convert_single

# That's it! Start converting
result = convert_single('document.docx', 'output.pdf')
print(f"Success: {result.success}")
```

## 📦 What You Get

After copying the `src` folder, you get a complete, standalone converter that:
- ✅ Works **without** any config file
- ✅ Converts Excel, Word, and PowerPoint to PDF
- ✅ Has sensible default settings
- ✅ Can be customized via code (no YAML needed)
- ✅ No external dependencies on the original project

## 💡 Simple Examples

### Convert a Single File
```python
from office2pdf.interface import convert_single

result = convert_single('report.docx', 'report.pdf')
if result.success:
    print(f"✓ Converted in {result.duration:.2f} seconds")
```

### Convert Multiple Files
```python
from office2pdf.interface import convert_batch

files = ['doc1.docx', 'sheet1.xlsx', 'pres1.pptx']
results = convert_batch(files, 'output_folder')

print(f"{sum(r.success for r in results)} files converted")
```

### With Custom Configuration
```python
from office2pdf.interface import OfficeConverter

config = {
    'word_options': {'create_bookmarks': True},
    'excel': {'prepare_for_print': False},
    'pdf_trim': {'enabled': False}
}

converter = OfficeConverter(config)
result = converter.convert('file.xlsx', 'output.pdf')
```

## 🔧 Configuration (Optional)

You can pass configuration as a Python dictionary - **no YAML file needed**:

```python
from office2pdf.interface import OfficeConverter

my_config = {
    # Excel settings
    'excel': {
        'prepare_for_print': True,
        'enhanced_dir': 'enhanced_files'
    },
    
    # Excel print options
    'print_options': {
        'mode': 'auto',          # auto, one_page, table_row_break
        'page_size': 'A4',       # A4, LETTER, A3, etc.
        'orientation': 'auto',    # auto, portrait, landscape
        'scaling': 'fit_columns',
        'margins': 'normal'      # normal, wide, narrow
    },
    
    # PDF trimming (for Excel)
    'pdf_trim': {
        'enabled': True,
        'margin_threshold': 10,
        'min_margin': 5
    },
    
    # Word settings
    'word_options': {
        'create_bookmarks': True,
        'optimize_for_print': True,
        'include_doc_properties': True,
        'keep_form_fields': True
    },
    
    # PowerPoint settings
    'powerpoint_options': {
        'output_type': 'slides',      # slides, notes, handouts
        'handout_order': 'vertical',  # vertical, horizontal
        'slides_per_page': 1,
        'include_hidden_slides': False,
        'frame_slides': False,
        'print_comments': False
    }
}

converter = OfficeConverter(config=my_config)
```

## 📋 Supported File Types

- **Excel**: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`
- **Word**: `.docx`, `.doc`, `.docm`, `.dotx`, `.dotm`
- **PowerPoint**: `.pptx`, `.ppt`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, `.potm`

## 🎯 Complete Integration Example

```python
import os
from office2pdf.interface import OfficeConverter

def convert_all_office_files(input_dir, output_dir):
    """Convert all Office files in a directory to PDF"""
    
    # Initialize with custom config (or use defaults)
    converter = OfficeConverter({
        'excel': {'prepare_for_print': False},  # Faster
        'pdf_trim': {'enabled': False}          # Keep margins
    })
    
    # Scan for Office files
    files = []
    for root, dirs, filenames in os.walk(input_dir):
        for filename in filenames:
            filepath = os.path.join(root, filename)
            if converter.is_supported(filepath):
                files.append(filepath)
    
    print(f"Found {len(files)} Office files")
    
    # Convert all files
    results = converter.convert_batch(
        input_files=files,
        output_dir=output_dir,
        preserve_structure=True,
        base_dir=input_dir
    )
    
    # Show statistics
    stats = converter.get_conversion_statistics(results)
    print(f"\n✓ Converted: {stats['successful']}/{stats['total']}")
    print(f"  Success rate: {stats['success_rate']:.1f}%")
    print(f"  Total time: {stats['total_duration']:.2f}s")
    
    return results

# Usage
results = convert_all_office_files('./input', './output')
```

## 🔌 Minimal Integration Function

Copy this function to your project for minimal integration:

```python
from office2pdf.interface import convert_single

def to_pdf(office_file, pdf_file=None):
    """
    Convert any Office file to PDF.
    
    Args:
        office_file: Path to Office file (.docx, .xlsx, .pptx)
        pdf_file: Output path (optional, auto-generated if None)
    
    Returns:
        True if successful, False otherwise
    """
    if pdf_file is None:
        pdf_file = os.path.splitext(office_file)[0] + '.pdf'
    
    result = convert_single(office_file, pdf_file)
    return result.success

# Usage - super simple!
success = to_pdf('document.docx')
success = to_pdf('spreadsheet.xlsx', 'custom_output.pdf')
```

## 📦 Dependencies

Install these packages in your project:

```bash
pip install pywin32 rich pypdf psutil pyyaml
```

Or add to your `requirements.txt`:
```
pywin32>=300
rich>=10.0.0
pypdf>=3.0.0
psutil>=5.8.0
pyyaml>=5.4.0
```

## 🎓 API Quick Reference

### Simple Functions
```python
from office2pdf.interface import convert_single, convert_batch

# Single file
result = convert_single(input_path, output_path)

# Batch
results = convert_batch(files, output_dir)
```

### Converter Class
```python
from office2pdf.interface import OfficeConverter

converter = OfficeConverter(config=None)  # config is optional!

# Convert
result = converter.convert(input_path, output_path)

# Batch convert
results = converter.convert_batch(files, output_dir)

# Check support
supported = converter.is_supported('file.xlsx')

# Get statistics
stats = converter.get_conversion_statistics(results)
```

### ConversionResult Object
```python
result = convert_single('file.docx', 'output.pdf')

print(result.success)      # True/False
print(result.error)        # Error message if failed
print(result.duration)     # Time in seconds
print(result.file_type)    # 'word', 'excel', or 'powerpoint'
print(result.input_path)   # Original file path
print(result.output_path)  # PDF output path
```

## ⚠️ Platform Support

- **Windows Only** (uses Microsoft Office COM automation)
- Requires Microsoft Office installed
- Python 3.7+

## 🎉 That's It!

No complex setup, no config files required, just copy and use!

```python
from office2pdf.interface import convert_single
convert_single('any_file.docx', 'output.pdf')  # Done! 🎊
```

## 📚 More Examples

See `standalone_example.py` for more usage examples including:
- Error handling
- Custom configuration
- Scanning directories
- Creating wrapper classes
- And more!

---

**Need Help?**
- Check `docs/API_REFERENCE.md` for complete API documentation
- See `examples.py` for comprehensive examples
- Read `docs/PORTABLE_USAGE.md` for detailed integration guide
