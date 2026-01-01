# API Reference - Office to PDF Converter

## Quick Import Guide

```python
# Simple functions (recommended for most use cases)
from src.interface import convert_single, convert_batch

# Main converter class
from src.interface import OfficeConverter, ConversionResult

# Specific converters (for advanced use)
from src.features.excel import ExcelConverter
from src.features.word import WordConverter
from src.features.powerpoint import PowerPointConverter

# Core utilities
from src.core import load_config, setup_logger
```

---

## Interface Layer Functions

### `convert_single(input_path, output_path, config=None)`

Convert a single Office file to PDF.

**Parameters:**
- `input_path` (str): Path to input Office file
- `output_path` (str): Path where PDF should be saved
- `config` (dict, optional): Custom configuration

**Returns:** `ConversionResult` object

**Example:**
```python
from src.interface import convert_single

result = convert_single('document.docx', 'output.pdf')
if result.success:
    print(f"Converted in {result.duration:.2f}s")
else:
    print(f"Failed: {result.error}")
```

---

### `convert_batch(input_files, output_dir, config=None, preserve_structure=True, base_dir=None, max_workers=None)`

Convert multiple Office files to PDF in batch.

**Parameters:**
- `input_files` (list): List of input file paths
- `output_dir` (str): Output directory for PDFs
- `config` (dict, optional): Custom configuration
- `preserve_structure` (bool): Keep folder structure (default: True)
- `base_dir` (str, optional): Base directory for relative paths
- `max_workers` (int, optional): Number of parallel workers

**Returns:** List of `ConversionResult` objects

**Example:**
```python
from src.interface import convert_batch

files = ['doc1.docx', 'sheet1.xlsx', 'pres1.pptx']
results = convert_batch(files, 'output_folder')

for result in results:
    print(f"{result.input_path}: {'✓' if result.success else '✗'}")
```

---

## OfficeConverter Class

Main converter class with automatic file type detection.

### Constructor

```python
OfficeConverter(config=None)
```

**Parameters:**
- `config` (dict, optional): Configuration dictionary. Loads from config.yaml if None.

**Example:**
```python
from src.interface import OfficeConverter

converter = OfficeConverter()
# or with custom config
converter = OfficeConverter({'word_options': {'create_bookmarks': True}})
```

---

### Methods

#### `convert(input_path, output_path, pid_queue=None)`

Convert a single file.

**Parameters:**
- `input_path` (str): Input file path
- `output_path` (str): Output PDF path
- `pid_queue` (Queue, optional): Multiprocessing queue for PID

**Returns:** `ConversionResult`

**Example:**
```python
result = converter.convert('file.xlsx', 'output.pdf')
```

---

#### `convert_batch(input_files, output_dir, preserve_structure=True, base_dir=None, max_workers=None)`

Convert multiple files.

**Returns:** List of `ConversionResult` objects

**Example:**
```python
results = converter.convert_batch(
    input_files=['file1.xlsx', 'file2.docx'],
    output_dir='output',
    preserve_structure=True
)
```

---

#### `is_supported(file_path)`

Check if file type is supported.

**Parameters:**
- `file_path` (str): Path to file

**Returns:** bool

**Example:**
```python
if converter.is_supported('document.docx'):
    converter.convert('document.docx', 'output.pdf')
```

---

#### `get_converter(file_path)`

Get appropriate converter for a file.

**Parameters:**
- `file_path` (str): Path to file

**Returns:** Tuple of (converter_instance, file_type) or (None, None)

**Example:**
```python
converter_obj, file_type = converter.get_converter('file.xlsx')
if converter_obj:
    print(f"File type: {file_type}")  # "excel"
```

---

#### `get_conversion_statistics(results)`

Calculate statistics from conversion results.

**Parameters:**
- `results` (list): List of ConversionResult objects

**Returns:** Dictionary with statistics

**Example:**
```python
results = converter.convert_batch(files, 'output')
stats = converter.get_conversion_statistics(results)

print(f"Total: {stats['total']}")
print(f"Successful: {stats['successful']}")
print(f"Failed: {stats['failed']}")
print(f"Success rate: {stats['success_rate']:.1f}%")
print(f"By type: {stats['by_type']}")
```

---

#### `get_supported_extensions()` (static)

Get list of all supported file extensions.

**Returns:** List of extensions

**Example:**
```python
extensions = OfficeConverter.get_supported_extensions()
print(extensions)  # ['.xlsx', '.docx', '.pptx', ...]
```

---

## ConversionResult Class

Result object returned by conversion operations.

### Attributes

- `input_path` (str): Input file path
- `output_path` (str): Output PDF path (None if failed)
- `success` (bool): Whether conversion succeeded
- `error` (str): Error message if failed (None if successful)
- `duration` (float): Time taken in seconds
- `file_type` (str): Type of file ('excel', 'word', 'powerpoint')
- `timestamp` (datetime): When conversion occurred

### Methods

#### `__str__()`
String representation of result.

#### `to_dict()`
Convert to dictionary.

**Example:**
```python
result = convert_single('file.xlsx', 'output.pdf')

print(result)  # String representation
print(result.success)  # True/False
print(result.duration)  # 2.5
print(result.to_dict())  # Dictionary
```

---

## Specific Converters

### ExcelConverter

Excel-specific converter with advanced options.

```python
from src.features.excel import ExcelConverter

converter = ExcelConverter(config)
converter.convert('spreadsheet.xlsx', 'output.pdf')
```

**Supported Extensions:** `.xlsx`, `.xls`, `.xlsm`, `.xlsb`

**Configuration Options:**
```yaml
excel:
  prepare_for_print: true
print_options:
  mode: 'auto'  # auto, one_page, table_row_break, etc.
  page_size: 'A4'
  orientation: 'auto'
pdf_trim:
  enabled: true
```

---

### WordConverter

Word document converter.

```python
from src.features.word import WordConverter

converter = WordConverter(config)
converter.convert('document.docx', 'output.pdf')
```

**Supported Extensions:** `.docx`, `.doc`, `.docm`, `.dotx`, `.dotm`

**Configuration Options:**
```yaml
word_options:
  create_bookmarks: true
  optimize_for_print: true
  include_doc_properties: true
  keep_form_fields: true
```

---

### PowerPointConverter

PowerPoint presentation converter.

```python
from src.features.powerpoint import PowerPointConverter

converter = PowerPointConverter(config)
converter.convert('presentation.pptx', 'output.pdf')
```

**Supported Extensions:** `.pptx`, `.ppt`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, `.potm`

**Configuration Options:**
```yaml
powerpoint_options:
  output_type: 'slides'  # slides, notes, handouts
  handout_order: 'vertical'  # vertical, horizontal
  slides_per_page: 1  # For handouts: 1, 2, 3, 4, 6, 9
  include_hidden_slides: false
  frame_slides: false
  print_comments: false
```

---

## BaseConverter (Abstract Class)

Base class for all converters. Not used directly but useful for creating custom converters.

### Abstract Methods

Must be implemented by subclasses:

#### `convert(input_path, output_path, pid_queue=None)`
Perform the conversion.

#### `validate_input(input_path)`
Validate input file.

#### `supported_extensions` (property)
Return list of supported extensions.

### Utility Methods

Available to all converters:

#### `can_convert(file_path)`
Check if converter supports a file.

#### `_ensure_output_directory(output_path)`
Ensure output directory exists.

#### `_cleanup_temp_files(*file_paths)`
Clean up temporary files.

---

## Core Utilities

### `load_config(config_path='config.yaml')`

Load configuration from YAML file.

```python
from src.core import load_config

config = load_config('config.yaml')
```

---

### `setup_logger(log_file, error_file, log_level='INFO', logs_folder='logs')`

Setup logging system.

```python
from src.core import setup_logger

logger, log_path, error_path = setup_logger(
    'conversion.log',
    'errors.log',
    log_level='INFO'
)
```

---

### `ensure_dir(file_path)`

Ensure directory for a file exists.

```python
from src.core.utils import ensure_dir

ensure_dir('output/subfolder/file.pdf')
```

---

## Command Line Interface

```bash
python main.py [OPTIONS]
```

**Options:**
- `--input DIR` - Input directory (default: ./input)
- `--output DIR` - Output directory (default: ./output)
- `--config FILE` - Config file (default: config.yaml)
- `--file-types TYPES` - File types to convert (default: all)

**File Types:**
- `all` - All Office files
- `excel` - Excel files only
- `word` - Word documents only
- `powerpoint` - PowerPoint presentations only
- Comma-separated: `"excel,word"`, `"word,powerpoint"`, etc.

**Examples:**
```bash
# Convert all Office files
python main.py --input ./input --output ./output

# Convert only Excel files
python main.py --file-types excel

# Convert Word and PowerPoint
python main.py --file-types "word,powerpoint"

# Custom config
python main.py --config custom_config.yaml
```

---

## Error Handling

All conversions return `ConversionResult` with error information:

```python
result = convert_single('file.xlsx', 'output.pdf')

if not result.success:
    print(f"Error: {result.error}")
    # Handle error
else:
    print(f"Success! Duration: {result.duration:.2f}s")
```

Common errors:
- File not found
- Unsupported file type
- Conversion failure (corrupt file, missing dependencies, etc.)
- Permission errors

---

## Best Practices

1. **Use the interface layer** for most tasks
2. **Check `is_supported()`** before conversion
3. **Handle errors** via `ConversionResult.success`
4. **Use batch conversion** for multiple files
5. **Configure via YAML** for maintainability
6. **Use specific converters** only when needed

---

## Complete Example

```python
from src.interface import OfficeConverter
import os

# Initialize converter
converter = OfficeConverter()

# Scan directory
files = []
for root, dirs, filenames in os.walk('input'):
    for filename in filenames:
        filepath = os.path.join(root, filename)
        if converter.is_supported(filepath):
            files.append(filepath)

print(f"Found {len(files)} Office files")

# Convert all
results = converter.convert_batch(
    input_files=files,
    output_dir='output',
    preserve_structure=True,
    base_dir='input'
)

# Show statistics
stats = converter.get_conversion_statistics(results)
print(f"\nConversion complete:")
print(f"  Success rate: {stats['success_rate']:.1f}%")
print(f"  Total time: {stats['total_duration']:.2f}s")
print(f"  Average time: {stats['avg_duration']:.2f}s")

# Show failed files
for result in results:
    if not result.success:
        print(f"  Failed: {result.input_path} - {result.error}")
```

---

**Version:** 2.0.0
**Last Updated:** January 1, 2026
