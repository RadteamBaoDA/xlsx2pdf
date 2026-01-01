# Office to PDF Converter - Quick Start Guide

## 🚀 Quick Usage

### 1. Single File Conversion
```python
from src.interface import convert_single

# Convert any Office file
result = convert_single('document.docx', 'output.pdf')
print(f"✓ Converted: {result.success}")
```

### 2. Batch Conversion
```python
from src.interface import convert_batch

files = ['report.docx', 'data.xlsx', 'slides.pptx']
results = convert_batch(files, 'output_folder')
```

### 3. Command Line
```bash
# Convert all Office files
python main.py --input ./input --output ./output

# Convert only Excel files
python main.py --file-types excel

# Convert Word and PowerPoint only
python main.py --file-types "word,powerpoint"
```

## 📋 Supported File Types

| Type | Extensions |
|------|-----------|
| **Excel** | .xlsx, .xls, .xlsm, .xlsb |
| **Word** | .docx, .doc, .docm, .dotx |
| **PowerPoint** | .pptx, .ppt, .pptm, .ppsx |

## 🎯 Common Use Cases

### Use Case 1: Simple Integration
```python
from src.interface import convert_single

def my_function(office_file):
    result = convert_single(office_file, 'output.pdf')
    return result.success
```

### Use Case 2: Batch with Error Handling
```python
from src.interface import OfficeConverter

converter = OfficeConverter()
results = converter.convert_batch(files, 'output')

for result in results:
    if result.success:
        print(f"✓ {result.input_path}")
    else:
        print(f"✗ {result.input_path}: {result.error}")
```

### Use Case 3: Statistics
```python
converter = OfficeConverter()
results = converter.convert_batch(files, 'output')
stats = converter.get_conversion_statistics(results)

print(f"Success rate: {stats['success_rate']:.1f}%")
```

## 🔧 Configuration

Edit `config.yaml` for options:

```yaml
# Excel specific
excel:
  prepare_for_print: true

print_options:
  mode: 'auto'
  page_size: 'A4'
  orientation: 'auto'

# Word specific
word_options:
  create_bookmarks: true
  optimize_for_print: true

# PowerPoint specific  
powerpoint_options:
  output_type: 'slides'
  include_hidden_slides: false
```

## 📦 Installation

```bash
pip install -r requirements.txt
```

## 🏗️ Architecture

```
src/
├── core/          # Base classes & utilities
├── features/      # Converters (excel, word, powerpoint)
└── interface/     # Easy-to-use API
```

## 🔍 Check Supported Files

```python
from src.interface import OfficeConverter

converter = OfficeConverter()
if converter.is_supported('myfile.xlsx'):
    converter.convert('myfile.xlsx', 'output.pdf')
```

## 📖 More Examples

See `examples.py` for complete examples.

## ❓ FAQ

**Q: Can I convert multiple file types at once?**
A: Yes! Use `convert_batch()` with mixed file types.

**Q: How do I customize conversion settings?**
A: Edit `config.yaml` or pass custom config to converters.

**Q: What if a file fails to convert?**
A: Check the `ConversionResult.error` field for details.

**Q: Can I use this in my own project?**
A: Yes! Import from `src.interface` and use the provided functions.

## 🔗 Related Docs

- [Architecture Guide](docs/architecture_v2.md) - Full architecture details
- [Configuration Guide](docs/configuration.md) - All config options
- [API Reference](examples.py) - Code examples
