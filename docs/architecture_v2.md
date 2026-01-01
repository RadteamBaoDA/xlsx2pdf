# Office to PDF Converter - Architecture Guide

## 📁 New Project Structure

The project has been refactored to follow a **feature-based architecture** with Python best practices:

```
src/
├── core/                      # Core functionality and utilities
│   ├── __init__.py
│   ├── base_converter.py     # Abstract base class for all converters
│   ├── utils.py              # Common utilities
│   ├── logger.py             # Logging configuration
│   └── language_detector.py  # Language detection
│
├── features/                  # Feature-based converters
│   ├── __init__.py
│   ├── excel/                # Excel conversion feature
│   │   ├── __init__.py
│   │   ├── excel_converter.py
│   │   └── pdf_trimmer.py
│   ├── word/                 # Word conversion feature
│   │   ├── __init__.py
│   │   └── word_converter.py
│   └── powerpoint/           # PowerPoint conversion feature
│       ├── __init__.py
│       └── powerpoint_converter.py
│
├── interface/                # Clean API for external integration
│   ├── __init__.py
│   └── converter_interface.py
│
├── __init__.py               # Package initialization
└── ui.py                     # UI components (unchanged)
```

## 🎯 Design Patterns Used

### 1. **Strategy Pattern** (Base Converter)
- `BaseConverter` abstract class defines the conversion interface
- Each converter implements the strategy for its file type
- Easy to add new converters without modifying existing code

### 2. **Factory Pattern** (Converter Selection)
- `OfficeConverter` automatically selects the right converter based on file extension
- Centralized converter management

### 3. **Template Method Pattern** (Conversion Workflow)
- Base class defines the conversion workflow structure
- Subclasses implement specific steps

### 4. **Facade Pattern** (Interface Layer)
- `OfficeConverter` provides a simple interface to complex conversion logic
- Hides implementation details from external code

## 🚀 Usage

### Simple Single File Conversion

```python
from src.interface import convert_single

result = convert_single('document.docx', 'output.pdf')
print(f"Success: {result.success}")
```

### Batch Conversion

```python
from src.interface import convert_batch

files = ['file1.xlsx', 'file2.docx', 'file3.pptx']
results = convert_batch(files, 'output_folder')

for result in results:
    print(result)
```

### Using the OfficeConverter Class

```python
from src.interface import OfficeConverter

converter = OfficeConverter()

# Single conversion
result = converter.convert('input.pptx', 'output.pdf')

# Batch conversion with options
results = converter.convert_batch(
    input_files=['file1.xlsx', 'file2.docx'],
    output_dir='output',
    preserve_structure=True
)

# Get statistics
stats = converter.get_conversion_statistics(results)
print(f"Success rate: {stats['success_rate']:.1f}%")
```

### Direct Converter Access

```python
from src.features.excel import ExcelConverter
from src.features.word import WordConverter
from src.features.powerpoint import PowerPointConverter

config = {'excel': {...}, 'word_options': {...}}

# Use specific converter
excel_conv = ExcelConverter(config)
excel_conv.convert('spreadsheet.xlsx', 'output.pdf')

word_conv = WordConverter(config)
word_conv.convert('document.docx', 'output.pdf')

ppt_conv = PowerPointConverter(config)
ppt_conv.convert('presentation.pptx', 'output.pdf')
```

## 📊 Supported File Types

### Excel
- `.xlsx` - Excel Workbook
- `.xls` - Excel 97-2003 Workbook
- `.xlsm` - Excel Macro-Enabled Workbook
- `.xlsb` - Excel Binary Workbook

### Word
- `.docx` - Word Document
- `.doc` - Word 97-2003 Document
- `.docm` - Word Macro-Enabled Document
- `.dotx` - Word Template
- `.dotm` - Word Macro-Enabled Template

### PowerPoint
- `.pptx` - PowerPoint Presentation
- `.ppt` - PowerPoint 97-2003 Presentation
- `.pptm` - PowerPoint Macro-Enabled Presentation
- `.ppsx` - PowerPoint Show
- `.ppsm` - PowerPoint Macro-Enabled Show
- `.potx` - PowerPoint Template
- `.potm` - PowerPoint Macro-Enabled Template

## 🔧 Configuration

### Excel Options (config.yaml)
```yaml
excel:
  prepare_for_print: true
  enhanced_dir: 'enhanced_files'
  
print_options:
  mode: 'auto'  # auto, one_page, table_row_break, etc.
  page_size: 'A4'
  orientation: 'auto'
  
pdf_trim:
  enabled: true
  margin_threshold: 10
```

### Word Options
```yaml
word_options:
  create_bookmarks: true
  optimize_for_print: true
  include_doc_properties: true
  keep_form_fields: true
```

### PowerPoint Options
```yaml
powerpoint_options:
  output_type: 'slides'  # slides, notes, handouts
  handout_order: 'vertical'  # vertical, horizontal
  slides_per_page: 1
  include_hidden_slides: false
  frame_slides: false
  print_comments: false
```

## 🧪 Testing

Run examples:
```bash
python examples.py
```

Run main converter:
```bash
# Convert all Office files
python main.py --input ./input --output ./output --file-types all

# Convert only Excel files
python main.py --input ./input --output ./output --file-types excel

# Convert specific types
python main.py --input ./input --output ./output --file-types "excel,word"
```

## 🔄 Migration Guide

### Old Code
```python
from src.converter import ExcelConverter
converter = ExcelConverter(config)
converter.convert(input_path, output_path)
```

### New Code
```python
# Option 1: Use unified interface (recommended)
from src.interface import convert_single
result = convert_single(input_path, output_path)

# Option 2: Use OfficeConverter
from src.interface import OfficeConverter
converter = OfficeConverter(config)
result = converter.convert(input_path, output_path)

# Option 3: Use specific converter (for Excel-specific features)
from src.features.excel import ExcelConverter
converter = ExcelConverter(config)
converter.convert(input_path, output_path)
```

## 🎓 Benefits of New Architecture

1. **Separation of Concerns**: Each module has a single, well-defined responsibility
2. **Open/Closed Principle**: Easy to extend with new converters without modifying existing code
3. **Single Responsibility**: Each converter handles only its file type
4. **Interface Segregation**: Clean, simple interfaces for different use cases
5. **Dependency Inversion**: High-level interface doesn't depend on low-level converter details
6. **Easy Testing**: Each component can be tested independently
7. **Better Maintainability**: Feature-based structure makes code easier to understand and maintain
8. **Reusability**: Core utilities and base classes can be reused across converters

## 📝 Adding New Converters

To add a new converter (e.g., for Outlook messages):

1. Create feature folder: `src/features/outlook/`
2. Create converter class:
```python
from ...core.base_converter import BaseConverter

class OutlookConverter(BaseConverter):
    @property
    def supported_extensions(self):
        return ['.msg', '.eml']
    
    def validate_input(self, input_path):
        return self.can_convert(input_path)
    
    def convert(self, input_path, output_path, pid_queue=None):
        # Implementation
        pass
```
3. Register in `interface/converter_interface.py`:
```python
CONVERTER_MAP = {
    # ... existing
    '.msg': ('outlook', OutlookConverter),
    '.eml': ('outlook', OutlookConverter),
}
```

## 🤝 Integration with External Code

The interface layer makes it easy for other projects to use this converter:

```python
# In external project
from xlsx2pdf.src.interface import convert_single, convert_batch

# Convert file
convert_single('my_file.xlsx', 'output.pdf')

# Batch conversion
files = get_files_from_somewhere()
results = convert_batch(files, 'output_folder')
```

## 📚 API Reference

See `examples.py` for comprehensive usage examples of all features.

## 🐛 Debugging

Enable detailed logging:
```python
import logging
logging.basicConfig(level=logging.DEBUG)

from src.interface import convert_single
result = convert_single('file.xlsx', 'output.pdf')
```

## 🔒 Thread Safety

Each converter creates its own COM instance, making them safe for sequential use. For parallel processing, use multiprocessing (not threading) due to COM limitations.

## 📦 Dependencies

- **pywin32**: Windows COM automation
- **pypdf**: PDF manipulation
- **psutil**: Process management
- **rich**: Terminal UI
- **pyyaml**: Configuration
- **langdetect**: Language detection (optional)
