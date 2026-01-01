# Migration Guide: v1 to v2

## 📦 What Changed?

The codebase has been refactored with:
- ✅ Feature-based architecture
- ✅ Support for Word and PowerPoint (in addition to Excel)
- ✅ Clean interface layer for easy integration
- ✅ Following Python best practices and design patterns

## 🔄 Import Changes

### Old Structure
```
src/
├── converter.py
├── utils.py
├── logger.py
├── language_detector.py
├── pdf_trimmer.py
└── ui.py
```

### New Structure
```
src/
├── core/
│   ├── base_converter.py
│   ├── utils.py
│   ├── logger.py
│   └── language_detector.py
├── features/
│   ├── excel/
│   │   ├── excel_converter.py
│   │   └── pdf_trimmer.py
│   ├── word/
│   │   └── word_converter.py
│   └── powerpoint/
│       └── powerpoint_converter.py
└── interface/
    └── converter_interface.py
```

## 📝 Code Migration

### Scenario 1: Basic Excel Conversion

**Before:**
```python
from src.converter import ExcelConverter

converter = ExcelConverter(config)
converter.convert(input_path, output_path)
```

**After (Option 1 - Recommended):**
```python
from src.interface import convert_single

result = convert_single(input_path, output_path)
if not result.success:
    print(f"Error: {result.error}")
```

**After (Option 2 - Direct Access):**
```python
from src.features.excel import ExcelConverter

converter = ExcelConverter(config)
converter.convert(input_path, output_path)
```

### Scenario 2: Utilities Import

**Before:**
```python
from src.utils import load_config, ensure_dir
from src.logger import setup_logger
```

**After:**
```python
from src.core.utils import load_config, ensure_dir
from src.core.logger import setup_logger
```

### Scenario 3: Language Detector

**Before:**
```python
from src.language_detector import LanguageDetector
```

**After:**
```python
from src.core.language_detector import LanguageDetector
```

### Scenario 4: Custom Configuration

**Before:**
```python
from src.utils import load_config
from src.converter import ExcelConverter

config = load_config()
converter = ExcelConverter(config)
```

**After:**
```python
from src.core import load_config
from src.interface import OfficeConverter

config = load_config()
converter = OfficeConverter(config)
# Now supports Excel, Word, AND PowerPoint!
```

## 🆕 New Features

### 1. Word Conversion
```python
from src.interface import convert_single

# Just works!
result = convert_single('document.docx', 'output.pdf')
```

### 2. PowerPoint Conversion
```python
from src.interface import convert_single

result = convert_single('presentation.pptx', 'output.pdf')
```

### 3. Mixed Batch Conversion
```python
from src.interface import convert_batch

files = [
    'report.docx',      # Word
    'data.xlsx',        # Excel
    'slides.pptx'       # PowerPoint
]

results = convert_batch(files, 'output_folder')
```

### 4. Unified Interface
```python
from src.interface import OfficeConverter

converter = OfficeConverter()

# Works for all Office file types
converter.convert('any_office_file.docx', 'output.pdf')
converter.convert('any_office_file.xlsx', 'output.pdf')
converter.convert('any_office_file.pptx', 'output.pdf')
```

### 5. Statistics and Reporting
```python
from src.interface import OfficeConverter

converter = OfficeConverter()
results = converter.convert_batch(files, 'output')

stats = converter.get_conversion_statistics(results)
print(f"Success rate: {stats['success_rate']:.1f}%")
print(f"By type: {stats['by_type']}")
```

## 🔧 main.py Changes

### Command Line Options

**Before:**
```bash
python main.py --input ./input --output ./output
# Only converted Excel files
```

**After:**
```bash
# Convert all Office files
python main.py --input ./input --output ./output --file-types all

# Convert only Excel
python main.py --file-types excel

# Convert Word and Excel
python main.py --file-types "word,excel"

# Convert PowerPoint only
python main.py --file-types powerpoint
```

## 📋 Configuration File

The `config.yaml` structure is backward compatible! Old Excel configurations still work.

**New Options Available:**

```yaml
# Word-specific options (NEW)
word_options:
  create_bookmarks: true
  optimize_for_print: true
  include_doc_properties: true

# PowerPoint-specific options (NEW)
powerpoint_options:
  output_type: 'slides'  # slides, notes, handouts
  include_hidden_slides: false
  frame_slides: false

# Excel options (UNCHANGED)
excel:
  prepare_for_print: true

print_options:
  mode: 'auto'
  page_size: 'A4'
```

## 🔍 Testing Migration

### Step 1: Update Imports
Run your code and fix import errors:
- `src.converter` → `src.features.excel` or `src.interface`
- `src.utils` → `src.core.utils`
- `src.logger` → `src.core.logger`

### Step 2: Test Functionality
Ensure your conversions still work:
```python
from src.interface import convert_single

# Test with existing Excel file
result = convert_single('test.xlsx', 'test.pdf')
assert result.success
```

### Step 3: Update Configuration
Review and add new options to `config.yaml` if needed.

### Step 4: Test New Features
Try converting Word or PowerPoint files:
```python
result = convert_single('test.docx', 'test.pdf')
result = convert_single('test.pptx', 'test.pdf')
```

## ⚠️ Breaking Changes

### Import Paths
All direct imports from `src.converter`, `src.utils`, etc. need to be updated to new paths.

### No Functional Breaking Changes
The actual conversion functionality remains the same. Only the organization changed.

## 💡 Recommended Approach

**For New Projects:**
Use the interface layer:
```python
from src.interface import convert_single, convert_batch, OfficeConverter
```

**For Existing Projects:**
Update imports gradually:
1. Update utility imports first (`src.core.*`)
2. Then update converter imports (`src.features.*` or `src.interface`)
3. Test thoroughly

**For Maximum Compatibility:**
Use the unified interface which supports all file types:
```python
from src.interface import OfficeConverter

converter = OfficeConverter(config)
converter.convert(any_office_file, output_pdf)
```

## 📞 Need Help?

- Check `examples.py` for code samples
- Review `docs/architecture_v2.md` for architecture details
- See `docs/quick_start_v2.md` for quick reference

## ✅ Migration Checklist

- [ ] Update all imports to new structure
- [ ] Test Excel conversions still work
- [ ] (Optional) Try Word conversions
- [ ] (Optional) Try PowerPoint conversions
- [ ] Update `config.yaml` with new options
- [ ] Update any custom scripts/integrations
- [ ] Run full test suite
- [ ] Update documentation/comments

## 🎉 Benefits After Migration

✅ Support for Word and PowerPoint files
✅ Cleaner, more maintainable code structure
✅ Better error handling with `ConversionResult`
✅ Easy integration via interface layer
✅ Statistics and batch processing improvements
✅ Following Python best practices
✅ Ready for future extensions
