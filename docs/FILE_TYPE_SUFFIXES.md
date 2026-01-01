# File-Type Specific Output Suffixes

## Overview

The system now supports different output suffixes for each Office file type, making it easier to distinguish PDF outputs by their original format.

## Configuration

In `config.yaml`, suffixes are defined per file type:

```yaml
# Excel configuration
excel:
  output_suffix: "_x"       # Excel files: report.xlsx → report_x.pdf

# Word configuration  
word_options:
  output_suffix: "_d"       # Word files: document.docx → document_d.pdf

# PowerPoint configuration
powerpoint_options:
  output_suffix: "_p"       # PowerPoint files: slides.pptx → slides_p.pdf
```

## Default Suffixes

| File Type | Default Suffix | Example |
|-----------|---------------|---------|
| Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`) | `_x` | `report.xlsx` → `report_x.pdf` |
| Word (`.docx`, `.doc`, `.docm`, `.dotx`, `.dotm`) | `_d` | `document.docx` → `document_d.pdf` |
| PowerPoint (`.pptx`, `.ppt`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, `.potm`) | `_p` | `slides.pptx` → `slides_p.pdf` |

## Usage

### Single File Conversion

```python
from src.interface import convert_single

# Excel file - will use _x suffix
convert_single('report.xlsx', 'output/report_x.pdf')

# Word file - will use _d suffix  
convert_single('document.docx', 'output/document_d.pdf')

# PowerPoint file - will use _p suffix
convert_single('slides.pptx', 'output/slides_p.pdf')
```

### Batch Conversion

The batch converter automatically applies the correct suffix based on file type:

```python
from src.interface import convert_batch

files = ['report.xlsx', 'document.docx', 'slides.pptx']
results = convert_batch(files, 'output_folder')

# Output files:
# - output_folder/report_x.pdf
# - output_folder/document_d.pdf
# - output_folder/slides_p.pdf
```

### Using OfficeConverter Class

```python
from src.interface import OfficeConverter

converter = OfficeConverter(config)

# Get the appropriate suffix for a file
excel_suffix = converter.get_output_suffix('report.xlsx')  # Returns '_x'
word_suffix = converter.get_output_suffix('document.docx')  # Returns '_d'
ppt_suffix = converter.get_output_suffix('slides.pptx')    # Returns '_p'
```

## Customization

You can customize the suffixes in `config.yaml`:

```yaml
excel:
  output_suffix: "_excel"

word_options:
  output_suffix: "_word"

powerpoint_options:
  output_suffix: "_slides"
```

This produces:
- `report.xlsx` → `report_excel.pdf`
- `document.docx` → `document_word.pdf`
- `slides.pptx` → `slides_slides.pdf`

## Implementation Details

### Components Updated

1. **config.yaml**: Added `output_suffix` to `excel`, `word_options`, and `powerpoint_options`
2. **src/interface/converter_interface.py**: Added `get_output_suffix()` method
3. **main.py**: Added `get_file_type_suffix()` helper function
4. **src/core/language_detector.py**: Added `get_output_path_with_suffix()` method

### Backward Compatibility

The system maintains backward compatibility:
- If no suffix is defined in config, defaults are used (`_x`, `_d`, `_p`)
- Old code using global `output_suffix` still works (now commented in config)
- All default configs include the new suffix settings

## Testing

Run the test suite to verify suffix behavior:

```bash
python test_suffix.py
```

Expected output:
```
Excel (.xlsx) suffix: _x
Excel (.xls) suffix: _x
Word (.docx) suffix: _d
Word (.doc) suffix: _d
PowerPoint (.pptx) suffix: _p
PowerPoint (.ppt) suffix: _p

✅ All suffix tests passed!
```

## Benefits

1. **Clear Identification**: Easy to identify the source file type from PDF name
2. **Organized Outputs**: Mixed conversions are easier to sort and manage
3. **Flexible Configuration**: Per-type suffixes allow custom naming schemes
4. **Scenario Support**: Each scenario group can have different suffixes
5. **Language Detection Compatible**: Works seamlessly with language classification

## Examples

### Mixed File Batch

```python
files = [
    'Q1_report.xlsx',
    'Q1_summary.docx', 
    'Q1_presentation.pptx',
    'Q2_report.xlsx',
    'Q2_summary.docx'
]

convert_batch(files, 'output')
```

Output structure:
```
output/
├── Q1_report_x.pdf
├── Q1_summary_d.pdf
├── Q1_presentation_p.pdf
├── Q2_report_x.pdf
└── Q2_summary_d.pdf
```

### Scenario Mode with Custom Suffixes

```yaml
# configs/finance_config.yaml
excel:
  output_suffix: "_finance"

# configs/hr_config.yaml  
word_options:
  output_suffix: "_hr"
```

Each department's files get their appropriate suffix automatically based on the scenario configuration.
