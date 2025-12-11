# xlsx2pdf

Excel to PDF converter with advanced features including language classification and distribution.

## Features

- Convert Excel files (.xls, .xlsx, .xlsm) to PDF
- Multiple print modes (auto, one_page, table_row_break, etc.)
- Custom headers and footers with sheet name and row range
- **Language classification and automatic distribution**
- Preserves folder structure in output
- Batch processing with progress tracking

## Language Classification

The converter can automatically classify and distribute PDF files based on language:

### Configuration

Edit `config.yaml`:

```yaml
language_classification:
  enabled: true              # Enable/disable language classification
  mode: "auto"               # Mode: "auto" or "filename"
  
  # Filename patterns (when mode: filename)
  filename_patterns:
    vi: ["_VN"]              # Vietnamese
    en: ["_EN", "_t"]        # English
    ja: ["_Ja", ""]          # Japanese
  
  output_suffix_format: "output-{lang}"  # Output folder format
  keep_folder_structure: true            # Maintain input folder structure
```

### Modes

1. **Auto Mode** (`mode: "auto"`)
   - Detects language from cell content using langdetect library
   - Analyzes text in all sheets
   - Automatically distributes files to `output-<lang>` folders
   - Supported languages: vi, en, ja, zh, ko, th, fr, de, es

2. **Filename Mode** (`mode: "filename"`)
   - Classifies based on filename patterns
   - Example: `Report_VN.xlsx` → `output-vi/`
   - Example: `Data_EN.xlsx` → `output-en/`
   - Example: `Info_Ja.xlsx` → `output-ja/`
   - Files without pattern → `output/` (default)

### Output Structure

With `keep_folder_structure: true`:
```
input/
  ├── folder1/
  │   ├── file_VN.xlsx
  │   └── file_EN.xlsx
  └── folder2/
      └── file_Ja.xlsx

output-vi/
  └── folder1/
      └── file_VN_x.pdf

output-en/
  └── folder1/
      └── file_EN_x.pdf

output-ja/
  └── folder2/
      └── file_Ja_x.pdf
```

## Installation

```bash
pip install -r requirements.txt
```

For language detection support:
```bash
pip install langdetect
```

## Usage

```bash
python main.py
```