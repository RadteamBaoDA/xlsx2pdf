# xlsx2pdf

Professional Excel to PDF converter with advanced layout control, multi-sheet configuration, and language classification.

## Key Features

- 🔄 **Batch Processing** - Convert multiple Excel files in one run
- 🎯 **Sheet-Specific Configuration** - Different settings for different sheets with priority system
- 📏 **Layout Preservation** - Maintains exact Excel dimensions (row heights, column widths)
- 📄 **Page Break Control** - Automatic page breaks by rows/columns
- 🌍 **Language Classification** - Auto-detect and distribute files by language
- 🎨 **Flexible Scaling** - Multiple scaling modes (fit_columns, fit_sheet, no_scaling, etc.)
- 📊 **Custom Margins & Headers** - Full control over PDF layout

## Quick Start

### 1. Installation

```bash
pip install -r requirements.txt
```

### 2. Setup Input Files

Place Excel files in the `input/` folder:
```
input/
  ├── report1.xlsx
  ├── data.xlsx
  └── subfolder/
      └── sheet.xlsx
```

### 3. Configure (Optional)

Edit `config.yaml` to customize settings. See [Configuration Guide](docs/configuration.md) for details.

### 4. Run

```bash
python main.py
```

Output PDFs will be in `output/` folder with `_x` suffix by default.

## Core Capabilities

### Multi-Sheet Configuration
Apply different print settings to different sheets based on sheet name matching:
- **Priority-based**: Lower priority number wins when sheet matches multiple configs
- **Sheet targeting**: Specify exact sheet names or use default config
- **Independent settings**: Each config has its own scaling, margins, page breaks

### Layout Control
- **Dimension preservation**: Keep original Excel cell dimensions (prepare_for_print: false)
- **Scaling modes**: no_scaling, fit_columns, fit_rows, fit_sheet, custom
- **Page breaks**: Automatic row/column-based pagination
- **Margins**: Normal, wide, narrow, or custom margins

### Language Classification
Automatically detect language and distribute files:
- **Auto mode**: Detects from cell content (vi, en, ja, zh, etc.)
- **Filename mode**: Classifies by filename patterns (_VN, _EN, _Ja)
- **Output organization**: Separate folders per language

## Documentation

- 📖 [Configuration Guide](docs/configuration.md) - Detailed config.yaml reference
- 🚀 [Quick Setup Guide](docs/quick-setup.md) - Step-by-step setup instructions
- 🏗️ [Architecture](docs/architecture.md) - Design and system overview

## Common Use Cases

**1. RAG/AI Document Processing**
```yaml
excel:
  prepare_for_print: false  # Preserve exact dimensions
print_options:
  scaling: "no_scaling"     # No modification
```

**2. Professional Reports**
```yaml
print_options:
  scaling: "fit_columns"
  margins: "normal"
  print_header_footer: true
```

**3. Large Data Sheets**
```yaml
print_options:
  rows_per_page: 100        # Page break every 100 rows
  columns_per_page: 15      # Page break every 15 columns
```

## Requirements

- Python 3.7+
- Windows OS (uses win32com for Excel automation)
- Microsoft Excel installed

## License

See LICENSE file for details.