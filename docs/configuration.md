# Configuration Guide

Complete reference for `config.yaml` settings.

## Table of Contents
- [General Settings](#general-settings)
- [Excel Settings](#excel-settings)
- [Print Options](#print-options)
- [Language Classification](#language-classification)
- [Multiple Print Configurations](#multiple-print-configurations)

---

## General Settings

### `timeout_minutes`
Maximum time in minutes for Excel conversion before timeout.
```yaml
timeout_minutes: 45  # Default: 45 minutes
```

### `output_suffix`
Suffix appended to output PDF filenames.
```yaml
output_suffix: "_x"  # input.xlsx → input_x.pdf
```

---

## Excel Settings

### `prepare_for_print`
Controls whether to modify Excel content before PDF export.

```yaml
excel:
  prepare_for_print: false  # Recommended for exact dimension preservation
```

**Options:**
- `true`: Auto-adjusts row heights, column widths, text wrapping
- `false`: Preserves original Excel dimensions (no modifications)

**Note:** `print_options` (scaling, margins, headers) are ALWAYS applied regardless of this setting.

### `enhanced_dir`
Directory for intermediate enhanced Excel files.
```yaml
excel:
  enhanced_dir: "enhanced_files"
```

---

## Print Options

Print options control PDF output formatting and layout. Supports both **single configuration** (applies to all sheets) and **multiple configurations** (sheet-specific settings).

### Basic Structure (Single Config)

```yaml
print_options:
  sheets: null              # null = applies to all sheets
  priority: 1               # Only matters with multiple configs
  mode: "auto"
  page_size: "auto"
  scaling: "fit_columns"
  margins: "normal"
  print_header_footer: true
  print_row_col_headings: false
```

### Print Modes

#### `mode`
Determines page layout strategy.

**Options:**
- `auto` (default): Auto orientation and page size based on content
- `one_page`: Fit entire sheet on one page
- `table_row_break`: Insert page breaks at table boundaries
- `auto_page_size`: Auto-select paper size to fit content
- `native_print`: Preserve exact Excel dimensions (best for RAG)
- `uniform_page_size`: Apply same page size to all sheets

```yaml
print_options:
  mode: "auto"
```

### Page Size

#### `page_size`
Paper size for PDF output.

**Options:** `auto`, `letter`, `tabloid`, `legal`, `statement`, `executive`, `A1`, `A2`, `A3`, `A4`, `A5`, `B4`, `B5`

```yaml
print_options:
  page_size: "A4"  # or "auto" for automatic selection
```

### Page Breaks

#### `rows_per_page`
Add horizontal page break every N rows.

```yaml
print_options:
  rows_per_page: 100  # New page every 100 rows
  # rows_per_page: null  # No automatic row breaks (default)
```

**Use Cases:**
- Large data sheets with consistent pagination
- Reports requiring fixed rows per page
- Splitting long sheets into manageable pages

#### `columns_per_page`
Add vertical page break every N columns.

```yaml
print_options:
  columns_per_page: 15  # New page every 15 columns
  # columns_per_page: null  # No automatic column breaks (default)
```

**Use Cases:**
- Wide spreadsheets with many columns
- Grid-style page breaking
- Horizontal pagination for large tables

**Example - Both Row and Column Breaks:**
```yaml
print_options:
  rows_per_page: 50
  columns_per_page: 10
  # Creates grid pagination: 50 rows × 10 columns per page
```

### Scaling Options

#### `scaling`
Controls how content fits on pages.

**Options:**
- `no_scaling`: 100% zoom, no fitting (preserves exact dimensions)
- `fit_sheet`: Fit entire sheet on one page
- `fit_columns`: Fit all columns on page width (rows span multiple pages)
- `fit_rows`: Fit all rows on page height (columns span multiple pages)
- `custom`: Custom zoom percentage (use with `scaling_percent`)

```yaml
print_options:
  scaling: "fit_columns"  # Most common
  scaling_percent: 100    # Only used when scaling: "custom"
```

**When to Use:**
- `no_scaling`: RAG/AI processing, exact dimension preservation
- `fit_columns`: Standard reports, readable width
- `fit_sheet`: Summaries, dashboards
- `custom`: Specific zoom requirements

### Margins

#### `margins`
Page margin preset.

**Options:** `normal`, `wide`, `narrow`, `custom`

```yaml
print_options:
  margins: "normal"
```

#### `custom_margins`
Custom margin values in centimeters (used when `margins: custom`).

```yaml
print_options:
  margins: "custom"
  custom_margins:
    top: 2.0
    bottom: 2.0
    left: 1.5
    right: 1.5
    header: 0.8
    footer: 0.8
```

### Headers and Footers

#### `print_header_footer`
Enable/disable automatic headers and footers.

```yaml
print_options:
  print_header_footer: true
```

**Default Output:**
- Header: Sheet name
- Footer: Row range and page number

#### `print_row_col_headings`
Print Excel row numbers (1,2,3...) and column letters (A,B,C...).

```yaml
print_options:
  print_row_col_headings: false
```

---

## Multiple Print Configurations

Apply different settings to different sheets based on sheet name matching.

### Structure

```yaml
print_options:
  - sheets: ['Summary', 'Dashboard']  # Sheet name list
    priority: 1                        # Lower = higher priority
    mode: "one_page"
    scaling: "fit_sheet"
    margins: "narrow"
    print_header_footer: false
    
  - sheets: ['DataSheet1', 'DataSheet2']
    priority: 2
    mode: "auto"
    scaling: "fit_columns"
    margins: "normal"
    rows_per_page: 100
    
  - sheets: null                       # Default for unmatched sheets
    priority: 99                       # Lowest priority
    mode: "auto"
    scaling: "fit_columns"
    margins: "normal"
```

### Priority System

When a sheet name matches multiple configurations, the one with the **lowest priority number wins**.

**Example:**
```yaml
print_options:
  - sheets: ['Report']
    priority: 1         # This wins for 'Report' sheet
    scaling: "fit_sheet"
    
  - sheets: ['Report', 'Data']
    priority: 2         # Lower priority, ignored for 'Report'
    scaling: "fit_columns"
```

### Sheet Matching

#### Exact Match
```yaml
- sheets: ['Sheet1', 'Sheet2']  # Matches only these exact names
```

#### Default Config
```yaml
- sheets: null                   # Matches all unmatched sheets
  priority: 99                   # Typically lowest priority
```

### Complete Example

```yaml
print_options:
  # Summary sheets - fit on one page
  - sheets: ['Summary', 'Overview', 'Dashboard']
    priority: 1
    mode: "one_page"
    scaling: "fit_sheet"
    margins: "narrow"
    print_header_footer: false
    
  # Data sheets - columnar layout with page breaks
  - sheets: ['Data', 'Details', 'Records']
    priority: 2
    mode: "auto"
    scaling: "fit_columns"
    margins: "normal"
    rows_per_page: 100
    columns_per_page: 15
    print_header_footer: true
    
  # RAG processing sheets - exact dimensions
  - sheets: ['RAG_Input', 'AI_Data']
    priority: 3
    mode: "native_print"
    scaling: "no_scaling"
    margins: "narrow"
    print_header_footer: false
    
  # Default config for all other sheets
  - sheets: null
    priority: 99
    mode: "auto"
    scaling: "fit_columns"
    margins: "normal"
    print_header_footer: true
```

---

## Language Classification

Automatically detect language and distribute PDF files to language-specific folders.

### Basic Configuration

```yaml
language_classification:
  enabled: true               # Enable/disable feature
  mode: "auto"                # "auto" or "filename"
  output_suffix_format: "output-{lang}"
  keep_folder_structure: true
```

### Detection Modes

#### Auto Mode
Detects language from cell content using NLP.

```yaml
language_classification:
  mode: "auto"
```

**Supported Languages:**
- `vi` - Vietnamese
- `en` - English
- `ja` - Japanese
- `zh` - Chinese
- `ko` - Korean
- `th` - Thai
- `fr` - French
- `de` - German
- `es` - Spanish

#### Filename Mode
Classifies based on filename patterns.

```yaml
language_classification:
  mode: "filename"
  filename_patterns:
    vi: ["_VN"]              # Matches *_VN.xlsx
    en: ["_EN", "_t"]        # Matches *_EN.xlsx or *_t.xlsx
    ja: ["_Ja", ""]          # Matches *_Ja.xlsx or files without pattern
```

### Output Structure

#### `output_suffix_format`
Template for language-specific output folders.

```yaml
output_suffix_format: "output-{lang}"  # {lang} replaced with language code
```

**Examples:**
- `output-vi/` for Vietnamese files
- `output-en/` for English files
- `output-ja/` for Japanese files

#### `keep_folder_structure`
Maintain input folder structure in output.

```yaml
keep_folder_structure: true
```

**Example:**
```
Input:
  input/
    ├── reports/
    │   ├── data_VN.xlsx
    │   └── summary_EN.xlsx
    └── archive/
        └── old_Ja.xlsx

Output (with keep_folder_structure: true):
  output-vi/
    └── reports/
        └── data_VN_x.pdf
  output-en/
    └── reports/
        └── summary_EN_x.pdf
  output-ja/
    └── archive/
        └── old_Ja_x.pdf
```

---

## Logging Settings

```yaml
logging:
  log_file: "conversion.log"     # Main log file
  error_file: "errors.log"       # Error-only log file
  log_level: "INFO"              # DEBUG, INFO, WARNING, ERROR
  log_console_lines: 20          # Number of lines in UI console
```

**Log Levels:**
- `DEBUG`: Detailed information for debugging
- `INFO`: General information about conversion progress
- `WARNING`: Warning messages (non-critical issues)
- `ERROR`: Error messages only

---

## Configuration Templates

### Template 1: RAG/AI Processing
Preserve exact Excel dimensions for data processing.

```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "native_print"
  scaling: "no_scaling"
  margins: "narrow"
  print_header_footer: false
  print_row_col_headings: false
```

### Template 2: Professional Reports
Readable output with proper formatting.

```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "auto"
  page_size: "A4"
  scaling: "fit_columns"
  margins: "normal"
  print_header_footer: true
  print_row_col_headings: false
```

### Template 3: Large Data Sheets
Consistent pagination for large datasets.

```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "auto"
  page_size: "A4"
  scaling: "fit_columns"
  margins: "narrow"
  rows_per_page: 100
  columns_per_page: 15
  print_header_footer: true
```

### Template 4: Multi-Sheet Workbooks
Different settings per sheet type.

```yaml
excel:
  prepare_for_print: false

print_options:
  - sheets: ['Summary']
    priority: 1
    mode: "one_page"
    scaling: "fit_sheet"
    margins: "narrow"
    
  - sheets: ['Data1', 'Data2']
    priority: 2
    scaling: "fit_columns"
    rows_per_page: 50
    
  - sheets: null
    priority: 99
    mode: "auto"
    scaling: "fit_columns"
```

---

## Troubleshooting

### Issue: PDF dimensions don't match Excel
**Solution:**
```yaml
excel:
  prepare_for_print: false
print_options:
  scaling: "no_scaling"
```

### Issue: Content cut off on pages
**Solution:**
```yaml
print_options:
  rows_per_page: 50      # Add page breaks
  columns_per_page: 10   # Split columns
```

### Issue: Text too small in PDF
**Solution:**
```yaml
print_options:
  scaling: "custom"
  scaling_percent: 150   # Increase zoom
```

### Issue: Multiple sheets need different settings
**Solution:** Use multiple print_options with sheet matching (see [Multiple Print Configurations](#multiple-print-configurations))

---

## See Also

- [Quick Setup Guide](quick-setup.md) - Step-by-step setup instructions
- [Architecture](architecture.md) - System design and components
