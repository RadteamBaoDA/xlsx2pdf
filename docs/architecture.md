# Architecture Documentation

System design and technical overview of xlsx2pdf converter.

## System Overview

xlsx2pdf is a Python-based Excel to PDF converter that uses Windows COM automation to leverage Microsoft Excel's native rendering engine for high-fidelity PDF generation.

### Design Philosophy

1. **Preservation First**: Maintain original Excel formatting and dimensions
2. **Flexibility**: Support multiple configuration strategies per workbook
3. **Automation**: Batch processing with minimal user intervention
4. **Extensibility**: Modular design for easy feature additions

---

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                        User Interface                        │
│                         (main.py)                            │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                    Configuration Layer                       │
│                     (config.yaml)                            │
│  • Print Options (single/multiple)                           │
│  • Language Classification                                   │
│  • Logging Settings                                          │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                     Core Converter                           │
│                  (src/converter.py)                          │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │         Sheet Configuration Resolver                  │  │
│  │  • Priority-based matching                            │  │
│  │  • Sheet name pattern matching                        │  │
│  │  • Default fallback                                   │  │
│  └──────────────────────────────────────────────────────┘  │
│                        │                                     │
│                        ▼                                     │
│  ┌──────────────────────────────────────────────────────┐  │
│  │         Layout Optimization Pipeline                  │  │
│  │  1. Expand hidden content                             │  │
│  │  2. Fix shape placement                               │  │
│  │  3. Apply print mode                                  │  │
│  │  4. Preserve dimensions                               │  │
│  │  5. Apply scaling                                     │  │
│  │  6. Apply margins                                     │  │
│  │  6.5 Apply page breaks                                │  │
│  │  7. Setup headers/footers                             │  │
│  │  8. Configure headings                                │  │
│  │  9. Enhanced preparation (optional)                   │  │
│  └──────────────────────────────────────────────────────┘  │
│                        │                                     │
│                        ▼                                     │
│  ┌──────────────────────────────────────────────────────┐  │
│  │           Excel COM Automation                        │  │
│  │  • Open workbook                                      │  │
│  │  • Configure page setup                               │  │
│  │  • Export to PDF                                      │  │
│  └──────────────────────────────────────────────────────┘  │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                  Language Classifier                         │
│                                                              │
│  ┌─────────────────┐        ┌─────────────────┐            │
│  │   Auto Mode     │        │  Filename Mode  │            │
│  │                 │        │                 │            │
│  │ • Content scan  │        │ • Pattern match │            │
│  │ • NLP detection │        │ • Rules-based   │            │
│  └─────────────────┘        └─────────────────┘            │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                    Output Management                         │
│                                                              │
│  • Language-based distribution                               │
│  • Folder structure preservation                            │
│  • Filename suffixing                                        │
└─────────────────────────────────────────────────────────────┘
```

---

## Core Components

### 1. Main Controller (`main.py`)

**Responsibilities:**
- User interface (console-based)
- File discovery and scanning
- Batch processing orchestration
- Progress tracking
- Error handling and logging

**Key Functions:**
- `find_excel_files()`: Recursively scan input directory
- `process_file()`: Handle single file conversion
- `classify_and_move()`: Language-based distribution

### 2. Excel Converter (`src/converter.py`)

**Class:** `ExcelConverter`

**Core Methods:**

#### `convert(input_path, output_path)`
Main entry point for conversion. Orchestrates the entire conversion pipeline.

#### `_get_sheet_print_options(sheet_name)`
Sheet configuration resolver with priority-based matching.

**Algorithm:**
```python
1. Load print_options from config
2. If single dict → return it (backward compatible)
3. If list:
   a. Iterate through configs
   b. Match sheet name against 'sheets' key
   c. Collect all matches with priorities
   d. Sort by priority (lower = higher)
   e. Return highest priority match
4. Fallback to default config
```

**Priority Logic:**
- Priority 1 = highest (applied first)
- When sheet matches multiple configs, lowest number wins
- `sheets: null` = default fallback (typically priority 99)

#### `_optimize_layout(workbook, print_mode)`
Multi-step pipeline for preparing workbook for PDF export.

**Pipeline Steps:**

1. **Expand Hidden Content**
   - Unhide all rows and columns
   - Expand grouped sections
   - Ensure all data visible

2. **Fix Shape Placement**
   - Set shapes to move but not resize with cells
   - Ensure shapes are visible and printable

3. **Apply Print Mode**
   - Configure page layout based on mode
   - Set orientation, paper size
   - Apply mode-specific settings

4. **Preserve Dimensions**
   - Skip dimension-modifying functions when `prepare_for_print: false`
   - Maintain original row heights and column widths

5. **Apply Scaling**
   - Configure zoom and page fitting
   - Options: no_scaling, fit_columns, fit_rows, fit_sheet, custom

6. **Apply Margins**
   - Set page margins (normal, wide, narrow, custom)
   - Configure header/footer margins

6.5. **Apply Page Breaks**
   - Insert horizontal breaks (rows_per_page)
   - Insert vertical breaks (columns_per_page)

7. **Setup Headers/Footers**
   - Add sheet name to header
   - Add row range and page numbers to footer

8. **Configure Headings**
   - Enable/disable row numbers and column letters

9. **Enhanced Preparation** (optional)
   - Additional layout fixes if `prepare_for_print: true`
   - Currently disabled to preserve dimensions

#### `_apply_scaling(sheet, workbook_name, scaling, scaling_percent)`
Configure how content fits on pages.

**Scaling Modes:**

| Mode | Zoom | FitToPagesWide | FitToPagesTall | Effect |
|------|------|----------------|----------------|--------|
| `no_scaling` | 100 | False | False | Exact size, no fitting |
| `fit_sheet` | False | 1 | 1 | Fit entire sheet on one page |
| `fit_columns` | False | 1 | False | Fit width, rows span pages |
| `fit_rows` | False | False | 1 | Fit height, columns span pages |
| `custom` | % | False | False | Custom zoom percentage |

#### `_insert_page_breaks_by_rows(sheet, workbook_name, rows_per_page)`
Add horizontal page breaks at regular row intervals.

**Algorithm:**
```python
1. Get used range row count and start row
2. For each interval (start + N, start + 2N, ...):
   a. Insert HPageBreak at row position
3. Log total breaks inserted
```

#### `_insert_page_breaks_by_columns(sheet, workbook_name, columns_per_page)`
Add vertical page breaks at regular column intervals.

**Algorithm:**
```python
1. Get used range column count and start column
2. For each interval (start + N, start + 2N, ...):
   a. Insert VPageBreak at column position
3. Log total breaks inserted
```

#### `_export_to_pdf(workbook, output_path)`
Final PDF export using Excel COM API.

**Parameters:**
- `Type`: xlTypePDF (0)
- `Quality`: xlQualityStandard (0)
- `IncludeDocProperties`: True (for metadata)
- `IgnorePrintAreas`: False (respect print areas)
- `OpenAfterPublish`: False (don't open PDF)

### 3. Logger (`src/logger.py`)

**Responsibilities:**
- Centralized logging configuration
- Dual output (file and console)
- Log level management
- Structured log formatting

**Log Files:**
- `conversion.log`: All conversion activity
- `errors.log`: Errors only (easier troubleshooting)

### 4. UI Module (`src/ui.py`)

**Components:**
- Console output formatting
- Progress indicators
- User prompts
- Error display

---

## Data Flow

### Single File Conversion Flow

```
1. User runs main.py
   │
   ▼
2. Scan input/ for Excel files
   │
   ▼
3. For each file:
   │
   ├─▶ Load config.yaml
   │
   ├─▶ Create ExcelConverter instance
   │
   ├─▶ Open Excel application (COM)
   │
   ├─▶ Open workbook
   │
   ├─▶ For each sheet:
   │   │
   │   ├─▶ Get sheet-specific print_options
   │   │   (priority-based matching)
   │   │
   │   ├─▶ Run optimization pipeline
   │   │   (9 steps)
   │   │
   │   └─▶ Configure PageSetup
   │
   ├─▶ Export to PDF
   │
   ├─▶ Close workbook (no save)
   │
   ├─▶ Quit Excel
   │
   ├─▶ [Optional] Classify by language
   │
   └─▶ Move to output folder
```

### Multi-Config Sheet Matching Flow

```
Input: sheet_name = "DataSheet"
Config: [Config1, Config2, Config3]

1. Initialize matches = []
   │
   ▼
2. For each config in config_list:
   │
   ├─▶ sheets = config.get('sheets')
   │   priority = config.get('priority')
   │
   ├─▶ If sheets is None or []:
   │   └─▶ matches.append((priority, config))  # Default
   │
   ├─▶ If "DataSheet" in sheets:
   │   └─▶ matches.append((priority, config))  # Match
   │
   └─▶ Continue
   │
   ▼
3. Sort matches by priority (ascending)
   │
   ▼
4. Return matches[0][1]  # Highest priority config
```

---

## Key Design Patterns

### 1. Strategy Pattern
Different print modes are strategies applied to sheets:
- `_apply_auto_mode()`
- `_apply_one_page_mode()`
- `_apply_table_row_break_mode()`
- etc.

### 2. Template Method Pattern
`_optimize_layout()` defines the algorithm structure; subcomponents fill in specifics.

### 3. Configuration Hierarchy
Priority-based configuration with fallback:
1. Sheet-specific config (lowest priority number)
2. Default config (highest priority number)
3. Hardcoded defaults (fallback)

### 4. Separation of Concerns
- `converter.py`: Core conversion logic
- `logger.py`: Logging infrastructure
- `ui.py`: User interface
- `main.py`: Application orchestration

---

## Configuration System

### Structure

```yaml
┌─────────────────────────────────────┐
│         config.yaml                 │
│                                     │
│  ┌───────────────────────────────┐ │
│  │   General Settings            │ │
│  │   • timeout_minutes           │ │
│  │   • output_suffix             │ │
│  └───────────────────────────────┘ │
│                                     │
│  ┌───────────────────────────────┐ │
│  │   Excel Settings              │ │
│  │   • prepare_for_print         │ │
│  │   • enhanced_dir              │ │
│  └───────────────────────────────┘ │
│                                     │
│  ┌───────────────────────────────┐ │
│  │   Print Options               │ │
│  │   (single or list)            │ │
│  │                               │ │
│  │   ┌─────────────────────────┐ │ │
│  │   │ Config 1                │ │ │
│  │   │ • sheets: [...]         │ │ │
│  │   │ • priority: N           │ │ │
│  │   │ • mode, scaling, etc.   │ │ │
│  │   └─────────────────────────┘ │ │
│  │                               │ │
│  │   ┌─────────────────────────┐ │ │
│  │   │ Config 2                │ │ │
│  │   └─────────────────────────┘ │ │
│  │                               │ │
│  │   ┌─────────────────────────┐ │ │
│  │   │ Default Config          │ │ │
│  │   │ • sheets: null          │ │ │
│  │   │ • priority: 99          │ │ │
│  │   └─────────────────────────┘ │ │
│  └───────────────────────────────┘ │
│                                     │
│  ┌───────────────────────────────┐ │
│  │   Language Classification     │ │
│  │   • enabled                   │ │
│  │   • mode                      │ │
│  │   • patterns                  │ │
│  └───────────────────────────────┘ │
│                                     │
│  ┌───────────────────────────────┐ │
│  │   Logging Settings            │ │
│  │   • log_file                  │ │
│  │   • log_level                 │ │
│  └───────────────────────────────┘ │
└─────────────────────────────────────┘
```

### Loading Process

1. Parse YAML file
2. Validate structure
3. Apply defaults for missing values
4. Store in memory for conversion process

---

## Excel COM Integration

### COM Objects Used

```python
excel = win32com.client.DispatchEx("Excel.Application")
workbook = excel.Workbooks.Open(path)
sheet = workbook.Sheets(index)
pageSetup = sheet.PageSetup
```

### Key Properties Modified

**PageSetup:**
- `Zoom`: Scaling percentage
- `FitToPagesWide`: Horizontal page fitting
- `FitToPagesTall`: Vertical page fitting
- `PaperSize`: Paper size constant
- `Orientation`: Portrait/Landscape
- `TopMargin`, `BottomMargin`, etc.: Margins
- `PrintHeadings`: Row/column headings
- `CenterHeader`, `CenterFooter`: Headers/footers

**Sheet:**
- `UsedRange`: Active data range
- `Shapes`: Embedded objects
- `ListObjects`: Excel tables
- `HPageBreaks`: Horizontal page breaks
- `VPageBreaks`: Vertical page breaks

### Export Method

```python
workbook.ExportAsFixedFormat(
    Type=0,                      # PDF
    Filename=output_path,
    Quality=0,                   # Standard
    IncludeDocProperties=True,
    IgnorePrintAreas=False,
    OpenAfterPublish=False
)
```

---

## Language Classification

### Auto Mode Algorithm

```
1. Open Excel file
2. For each sheet:
   a. Read all cells with text
   b. Concatenate text samples
3. Detect language using langdetect
4. Assign language code (en, vi, ja, etc.)
5. Return detected language
```

### Filename Mode Algorithm

```
1. Get filename
2. For each language pattern:
   a. Check if pattern in filename
   b. If match, return language code
3. If no match, return default (empty string)
```

---

## Error Handling

### Strategy

1. **Fail Gracefully**: Continue processing remaining files on error
2. **Detailed Logging**: Log all errors to `errors.log`
3. **User Notification**: Display errors in console
4. **Cleanup**: Always close Excel COM objects

### Error Types

**File Errors:**
- File not found
- File locked
- Permissions issue

**Excel Errors:**
- Corrupted workbook
- COM initialization failure
- Timeout during conversion

**Configuration Errors:**
- Invalid YAML syntax
- Missing required keys
- Invalid value types

---

## Performance Considerations

### Optimization Strategies

1. **COM Object Reuse**: One Excel instance per batch
2. **Parallel Processing**: Could be added (currently sequential)
3. **Selective Processing**: Skip already converted files
4. **Timeout Management**: Configurable timeout per file

### Resource Management

- **Memory**: Excel COM objects released after each file
- **Disk**: Temporary files cleaned up
- **CPU**: Excel does heavy lifting (native engine)

### Scalability

**Current Limits:**
- Sequential processing (one file at a time)
- Windows-only (COM dependency)
- Excel installation required

**Future Enhancements:**
- Multi-threading support
- Cloud processing option
- Alternative rendering engines

---

## Extension Points

### Adding New Print Modes

1. Add constant: `PRINT_MODE_CUSTOM = "custom"`
2. Create method: `_apply_custom_mode(sheet, workbook_name)`
3. Add to mode switch in `_optimize_layout()`

### Adding New Scaling Options

1. Add case to `_apply_scaling()` method
2. Set appropriate PageSetup properties
3. Document in config reference

### Adding New Languages

**Auto Mode:**
- Supported automatically by langdetect library

**Filename Mode:**
- Add pattern to `filename_patterns` in config

---

## Testing Strategy

### Manual Testing Checklist

- [ ] Single file conversion
- [ ] Batch processing
- [ ] Multi-sheet workbooks
- [ ] Different scaling modes
- [ ] Page break functionality
- [ ] Language classification
- [ ] Error handling
- [ ] Large files (timeout)
- [ ] Nested folders
- [ ] Filename patterns

### Test Files Needed

- Simple single-sheet workbook
- Multi-sheet workbook with varied content
- Large data sheet (1000+ rows)
- Wide sheet (50+ columns)
- Files with embedded images
- Files with different languages
- Files with special characters

---

## Future Enhancements

### Planned Features

1. **Parallel Processing**: Multi-threaded conversion
2. **Web Interface**: Browser-based UI
3. **API Mode**: REST API for integration
4. **Cloud Storage**: Direct S3/Azure upload
5. **Template System**: Reusable configuration templates
6. **Watermarking**: Add watermarks to PDFs
7. **Encryption**: Password-protect PDFs
8. **Compression**: PDF size optimization

### Architecture Changes

1. **Plugin System**: Extensible conversion pipeline
2. **Event System**: Hooks for custom processing
3. **Cache Layer**: Speed up repeated conversions
4. **Database**: Track conversion history

---

## Troubleshooting Guide

### Common Issues

**Issue: COM Automation Fails**
- Check Excel installation
- Run as administrator
- Clear COM cache: `gen_py` folder

**Issue: Slow Conversion**
- Increase timeout
- Reduce file size
- Disable prepare_for_print

**Issue: Layout Issues**
- Adjust scaling mode
- Check page size setting
- Add page breaks

**Issue: Memory Usage High**
- Process in smaller batches
- Restart between batches
- Close other Excel instances

---

## Development Guidelines

### Code Style

- Follow PEP 8
- Use type hints where possible
- Document public methods
- Keep functions focused (single responsibility)

### Logging Practices

- Use appropriate log levels
- Include context (workbook/sheet name)
- Log both success and failure
- Avoid sensitive data in logs

### Configuration Changes

- Update config.yaml
- Update configuration.md
- Update quick-setup.md
- Add migration notes if breaking changes

### Version Control

- Meaningful commit messages
- Branch for features
- Tag releases
- Update CHANGELOG

---

## References

### External Dependencies

- **pywin32**: Windows COM automation
- **PyYAML**: YAML parsing
- **langdetect**: Language detection (optional)

### Microsoft Excel API

- [Excel VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [PageSetup Object](https://docs.microsoft.com/en-us/office/vba/api/excel.pagesetup)

### Related Documentation

- [Configuration Guide](configuration.md)
- [Quick Setup Guide](quick-setup.md)
- [README](../README.md)
