# Quick Setup Guide

Step-by-step guide to get started with xlsx2pdf converter.

## Prerequisites

- Windows OS
- Microsoft Excel installed
- Python 3.7 or higher

## Installation Steps

### 1. Clone or Download

Download the project to your local machine.

### 2. Install Dependencies

```bash
cd xlsx2pdf
pip install -r requirements.txt
```

**Optional:** For language classification:
```bash
pip install langdetect
```

### 3. Verify Installation

Check that all required packages are installed:
```bash
python -c "import win32com.client; print('OK')"
```

You should see "OK" printed. If there's an error, reinstall:
```bash
pip install pywin32
```

---

## Basic Usage

### Step 1: Prepare Input Files

Create an `input` folder in the project directory (if it doesn't exist):

```
xlsx2pdf/
  ├── input/           ← Create this folder
  ├── output/          ← Will be created automatically
  ├── config.yaml
  └── main.py
```

Place your Excel files in the `input` folder:

```
input/
  ├── report.xlsx
  ├── data.xlsx
  └── subfolder/
      └── sheet.xlsx
```

**Supported Formats:**
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)
- `.xlsm` (Excel with macros)

### Step 2: Configure (Optional)

The tool works with default settings, but you can customize `config.yaml` for specific needs.

**Default behavior:**
- Converts all `.xlsx`, `.xls`, `.xlsm` files in `input/`
- Outputs to `output/` folder
- Adds `_x` suffix to filenames
- Fits columns to page width

**To customize:** Edit `config.yaml` (see [Configuration Guide](configuration.md) for all options)

### Step 3: Run Conversion

```bash
python main.py
```

**What happens:**
1. Scans `input/` folder for Excel files
2. Converts each file to PDF
3. Saves PDFs to `output/` folder
4. Shows progress in console
5. Logs details to `conversion.log`

### Step 4: Check Output

Find converted PDFs in the `output/` folder:

```
output/
  ├── report_x.pdf
  ├── data_x.pdf
  └── subfolder/
      └── sheet_x.pdf
```

---

## Common Configurations

### Configuration 1: Simple Batch Conversion

**Goal:** Convert all Excel files with default settings.

**Setup:**
1. Place files in `input/`
2. Run: `python main.py`

**Config (default):**
```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "auto"
  scaling: "fit_columns"
  margins: "normal"
```

---

### Configuration 2: Exact Dimension Preservation

**Goal:** Preserve exact Excel cell dimensions for RAG/AI processing.

**Setup:**

1. Edit `config.yaml`:
```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "native_print"
  scaling: "no_scaling"
  margins: "narrow"
  print_header_footer: false
```

2. Place files in `input/`
3. Run: `python main.py`

**Result:** PDFs match Excel dimensions exactly (no scaling).

---

### Configuration 3: Large Data Sheets with Page Breaks

**Goal:** Split large sheets into pages with consistent breaks.

**Setup:**

1. Edit `config.yaml`:
```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "auto"
  scaling: "fit_columns"
  rows_per_page: 100        # New page every 100 rows
  columns_per_page: 15      # New page every 15 columns
  margins: "narrow"
```

2. Place files in `input/`
3. Run: `python main.py`

**Result:** Large sheets paginated consistently.

---

### Configuration 4: Language Classification

**Goal:** Auto-detect language and distribute files to separate folders.

**Setup:**

1. Edit `config.yaml`:
```yaml
excel:
  prepare_for_print: false

print_options:
  mode: "auto"
  scaling: "fit_columns"

language_classification:
  enabled: true
  mode: "auto"              # or "filename" for pattern matching
  output_suffix_format: "output-{lang}"
  keep_folder_structure: true
```

2. Place files in `input/`
3. Run: `python main.py`

**Result:**
```
output-en/    ← English files
output-vi/    ← Vietnamese files
output-ja/    ← Japanese files
```

---

### Configuration 5: Multi-Sheet Workbooks

**Goal:** Apply different settings to different sheets.

**Setup:**

1. Edit `config.yaml`:
```yaml
excel:
  prepare_for_print: false

print_options:
  # Summary sheets - fit on one page
  - sheets: ['Summary', 'Dashboard']
    priority: 1
    mode: "one_page"
    scaling: "fit_sheet"
    margins: "narrow"
    
  # Data sheets - columnar layout
  - sheets: ['Data', 'Details']
    priority: 2
    mode: "auto"
    scaling: "fit_columns"
    rows_per_page: 50
    
  # Default for other sheets
  - sheets: null
    priority: 99
    mode: "auto"
    scaling: "fit_columns"
```

2. Place files in `input/`
3. Run: `python main.py`

**Result:** Each sheet type gets appropriate formatting.

---

## Folder Structure

### Recommended Setup

```
xlsx2pdf/
  ├── input/              # Your Excel files here
  │   ├── file1.xlsx
  │   ├── file2.xlsx
  │   └── subfolder/
  │       └── file3.xlsx
  ├── output/             # PDFs created here (auto-created)
  ├── enhanced_files/     # Temporary files (auto-created)
  ├── config.yaml         # Configuration file
  ├── main.py             # Run this file
  ├── requirements.txt
  └── src/
      ├── converter.py
      ├── logger.py
      └── ui.py
```

### Input Folder Organization

**Simple Structure:**
```
input/
  ├── file1.xlsx
  ├── file2.xlsx
  └── file3.xlsx
```

**Nested Structure:**
```
input/
  ├── 2024/
  │   ├── Q1/
  │   │   ├── jan.xlsx
  │   │   └── feb.xlsx
  │   └── Q2/
  │       └── apr.xlsx
  └── archive/
      └── old.xlsx
```

**With Language Patterns:**
```
input/
  ├── report_EN.xlsx      # English
  ├── data_VN.xlsx        # Vietnamese
  └── summary_Ja.xlsx     # Japanese
```

---

## Usage Tips

### Tip 1: Test with One File First

Before batch processing:
1. Place one test file in `input/`
2. Run conversion
3. Check output quality
4. Adjust `config.yaml` if needed
5. Process remaining files

### Tip 2: Monitor Conversion Log

Watch `conversion.log` for details:
```bash
# Windows PowerShell
Get-Content conversion.log -Tail 20 -Wait

# Command Prompt
type conversion.log
```

### Tip 3: Handle Large Batches

For many files:
1. Process in smaller batches
2. Increase timeout: `timeout_minutes: 60`
3. Check `errors.log` for failures

### Tip 4: Backup Original Files

Always keep original Excel files:
- Conversion is non-destructive
- Enhanced files stored separately
- Can reprocess with different settings

### Tip 5: Use Descriptive Suffixes

Change output suffix for different runs:
```yaml
output_suffix: "_draft"   # report.xlsx → report_draft.pdf
output_suffix: "_final"   # report.xlsx → report_final.pdf
output_suffix: "_v2"      # report.xlsx → report_v2.pdf
```

---

## Troubleshooting

### Issue: "No Excel files found"

**Cause:** No `.xlsx`, `.xls`, or `.xlsm` files in `input/` folder.

**Solution:**
1. Check folder exists: `xlsx2pdf/input/`
2. Verify files have correct extensions
3. Check files aren't in a subfolder (unless intended)

---

### Issue: "Excel is not installed"

**Cause:** Microsoft Excel not found on system.

**Solution:**
- Install Microsoft Excel (required for COM automation)
- This tool requires actual Excel, not alternatives like LibreOffice

---

### Issue: Conversion very slow

**Cause:** Large files or many sheets.

**Solution:**
1. Increase timeout:
```yaml
timeout_minutes: 90
```

2. Process fewer files at once
3. Disable header/footer if not needed:
```yaml
print_options:
  print_header_footer: false
```

---

### Issue: PDF layout doesn't look right

**Cause:** Scaling or margin settings.

**Solution:** Try different scaling modes:
```yaml
print_options:
  scaling: "no_scaling"     # Exact dimensions
  # or
  scaling: "fit_sheet"      # Fit on one page
  # or
  scaling: "fit_columns"    # Fit width only
```

---

### Issue: Text cut off in PDF

**Cause:** Page breaks needed.

**Solution:** Add automatic page breaks:
```yaml
print_options:
  rows_per_page: 50
  columns_per_page: 10
```

---

### Issue: Language detection wrong

**Cause:** Not enough text or mixed languages.

**Solution:** Use filename mode instead:
```yaml
language_classification:
  mode: "filename"
  filename_patterns:
    en: ["_EN"]
    vi: ["_VN"]
```

---

## Next Steps

- **Customize Settings:** Read [Configuration Guide](configuration.md)
- **Learn Architecture:** See [Architecture Documentation](architecture.md)
- **Advanced Features:** Explore multi-sheet configurations
- **Automation:** Schedule batch processing with Task Scheduler

---

## Getting Help

**Check Logs:**
- `conversion.log` - Detailed conversion info
- `errors.log` - Error messages only

**Common Issues:**
- Make sure Excel is installed
- Verify Python 3.7+
- Check file permissions
- Ensure input folder exists

**Configuration Help:**
- See [Configuration Guide](configuration.md)
- Check example templates in config file
- Test with simple config first
