# Row Index Header System for RAG Metadata

## Overview

The system now automatically calculates page break structure and writes the row index information directly into:
1. **PDF Header** - Shows page structure summary (e.g., "P1:R2-31, P2:R32-61, P3:R62-91")
2. **Added Column** - New "Page-Row Index" column in the Excel sheet with labels at each page
3. **Conversion Logs** - Complete RAG metadata mapping for all pages

This replaces the old `print_row_col_headings` system with calculated, accurate row indices.

## How It Works

### Automatic Page Break Calculation

When `rows_per_page` is configured, the system:

1. **Calculates page breaks** based on actual row heights and page dimensions
2. **Records page structure** with exact row index ranges for each page
3. **Writes to header** - Adds page structure to PDF center header
4. **Adds column** - Inserts "Page-Row Index" column with labels
5. **Logs metadata** - Outputs complete RAG mapping to conversion logs

### Output Format

**PDF Header (Center):**
```
P1:R2-31, P2:R32-61, P3:R62-91, ...+5pages
```

**Added Column in Excel:**
```
| Page-Row Index |
|----------------|
| P1: R2-31      |  <- At row 2 (page 1 start)
| P2: R32-61     |  <- At row 32 (page 2 start)
| P3: R62-91     |  <- At row 62 (page 3 start)
```

**Conversion Logs:**
```
[File.xlsx] Sheet1: RAG METADATA - Page-to-Row Index Mapping:
  Page 1: Rows 2-31 (30 rows)
  Page 2: Rows 32-61 (30 rows)
  Page 3: Rows 62-91 (30 rows)
  Page 4: Rows 92-121 (30 rows)
  Page 5: Rows 122-150 (29 rows)
```

## Configuration

### Enable Row Index Headers

```yaml
print_options:
  rows_per_page: 30                  # Trigger automatic page breaks
  print_header_footer: true          # Enable header with row index info
  print_row_col_headings: false      # Deprecated - no longer used
```

### Example Configurations

**Dense Data (50 rows/page):**
```yaml
print_options:
  mode: "auto"
  rows_per_page: 50
  page_size: "TABLOID"
  orientation: "landscape"
  print_header_footer: true
```

Output: Pages with rows 2-51, 52-101, 102-151, etc.

**Sparse Data (20 rows/page):**
```yaml
print_options:
  mode: "auto"
  rows_per_page: 20
  page_size: "LETTER"
  orientation: "portrait"
  print_header_footer: true
```

Output: Pages with rows 2-21, 22-41, 42-61, etc.

## RAG System Integration

### Extracting Row Indices

**From PDF Header:**
- Parse center header text: "P1:R2-31, P2:R32-61, P3:R62-91"
- Extract page number (P1, P2, P3) and row range (R2-31, R32-61, R62-91)

**From Added Column:**
- Look for "Page-Row Index" column in extracted table
- Find labels like "P1: R2-31" at the start of each page section

**From Conversion Logs:**
- Parse log lines: "Page 1: Rows 2-31 (30 rows)"
- Build complete page-to-row mapping dictionary

### Metadata Structure

```python
# Build from header text
page_structure = {
    1: {"start_row": 2, "end_row": 31, "row_count": 30},
    2: {"start_row": 32, "end_row": 61, "row_count": 30},
    3: {"start_row": 62, "end_row": 91, "row_count": 30},
}

# For each extracted content chunk
metadata = {
    "source_file": "report.xlsx",
    "sheet_name": "Sheet1",
    "pdf_page": 2,
    "excel_start_row": 32,
    "excel_end_row": 61,
    "row_count": 30,
    "content": "extracted text..."
}
```

### Validation Example

```python
import re

def parse_header_structure(header_text):
    """Parse row indices from PDF header text"""
    # Example: "P1:R2-31, P2:R32-61, P3:R62-91"
    pattern = r'P(\d+):R(\d+)-(\d+)'
    matches = re.findall(pattern, header_text)
    
    page_map = {}
    for page_num, start_row, end_row in matches:
        page_map[int(page_num)] = {
            'start_row': int(start_row),
            'end_row': int(end_row)
        }
    
    return page_map

# Example usage
header = "P1:R2-31, P2:R32-61, P3:R62-91"
structure = parse_header_structure(header)
print(structure)
# {1: {'start_row': 2, 'end_row': 31}, 
#  2: {'start_row': 32, 'end_row': 61},
#  3: {'start_row': 62, 'end_row': 91}}
```

## Advantages Over Old System

### Old System (print_row_col_headings)
- ❌ Always showed 1, 2, 3, 4... regardless of page breaks
- ❌ Couldn't show which actual rows were on each page
- ❌ No page-to-row mapping available
- ❌ Required manual calculation to determine original row

### New System (Calculated Row Index)
- ✅ Shows exact Excel row indices (2-31, 32-61, etc.)
- ✅ Visible in PDF header on every page
- ✅ Added column with page labels for easy reference
- ✅ Complete RAG metadata in conversion logs
- ✅ No calculation needed - direct mapping

## Testing

### Test Script

```bash
python test_row_index_header.py
```

This creates a test Excel file and converts it to PDF, demonstrating:
- Header with page structure
- Added "Page-Row Index" column
- Log output with RAG metadata

### Manual Verification

1. **Create test file** with 50+ rows
2. **Set config**: `rows_per_page: 10`
3. **Convert to PDF**
4. **Check PDF**:
   - Open PDF and look at header (top center)
   - Should show: "P1:R2-11, P2:R12-21, P3:R22-31, ..."
   - Check rightmost column for "Page-Row Index" labels
5. **Check logs**:
   - Search for "RAG METADATA - Page-to-Row Index Mapping"
   - Verify row ranges match PDF pages

## Troubleshooting

### Row Index Column Not Appearing

**Problem**: No "Page-Row Index" column in PDF

**Solutions**:
1. Ensure `rows_per_page` is set in config
2. Verify `print_header_footer: true`
3. Check logs for "Added row index column" message
4. Ensure sheet has enough rows to trigger page breaks

### Header Shows Wrong Information

**Problem**: Header text doesn't match actual pages

**Possible Causes**:
1. Page break calculation changed due to content
2. Hidden rows affecting calculation
3. Merged cells causing height issues

**Solution**: Check conversion logs for actual calculated page structure

### Missing in Conversion Logs

**Problem**: No "RAG METADATA" in logs

**Solution**: Set `log_level: "INFO"` in config logging section

## Implementation Details

### Code Changes

1. **`_setup_header_footer`** - Now calls `_add_row_index_to_header` when page_ranges available
2. **`_setup_enhanced_header`** - Formats page structure as compact summary for header
3. **`_add_row_index_to_header`** - NEW method that:
   - Adds "Page-Row Index" column to sheet
   - Inserts labels at each page break
   - Sets PrintTitleRows to repeat header
4. **`_set_row_col_headings`** - Deprecated, explicitly disables old PrintHeadings

### Page Break Calculation

Page breaks are calculated in `_insert_page_breaks_by_rows`:
- Analyzes actual row heights
- Compares against printable page height
- Inserts breaks at optimal positions
- Returns `page_ranges` array with structure:
  ```python
  [
      {'page': 1, 'start_row': 2, 'end_row': 31, 'row_count': 30},
      {'page': 2, 'start_row': 32, 'end_row': 61, 'row_count': 30},
      ...
  ]
  ```

## Best Practices

### For RAG Systems

1. **Always parse header** - Contains page structure summary
2. **Check added column** - Provides per-page row index labels
3. **Use conversion logs** - Most reliable source of complete mapping
4. **Validate ranges** - Ensure extracted content matches expected row range
5. **Track sheet names** - Different sheets have independent row numbering

### For Configuration

1. **Set appropriate `rows_per_page`** based on content density
2. **Keep `print_header_footer: true`** to enable row index display
3. **Use `log_level: "INFO"`** to capture RAG metadata in logs
4. **Test with sample files** before production deployment

## Migration from Old System

If you were using `print_row_col_headings: true`:

### Before:
```yaml
print_row_col_headings: true  # Showed 1, 2, 3, 4... on every page
```

**Issues:**
- Always started from 1 on every page
- Couldn't determine actual Excel row numbers
- No page-to-row mapping

### After:
```yaml
rows_per_page: 30            # Calculate page breaks
print_header_footer: true    # Show row index in header
print_row_col_headings: false  # Not needed anymore
```

**Benefits:**
- Shows actual Excel row numbers (2-31, 32-61, etc.)
- Page structure visible in header
- Complete RAG metadata in logs

## Summary

✅ **Automatic**: Row indices calculated from page breaks  
✅ **Accurate**: Shows exact Excel row numbers per page  
✅ **Visible**: Appears in PDF header and added column  
✅ **Complete**: Full RAG metadata in conversion logs  
✅ **Simple**: No manual configuration needed  

The new system provides everything needed for accurate RAG metadata extraction and source traceability.
