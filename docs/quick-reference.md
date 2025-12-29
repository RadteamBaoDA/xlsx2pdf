# Quick Reference: Paper Sizes & Orientation

## Quick Config Examples

### Force Landscape A4 for All Sheets
```yaml
print_options:
  page_size: "A4"
  orientation: "landscape"
  scaling: "fit_columns"
```

### Force Portrait Letter Size
```yaml
print_options:
  page_size: "letter"
  orientation: "portrait"
  scaling: "fit_columns"
```

### Auto-Detect Everything (Recommended)
```yaml
print_options:
  page_size: "auto"          # Automatically selects best paper size
  orientation: "auto"         # Automatically selects portrait/landscape
  scaling: "fit_columns"
```

### Different Settings Per Sheet
```yaml
print_options:
  # Wide dashboard sheets in landscape
  - sheets: ['Dashboard', 'Summary']
    priority: 1
    page_size: "A3"
    orientation: "landscape"
  
  # Data sheets in portrait
  - sheets: ['Data', 'Details']
    priority: 2
    page_size: "A4"
    orientation: "portrait"
  
  # Default for everything else
  - sheets: null
    priority: 99
    page_size: "auto"
    orientation: "auto"
```

## Common Paper Sizes

| Size | Dimensions | Best For |
|------|-----------|----------|
| **A4** | 21 x 29.7 cm | Standard documents (International) |
| **Letter** | 8.5 x 11 in | Standard documents (US) |
| **A3** | 29.7 x 42 cm | Large spreadsheets, posters |
| **Legal** | 8.5 x 14 in | Legal documents (US) |
| **Tabloid** | 11 x 17 in | Large format printing |
| **Ledger** | 17 x 11 in | Wide spreadsheets |
| **A5** | 14.8 x 21 cm | Small booklets, notes |
| **Executive** | 7.25 x 10.5 in | Executive summaries |
| **B4** | 25.7 x 36.4 cm | Large format (JIS) |
| **Envelope_10** | 4.125 x 9.5 in | Business envelopes |

## Orientation Quick Guide

### Auto (Default)
- Wide content (width > height) → **Landscape**
- Tall content (height ≥ width) → **Portrait**

### Portrait
- Vertical orientation (tall)
- Best for: Reports, documents, lists

### Landscape
- Horizontal orientation (wide)
- Best for: Spreadsheets, charts, dashboards

## Scaling Options

```yaml
scaling: "fit_columns"      # Fit all columns on one page width (recommended)
scaling: "fit_sheet"        # Fit entire sheet on one page
scaling: "fit_rows"         # Fit all rows on one page height
scaling: "no_scaling"       # Use original size (100% zoom)
scaling: "custom"           # Use custom percentage
scaling_percent: 85         # When scaling: custom
```

## Margin Options

```yaml
margins: "normal"           # Standard margins (1.91 cm top/bottom, 1.78 cm left/right)
margins: "narrow"           # Narrow margins (1.91 cm top/bottom, 0.64 cm left/right)
margins: "wide"             # Wide margins (2.54 cm all around)
margins: "custom"           # Use custom_margins values
```

## Complete Template

```yaml
print_options:
  sheets: null                      # null = all sheets, or ['Sheet1', 'Sheet2']
  priority: 1                       # Lower = higher priority
  mode: "auto"                      # auto, one_page, table_row_break, native_print
  
  # Paper and Orientation
  page_size: "auto"                 # auto, letter, A4, A3, legal, etc.
  orientation: "auto"               # auto, portrait, landscape
  
  # Scaling and Layout
  scaling: "fit_columns"            # fit_columns, fit_sheet, fit_rows, no_scaling, custom
  scaling_percent: 100              # Used when scaling: custom
  
  # Margins
  margins: "normal"                 # normal, narrow, wide, custom
  
  # Page Breaks
  rows_per_page: null               # Insert page break every N rows (null = none)
  columns_per_page: null            # Insert page break every N columns (null = none)
  
  # Headers and Footers
  print_header_footer: true         # Include sheet name in header, row range in footer
  print_row_col_headings: false     # Print row numbers and column letters
```

## Page Break Controls

### Smart Row-Based Page Breaks
The `rows_per_page` setting now intelligently calculates page breaks based on:
1. **Actual row heights** - Measures each row's real height in points
2. **Printable page area** - Considers paper size, orientation, and margins
3. **Maximum row limit** - Respects the rows_per_page value as an upper limit

```yaml
rows_per_page: 50           # Maximum 50 rows per page
                            # BUT breaks earlier if content height exceeds page
                            # System analyzes actual row heights dynamically
```

### How It Works
- System calculates printable height from: paper size + orientation + margins
- Accumulates actual row heights as it processes each row
- Inserts page break when:
  - Content height would exceed printable area, OR
  - Row count reaches rows_per_page limit
- Result: Content always fits properly on each page

### Examples

**Scenario 1: Small rows, high limit**
```yaml
rows_per_page: 100          # Limit: 100 rows
page_size: "A4"             # Printable: ~750 points
```
Result: If rows average 5 points each, fits ~150 rows. Page breaks at 100 rows (limit reached first).

**Scenario 2: Large rows, high limit**
```yaml
rows_per_page: 100          # Limit: 100 rows
page_size: "A4"             # Printable: ~750 points
```
Result: If rows average 15 points each, only ~50 rows fit. Page breaks at 50 rows (height limit reached first).

**Scenario 3: No row limit (height-based only)**
```yaml
rows_per_page: null         # No row limit
page_size: "A4"
```
Result: Pages break only when content height exceeds printable area. Number of rows per page varies based on actual row heights.

### Column-Based Page Breaks
```yaml
columns_per_page: 10        # Insert vertical page break every 10 columns
```
Simple column count-based breaks (does not calculate widths).

### Problem: Content is cut off
**Solutions:**
1. Increase paper size: `page_size: "A3"` or `page_size: "tabloid"`
2. Change to landscape: `orientation: "landscape"`
3. Use fit to sheet: `scaling: "fit_sheet"`

### Problem: Too much whitespace
**Solutions:**
1. Decrease paper size: `page_size: "A5"`
2. Use custom scaling: `scaling: "custom"` with `scaling_percent: 85`
3. Use narrow margins: `margins: "narrow"`

### Problem: Wrong orientation
**Solutions:**
1. Force orientation: `orientation: "portrait"` or `orientation: "landscape"`
2. Check content dimensions in Excel

### Problem: Different sheets need different settings
**Solution:** Use multiple configurations with priorities:
```yaml
print_options:
  - sheets: ['Sheet1']
    priority: 1
    page_size: "A4"
    orientation: "landscape"
  - sheets: null
    priority: 99
    page_size: "letter"
    orientation: "portrait"
```

## Tips

1. **Use "auto" first** - Let the system figure it out, then adjust if needed
2. **Test with sample file** - Run one file to verify settings before batch processing
3. **Match paper to content** - Wide spreadsheets → landscape, documents → portrait
4. **Consider printing** - Choose paper sizes your printer supports
5. **Preserve layout** - The converter maintains original Excel formatting automatically

## All Supported Paper Sizes

### Standard Sizes
letter, letter_small, legal, tabloid, ledger, statement, executive, folio, quarto, note, 10x14, 11x17

### A Series
A1, A2, A3, A4, A4_small, A5, A6

### B Series
B4, B5, B6

### Envelopes
envelope_9, envelope_10, envelope_11, envelope_12, envelope_14, envelope_dl, envelope_c3, envelope_c4, envelope_c5, envelope_c6, envelope_c65, envelope_b4, envelope_b5, envelope_b6, envelope_monarch

## Need More Help?

See full documentation: `docs/paper-sizes-and-orientation.md`
