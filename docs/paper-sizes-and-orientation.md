# Paper Sizes and Page Orientation Configuration

## Overview
The xlsx2pdf converter now supports all Microsoft Print to PDF paper sizes and configurable page orientations. The converter preserves the original Excel layout and style without modifications, ensuring accurate PDF output.

## Supported Paper Sizes

### Standard Letter Sizes
- **LETTER** - 8.5 x 11 in (21.59 x 27.94 cm)
- **LETTER_SMALL** - 8.5 x 11 in with smaller margins
- **LEGAL** - 8.5 x 14 in (21.59 x 35.56 cm)
- **TABLOID** - 11 x 17 in (27.94 x 43.18 cm)
- **LEDGER** - 17 x 11 in (43.18 x 27.94 cm)
- **STATEMENT** - 5.5 x 8.5 in (13.97 x 21.59 cm)
- **EXECUTIVE** - 7.25 x 10.5 in (18.41 x 26.67 cm)
- **FOLIO** - 8.5 x 13 in (21.59 x 33.02 cm)
- **QUARTO** - 8.47 x 10.8 in (21.59 x 27.43 cm)
- **NOTE** - 8.5 x 11 in (21.59 x 27.94 cm)
- **10X14** - 10 x 14 in (25.4 x 35.56 cm)
- **11X17** - 11 x 17 in (27.94 x 43.18 cm)

### A Series (ISO 216)
- **A1** - 59.4 x 84.1 cm
- **A2** - 42 x 59.4 cm
- **A3** - 29.7 x 42 cm
- **A4** - 21 x 29.7 cm (default)
- **A4_SMALL** - 21 x 29.7 cm with smaller margins
- **A5** - 14.8 x 21 cm
- **A6** - 10.5 x 14.8 cm

### B Series (JIS)
- **B4** - 25.7 x 36.4 cm (JIS)
- **B5** - 18.2 x 25.7 cm (JIS)
- **B6** - 12.5 x 17.6 cm

### Envelopes
- **ENVELOPE_9** - 3.875 x 8.875 in (9.84 x 22.54 cm)
- **ENVELOPE_10** - 4.125 x 9.5 in (10.48 x 24.13 cm)
- **ENVELOPE_11** - 4.5 x 10.375 in (11.43 x 26.35 cm)
- **ENVELOPE_12** - 4.75 x 11 in (12.07 x 27.94 cm)
- **ENVELOPE_14** - 5 x 11.5 in (12.7 x 29.21 cm)
- **ENVELOPE_DL** - 11 x 22 cm (DL Envelope)
- **ENVELOPE_C3** - 32.4 x 45.8 cm (C3 Envelope)
- **ENVELOPE_C4** - 22.9 x 32.4 cm (C4 Envelope)
- **ENVELOPE_C5** - 16.2 x 22.9 cm (C5 Envelope)
- **ENVELOPE_C6** - 11.4 x 16.2 cm (C6 Envelope)
- **ENVELOPE_C65** - 11.4 x 22.9 cm (C65 Envelope)
- **ENVELOPE_B4** - 25 x 35.3 cm (B4 Envelope)
- **ENVELOPE_B5** - 17.6 x 25 cm (B5 Envelope)
- **ENVELOPE_B6** - 17.6 x 12.5 cm (B6 Envelope)
- **ENVELOPE_MONARCH** - 3.875 x 7.5 in (9.84 x 19.05 cm)

## Page Orientation Options

### Auto (Default)
Automatically determines orientation based on content dimensions:
- **Wide content** (width > height) → Landscape
- **Tall/square content** (height ≥ width) → Portrait

### Portrait
Forces portrait orientation (vertical) for all sheets, regardless of content dimensions.

### Landscape
Forces landscape orientation (horizontal) for all sheets, regardless of content dimensions.

## Configuration Examples

### Single Configuration (All Sheets)
```yaml
print_options:
  sheets: null              # Applies to all sheets
  priority: 1
  mode: "auto"
  page_size: "auto"         # Auto-detect best size
  orientation: "auto"       # Auto-detect orientation
  scaling: "fit_columns"
  margins: "normal"
  print_header_footer: true
  print_row_col_headings: false
```

### Multiple Configurations (Per Sheet)
```yaml
print_options:
  # Landscape summaries
  - sheets: ['Summary', 'Dashboard']
    priority: 1             # Highest priority
    mode: "one_page"
    page_size: "A4"
    orientation: "landscape"  # Force landscape
    scaling: "fit_sheet"
    margins: "narrow"
    print_header_footer: false
    print_row_col_headings: false
  
  # Portrait data sheets
  - sheets: ['DataSheet1', 'DataSheet2']
    priority: 2
    mode: "auto"
    page_size: "letter"
    orientation: "portrait"   # Force portrait
    scaling: "fit_columns"
    margins: "normal"
    print_header_footer: true
    print_row_col_headings: false
  
  # Default for all other sheets
  - sheets: null
    priority: 99            # Lowest priority (fallback)
    mode: "auto"
    page_size: "auto"
    orientation: "auto"     # Auto-detect
    scaling: "fit_columns"
    margins: "normal"
    print_header_footer: true
    print_row_col_headings: false
```

### Specific Paper Size Examples
```yaml
# US Letter size with portrait orientation
print_options:
  page_size: "letter"
  orientation: "portrait"
  
# Legal size with landscape orientation
print_options:
  page_size: "legal"
  orientation: "landscape"
  
# Envelope printing
print_options:
  page_size: "envelope_10"
  orientation: "portrait"
  scaling: "fit_sheet"
```

## Key Features

### 1. Priority-Based Configuration
When using multiple configurations, lower priority numbers take precedence:
- Priority 1 = Highest priority
- Priority 99 = Lowest priority (typically the default/fallback)

When a sheet matches multiple configurations, only the highest priority one is applied.

### 2. Original Layout Preservation
The converter now **preserves the original Excel layout and style**:
- ✅ Original row heights maintained
- ✅ Original column widths maintained
- ✅ Original cell formatting preserved
- ✅ Original shape/image positions preserved
- ✅ No automatic resizing or refitting

### 3. Sheet-Specific Configuration
You can apply different settings to different sheets:
```yaml
print_options:
  - sheets: ['Sheet1', 'Sheet2']  # Specific sheets
    page_size: "A4"
    orientation: "landscape"
  - sheets: null                   # All other sheets
    page_size: "letter"
    orientation: "portrait"
```

### 4. Auto-Detection
The "auto" option intelligently selects:
- **page_size: "auto"** - Chooses smallest paper size that fits content
- **orientation: "auto"** - Selects orientation based on content dimensions

## Migration from Previous Version

### Removed: prepare_for_print Configuration
The `prepare_for_print` configuration option has been removed. The converter now:
- Always preserves original Excel dimensions
- Never modifies row heights or column widths
- Never adjusts text wrapping or cell formatting

### What Changed
**Before:**
```yaml
excel:
  prepare_for_print: false
  enhanced_dir: "enhanced_files"
```

**After:**
```yaml
# prepare_for_print section removed entirely
# Original layout is always preserved
```

**Print options scaling, margins, and headers are still fully configurable and applied during PDF export.**

## Best Practices

1. **Use "auto" for mixed content** - Let the system choose the best paper size and orientation
2. **Use specific sizes for consistent output** - When you need uniform paper sizes across all documents
3. **Use priority-based configs** - When different sheet types need different settings
4. **Test with sample files** - Verify your configuration produces the desired output

## Troubleshooting

### Content Gets Cut Off
- Increase paper size (e.g., from A4 to A3)
- Change orientation to landscape
- Adjust scaling to "fit_sheet" instead of "fit_columns"

### Wrong Orientation
- Set explicit orientation: "portrait" or "landscape"
- Check if content dimensions match expected orientation

### Inconsistent Paper Sizes
- Use uniform_page_size mode
- Or set explicit page_size instead of "auto"

## Technical Notes

- Paper sizes use Excel xlPaperSize enumeration constants
- Orientation uses xlPortrait (1) and xlLandscape (2) constants
- All dimensions converted from points (72 points = 1 inch)
- Printable area accounts for typical printer margins
