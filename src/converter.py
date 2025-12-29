import win32com.client
import win32com, shutil
from pathlib import Path
import pythoncom
import win32process
import os
import logging
from .utils import ensure_dir

# Excel Constants
xlTypePDF = 0
xlQualityStandard = 0
xlLandscape = 2
xlPortrait = 1

# Paper Size Constants (Excel xlPaperSize enumeration)
# All paper sizes supported by Microsoft Print to PDF
xlPaperLetter = 1       # 21.59 x 27.94 cm (8.5 x 11 in)
xlPaperLetterSmall = 2  # 21.59 x 27.94 cm (8.5 x 11 in) - Small margins
xlPaperTabloid = 3      # 27.94 x 43.18 cm (11 x 17 in)
xlPaperLedger = 4       # 43.18 x 27.94 cm (17 x 11 in)
xlPaperLegal = 5        # 21.59 x 35.56 cm (8.5 x 14 in)
xlPaperStatement = 6    # 13.97 x 21.59 cm (5.5 x 8.5 in)
xlPaperExecutive = 7    # 18.41 x 26.67 cm (7.25 x 10.5 in)
xlPaperA3 = 8           # 29.7 x 42 cm
xlPaperA4 = 9           # 21 x 29.7 cm
xlPaperA4Small = 10     # 21 x 29.7 cm - Small margins
xlPaperA5 = 11          # 14.8 x 21 cm
xlPaperB4 = 12          # 25.7 x 36.4 cm (JIS)
xlPaperB5 = 13          # 18.2 x 25.7 cm (JIS)
xlPaperFolio = 14       # 21.59 x 33.02 cm (8.5 x 13 in)
xlPaperQuarto = 15      # 21.59 x 27.43 cm (8.47 x 10.8 in)
xlPaper10x14 = 16       # 25.4 x 35.56 cm (10 x 14 in)
xlPaper11x17 = 17       # 27.94 x 43.18 cm (11 x 17 in)
xlPaperNote = 18        # 21.59 x 27.94 cm (8.5 x 11 in)
xlPaperEnvelope9 = 19   # 9.84 x 22.54 cm (3.875 x 8.875 in)
xlPaperEnvelope10 = 20  # 10.48 x 24.13 cm (4.125 x 9.5 in)
xlPaperEnvelope11 = 21  # 11.43 x 26.35 cm (4.5 x 10.375 in)
xlPaperEnvelope12 = 22  # 12.07 x 27.94 cm (4.75 x 11 in)
xlPaperEnvelope14 = 23  # 12.7 x 29.21 cm (5 x 11.5 in)
xlPaperCSheet = 24      # 43.18 x 55.88 cm (17 x 22 in)
xlPaperDSheet = 25      # 55.88 x 86.36 cm (22 x 34 in)
xlPaperESheet = 26      # 86.36 x 111.76 cm (34 x 44 in)
xlPaperEnvelopeDL = 27  # 11 x 22 cm (DL Envelope)
xlPaperEnvelopeC5 = 28  # 16.2 x 22.9 cm (C5 Envelope)
xlPaperEnvelopeC3 = 29  # 32.4 x 45.8 cm (C3 Envelope)
xlPaperEnvelopeC4 = 30  # 22.9 x 32.4 cm (C4 Envelope)
xlPaperEnvelopeC6 = 31  # 11.4 x 16.2 cm (C6 Envelope)
xlPaperEnvelopeC65 = 32 # 11.4 x 22.9 cm (C65 Envelope)
xlPaperEnvelopeB4 = 33  # 25 x 35.3 cm (B4 Envelope)
xlPaperEnvelopeB5 = 34  # 17.6 x 25 cm (B5 Envelope)
xlPaperEnvelopeB6 = 35  # 17.6 x 12.5 cm (B6 Envelope)
xlPaperEnvelopeItaly = 36  # 11 x 23 cm (Italy Envelope)
xlPaperEnvelopeMonarch = 37  # 9.84 x 19.05 cm (3.875 x 7.5 in)
xlPaperEnvelopePersonal = 38  # 9.21 x 16.51 cm (3.625 x 6.5 in)
xlPaperFanfoldUS = 39   # 37.83 x 27.94 cm (14.875 x 11 in)
xlPaperFanfoldStdGerman = 40  # 21.59 x 30.48 cm (8.5 x 12 in)
xlPaperFanfoldLegalGerman = 41  # 21.59 x 33.02 cm (8.5 x 13 in)
xlPaperA6 = 70          # 10.5 x 14.8 cm

# Note: A1 (67) and A2 (66) are NOT supported by Microsoft Print to PDF
# Use C Sheet, D Sheet, or E Sheet for large formats instead

# Print Mode Constants
PRINT_MODE_AUTO = "auto"
PRINT_MODE_ONE_PAGE = "one_page"
PRINT_MODE_TABLE_ROW_BREAK = "table_row_break"
PRINT_MODE_AUTO_PAGE_SIZE = "auto_page_size"
PRINT_MODE_NATIVE_PRINT = "native_print"
PRINT_MODE_UNIFORM_PAGE_SIZE = "uniform_page_size"

# Page sizes in points (1 inch = 72 points, 1 cm = 28.35 points)
# These are printable area estimates (minus typical margins)
# All paper sizes supported by Microsoft Print to PDF
PAGE_SIZES = {
    "LETTER": {"width": 612, "height": 792, "printable_height": 700, "xl_const": xlPaperLetter},
    "LETTER_SMALL": {"width": 612, "height": 792, "printable_height": 700, "xl_const": xlPaperLetterSmall},
    "TABLOID": {"width": 792, "height": 1224, "printable_height": 1130, "xl_const": xlPaperTabloid},
    "LEDGER": {"width": 1224, "height": 792, "printable_height": 700, "xl_const": xlPaperLedger},
    "LEGAL": {"width": 612, "height": 1008, "printable_height": 915, "xl_const": xlPaperLegal},
    "STATEMENT": {"width": 396, "height": 612, "printable_height": 520, "xl_const": xlPaperStatement},
    "EXECUTIVE": {"width": 522, "height": 756, "printable_height": 665, "xl_const": xlPaperExecutive},
    "FOLIO": {"width": 612, "height": 936, "printable_height": 845, "xl_const": xlPaperFolio},
    "QUARTO": {"width": 610, "height": 780, "printable_height": 690, "xl_const": xlPaperQuarto},
    "10X14": {"width": 720, "height": 1008, "printable_height": 915, "xl_const": xlPaper10x14},
    "11X17": {"width": 792, "height": 1224, "printable_height": 1130, "xl_const": xlPaper11x17},
    "NOTE": {"width": 612, "height": 792, "printable_height": 700, "xl_const": xlPaperNote},
    "ENVELOPE_9": {"width": 279, "height": 639, "printable_height": 550, "xl_const": xlPaperEnvelope9},
    "ENVELOPE_10": {"width": 297, "height": 684, "printable_height": 595, "xl_const": xlPaperEnvelope10},
    "ENVELOPE_11": {"width": 324, "height": 747, "printable_height": 660, "xl_const": xlPaperEnvelope11},
    "ENVELOPE_12": {"width": 342, "height": 792, "printable_height": 700, "xl_const": xlPaperEnvelope12},
    "ENVELOPE_14": {"width": 360, "height": 828, "printable_height": 735, "xl_const": xlPaperEnvelope14},
    "ENVELOPE_DL": {"width": 312, "height": 624, "printable_height": 535, "xl_const": xlPaperEnvelopeDL},
    "ENVELOPE_C3": {"width": 918, "height": 1296, "printable_height": 1205, "xl_const": xlPaperEnvelopeC3},
    "ENVELOPE_C4": {"width": 649, "height": 918, "printable_height": 830, "xl_const": xlPaperEnvelopeC4},
    "ENVELOPE_C5": {"width": 459, "height": 649, "printable_height": 560, "xl_const": xlPaperEnvelopeC5},
    "ENVELOPE_C6": {"width": 323, "height": 459, "printable_height": 370, "xl_const": xlPaperEnvelopeC6},
    "ENVELOPE_C65": {"width": 323, "height": 649, "printable_height": 560, "xl_const": xlPaperEnvelopeC65},
    "ENVELOPE_B4": {"width": 709, "height": 1001, "printable_height": 910, "xl_const": xlPaperEnvelopeB4},
    "ENVELOPE_B5": {"width": 499, "height": 709, "printable_height": 620, "xl_const": xlPaperEnvelopeB5},
    "ENVELOPE_B6": {"width": 354, "height": 499, "printable_height": 410, "xl_const": xlPaperEnvelopeB6},
    "ENVELOPE_MONARCH": {"width": 279, "height": 540, "printable_height": 450, "xl_const": xlPaperEnvelopeMonarch},
    # Note: A1 and A2 not reliably supported - use Architecture sizes or A3 instead
    # "A1": {"width": 1684, "height": 2384, "printable_height": 2290, "xl_const": 67},
    # "A2": {"width": 1191, "height": 1684, "printable_height": 1590, "xl_const": 66},
    "A3": {"width": 842, "height": 1191, "printable_height": 1100, "xl_const": xlPaperA3},
    "A4": {"width": 595, "height": 842, "printable_height": 750, "xl_const": xlPaperA4},
    "A4_SMALL": {"width": 595, "height": 842, "printable_height": 750, "xl_const": xlPaperA4Small},
    "A5": {"width": 420, "height": 595, "printable_height": 505, "xl_const": xlPaperA5},
    "A6": {"width": 298, "height": 420, "printable_height": 330, "xl_const": xlPaperA6},
    "B4": {"width": 729, "height": 1032, "printable_height": 940, "xl_const": xlPaperB4},
    "B5": {"width": 516, "height": 729, "printable_height": 640, "xl_const": xlPaperB5},
    "B6": {"width": 363, "height": 516, "printable_height": 425, "xl_const": xlPaperEnvelopeB6},
    # Large Format Engineering Sizes (C/D/E Sheet) - Fully supported by Microsoft Print to PDF
    "C_SHEET": {"width": 1224, "height": 1584, "printable_height": 1490, "xl_const": xlPaperCSheet},
    "D_SHEET": {"width": 1584, "height": 2448, "printable_height": 2355, "xl_const": xlPaperDSheet},
    "E_SHEET": {"width": 2448, "height": 3168, "printable_height": 3075, "xl_const": xlPaperESheet}
}

class ExcelConverter:
    def __init__(self, config):
        self.config = config

    def convert(self, input_path, output_path, pid_queue=None):
        """
        Converts an Excel file to PDF.
        """
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        
        try:
            # Local thge gen_py cache directory
            gen_path = Path(win32com.__gen_path__)
            # Remove the problem gen_py cache directory
            shutil.rmtree(gen_path, ignore_errors=True)
            # Force new instance for isolation
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Send PID back to parent if queue provided
            if pid_queue:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
                    pid_queue.put(pid)
                except Exception as e:
                    logging.warning(f"Failed to get Excel PID: {e}")

            # Handle ReadOnly attribute (remove it if present to allow editing/saving if needed, 
            # though we primarily need it for 'Edit Mode' as requested)
            if os.path.exists(input_path):
                try:
                    # Check if file is read-only
                    if not os.access(input_path, os.W_OK):
                        # Try to remove read-only attribute
                        os.chmod(input_path, 0o777)
                except Exception as e:
                    logging.warning(f"Could not change file permissions for {input_path}: {e}")

            try:
                # Open workbook
                # UpdateLinks=0 (Don't update), ReadOnly=False (Edit mode), IgnoreReadOnlyRecommended=True
                workbook = excel.Workbooks.Open(input_path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
            except Exception as e:
                logging.error(f"Failed to open workbook {input_path}: {e}")
                raise e

            # Optimize Layout and apply print mode
            # Get default print mode from first sheet's config (handles both dict and list formats)
            default_print_options = self._get_sheet_print_options("")  # Empty string to get default config
            print_mode = default_print_options.get('mode', PRINT_MODE_AUTO)
            logging.info(f"[{workbook.Name}] Using print mode: {print_mode}")
            
            self._optimize_layout(workbook, print_mode)
            
            # Ensure output directory exists
            ensure_dir(output_path)
            
            # Export to PDF using ExportAsFixedFormat (reliable, no dialog)
            self._export_to_pdf(workbook, output_path)
            
            return True

        except Exception as e:
            logging.error(f"Error converting {input_path}: {e}")
            raise e
        finally:
            # Cleanup
            if workbook:
                try:
                    workbook.Close(SaveChanges=False)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()



    def _fix_shape_placement(self, sheet):
        """
        Sets shapes to move with cells but NOT resize.
        This prevents objects from being distorted when row heights change.
        """
        try:
            # xlMove = 2 (Move but don't size with cells)
            # This keeps original object dimensions while following cell layout
            for shape in sheet.Shapes:
                try:
                    shape.Placement = 2
                except:
                    pass
        except Exception as e:
            logging.warning(f"Could not fix shape placement in {sheet.Name}: {e}")

    def _fit_shapes_to_cells(self, sheet, workbook_name):
        """
        DISABLED: Preserves original shape positions and sizes.
        Does not modify shapes to preserve exact Excel layout.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Preserving original shape layout")

    def _ensure_shapes_visible(self, sheet, workbook_name):
        """
        Ensures all shapes (images, charts, objects) are visible and properly sized.
        Fixes issues where images/objects may be hidden or clipped.
        """
        try:
            shape_count = 0
            fixed_count = 0
            
            for shape in sheet.Shapes:
                try:
                    shape_count += 1
                    
                    # Ensure shape is visible
                    if hasattr(shape, 'Visible'):
                        if not shape.Visible:
                            shape.Visible = True
                            fixed_count += 1
                    
                    # Ensure shape prints
                    try:
                        shape.PrintObject = True
                    except:
                        pass
                    
                    # Get the cell range the shape is in
                    try:
                        top_left_cell = shape.TopLeftCell
                        
                        # DISABLED: Do not modify row heights
                        # Keep original Excel layout as-is
                        pass
                            
                    except:
                        pass
                        
                except Exception as e:
                    continue
            
            if fixed_count > 0:
                logging.info(f"[{workbook_name}] {sheet.Name}: Fixed {fixed_count} shapes/images visibility")
            elif shape_count > 0:
                logging.info(f"[{workbook_name}] {sheet.Name}: Checked {shape_count} shapes/images")
                
        except Exception as e:
            logging.warning(f"Could not ensure shapes visible in {sheet.Name}: {e}")

    def _fix_tables_for_print(self, sheet, workbook_name):
        """
        Ensures Excel Tables (ListObjects) are properly formatted for printing.
        Preserves fonts, styles, and ensures all content is visible.
        """
        try:
            table_count = sheet.ListObjects.Count
            if table_count == 0:
                return
            
            for i in range(1, table_count + 1):
                try:
                    table = sheet.ListObjects(i)
                    table_range = table.Range
                    
                    # DISABLED: Do not autofit table rows - preserve original heights
                    
                    # Ensure table style allows printing
                    try:
                        # Make sure borders are visible for print
                        table_range.Borders.LineStyle = 1  # xlContinuous
                    except:
                        pass
                    
                    logging.info(f"[{workbook_name}] {sheet.Name}: Prepared table '{table.Name}' for print")
                    
                except Exception as e:
                    logging.warning(f"Could not fix table {i} in {sheet.Name}: {e}")
                    
        except Exception as e:
            logging.warning(f"Could not fix tables in {sheet.Name}: {e}")

    def _ensure_text_visible(self, sheet, workbook_name):
        """
        DISABLED: Preserves original Excel text formatting and cell dimensions.
        Does not modify cell properties to preserve exact Excel layout.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Preserving original text layout")

    def _fix_cell_layout(self, sheet, workbook_name):
        """
        DISABLED: Preserves original cell alignment and layout.
        Does not modify cell properties to preserve exact Excel layout.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Preserving original cell layout")

    def _autofit_columns_smart(self, sheet, workbook_name):
        """
        DISABLED: Preserves original column widths.
        Does not modify column dimensions to preserve exact Excel layout.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Preserving original column widths")

    def _get_sheet_print_options(self, sheet_name):
        """
        Get the appropriate print_options for a specific sheet based on sheet name matching.
        Supports both single print_options dict and list of print_options with priority.
        Returns the matched print_options dict.
        """
        print_options_config = self.config.get('print_options', {})
        
        # Handle backward compatibility: single dict format
        if isinstance(print_options_config, dict):
            return print_options_config
        
        # Handle list format: multiple print_options with sheet matching
        if isinstance(print_options_config, list):
            matched_configs = []
            
            for config in print_options_config:
                sheets = config.get('sheets', None)
                priority = config.get('priority', 999)
                
                # If sheets is None or empty, it's a default config (matches all)
                if sheets is None or sheets == []:
                    matched_configs.append((priority, config))
                # If sheet name matches any in the list
                elif isinstance(sheets, list) and sheet_name in sheets:
                    matched_configs.append((priority, config))
            
            # Sort by priority (lower number = higher priority)
            if matched_configs:
                matched_configs.sort(key=lambda x: x[0])
                return matched_configs[0][1]  # Return highest priority config
        
        # Fallback: return default config
        return {
            'mode': 'auto',
            'page_size': 'auto',
            'orientation': 'auto',
            'scaling': 'fit_columns',
            'scaling_percent': 100,
            'margins': 'normal',
            'print_header_footer': True,
            'print_row_col_headings': False
        }

    def _determine_orientation(self, sheet, orientation_setting):
        """
        Determine the appropriate page orientation based on content and config.
        
        Args:
            sheet: Excel sheet object
            orientation_setting: Config value ('auto', 'portrait', or 'landscape')
        
        Returns:
            xlPortrait (1) or xlLandscape (2)
        """
        orientation_setting = str(orientation_setting).lower() if orientation_setting else 'auto'
        
        # Forced orientations
        if orientation_setting == 'portrait':
            return xlPortrait
        elif orientation_setting == 'landscape':
            return xlLandscape
        
        # Auto-detect based on content dimensions
        try:
            total_width_pts = sheet.UsedRange.Width
            total_height_pts = sheet.UsedRange.Height
            
            if total_width_pts > total_height_pts:
                return xlLandscape  # Wide content -> Landscape
            else:
                return xlPortrait   # Tall/Square content -> Portrait
        except:
            return xlPortrait  # Default to portrait if dimensions can't be read

    def _optimize_layout(self, workbook, print_mode=PRINT_MODE_AUTO):
        """
        Prepares workbook for print - merged logic from optimize_layout and enhance_layout.
        Ensures no content is hidden: expands collapsed groups, fixes row heights, fixes images.
        """

        # For native_print mode, set basic PageSetup to preserve exact Excel dimensions
        if print_mode == PRINT_MODE_NATIVE_PRINT:
            logging.info(f"[{workbook.Name}] Native print mode - preserving exact Excel dimensions for RAG")
            for sheet in workbook.Sheets:
                try:
                    # Set to no scaling - preserve exact dimensions (100% zoom)
                    sheet.PageSetup.Zoom = 100
                    sheet.PageSetup.FitToPagesWide = False
                    sheet.PageSetup.FitToPagesTall = False
                    
                    # Set print area to used range only (removes white space)
                    try:
                        sheet.PageSetup.PrintArea = sheet.UsedRange.Address
                    except:
                        pass
                    
                    logging.info(f"[{workbook.Name}] {sheet.Name}: Native print - exact dimensions preserved (no scaling)")
                except Exception as e:
                    logging.warning(f"Could not set native print mode for {sheet.Name}: {e}")
            return

        for sheet in workbook.Sheets:
            try:
                logging.info(f"[{workbook.Name}] Processing Sheet: {sheet.Name}")
                
                # Get sheet-specific print_options (supports multiple configs with sheet matching)
                print_options = self._get_sheet_print_options(sheet.Name)
                sheet_print_mode = print_options.get('mode', print_mode)
                page_size = print_options.get('page_size', 'A4').upper()
                rows_per_page = print_options.get('rows_per_page')
                orientation = print_options.get('orientation', 'auto')
                
                logging.info(f"[{workbook.Name}] {sheet.Name}: Using print mode '{sheet_print_mode}' (priority-based config)")
                
                # ========================================
                # STEP 1: EXPAND ALL HIDDEN CONTENT (ALWAYS)
                # Critical for ExportAsFixedFormat - must show all content
                # ========================================
                self._expand_all_groups(sheet, workbook.Name)
                self._unhide_rows_columns(sheet, workbook.Name)
                
                # ========================================
                # STEP 2: FIX SHAPE/IMAGE PLACEMENT (ALWAYS)
                # Prevent images from being hidden or distorted
                # ========================================
                self._fix_shape_placement(sheet)
                self._ensure_shapes_visible(sheet, workbook.Name)

                # ========================================
                # STEP 3: PRINT MODE SPECIFIC SETUP
                # ========================================
                if sheet_print_mode == PRINT_MODE_ONE_PAGE:
                    self._apply_one_page_mode(sheet, workbook.Name, orientation, page_size)
                elif sheet_print_mode == PRINT_MODE_TABLE_ROW_BREAK:
                    self._apply_table_row_break_mode(sheet, workbook.Name, rows_per_page, orientation, page_size)
                elif sheet_print_mode == PRINT_MODE_AUTO_PAGE_SIZE:
                    self._apply_auto_page_size_mode(sheet, workbook.Name, page_size, orientation)
                elif sheet_print_mode == PRINT_MODE_UNIFORM_PAGE_SIZE:
                    # Uniform page size is handled at workbook level, not per-sheet
                    # Skip here - will be applied after all sheets are processed
                    pass
                else:
                    # Default AUTO mode
                    self._apply_auto_mode(sheet, workbook.Name, orientation, page_size)

                # ========================================
                # STEP 4: PRESERVE ORIGINAL DIMENSIONS (NO MODIFICATIONS)
                # Keep original row heights and column widths from Excel file
                # ========================================
                logging.info(f"[{workbook.Name}] {sheet.Name}: Preserving original row/column dimensions")

                # ========================================
                # STEP 5: APPLY SCALING (ALWAYS APPLIED)
                # Match Excel's print scaling options from config
                # This controls how Excel fits content to pages during PDF export
                # ========================================
                scaling = print_options.get('scaling', 'fit_columns')
                scaling_percent = print_options.get('scaling_percent', 100)
                self._apply_scaling(sheet, workbook.Name, scaling, scaling_percent)

                # ========================================
                # STEP 6: APPLY MARGINS
                # Match Excel's print margin options
                # ========================================
                margins = print_options.get('margins', 'normal')
                custom_margins = print_options.get('custom_margins', {})
                self._apply_margins(sheet, workbook.Name, margins, custom_margins)

                # ========================================
                # STEP 6.5: APPLY CUSTOM PAGE BREAKS
                # Insert page breaks based on row/column limits
                # ========================================
                rows_per_page_custom = print_options.get('rows_per_page')
                columns_per_page_custom = print_options.get('columns_per_page')
                
                if rows_per_page_custom:
                    self._insert_page_breaks_by_rows(sheet, workbook.Name, rows_per_page_custom)
                
                if columns_per_page_custom:
                    self._insert_page_breaks_by_columns(sheet, workbook.Name, columns_per_page_custom)

                # ========================================
                # STEP 7: SETUP HEADER AND FOOTER
                # Add sheet name to header and row range to footer
                # ========================================
                if print_options.get('print_header_footer', True):
                    self._setup_header_footer(sheet, workbook.Name)
                else:
                    # Clear any existing header/footer
                    self._clear_header_footer(sheet, workbook.Name)

                # ========================================
                # STEP 8: ROW AND COLUMN HEADINGS
                # Print Excel row numbers (1,2,3...) and column letters (A,B,C...)
                # ========================================
                print_headings = print_options.get('print_row_col_headings', False)
                self._set_row_col_headings(sheet, workbook.Name, print_headings)

            except Exception as e:
                logging.warning(f"Could not prepare sheet {sheet.Name}: {e}")

        # Handle uniform_page_size mode at workbook level (after all sheets processed)
        if print_mode == PRINT_MODE_UNIFORM_PAGE_SIZE:
            self._apply_uniform_page_size_mode(workbook, page_size)

    def _expand_all_groups(self, sheet, workbook_name):
        """
        Expands all collapsed row and column groups (outline groups).
        Ensures grouped/collapsed sections are visible in PDF.
        """
        try:
            # Show all outline levels (expand all groups)
            # Level 8 is the maximum outline level in Excel
            try:
                sheet.Outline.ShowLevels(RowLevels=8, ColumnLevels=8)
                logging.info(f"[{workbook_name}] {sheet.Name}: Expanded all outline groups")
            except:
                # Sheet may not have outlines
                pass
        except Exception as e:
            logging.warning(f"Could not expand groups in {sheet.Name}: {e}")

    def _unhide_rows_columns(self, sheet, workbook_name):
        """
        Unhides all hidden rows and columns to ensure all data is visible.
        """
        try:
            hidden_rows = 0
            hidden_cols = 0
            
            # Unhide all rows in used range
            for row in sheet.UsedRange.Rows:
                try:
                    if row.Hidden:
                        row.Hidden = False
                        hidden_rows += 1
                except:
                    pass
            
            # Unhide all columns in used range
            for col in sheet.UsedRange.Columns:
                try:
                    if col.Hidden:
                        col.Hidden = False
                        hidden_cols += 1
                except:
                    pass
            
            if hidden_rows > 0 or hidden_cols > 0:
                logging.info(f"[{workbook_name}] {sheet.Name}: Unhid {hidden_rows} rows, {hidden_cols} columns")
        except Exception as e:
            logging.warning(f"Could not unhide rows/columns in {sheet.Name}: {e}")

    def _apply_scaling(self, sheet, workbook_name, scaling, scaling_percent=100):
        """
        Apply Excel-like scaling options to sheet.
        Options: no_scaling, fit_sheet, fit_columns, fit_rows, custom
        """
        try:
            scaling = scaling.lower() if scaling else 'fit_columns'
            
            if scaling == 'no_scaling':
                # Print at actual size (100% zoom, no fitting)
                sheet.PageSetup.Zoom = 100
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False
                logging.info(f"[{workbook_name}] {sheet.Name}: Scaling -> No Scaling (actual size)")
                
            elif scaling == 'fit_sheet':
                # Fit entire sheet on one page
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
                logging.info(f"[{workbook_name}] {sheet.Name}: Scaling -> Fit Sheet on One Page")
                
            elif scaling == 'fit_columns':
                # Fit all columns on one page (rows can span multiple pages)
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
                logging.info(f"[{workbook_name}] {sheet.Name}: Scaling -> Fit All Columns on One Page")
                
            elif scaling == 'fit_rows':
                # Fit all rows on one page (columns can span multiple pages)
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = 1
                logging.info(f"[{workbook_name}] {sheet.Name}: Scaling -> Fit All Rows on One Page")
                
            elif scaling == 'custom':
                # Custom scaling percentage (1-400%)
                zoom = max(1, min(400, scaling_percent))
                sheet.PageSetup.Zoom = zoom
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False
                logging.info(f"[{workbook_name}] {sheet.Name}: Scaling -> Custom {zoom}%")
                
            else:
                # Default: fit columns
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
                logging.info(f"[{workbook_name}] {sheet.Name}: Scaling -> Fit All Columns (default)")
                
        except Exception as e:
            logging.warning(f"Could not apply scaling to {sheet.Name}: {e}")

    def _apply_margins(self, sheet, workbook_name, margins, custom_margins=None):
        """
        Apply Excel-like margin options to sheet.
        Options: normal, wide, narrow, custom
        Margin values are in centimeters, converted to points (1 cm = 28.35 points)
        """
        CM_TO_POINTS = 28.35
        
        # Predefined margin presets (matching Excel exactly)
        MARGIN_PRESETS = {
            'normal': {'top': 1.91, 'bottom': 1.91, 'left': 1.78, 'right': 1.78, 'header': 0.76, 'footer': 0.76},
            'wide': {'top': 2.54, 'bottom': 2.54, 'left': 2.54, 'right': 2.54, 'header': 1.27, 'footer': 1.27},
            'narrow': {'top': 1.91, 'bottom': 1.91, 'left': 0.64, 'right': 0.64, 'header': 0.76, 'footer': 0.76}
        }
        
        try:
            margins = margins.lower() if margins else 'normal'
            
            if margins == 'custom' and custom_margins:
                # Use custom margin values
                m = custom_margins
            elif margins in MARGIN_PRESETS:
                m = MARGIN_PRESETS[margins]
            else:
                m = MARGIN_PRESETS['normal']
            
            # Apply margins (convert cm to points)
            sheet.PageSetup.TopMargin = m.get('top', 1.91) * CM_TO_POINTS
            sheet.PageSetup.BottomMargin = m.get('bottom', 1.91) * CM_TO_POINTS
            sheet.PageSetup.LeftMargin = m.get('left', 1.78) * CM_TO_POINTS
            sheet.PageSetup.RightMargin = m.get('right', 1.78) * CM_TO_POINTS
            sheet.PageSetup.HeaderMargin = m.get('header', 0.76) * CM_TO_POINTS
            sheet.PageSetup.FooterMargin = m.get('footer', 0.76) * CM_TO_POINTS
            
            logging.info(f"[{workbook_name}] {sheet.Name}: Margins -> {margins.capitalize()}")
            
        except Exception as e:
            logging.warning(f"Could not apply margins to {sheet.Name}: {e}")

    def _apply_auto_mode(self, sheet, workbook_name, orientation='auto', page_size='auto'):
        """
        Default AUTO mode - fit columns to page, configurable orientation and page size.
        NOTE: This mode applies fitting which may alter dimensions.
        For exact dimension preservation, use native_print mode instead.
        """
        try:
            total_width_pts = sheet.UsedRange.Width
            total_height_pts = sheet.UsedRange.Height
        except:
            total_width_pts = 0
            total_height_pts = 0

        # Determine orientation
        page_orientation = self._determine_orientation(sheet, orientation)
        sheet.PageSetup.Orientation = page_orientation
        
        # Set paper size based on config or auto-detect
        page_size_upper = page_size.upper() if page_size else "AUTO"
        
        if page_size_upper == "AUTO":
            # Auto-detect paper size based on content and orientation
            if page_orientation == xlLandscape:
                # Landscape orientation
                if total_width_pts < 900:
                    sheet.PageSetup.PaperSize = xlPaperA4
                    logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> A4 Landscape")
                else:
                    sheet.PageSetup.PaperSize = xlPaperA3
                    logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> A3 Landscape")
            else:
                # Portrait orientation
                sheet.PageSetup.PaperSize = xlPaperA4
                logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> A4 Portrait")
        else:
            # Use configured page size
            if page_size_upper in PAGE_SIZES:
                page_info = PAGE_SIZES[page_size_upper]
                try:
                    sheet.PageSetup.PaperSize = page_info["xl_const"]
                    logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> {page_size_upper} {('Landscape' if page_orientation == xlLandscape else 'Portrait')}")
                except Exception as e:
                    # Fallback to A3
                    sheet.PageSetup.PaperSize = xlPaperA3
                    logging.warning(f"[{workbook_name}] {sheet.Name}: Paper size '{page_size_upper}' not supported, using A3. For large formats, use C_SHEET, D_SHEET, or E_SHEET instead.")
            else:
                # Fallback to A4 if invalid page size
                sheet.PageSetup.PaperSize = xlPaperA4
                logging.warning(f"[{workbook_name}] {sheet.Name}: Invalid page size '{page_size}', using A4")

        # Force Fit to 1 Page Wide (keeps original column proportions)
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False 
        
        # Set print area to used range only (removes white space)
        try:
            sheet.PageSetup.PrintArea = sheet.UsedRange.Address
        except:
            pass

    def _apply_one_page_mode(self, sheet, workbook_name, orientation='auto', page_size='auto'):
        """
        ONE PAGE mode - fit entire sheet content to a single page.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Applying One Page mode")
        
        # Set page size
        page_size_upper = page_size.upper() if page_size else "AUTO"
        if page_size_upper != "AUTO" and page_size_upper in PAGE_SIZES:
            try:
                sheet.PageSetup.PaperSize = PAGE_SIZES[page_size_upper]["xl_const"]
            except Exception as e:
                # Fallback to A3
                sheet.PageSetup.PaperSize = xlPaperA3
                logging.warning(f"[{workbook_name}] {sheet.Name}: Paper size '{page_size_upper}' not supported, using A3. For large formats, use ARCH_C, ARCH_D, or ARCH_E.")
        else:
            sheet.PageSetup.PaperSize = xlPaperA4
        
        # Determine and set orientation
        page_orientation = self._determine_orientation(sheet, orientation)
        sheet.PageSetup.Orientation = page_orientation
        
        # Fit BOTH width and height to 1 page
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = 1
        
        # Set print area to used range only (removes white space)
        try:
            sheet.PageSetup.PrintArea = sheet.UsedRange.Address
        except:
            pass

    def _apply_table_row_break_mode(self, sheet, workbook_name, rows_per_page=None, orientation='auto', page_size='auto'):
        """
        TABLE ROW BREAK mode - insert page breaks after tables or every N rows.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Applying Table Row Break mode")
        
        # Set paper size
        page_size_upper = page_size.upper() if page_size else "AUTO"
        if page_size_upper != "AUTO" and page_size_upper in PAGE_SIZES:
            try:
                sheet.PageSetup.PaperSize = PAGE_SIZES[page_size_upper]["xl_const"]
            except Exception as e:
                # Fallback to A3
                sheet.PageSetup.PaperSize = xlPaperA3
                logging.warning(f"[{workbook_name}] {sheet.Name}: Paper size '{page_size_upper}' not supported, using A3. For large formats, use ARCH_C, ARCH_D, or ARCH_E.")
        else:
            sheet.PageSetup.PaperSize = xlPaperA4
        
        # Set orientation
        page_orientation = self._determine_orientation(sheet, orientation)
        sheet.PageSetup.Orientation = page_orientation
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
        
        # Clear existing page breaks
        try:
            sheet.ResetAllPageBreaks()
        except Exception as e:
            logging.warning(f"Could not reset page breaks: {e}")
        
        if rows_per_page:
            # Fixed rows per page mode
            self._insert_page_breaks_by_rows(sheet, workbook_name, rows_per_page)
        else:
            # Auto-detect tables (ListObjects) and break after each
            self._insert_page_breaks_by_tables(sheet, workbook_name)

    def _insert_page_breaks_by_tables(self, sheet, workbook_name):
        """
        Insert page breaks after each Excel Table (ListObject).
        """
        try:
            table_count = sheet.ListObjects.Count
            if table_count == 0:
                logging.info(f"[{workbook_name}] {sheet.Name}: No tables found, using data region breaks")
                # Fallback: try to detect data regions by used range
                self._insert_page_breaks_by_rows(sheet, workbook_name, 50)  # Default 50 rows
                return
            
            for i in range(1, table_count + 1):
                try:
                    table = sheet.ListObjects(i)
                    table_range = table.Range
                    last_row = table_range.Row + table_range.Rows.Count
                    
                    # Insert page break after the table
                    sheet.HPageBreaks.Add(Before=sheet.Rows(last_row + 1))
                    logging.info(f"[{workbook_name}] {sheet.Name}: Page break after table '{table.Name}' at row {last_row}")
                except Exception as e:
                    logging.warning(f"Could not add page break for table: {e}")
                    
        except Exception as e:
            logging.warning(f"Error detecting tables in {sheet.Name}: {e}")

    def _insert_page_breaks_by_rows(self, sheet, workbook_name, rows_per_page):
        """
        Insert page breaks intelligently based on rows_per_page configuration.
        Calculates actual row heights and ensures content fits within the page's printable area.
        """
        try:
            used_range = sheet.UsedRange
            used_rows = used_range.Rows.Count
            start_row = used_range.Row
            
            # Get current page setup to calculate printable height
            try:
                page_size_const = sheet.PageSetup.PaperSize
                orientation = sheet.PageSetup.Orientation
                
                # Find matching page size info
                printable_height = None
                for size_name, size_info in PAGE_SIZES.items():
                    if size_info["xl_const"] == page_size_const:
                        if orientation == xlLandscape:
                            # For landscape, swap width and height
                            printable_height = size_info["width"] - 100  # Account for margins
                        else:
                            printable_height = size_info["printable_height"]
                        break
                
                # Fallback: use A4 portrait if page size not found
                if printable_height is None:
                    printable_height = PAGE_SIZES["A4"]["printable_height"]
                    logging.warning(f"[{workbook_name}] {sheet.Name}: Could not determine page size, using A4 defaults")
                
                # Account for margins from PageSetup
                try:
                    top_margin = sheet.PageSetup.TopMargin
                    bottom_margin = sheet.PageSetup.BottomMargin
                    header_margin = sheet.PageSetup.HeaderMargin
                    footer_margin = sheet.PageSetup.FooterMargin
                    
                    # Adjust printable height by subtracting margins
                    available_height = printable_height - (top_margin + bottom_margin + header_margin + footer_margin)
                    if available_height > 0:
                        printable_height = available_height
                except:
                    pass  # Use default printable_height if margin reading fails
                
                logging.info(f"[{workbook_name}] {sheet.Name}: Printable page height: {printable_height:.0f} points")
                
            except Exception as e:
                # Fallback to A4 portrait if page setup reading fails
                printable_height = PAGE_SIZES["A4"]["printable_height"]
                logging.warning(f"[{workbook_name}] {sheet.Name}: Could not read page setup, using A4 defaults: {e}")
            
            # Page break insertion logic
            rows_in_current_page = 0
            page_count = 1
            last_break_row = start_row
            
            # When rows_per_page is set, use ONLY row count (ignore height calculations)
            if rows_per_page:
                for i in range(used_rows):
                    try:
                        row_index = start_row + i
                        rows_in_current_page += 1
                        
                        # Insert page break after every rows_per_page rows
                        if rows_in_current_page >= rows_per_page and i + 1 < used_rows:
                            next_row_index = row_index + 1
                            try:
                                sheet.HPageBreaks.Add(Before=sheet.Rows(next_row_index))
                                page_count += 1
                                logging.debug(f"[{workbook_name}] {sheet.Name}: Page break at row {next_row_index} (every {rows_per_page} rows)")
                                rows_in_current_page = 0
                                last_break_row = next_row_index
                            except Exception as e:
                                logging.warning(f"Could not insert page break at row {next_row_index}: {e}")
                    except Exception as e:
                        logging.warning(f"Error processing row {i}: {e}")
                        continue
            else:
                # No rows_per_page limit - use height-based calculation only
                accumulated_height = 0
                for i in range(used_rows):
                    try:
                        row = used_range.Rows(i + 1)
                        row_height = row.Height
                        row_index = start_row + i
                        
                        # Check if adding this row would exceed page height
                        if accumulated_height + row_height > printable_height:
                            # Insert page break before this row
                            try:
                                sheet.HPageBreaks.Add(Before=sheet.Rows(row_index))
                                page_count += 1
                                logging.debug(f"[{workbook_name}] {sheet.Name}: Page break at row {row_index} (accumulated: {accumulated_height:.0f}pts, limit: {printable_height:.0f}pts)")
                                accumulated_height = row_height
                                last_break_row = row_index
                            except Exception as e:
                                logging.warning(f"Could not insert page break at row {row_index}: {e}")
                        else:
                            accumulated_height += row_height
                    except Exception as e:
                        logging.warning(f"Error processing row {i}: {e}")
                        continue
            
            avg_rows_per_page = used_rows / page_count if page_count > 0 else used_rows
            logging.info(f"[{workbook_name}] {sheet.Name}: Inserted {page_count - 1} smart page breaks ({avg_rows_per_page:.1f} avg rows/page, max {rows_per_page} rows/page)")
            
        except Exception as e:
            logging.warning(f"Error inserting smart row-based page breaks: {e}")

    def _insert_page_breaks_by_columns(self, sheet, workbook_name, columns_per_page):
        """
        Insert vertical page breaks every N columns.
        Automatically splits content into new pages when column count reaches limit.
        """
        try:
            used_cols = sheet.UsedRange.Columns.Count
            start_col = sheet.UsedRange.Column
            
            # Insert vertical page breaks at column intervals
            for col in range(start_col + columns_per_page, start_col + used_cols, columns_per_page):
                try:
                    sheet.VPageBreaks.Add(Before=sheet.Columns(col))
                except:
                    pass
            
            breaks_count = (used_cols // columns_per_page)
            logging.info(f"[{workbook_name}] {sheet.Name}: Inserted {breaks_count} vertical page breaks (every {columns_per_page} columns)")
        except Exception as e:
            logging.warning(f"Error inserting column-based page breaks: {e}")

    def _find_best_page_size(self, content_width, content_height):
        """
        Find the smallest page size that can fit the content.
        Considers both portrait and landscape orientations.
        Returns the page size name (e.g., 'A4', 'LETTER', etc.)
        """
        # Sort page sizes by area (smallest first) for efficient selection
        sorted_sizes = sorted(
            PAGE_SIZES.items(),
            key=lambda x: x[1]["width"] * x[1]["height"]
        )
        
        # Try to find smallest page that fits content width
        # (height can span multiple pages, but width should fit)
        for size_name, size_info in sorted_sizes:
            page_width = size_info["width"]
            page_height = size_info["height"]
            
            # Check portrait orientation
            if content_width <= page_width:
                return size_name
            
            # Check landscape orientation (swap width/height)
            if content_width <= page_height:
                return size_name
        
        # If nothing fits, return largest size (A1)
        return "A1"

    def _apply_auto_page_size_mode(self, sheet, workbook_name, page_size="A4", orientation='auto'):
        """
        AUTO PAGE SIZE mode - calculate page breaks based on selected page size.
        Supports: auto, letter, tabloid, legal, statement, executive, A1, A2, A3, A4, A5, A6, B4, B5, B6, and more
        """
        # Get content dimensions first for auto page size detection
        try:
            total_width_pts = sheet.UsedRange.Width
            total_height_pts = sheet.UsedRange.Height
        except:
            total_width_pts = 0
            total_height_pts = 0
        
        # Auto-detect page size based on content dimensions
        page_size_upper = page_size.upper() if page_size else "A4"
        if page_size_upper == "AUTO":
            page_size_upper = self._find_best_page_size(total_width_pts, total_height_pts)
            logging.info(f"[{workbook_name}] {sheet.Name}: Auto-selected {page_size_upper} (content: {total_width_pts:.0f}x{total_height_pts:.0f}pts)")
        
        logging.info(f"[{workbook_name}] {sheet.Name}: Applying Auto Page Size mode ({page_size_upper})")
        
        # Get page size info from dictionary, default to A4
        if page_size_upper in PAGE_SIZES:
            page_info = PAGE_SIZES[page_size_upper]
        else:
            page_info = PAGE_SIZES["A4"]
            page_size_upper = "A4"
        
        # Set paper size using Excel constant with error handling
        try:
            sheet.PageSetup.PaperSize = page_info["xl_const"]
        except Exception as e:
            # Fallback to A3
            sheet.PageSetup.PaperSize = xlPaperA3
            logging.warning(f"[{workbook_name}] {sheet.Name}: Paper size '{page_size_upper}' not supported, using A3. For large formats, use C_SHEET, D_SHEET, or E_SHEET.")
            # Update page_info to use fallback
            page_info = PAGE_SIZES["A3"]
        
        printable_height = page_info["printable_height"]
        
        # Determine and set orientation
        page_orientation = self._determine_orientation(sheet, orientation)
        sheet.PageSetup.Orientation = page_orientation
        
        # For landscape, swap printable height with width
        if page_orientation == xlLandscape:
            printable_height = page_info["width"] - 100  # Account for margins
        
        # Fit width to 1 page
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
        
        # Clear existing page breaks
        try:
            sheet.ResetAllPageBreaks()
        except:
            pass
        
        # Calculate and insert page breaks based on row heights
        try:
            accumulated_height = 0
            page_count = 1
            
            for row in sheet.UsedRange.Rows:
                try:
                    row_height = row.Height
                    accumulated_height += row_height
                    
                    if accumulated_height > printable_height:
                        # Insert page break before this row
                        sheet.HPageBreaks.Add(Before=row)
                        page_count += 1
                        accumulated_height = row_height
                except:
                    continue
            
            logging.info(f"[{workbook_name}] {sheet.Name}: Created {page_count} pages for {page_size}")
        except Exception as e:
            logging.warning(f"Error calculating auto page breaks: {e}")

    def _find_max_content_width(self, workbook):
        """
        Scan all sheets in the workbook and find the maximum content width.
        Returns (max_width, max_height) in points.
        """
        max_width = 0
        max_height = 0
        
        for sheet in workbook.Sheets:
            try:
                width = sheet.UsedRange.Width
                height = sheet.UsedRange.Height
                
                if width > max_width:
                    max_width = width
                if height > max_height:
                    max_height = height
                    
                logging.info(f"[{workbook.Name}] {sheet.Name}: Content size {width:.0f}x{height:.0f}pts")
            except Exception as e:
                logging.warning(f"Could not get dimensions for sheet {sheet.Name}: {e}")
                continue
        
        return max_width, max_height

    def _apply_uniform_page_size_mode(self, workbook, page_size="auto"):
        """
        UNIFORM PAGE SIZE mode - find the sheet with largest content width
        and apply that page size to ALL sheets in the workbook.
        
        If page_size is "auto", automatically selects the smallest paper size
        that can fit the widest sheet content.
        """
        logging.info(f"[{workbook.Name}] Applying Uniform Page Size mode")
        
        # Step 1: Find maximum content width across all sheets
        max_width, max_height = self._find_max_content_width(workbook)
        logging.info(f"[{workbook.Name}] Maximum content dimensions: {max_width:.0f}x{max_height:.0f}pts")
        
        # Step 2: Determine page size to use
        page_size_upper = page_size.upper() if page_size else "AUTO"
        if page_size_upper == "AUTO":
            page_size_upper = self._find_best_page_size(max_width, max_height)
            logging.info(f"[{workbook.Name}] Auto-selected page size: {page_size_upper}")
        
        # Get page info
        if page_size_upper in PAGE_SIZES:
            page_info = PAGE_SIZES[page_size_upper]
        else:
            page_info = PAGE_SIZES["A4"]
            page_size_upper = "A4"
        
        logging.info(f"[{workbook.Name}] Applying {page_size_upper} to all sheets")
        
        # Step 3: Apply uniform page size to ALL sheets
        for sheet in workbook.Sheets:
            try:
                # Set paper size with error handling
                try:
                    sheet.PageSetup.PaperSize = page_info["xl_const"]
                except Exception as e:
                    # Fallback to A3
                    sheet.PageSetup.PaperSize = xlPaperA3
                    logging.warning(f"[{workbook.Name}] {sheet.Name}: Paper size '{page_size_upper}' not supported, using A3. For large formats, use C_SHEET, D_SHEET, or E_SHEET.")
                    # Update page_info to use fallback
                    page_info = PAGE_SIZES["A3"]
                
                # Get sheet's own dimensions for orientation
                try:
                    sheet_width = sheet.UsedRange.Width
                    sheet_height = sheet.UsedRange.Height
                except:
                    sheet_width = 0
                    sheet_height = 0
                
                # Set orientation based on content
                if sheet_width > sheet_height:
                    sheet.PageSetup.Orientation = xlLandscape
                    printable_height = page_info["width"] - 100
                else:
                    sheet.PageSetup.Orientation = xlPortrait
                    printable_height = page_info["printable_height"]
                
                # Fit width to 1 page
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
                
                # Set print area to used range only (removes white space)
                try:
                    sheet.PageSetup.PrintArea = sheet.UsedRange.Address
                except:
                    pass
                
                # Clear existing page breaks
                try:
                    sheet.ResetAllPageBreaks()
                except:
                    pass
                
                # Calculate and insert page breaks based on row heights
                accumulated_height = 0
                page_count = 1
                
                for row in sheet.UsedRange.Rows:
                    try:
                        row_height = row.Height
                        accumulated_height += row_height
                        
                        if accumulated_height > printable_height:
                            sheet.HPageBreaks.Add(Before=row)
                            page_count += 1
                            accumulated_height = row_height
                    except:
                        continue
                
                logging.info(f"[{workbook.Name}] {sheet.Name}: Applied {page_size_upper}, {page_count} pages")
                
            except Exception as e:
                logging.warning(f"Could not apply uniform page size to sheet {sheet.Name}: {e}")

    def _setup_header_footer(self, sheet, workbook_name):
        """
        Setup header and footer for the sheet.
        Header: Sheet name
        Footer: Row range (begin to end) for the current page
        
        Note: Only includes data from Excel file - no date/time.
        
        Excel Header/Footer codes:
        &A - Sheet name
        &P - Current page number
        &N - Total pages
        &F - File name
        """
        try:
            # Get used range info for row tracking
            used_range = sheet.UsedRange
            start_row = used_range.Row
            end_row = start_row + used_range.Rows.Count - 1
            
            # Clear all header/footer sections first (remove any date/time)
            sheet.PageSetup.LeftHeader = ""
            sheet.PageSetup.RightHeader = ""
            sheet.PageSetup.LeftFooter = ""
            sheet.PageSetup.RightFooter = ""
            
            # Center Header: Sheet name only
            # Format: "Sheet: [SheetName]"
            sheet.PageSetup.CenterHeader = f"&\"Arial,Bold\"Sheet: {sheet.Name}"
            
            # Center Footer: Row range and page information only (no date/time)
            # Format: "Rows: [StartRow] - [EndRow] | Page X of Y"
            sheet.PageSetup.CenterFooter = f"&\"Arial\"Rows: {start_row} - {end_row} | Page &P of &N"
            
            logging.info(f"[{workbook_name}] {sheet.Name}: Set header/footer (Rows {start_row}-{end_row})")
            
        except Exception as e:
            logging.warning(f"Could not setup header/footer for {sheet.Name}: {e}")

    def _clear_header_footer(self, sheet, workbook_name):
        """
        Clear all header and footer content from the sheet.
        """
        try:
            sheet.PageSetup.LeftHeader = ""
            sheet.PageSetup.CenterHeader = ""
            sheet.PageSetup.RightHeader = ""
            sheet.PageSetup.LeftFooter = ""
            sheet.PageSetup.CenterFooter = ""
            sheet.PageSetup.RightFooter = ""
            logging.info(f"[{workbook_name}] {sheet.Name}: Cleared header/footer")
        except Exception as e:
            logging.warning(f"Could not clear header/footer for {sheet.Name}: {e}")

    def _set_row_col_headings(self, sheet, workbook_name, enable):
        """
        Enable or disable printing of row numbers (1,2,3...) on left
        and column letters (A,B,C...) on top of the printed page.
        """
        try:
            sheet.PageSetup.PrintHeadings = enable
            if enable:
                logging.info(f"[{workbook_name}] {sheet.Name}: Enabled row/column headings")
            else:
                logging.info(f"[{workbook_name}] {sheet.Name}: Disabled row/column headings")
        except Exception as e:
            logging.warning(f"Could not set row/column headings for {sheet.Name}: {e}")

    def _export_to_pdf(self, workbook, output_path):
        """
        Export workbook to PDF using ExportAsFixedFormat.
        This is the reliable method that works without dialogs.
        """
        logging.info(f"[{workbook.Name}] Exporting to PDF: {output_path}")
        
        try:
            # Ensure output path is absolute and properly formatted
            output_path = str(Path(output_path).resolve())
            
            # Export to PDF using ExportAsFixedFormat
            # IncludeDocProperties=True for RAG/AI Search metadata
            workbook.ExportAsFixedFormat(
                Type=xlTypePDF,
                Filename=output_path,
                Quality=xlQualityStandard,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
            # Verify output file was created
            if Path(output_path).exists() and Path(output_path).stat().st_size > 0:
                file_size = Path(output_path).stat().st_size
                logging.info(f"[{workbook.Name}] PDF export completed: {output_path} ({file_size} bytes)")
            else:
                logging.error(f"[{workbook.Name}] PDF file was not created or is empty: {output_path}")
                raise Exception(f"ExportAsFixedFormat did not create output file: {output_path}")
                
        except Exception as e:
            logging.error(f"PDF export failed: {e}")
            raise

    def _autofit_merged_cells(self, sheet, workbook_name):
        """DISABLED: Preserves original merged cell heights."""
        logging.info(f"[{workbook_name}] {sheet.Name}: Preserving original merged cell layout")

