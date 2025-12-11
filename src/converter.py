import win32com.client
import win32com, shutil
from pathlib import Path
import pythoncom
import win32process
import os
import logging
from .utils import ensure_dir

# Excel Constants
xlLandscape = 2
xlPortrait = 1

# Paper Size Constants (Excel xlPaperSize enumeration)
xlPaperLetter = 1       # 21.59 x 27.94 cm (8.5 x 11 in)
xlPaperTabloid = 3      # 27.94 x 43.18 cm (11 x 17 in)
xlPaperLegal = 5        # 21.59 x 35.56 cm (8.5 x 14 in)
xlPaperStatement = 6    # 13.97 x 21.59 cm (5.5 x 8.5 in)
xlPaperExecutive = 7    # 18.41 x 26.67 cm (7.25 x 10.5 in)
xlPaperA3 = 8           # 29.7 x 42 cm
xlPaperA4 = 9           # 21 x 29.7 cm
xlPaperA5 = 11          # 14.8 x 21 cm
xlPaperB4 = 12          # 25.7 x 36.4 cm (JIS)
xlPaperB5 = 13          # 18.2 x 25.7 cm (JIS)
xlPaperA2 = 66          # 42 x 59.4 cm (based on extended paper size constants)
xlPaperA1 = 67          # 59.4 x 84.1 cm (based on extended paper size constants)

# Print Mode Constants
PRINT_MODE_AUTO = "auto"
PRINT_MODE_ONE_PAGE = "one_page"
PRINT_MODE_TABLE_ROW_BREAK = "table_row_break"
PRINT_MODE_AUTO_PAGE_SIZE = "auto_page_size"
PRINT_MODE_NATIVE_PRINT = "native_print"
PRINT_MODE_UNIFORM_PAGE_SIZE = "uniform_page_size"

# Page sizes in points (1 inch = 72 points, 1 cm = 28.35 points)
# These are printable area estimates (minus typical margins)
PAGE_SIZES = {
    "LETTER": {"width": 612, "height": 792, "printable_height": 700, "xl_const": xlPaperLetter},
    "TABLOID": {"width": 792, "height": 1224, "printable_height": 1130, "xl_const": xlPaperTabloid},
    "LEGAL": {"width": 612, "height": 1008, "printable_height": 915, "xl_const": xlPaperLegal},
    "STATEMENT": {"width": 396, "height": 612, "printable_height": 520, "xl_const": xlPaperStatement},
    "EXECUTIVE": {"width": 522, "height": 756, "printable_height": 665, "xl_const": xlPaperExecutive},
    "A1": {"width": 1684, "height": 2384, "printable_height": 2290, "xl_const": xlPaperA1},
    "A2": {"width": 1191, "height": 1684, "printable_height": 1590, "xl_const": xlPaperA2},
    "A3": {"width": 842, "height": 1191, "printable_height": 1100, "xl_const": xlPaperA3},
    "A4": {"width": 595, "height": 842, "printable_height": 750, "xl_const": xlPaperA4},
    "A5": {"width": 420, "height": 595, "printable_height": 505, "xl_const": xlPaperA5},
    "B4": {"width": 729, "height": 1032, "printable_height": 940, "xl_const": xlPaperB4},
    "B5": {"width": 516, "height": 729, "printable_height": 640, "xl_const": xlPaperB5}
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
            print_mode = self.config.get('print_options', {}).get('mode', PRINT_MODE_AUTO)
            logging.info(f"[{workbook.Name}] Using print mode: {print_mode}")
            
            self._optimize_layout(workbook, print_mode)
            
            # Ensure output directory exists
            ensure_dir(output_path)
            
            # Use native PrintOut to virtual PDF printer for all conversions
            printer_name = self.config.get('print_options', {}).get('native_printer', 'Microsoft Print to PDF')
            self._print_to_pdf(workbook, output_path, printer_name)
            
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

    def _optimize_layout(self, workbook, print_mode=PRINT_MODE_AUTO):
        """
        Prepares workbook for print - merged logic from optimize_layout and enhance_layout.
        Ensures no content is hidden: expands collapsed groups, fixes row heights, fixes images.
        """
        prepare_for_print = self.config.get('excel', {}).get('prepare_for_print', True)
        print_options = self.config.get('print_options', {})
        page_size = print_options.get('page_size', 'A4').upper()
        rows_per_page = print_options.get('rows_per_page')

        # For native_print mode, skip optimization (preserve original formatting)
        if print_mode == PRINT_MODE_NATIVE_PRINT:
            logging.info(f"[{workbook.Name}] Native print mode - skipping layout optimization")
            return

        for sheet in workbook.Sheets:
            try:
                logging.info(f"[{workbook.Name}] Processing Sheet: {sheet.Name}")
                
                # ========================================
                # STEP 1: EXPAND ALL HIDDEN CONTENT
                # Ensure nothing is collapsed/hidden before print
                # ========================================
                if prepare_for_print:
                    self._expand_all_groups(sheet, workbook.Name)
                    self._unhide_rows_columns(sheet, workbook.Name)
                
                # ========================================
                # STEP 2: FIX SHAPE/IMAGE PLACEMENT
                # Prevent images from shrinking when rows resize
                # ========================================
                self._fix_shape_placement(sheet)

                # ========================================
                # STEP 3: PRINT MODE SPECIFIC SETUP
                # ========================================
                if print_mode == PRINT_MODE_ONE_PAGE:
                    self._apply_one_page_mode(sheet, workbook.Name)
                elif print_mode == PRINT_MODE_TABLE_ROW_BREAK:
                    self._apply_table_row_break_mode(sheet, workbook.Name, rows_per_page)
                elif print_mode == PRINT_MODE_AUTO_PAGE_SIZE:
                    self._apply_auto_page_size_mode(sheet, workbook.Name, page_size)
                elif print_mode == PRINT_MODE_UNIFORM_PAGE_SIZE:
                    # Uniform page size is handled at workbook level, not per-sheet
                    # Skip here - will be applied after all sheets are processed
                    pass
                elif prepare_for_print:
                    # Default AUTO mode
                    self._apply_auto_mode(sheet, workbook.Name)

                # ========================================
                # STEP 4: FIX HIDDEN TEXT IN CELLS
                # Auto-fit row heights to show all text
                # ========================================
                if prepare_for_print:
                    # Auto-fit all rows
                    sheet.UsedRange.Rows.AutoFit()
                    
                    # Handle Merged Cells (AutoFit doesn't work on merged cells)
                    self._autofit_merged_cells(sheet, workbook.Name)
                    
                    logging.info(f"[{workbook.Name}] {sheet.Name}: Adjusted row heights")

                # ========================================
                # STEP 5: APPLY SCALING
                # Match Excel's print scaling options
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

    def _apply_auto_mode(self, sheet, workbook_name):
        """
        Default AUTO mode - fit columns to page, auto orientation.
        """
        try:
            total_width_pts = sheet.UsedRange.Width
            total_height_pts = sheet.UsedRange.Height
        except:
            total_width_pts = 0
            total_height_pts = 0

        # Auto-Orientation based on content dimensions
        if total_width_pts > total_height_pts:
            # Wide content -> Landscape
            if total_width_pts < 900:
                sheet.PageSetup.PaperSize = xlPaperA4
                logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> A4 Landscape")
            else:
                sheet.PageSetup.PaperSize = xlPaperA3
                logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> A3 Landscape")
            sheet.PageSetup.Orientation = xlLandscape
        else:
            # Tall/Square content -> Portrait
            sheet.PageSetup.Orientation = xlPortrait
            sheet.PageSetup.PaperSize = xlPaperA4
            logging.info(f"[{workbook_name}] {sheet.Name}: Auto-Layout -> A4 Portrait")

        # Force Fit to 1 Page Wide (keeps original column proportions)
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False 
        
        # Clear Print Area to ensure entire sheet is printed
        try:
            sheet.PageSetup.PrintArea = ""
        except:
            pass

    def _apply_one_page_mode(self, sheet, workbook_name):
        """
        ONE PAGE mode - fit entire sheet content to a single page.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Applying One Page mode")
        
        # Set paper size and orientation
        try:
            total_width_pts = sheet.UsedRange.Width
            total_height_pts = sheet.UsedRange.Height
        except:
            total_width_pts = 0
            total_height_pts = 0

        if total_width_pts > total_height_pts:
            sheet.PageSetup.Orientation = xlLandscape
        else:
            sheet.PageSetup.Orientation = xlPortrait
        
        # Fit BOTH width and height to 1 page
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = 1
        
        # Clear Print Area
        try:
            sheet.PageSetup.PrintArea = ""
        except:
            pass

    def _apply_table_row_break_mode(self, sheet, workbook_name, rows_per_page=None):
        """
        TABLE ROW BREAK mode - insert page breaks after tables or every N rows.
        """
        logging.info(f"[{workbook_name}] {sheet.Name}: Applying Table Row Break mode")
        
        # Set basic page setup
        sheet.PageSetup.PaperSize = xlPaperA4
        sheet.PageSetup.Orientation = xlPortrait
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
        Insert page breaks every N rows.
        """
        try:
            used_rows = sheet.UsedRange.Rows.Count
            start_row = sheet.UsedRange.Row
            
            for row in range(start_row + rows_per_page, start_row + used_rows, rows_per_page):
                try:
                    sheet.HPageBreaks.Add(Before=sheet.Rows(row))
                except:
                    pass
            
            breaks_count = (used_rows // rows_per_page)
            logging.info(f"[{workbook_name}] {sheet.Name}: Inserted {breaks_count} page breaks (every {rows_per_page} rows)")
        except Exception as e:
            logging.warning(f"Error inserting row-based page breaks: {e}")

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

    def _apply_auto_page_size_mode(self, sheet, workbook_name, page_size="A4"):
        """
        AUTO PAGE SIZE mode - calculate page breaks based on selected page size.
        Supports: auto, letter, tabloid, legal, statement, executive, A1, A2, A3, A4, A5, B4, B5
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
        
        # Set paper size using Excel constant
        sheet.PageSetup.PaperSize = page_info["xl_const"]
        printable_height = page_info["printable_height"]
        
        # Set orientation based on content (dimensions already read above)
        if total_width_pts > total_height_pts:
            sheet.PageSetup.Orientation = xlLandscape
            # For landscape, swap printable height with width
            printable_height = page_info["width"] - 100  # Account for margins
        else:
            sheet.PageSetup.Orientation = xlPortrait
        
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
                # Set paper size
                sheet.PageSetup.PaperSize = page_info["xl_const"]
                
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
                
                # Clear print area to print entire sheet
                try:
                    sheet.PageSetup.PrintArea = ""
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

    def _print_to_pdf(self, workbook, output_path, printer_name):
        """
        NATIVE PRINT mode - use Excel's PrintOut to PDF virtual printer.
        This preserves all original formatting.
        """
        logging.info(f"[{workbook.Name}] Using native print to: {printer_name}")
        
        try:
            # Print entire workbook to PDF
            workbook.PrintOut(
                ActivePrinter=printer_name,
                PrintToFile=True,
                PrToFileName=output_path,
                Collate=True
            )
            logging.info(f"[{workbook.Name}] Native print completed: {output_path}")
        except Exception as e:
            logging.error(f"Native print failed: {e}")
            raise

    def _autofit_merged_cells(self, sheet, workbook_name):
        """
        Manually calculates and sets row height for merged cells with wrapped text.
        This is required because Excel's AutoFit ignores merged cells.
        """
        try:
            used_range = sheet.UsedRange
            # Performance safeguard: skip if too many rows
            if used_range.Rows.Count > 5000:
                logging.info(f"[{workbook_name}] {sheet.Name}: Skipping merged cell autofit (too many rows)")
                return

            merged_autofit_count = 0
            
            for row in used_range.Rows:
                # Calculate max required height for this row
                max_height = row.RowHeight
                row_changed = False
                
                # Check cells in this row
                # Limit columns to scan to avoid processing unused vast columns
                for cell in row.Cells:
                    try:
                        # Defensive property access for robustness against empty or error cells
                        try:
                            val = cell.Value
                            # If value is None or empty string, skip
                            if not val or str(val).strip() == "":
                                continue
                        except:
                            # If we can't get value, skip
                            continue

                        # Check MergeCells and WrapText safely
                        try:
                            is_merged = cell.MergeCells
                            is_wrapped = cell.WrapText
                        except:
                            continue
                            
                        if is_merged and is_wrapped:
                            # Found a candidate
                            merge_area = cell.MergeArea
                            
                            # Only process if this is the top-left cell of the merge range
                            if cell.Address == merge_area.Cells(1, 1).Address:
                                # Verify if it spans multiple rows? 
                                if merge_area.Rows.Count == 1 and merge_area.Columns.Count > 1:
                                    # Horizontal merge. We can simulate autofit.
                                    
                                    # Current Width of merged area
                                    merged_width = 0
                                    for col in merge_area.Columns:
                                        merged_width += col.ColumnWidth
                                                                    
                                    try:
                                        # Total width of the merged area in points?
                                        total_width_pts = 0
                                        # Iterate columns directly for clearer logic
                                        for col in merge_area.Columns:
                                            total_width_pts += col.Width
                                        
                                        # Let's use the scrape pad column (last column)
                                        # Safely find a far-out column
                                        temp_col_idx = sheet.Columns.Count
                                        temp_cell = sheet.Cells(row.Row, temp_col_idx)
                                        
                                        # CRITICAL: Copy the cell to preserve Font, Size, Bold, etc.
                                        cell.Copy(temp_cell)
                                        
                                        if temp_cell.MergeCells:
                                            temp_cell.UnMerge()

                                        # Force WrapText on temp cell to ensure accurate autofit calculation
                                        temp_cell.WrapText = True
                                        
                                        current_width_pts = temp_cell.Width
                                        current_cw = temp_cell.ColumnWidth
                                        
                                        if current_width_pts > 0:
                                            target_cw = (total_width_pts / current_width_pts) * current_cw
                                            # Buffer 0.85 - stricter width to force wrapping earlier
                                            temp_cell.ColumnWidth = target_cw * 0.85
                                            
                                        temp_cell.EntireRow.AutoFit()
                                        
                                        # Add constant buffer (15 points) - generous padding for font diffs
                                        needed_height = temp_cell.RowHeight + 15
                                        
                                        # Clean up
                                        temp_cell.Clear()
                                        temp_cell.ColumnWidth = 8.43 # Default
                                        
                                        # Check if we need to increase height
                                        if needed_height > max_height:
                                            max_height = needed_height
                                            row_changed = True
                                            
                                    except Exception as inner_e:
                                        logging.warning(f"Error calculating height for cell {cell.Address}: {inner_e}")
                                        pass
                    except Exception as e:
                        logging.warning(f"Error processing cell {cell.Address if 'cell' in locals() else 'unknown'}: {e}")

                if row_changed:
                    row.RowHeight = max_height
                    merged_autofit_count += 1
            
            if merged_autofit_count > 0:
                logging.info(f"[{workbook_name}] {sheet.Name}: Adjusted height for {merged_autofit_count} rows with merged cells")
               
        except Exception as e:
            logging.warning(f"Error in _autofit_merged_cells: {e}")
