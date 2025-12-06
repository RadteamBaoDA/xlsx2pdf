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

            # Optimize Layout
            self._optimize_layout(workbook)
            
            # Ensure output directory exists
            ensure_dir(output_path)
            
            # Export to PDF
            # IncludeDocProperties=True for RAG/AI Search metadata
            workbook.ExportAsFixedFormat(
                Type=xlTypePDF,
                Filename=output_path,
                Quality=xlQualityStandard,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
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

    def _optimize_layout(self, workbook):
        """
        Optimizes page setup for all sheets.
        """
        optimize = self.config.get('excel', {}).get('optimize_layout', True)
        enhance = self.config.get('excel', {}).get('enhance_layout', False)

        if not optimize and not enhance:
            return

        auto_orientation = self.config.get('excel', {}).get('auto_orientation', True)

        for sheet in workbook.Sheets:
            try:
                logging.info(f"[{workbook.Name}] Processing Sheet: {sheet.Name}")
                
                # Fix shape placement before any layout changes
                self._fix_shape_placement(sheet)

                # ========================================
                # PAGE SETUP (Always runs if optimize_layout: true)
                # This only changes print settings, NOT cell content
                # ========================================
                if optimize:
                    # Print-Ready Logic: Fit All Columns to 1 Page
                    try:
                        total_width_pts = sheet.UsedRange.Width
                        total_height_pts = sheet.UsedRange.Height
                    except:
                        total_width_pts = 0
                        total_height_pts = 0

                    # Constants
                    xlPortrait = 1
                    xlLandscape = 2
                    xlPaperA4 = 9
                    xlPaperA3 = 8
                    
                    # Auto-Orientation based on content dimensions
                    if total_width_pts > total_height_pts:
                        # Wide content -> Landscape
                        if total_width_pts < 900:
                            sheet.PageSetup.PaperSize = xlPaperA4
                            logging.info(f"[{workbook.Name}] {sheet.Name}: Auto-Layout -> A4 Landscape")
                        else:
                            sheet.PageSetup.PaperSize = xlPaperA3
                            logging.info(f"[{workbook.Name}] {sheet.Name}: Auto-Layout -> A3 Landscape")
                        sheet.PageSetup.Orientation = xlLandscape
                    else:
                        # Tall/Square content -> Portrait
                        sheet.PageSetup.Orientation = xlPortrait
                        sheet.PageSetup.PaperSize = xlPaperA4
                        logging.info(f"[{workbook.Name}] {sheet.Name}: Auto-Layout -> A4 Portrait")

                    # Force Fit to 1 Page Wide (keeps original column proportions)
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = False 
                    
                    # Clear Print Area to ensure entire sheet is printed
                    try:
                        sheet.PageSetup.PrintArea = ""
                    except:
                        pass

                # ========================================
                # ROW HEIGHT ADJUSTMENT (Fix hidden text)
                # Adjusts row heights only, preserves column widths
                # ========================================
                if enhance:
                    # 1. Standard Row AutoFit (all rows)
                    sheet.UsedRange.Rows.AutoFit()
                    
                    # 2. Handle Merged Cells (AutoFit doesn't work on merged cells)
                    self._autofit_merged_cells(sheet, workbook.Name)
                    
                    logging.info(f"[{workbook.Name}] {sheet.Name}: Adjusted row heights")

            except Exception as e:
                logging.warning(f"Could not optimize/enhance sheet {sheet.Name}: {e}")

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
