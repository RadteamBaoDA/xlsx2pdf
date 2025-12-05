import win32com.client
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

    def _optimize_layout(self, workbook):
        """
        Optimizes page setup for all sheets.
        """
        optimize = self.config.get('excel', {}).get('optimize_layout', True)
        enhance = self.config.get('excel', {}).get('enhance_layout', False)

        if not optimize and not enhance:
            return

        auto_orientation = self.config.get('excel', {}).get('auto_orientation', True)

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
                if enhance:
                    # 1. Global: Ensure wrap text is respected by fitting row heights
                    # We do NOT force WrapText=True everywhere anymore.
                    # Just AutoFit rows to ensure wrapped content is visible.
                    sheet.UsedRange.Rows.AutoFit()
                    logging.info(f"[{workbook.Name}] {sheet.Name}: Auto-fitted rows")
                    
                    # 2. Smart Column AutoFit
                    # "Enhance if cell is not set wrap text, resize column size fit with cell content."
                    # We iterate columns. If a column has ANY cell with WrapText=False, we want to AutoFit it.
                    # BUT we must treat Mixed columns carefully to avoid expanding based on Wrapped text.
                    
                    used_range = sheet.UsedRange
                    # Limit to reasonable size to prevent hanging on massive sheets
                    autofit_count = 0
                    if used_range.Columns.Count < 200: 
                        for col in used_range.Columns:
                            try:
                                wrap_status = col.WrapText
                                if wrap_status is False:
                                    # All cells No-Wrap -> Safe to AutoFit
                                    col.AutoFit()
                                    autofit_count += 1
                                elif wrap_status is True:
                                    # All cells Wrap -> Do NOT AutoFit Column (preserve width), Row AutoFit covers it.
                                    pass 
                                else:
                                    # Mixed (None). Some wrapped, some not.
                                    # We want to AutoFit based on the NON-WRAPPED cells only.
                                    # Iterate cells to find No-Wrap ones.
                                    # Optimization: If too many rows, fallback to safe strategy (don't autofit?)
                                    if used_range.Rows.Count > 1000:
                                        # Too big to iterate, skip AutoFit to be safe
                                        pass
                                    else:
                                        # Create a union of No-Wrap cells
                                        excel = workbook.Application
                                        no_wrap_cells = None
                                        
                                        # 1-based index for specific ranges is safer or simple iteration
                                        # Using Values is fast, but properties are slow.
                                        # We have to check WrapText property.
                                        
                                        count = 0
                                        # Iterate cells in this column only
                                        for cell in col.Cells:
                                            # Stop if outside used range
                                            if cell.Row > used_range.Row + used_range.Rows.Count - 1:
                                                break
                                                
                                            if not cell.WrapText:
                                                if no_wrap_cells is None:
                                                    no_wrap_cells = cell
                                                else:
                                                    no_wrap_cells = excel.Union(no_wrap_cells, cell)
                                                count += 1
                                        
                                        if no_wrap_cells:
                                            no_wrap_cells.EntireColumn.AutoFit()
                                            autofit_count += 1
                            except Exception as e:
                                pass # Ignore column errors
                    
                    if autofit_count > 0:
                        logging.info(f"[{workbook.Name}] {sheet.Name}: Enhanced width for {autofit_count} columns")
                    
                    # 3. Re-apply Row AutoFit in case Column AutoFit changed line breaks
                    sheet.UsedRange.Rows.AutoFit()
                    
                if optimize:
                    # Page Setup: Fit to 1 page wide, auto tall
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = False 
                    
                    # Auto Orientation
                    if auto_orientation:
                        try:
                            # Use current UsedRange dimensions
                            width = sheet.UsedRange.Width
                            height = sheet.UsedRange.Height
                            if width > height:
                                sheet.PageSetup.Orientation = xlLandscape
                                logging.info(f"[{workbook.Name}] {sheet.Name}: Set orientation to Landscape")
                            else:
                                sheet.PageSetup.Orientation = xlPortrait
                                logging.info(f"[{workbook.Name}] {sheet.Name}: Set orientation to Portrait")
                        except:
                            pass
            except Exception as e:
                logging.warning(f"Could not optimize/enhance sheet {sheet.Name}: {e}")
