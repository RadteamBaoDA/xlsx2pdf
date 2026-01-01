"""
Word to PDF Converter
Converts Microsoft Word documents (.docx, .doc) to PDF format.
"""
import win32com.client
import pythoncom
import os
import logging
import shutil
from pathlib import Path
from ...core.base_converter import BaseConverter
from ...core.utils import ensure_dir

# Word Constants
wdExportFormatPDF = 17
wdExportOptimizeForPrint = 0
wdExportQualityStandard = 0
wdExportDocumentContent = 0
wdExportCreateNoBookmarks = 0
wdExportCreateHeadingBookmarks = 1


class WordConverter(BaseConverter):
    """
    Word to PDF converter implementing the BaseConverter interface.
    Handles .docx, .doc files with proper formatting preservation.
    """
    
    @property
    def supported_extensions(self):
        """Return list of supported Word file extensions."""
        return ['.docx', '.doc', '.docm', '.dotx', '.dotm']
    
    def validate_input(self, input_path):
        """Validate that the input file is a Word document."""
        return self.can_convert(input_path) and os.path.exists(input_path)
    
    def __init__(self, config):
        """Initialize Word converter with configuration."""
        super().__init__(config)
        self.word_config = config.get('word_options', {})
        
        # Word-specific settings
        self.create_bookmarks = self.word_config.get('create_bookmarks', True)
        self.optimize_for_print = self.word_config.get('optimize_for_print', True)
        self.include_doc_properties = self.word_config.get('include_doc_properties', True)
        self.keep_form_fields = self.word_config.get('keep_form_fields', True)
    
    def convert(self, input_path, output_path, pid_queue=None):
        """
        Convert a Word document to PDF.
        
        Args:
            input_path: Path to the Word file
            output_path: Path where the PDF should be saved
            pid_queue: Optional queue for sending process ID to parent
            
        Returns:
            bool: True if conversion succeeded
            
        Raises:
            Exception: If conversion fails
        """
        if not self.validate_input(input_path):
            raise ValueError(f"Invalid Word file: {input_path}")
        
        pythoncom.CoInitialize()
        word = None
        doc = None
        
        try:
            # Clear Word COM cache
            gen_path = Path(win32com.__gen_path__)
            shutil.rmtree(gen_path, ignore_errors=True)
            
            # Create Word application instance
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone
            
            # Send PID back to parent if queue provided
            if pid_queue:
                try:
                    import win32process
                    import win32api
                    pid = win32process.GetWindowThreadProcessId(word.Hwnd)[1]
                    pid_queue.put(pid)
                except Exception as e:
                    logging.warning(f"Could not send Word PID: {e}")
            
            # Convert path to absolute path
            abs_input_path = os.path.abspath(input_path)
            abs_output_path = os.path.abspath(output_path)
            
            # Ensure output directory exists
            self._ensure_output_directory(abs_output_path)
            
            # Open the Word document
            logging.info(f"Opening Word document: {abs_input_path}")
            doc = word.Documents.Open(
                abs_input_path,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False
            )
            
            # Determine bookmark creation mode
            bookmark_mode = wdExportCreateHeadingBookmarks if self.create_bookmarks else wdExportCreateNoBookmarks
            
            # Export to PDF
            logging.info(f"Exporting Word document to PDF: {abs_output_path}")
            doc.ExportAsFixedFormat(
                OutputFileName=abs_output_path,
                ExportFormat=wdExportFormatPDF,
                OpenAfterExport=False,
                OptimizeFor=wdExportOptimizeForPrint if self.optimize_for_print else wdExportQualityStandard,
                CreateBookmarks=bookmark_mode,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
                IncludeDocProps=self.include_doc_properties,
                KeepIRM=False
            )
            
            logging.info(f"Successfully converted Word document: {abs_input_path} -> {abs_output_path}")
            return True
            
        except Exception as e:
            logging.error(f"Error converting Word document {input_path}: {e}")
            raise e
            
        finally:
            # Cleanup
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass
            
            if word:
                try:
                    word.Quit()
                except:
                    pass
            
            try:
                pythoncom.CoUninitialize()
            except:
                pass
