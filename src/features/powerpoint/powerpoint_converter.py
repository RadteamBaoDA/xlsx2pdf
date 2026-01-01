"""
PowerPoint to PDF Converter
Converts Microsoft PowerPoint presentations (.pptx, .ppt) to PDF format.
"""
import win32com.client
import pythoncom
import os
import logging
import shutil
from pathlib import Path
from ...core.base_converter import BaseConverter
from ...core.utils import ensure_dir

# PowerPoint Constants
ppFixedFormatTypePDF = 2
ppFixedFormatIntentPrint = 2
ppFixedFormatIntentScreen = 1
ppPrintOutputSlides = 1
ppPrintOutputNotesPages = 5
ppPrintOutputHandouts = 4
ppPrintHandoutVerticalFirst = 1
ppPrintHandoutHorizontalFirst = 2
ppQualityStandard = 0


class PowerPointConverter(BaseConverter):
    """
    PowerPoint to PDF converter implementing the BaseConverter interface.
    Handles .pptx, .ppt files with customizable output options.
    """
    
    @property
    def supported_extensions(self):
        """Return list of supported PowerPoint file extensions."""
        return ['.pptx', '.ppt', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm']
    
    def validate_input(self, input_path):
        """Validate that the input file is a PowerPoint presentation."""
        return self.can_convert(input_path) and os.path.exists(input_path)
    
    def __init__(self, config):
        """Initialize PowerPoint converter with configuration."""
        super().__init__(config)
        self.ppt_config = config.get('powerpoint_options', {})
        
        # PowerPoint-specific settings
        self.output_type = self.ppt_config.get('output_type', 'slides')  # 'slides', 'notes', 'handouts'
        self.handout_order = self.ppt_config.get('handout_order', 'vertical')  # 'vertical', 'horizontal'
        self.slides_per_page = self.ppt_config.get('slides_per_page', 1)  # For handouts: 1, 2, 3, 4, 6, 9
        self.include_hidden_slides = self.ppt_config.get('include_hidden_slides', False)
        self.frame_slides = self.ppt_config.get('frame_slides', False)
        self.print_comments = self.ppt_config.get('print_comments', False)
    
    def convert(self, input_path, output_path, pid_queue=None):
        """
        Convert a PowerPoint presentation to PDF.
        
        Args:
            input_path: Path to the PowerPoint file
            output_path: Path where the PDF should be saved
            pid_queue: Optional queue for sending process ID to parent
            
        Returns:
            bool: True if conversion succeeded
            
        Raises:
            Exception: If conversion fails
        """
        if not self.validate_input(input_path):
            raise ValueError(f"Invalid PowerPoint file: {input_path}")
        
        pythoncom.CoInitialize()
        powerpoint = None
        presentation = None
        
        try:
            # Clear PowerPoint COM cache
            gen_path = Path(win32com.__gen_path__)
            shutil.rmtree(gen_path, ignore_errors=True)
            
            # Create PowerPoint application instance
            powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
            powerpoint.Visible = 1  # PowerPoint needs to be visible for some operations
            powerpoint.DisplayAlerts = 0  # ppAlertsNone
            
            # Send PID back to parent if queue provided
            if pid_queue:
                try:
                    import win32process
                    import win32api
                    # PowerPoint doesn't have Hwnd, get by process name
                    import psutil
                    current_process = psutil.Process()
                    for child in current_process.children(recursive=True):
                        if 'POWERPNT.EXE' in child.name().upper():
                            pid_queue.put(child.pid)
                            break
                except Exception as e:
                    logging.warning(f"Could not send PowerPoint PID: {e}")
            
            # Convert path to absolute path
            abs_input_path = os.path.abspath(input_path)
            abs_output_path = os.path.abspath(output_path)
            
            # Ensure output directory exists
            self._ensure_output_directory(abs_output_path)
            
            # Open the PowerPoint presentation
            logging.info(f"Opening PowerPoint presentation: {abs_input_path}")
            presentation = powerpoint.Presentations.Open(
                abs_input_path,
                ReadOnly=1,
                Untitled=0,
                WithWindow=0
            )
            
            # Determine output type
            output_type_map = {
                'slides': ppPrintOutputSlides,
                'notes': ppPrintOutputNotesPages,
                'handouts': ppPrintOutputHandouts
            }
            output_type_value = output_type_map.get(self.output_type, ppPrintOutputSlides)
            
            # Export to PDF
            logging.info(f"Exporting PowerPoint presentation to PDF: {abs_output_path}")
            presentation.ExportAsFixedFormat(
                Path=abs_output_path,
                FixedFormatType=ppFixedFormatTypePDF,
                Intent=ppFixedFormatIntentPrint,
                FrameSlides=1 if self.frame_slides else 0,
                HandoutOrder=ppPrintHandoutVerticalFirst if self.handout_order == 'vertical' else ppPrintHandoutHorizontalFirst,
                OutputType=output_type_value,
                PrintHiddenSlides=1 if self.include_hidden_slides else 0,
                PrintComments=1 if self.print_comments else 0,
                PrintColorType=1,  # ppPrintColor
                RangeType=1,  # ppPrintAll
                IncludeDocProperties=True
            )
            
            logging.info(f"Successfully converted PowerPoint presentation: {abs_input_path} -> {abs_output_path}")
            return True
            
        except Exception as e:
            logging.error(f"Error converting PowerPoint presentation {input_path}: {e}")
            raise e
            
        finally:
            # Cleanup
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
            
            if powerpoint:
                try:
                    powerpoint.Quit()
                except:
                    pass
            
            try:
                pythoncom.CoUninitialize()
            except:
                pass
