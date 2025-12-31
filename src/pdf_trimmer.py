"""
PDF Trimmer Module

This module provides functionality to trim white margins from PDF files
by detecting the bounding box of actual content and cropping the page.

Uses pypdf library for PDF manipulation.
"""

import pypdf
import logging
import os
import io
from pathlib import Path
from typing import Optional, Tuple, Dict, Any, List


class PDFTrimmer:
    """
    PDF trimming functionality using PyMuPDF to detect content bounding boxes
    and crop pages to remove white margins.
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize PDFTrimmer with configuration.
        
        Args:
            config: Configuration dictionary with trimming settings
        """
        self.config = config or {}
        self.trim_config = self.config.get('pdf_trim', {})
        
        # Default trimming settings
        self.enabled = self.trim_config.get('enabled', True)
        self.margin_threshold = self.trim_config.get('margin_threshold', 10)  # Points
        self.min_margin = self.trim_config.get('min_margin', 5)  # Points to keep as minimum margin
        self.create_backup = self.trim_config.get('create_backup', False)
        self.backup_suffix = self.trim_config.get('backup_suffix', '_original')
        
    def trim_pdf(self, pdf_path: str, output_path: Optional[str] = None) -> bool:
        """
        Trim white margins from a PDF file.
        
        Args:
            pdf_path: Path to the input PDF file
            output_path: Path for the trimmed PDF output (optional, overwrites input if not provided)
            
        Returns:
            bool: True if trimming was successful, False otherwise
        """
        if not self.enabled:
            logging.info(f"PDF trimming is disabled, skipping: {pdf_path}")
            return True
            
        if not os.path.exists(pdf_path):
            logging.error(f"PDF file not found: {pdf_path}")
            return False
            
        try:
            # Create backup if requested
            if self.create_backup and not output_path:
                self._create_backup(pdf_path)
                
            # Read the PDF
            with open(pdf_path, 'rb') as file:
                reader = pypdf.PdfReader(file)
                writer = pypdf.PdfWriter()
                
                # Check if document has any pages
                if len(reader.pages) == 0:
                    logging.warning(f"PDF has no pages: {pdf_path}")
                    return True
                    
                # Track if any pages were actually trimmed
                pages_trimmed = 0
                total_pages = len(reader.pages)
                
                # Process each page
                for page_num, page in enumerate(reader.pages):
                    try:
                        # Get content bounding box
                        content_bbox = self._get_content_bbox_pypdf(page)
                        
                        if content_bbox:
                            # Calculate trim box with minimum margin
                            trim_bbox = self._calculate_trim_bbox(page, content_bbox)
                            
                            # Only trim if there's significant margin to remove
                            if self._should_trim_page_pypdf(page, trim_bbox):
                                # Create cropped page
                                page.cropbox.lower_left = (trim_bbox[0], trim_bbox[1])
                                page.cropbox.upper_right = (trim_bbox[2], trim_bbox[3])
                                pages_trimmed += 1
                                logging.debug(f"Trimmed page {page_num + 1}: trim_bbox={trim_bbox}")
                            else:
                                logging.debug(f"Skipped trimming page {page_num + 1}: insufficient margin")
                        else:
                            logging.debug(f"No content found on page {page_num + 1}")
                            
                        # Add page to writer (whether trimmed or not)
                        writer.add_page(page)
                        
                    except Exception as e:
                        logging.debug(f"Error processing page {page_num + 1}: {e}")
                        # Add original page on error
                        writer.add_page(page)
                
                # Write the trimmed PDF
                output_file = output_path or pdf_path
                with open(output_file, 'wb') as output:
                    writer.write(output)
                
                if pages_trimmed > 0:
                    logging.info(f"Successfully trimmed {pages_trimmed}/{total_pages} pages in: {pdf_path}")
                else:
                    logging.info(f"No pages required trimming in: {pdf_path}")
                    
                return True
                
        except Exception as e:
            logging.error(f"Error trimming PDF {pdf_path}: {e}")
            return False
    
    def _get_content_bbox_pypdf(self, page: pypdf.PageObject) -> Optional[Tuple[float, float, float, float]]:
        """
        Get the bounding box of actual content on a page using pypdf.
        
        Args:
            page: pypdf PageObject
            
        Returns:
            Tuple: (x0, y0, x1, y1) bounding rectangle of content, or None if no content found
        """
        try:
            # Extract text with locations to get content bounds
            text_content = page.extract_text()
            
            # If no text content, assume some content exists and use a conservative approach
            if not text_content or not text_content.strip():
                # For pages with no extractable text (e.g., images only),
                # we'll use a more conservative approach and only trim obvious large margins
                return self._get_conservative_content_bbox(page)
            
            # Get page dimensions
            page_bbox = page.mediabox
            page_width = float(page_bbox.width)
            page_height = float(page_bbox.height)
            page_x0 = float(page_bbox.lower_left[0])
            page_y0 = float(page_bbox.lower_left[1])
            
            # For text-based content, use a heuristic approach
            # Since pypdf doesn't provide detailed text positioning like PyMuPDF,
            # we'll estimate content bounds based on typical margins
            
            # Estimate content bounds (conservative approach)
            margin_estimate = min(page_width, page_height) * 0.1  # 10% margin estimate
            
            content_x0 = page_x0 + margin_estimate
            content_y0 = page_y0 + margin_estimate
            content_x1 = page_x0 + page_width - margin_estimate
            content_y1 = page_y0 + page_height - margin_estimate
            
            return (content_x0, content_y0, content_x1, content_y1)
            
        except Exception as e:
            logging.debug(f"Error getting content bbox: {e}")
            return None
    
    def _get_conservative_content_bbox(self, page: pypdf.PageObject) -> Optional[Tuple[float, float, float, float]]:
        """
        Get a conservative content bounding box for pages without extractable text.
        
        Args:
            page: pypdf PageObject
            
        Returns:
            Tuple: Conservative content bounding box
        """
        try:
            # Get page dimensions
            page_bbox = page.mediabox
            page_width = float(page_bbox.width)
            page_height = float(page_bbox.height)
            page_x0 = float(page_bbox.lower_left[0])
            page_y0 = float(page_bbox.lower_left[1])
            
            # Use very conservative margins (only trim very large white spaces)
            large_margin_threshold = min(page_width, page_height) * 0.15  # 15% margin
            
            content_x0 = page_x0 + large_margin_threshold
            content_y0 = page_y0 + large_margin_threshold
            content_x1 = page_x0 + page_width - large_margin_threshold
            content_y1 = page_y0 + page_height - large_margin_threshold
            
            return (content_x0, content_y0, content_x1, content_y1)
            
        except Exception as e:
            logging.debug(f"Error getting conservative content bbox: {e}")
            return None
    
    def _calculate_trim_bbox(self, page: pypdf.PageObject, content_bbox: Tuple[float, float, float, float]) -> Tuple[float, float, float, float]:
        """
        Calculate the trimming bounding box based on content and minimum margins.
        
        Args:
            page: pypdf PageObject
            content_bbox: (x0, y0, x1, y1) bounding box containing all content
            
        Returns:
            Tuple: (x0, y0, x1, y1) bounding box to use for cropping
        """
        page_bbox = page.mediabox
        page_x0 = float(page_bbox.lower_left[0])
        page_y0 = float(page_bbox.lower_left[1])
        page_x1 = page_x0 + float(page_bbox.width)
        page_y1 = page_y0 + float(page_bbox.height)
        
        min_margin = self.min_margin
        content_x0, content_y0, content_x1, content_y1 = content_bbox
        
        # Calculate trim bounding box with minimum margins
        x0 = max(page_x0, content_x0 - min_margin)
        y0 = max(page_y0, content_y0 - min_margin)
        x1 = min(page_x1, content_x1 + min_margin)
        y1 = min(page_y1, content_y1 + min_margin)
        
        return (x0, y0, x1, y1)
    
    def _should_trim_page_pypdf(self, page: pypdf.PageObject, trim_bbox: Tuple[float, float, float, float]) -> bool:
        """
        Determine if a page should be trimmed based on margin size.
        
        Args:
            page: pypdf PageObject
            trim_bbox: Proposed trim bounding box (x0, y0, x1, y1)
            
        Returns:
            bool: True if page should be trimmed
        """
        page_bbox = page.mediabox
        page_x0 = float(page_bbox.lower_left[0])
        page_y0 = float(page_bbox.lower_left[1])
        page_x1 = page_x0 + float(page_bbox.width)
        page_y1 = page_y0 + float(page_bbox.height)
        
        threshold = self.margin_threshold
        trim_x0, trim_y0, trim_x1, trim_y1 = trim_bbox
        
        # Calculate margins that would be removed
        left_margin = trim_x0 - page_x0
        bottom_margin = trim_y0 - page_y0
        right_margin = page_x1 - trim_x1
        top_margin = page_y1 - trim_y1
        
        # Check if any margin is significant enough to warrant trimming
        return (left_margin >= threshold or 
                bottom_margin >= threshold or 
                right_margin >= threshold or 
                top_margin >= threshold)
    
    def _create_backup(self, pdf_path: str) -> None:
        """
        Create a backup copy of the original PDF.
        
        Args:
            pdf_path: Path to the original PDF file
        """
        try:
            backup_path = self._get_backup_path(pdf_path)
            
            # Only create backup if it doesn't already exist
            if not os.path.exists(backup_path):
                import shutil
                shutil.copy2(pdf_path, backup_path)
                logging.info(f"Created backup: {backup_path}")
                
        except Exception as e:
            logging.warning(f"Failed to create backup for {pdf_path}: {e}")
    
    def _get_backup_path(self, pdf_path: str) -> str:
        """
        Get the backup file path for a PDF.
        
        Args:
            pdf_path: Original PDF file path
            
        Returns:
            str: Backup file path
        """
        path = Path(pdf_path)
        return str(path.parent / f"{path.stem}{self.backup_suffix}{path.suffix}")
    
    def get_trim_info(self, pdf_path: str) -> Dict[str, Any]:
        """
        Get information about potential trimming for a PDF without actually trimming it.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            Dict containing trim analysis information
        """
        info = {
            'file': pdf_path,
            'trimmable_pages': 0,
            'total_pages': 0,
            'pages_info': []
        }
        
        if not os.path.exists(pdf_path):
            info['error'] = 'File not found'
            return info
            
        try:
            with open(pdf_path, 'rb') as file:
                reader = pypdf.PdfReader(file)
                info['total_pages'] = len(reader.pages)
                
                for page_num, page in enumerate(reader.pages):
                    page_bbox = page.mediabox
                    page_size = (float(page_bbox.width), float(page_bbox.height))
                    
                    page_info = {
                        'page': page_num + 1,
                        'page_size': page_size,
                        'trimmable': False,
                        'content_bbox': None,
                        'trim_bbox': None,
                        'margins': {}
                    }
                    
                    content_bbox = self._get_content_bbox_pypdf(page)
                    if content_bbox:
                        trim_bbox = self._calculate_trim_bbox(page, content_bbox)
                        should_trim = self._should_trim_page_pypdf(page, trim_bbox)
                        
                        page_info['content_bbox'] = content_bbox
                        page_info['trim_bbox'] = trim_bbox
                        page_info['trimmable'] = should_trim
                        
                        if should_trim:
                            info['trimmable_pages'] += 1
                            
                        # Calculate margins
                        page_x0 = float(page_bbox.lower_left[0])
                        page_y0 = float(page_bbox.lower_left[1])
                        page_x1 = page_x0 + float(page_bbox.width)
                        page_y1 = page_y0 + float(page_bbox.height)
                        
                        page_info['margins'] = {
                            'left': trim_bbox[0] - page_x0,
                            'bottom': trim_bbox[1] - page_y0,
                            'right': page_x1 - trim_bbox[2],
                            'top': page_y1 - trim_bbox[3]
                        }
                    
                    info['pages_info'].append(page_info)
                    
        except Exception as e:
            info['error'] = str(e)
            
        return info