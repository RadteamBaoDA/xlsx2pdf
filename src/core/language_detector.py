"""
Language detection module for Excel files.
Detects language from cell content and classifies files for distribution.
"""
import logging
from pathlib import Path

# Try to import langdetect (optional dependency)
try:
    from langdetect import detect, DetectorFactory, LangDetectException
    # Set seed for consistent results
    DetectorFactory.seed = 0
    LANGDETECT_AVAILABLE = True
except ImportError:
    LANGDETECT_AVAILABLE = False
    logging.warning("langdetect not installed. Auto language detection will be limited. Install with: pip install langdetect")


class LanguageDetector:
    """
    Detects language from Excel file content or filename.
    """
    
    def __init__(self, config):
        self.config = config
        self.lang_config = config.get('language_classification', {})
        self.enabled = self.lang_config.get('enabled', False)
        self.mode = self.lang_config.get('mode', 'auto').lower()
        self.filename_patterns = self.lang_config.get('filename_patterns', {})
        self.keep_structure = self.lang_config.get('keep_folder_structure', True)
        self.output_format = self.lang_config.get('output_suffix_format', 'output-{lang}')
        
        # Language code mapping
        self.lang_map = {
            'vi': 'vi',
            'en': 'en', 
            'ja': 'ja',
            'zh': 'zh',
            'ko': 'ko',
            'th': 'th',
            'fr': 'fr',
            'de': 'de',
            'es': 'es',
            'other': 'other'
        }
    
    def is_enabled(self):
        """Check if language classification is enabled."""
        return self.enabled
    
    def detect_language_from_filename(self, filename):
        """
        Detect language based on filename patterns.
        
        Args:
            filename: Name of the file
            
        Returns:
            Language code (e.g., 'vi', 'en', 'ja') or 'other'
        """
        if not self.filename_patterns:
            return 'other'
        
        # Check each language pattern
        for lang_code, patterns in self.filename_patterns.items():
            for pattern in patterns:
                if pattern == "":
                    # Empty pattern - check if filename has no language suffix
                    # This is a fallback for files without any pattern
                    continue
                elif pattern in filename:
                    logging.info(f"Detected language '{lang_code}' from filename pattern '{pattern}' in '{filename}'")
                    return lang_code
        
        # Check for empty pattern (files without specific suffix)
        for lang_code, patterns in self.filename_patterns.items():
            if "" in patterns:
                # This language accepts files with no specific pattern
                has_other_pattern = False
                for other_lang, other_patterns in self.filename_patterns.items():
                    if other_lang == lang_code:
                        continue
                    for other_pattern in other_patterns:
                        if other_pattern != "" and other_pattern in filename:
                            has_other_pattern = True
                            break
                    if has_other_pattern:
                        break
                
                if not has_other_pattern:
                    logging.info(f"Detected language '{lang_code}' from filename (no specific pattern) '{filename}'")
                    return lang_code
        
        return 'other'
    
    def detect_language_from_content(self, workbook):
        """
        Detect language from workbook cell content.
        Samples text from all sheets and uses language detection.
        
        Args:
            workbook: Excel workbook COM object
            
        Returns:
            Language code (e.g., 'vi', 'en', 'ja') or 'other'
        """
        if not LANGDETECT_AVAILABLE:
            logging.warning("langdetect not available, cannot detect language from content")
            return 'other'
        
        try:
            # Collect sample text from all sheets
            sample_texts = []
            max_samples = 100  # Limit samples for performance
            sample_count = 0
            
            for sheet in workbook.Sheets:
                try:
                    used_range = sheet.UsedRange
                    
                    # Sample cells from the sheet
                    for row in used_range.Rows:
                        if sample_count >= max_samples:
                            break
                        
                        for cell in row.Cells:
                            if sample_count >= max_samples:
                                break
                            
                            try:
                                value = cell.Value
                                if value and isinstance(value, str):
                                    text = str(value).strip()
                                    # Only include text with actual words (not just numbers/symbols)
                                    if len(text) > 3 and any(c.isalpha() for c in text):
                                        sample_texts.append(text)
                                        sample_count += 1
                            except:
                                continue
                        
                        if sample_count >= max_samples:
                            break
                        
                except Exception as e:
                    logging.warning(f"Could not sample from sheet {sheet.Name}: {e}")
                    continue
            
            if not sample_texts:
                logging.warning("No text found in workbook for language detection")
                return 'other'
            
            # Combine samples and detect language
            combined_text = " ".join(sample_texts[:50])  # Use first 50 samples
            
            try:
                detected_lang = detect(combined_text)
                # Map to our language codes
                lang_code = self.lang_map.get(detected_lang, 'other')
                logging.info(f"Detected language '{lang_code}' from content (langdetect: {detected_lang})")
                return lang_code
            except LangDetectException as e:
                logging.warning(f"Language detection failed: {e}")
                return 'other'
                
        except Exception as e:
            logging.error(f"Error detecting language from content: {e}")
            return 'other'
    
    def classify_file(self, input_path, workbook=None):
        """
        Classify file to determine target language.
        
        Args:
            input_path: Path to input file
            workbook: Excel workbook COM object (optional, for content detection)
            
        Returns:
            Language code (e.g., 'vi', 'en', 'ja', 'other')
        """
        if not self.enabled:
            return 'other'
        
        filename = Path(input_path).stem
        
        if self.mode == 'filename':
            # Classify based on filename pattern
            return self.detect_language_from_filename(filename)
        
        elif self.mode == 'auto':
            # Classify based on content detection
            if workbook:
                return self.detect_language_from_content(workbook)
            else:
                logging.warning("Workbook not provided for auto language detection, falling back to filename")
                return self.detect_language_from_filename(filename)
        
        else:
            logging.warning(f"Unknown language classification mode: {self.mode}")
            return 'other'
    
    def get_output_path(self, input_path, base_output_dir, language_code):
        """
        Get the output path for a file based on its language classification.
        Maintains folder structure if configured.
        
        Args:
            input_path: Original input file path
            base_output_dir: Base output directory from config
            language_code: Detected language code
            
        Returns:
            Full output path for the PDF file
        """
        input_file = Path(input_path)
        base_output = Path(base_output_dir)
        
        # Determine output folder based on language
        if language_code and language_code != 'other':
            output_dir = Path(self.output_format.replace('{lang}', language_code))
        else:
            output_dir = base_output
        
        # Get relative path from input directory if maintaining structure
        if self.keep_structure and len(input_file.parents) > 1:
            # Try to find common base between input and output
            # Assume input files are in a subfolder of the workspace
            try:
                # Get the parent directories of the input file
                # We want to preserve the structure relative to the input root
                rel_parts = []
                current = input_file.parent
                
                # Walk up until we hit a known root or run out of parents
                # For now, just use the immediate parent structure
                if current.name and current.name not in ['input', 'Input', '.']:
                    rel_parts.insert(0, current.name)
                
                if rel_parts:
                    output_dir = output_dir / Path(*rel_parts)
            except:
                pass
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate output filename
        output_suffix = self.config.get('output_suffix', '_x')
        output_filename = f"{input_file.stem}{output_suffix}.pdf"
        
        return str(output_dir / output_filename)
    
    def get_output_path_with_suffix(self, input_path, output_root, lang_code, suffix='_x'):
        """
        Generates output path for a given input file, language code, and custom suffix.
        
        Args:
            input_path: Path to input file
            output_root: Root output directory (will be adjusted based on language)
            lang_code: Detected language code
            suffix: File type suffix (e.g., '_x' for Excel, '_d' for Word, '_p' for PowerPoint)
            
        Returns:
            Output file path in language-specific directory
        """
        input_file = Path(input_path)
        
        # Build language-specific output folder
        base_output = Path(output_root) / self.output_format.format(lang=lang_code)
        
        # Determine output directory
        if not self.keep_structure:
            output_dir = base_output
        else:
            output_dir = base_output
        
        # Get relative path from input directory if maintaining structure
        if self.keep_structure and len(input_file.parents) > 1:
            try:
                rel_parts = []
                current = input_file.parent
                
                if current.name and current.name not in ['input', 'Input', '.']:
                    rel_parts.insert(0, current.name)
                
                if rel_parts:
                    output_dir = output_dir / Path(*rel_parts)
            except:
                pass
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate output filename with provided suffix
        output_filename = f"{input_file.stem}{suffix}.pdf"
        
        return str(output_dir / output_filename)
