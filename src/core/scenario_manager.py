"""
Scenario Manager - Handle multi-folder conversion with different configs

This module allows you to define scenarios where different folders
use different conversion settings, all managed through a single scenario file.
"""

import os
import yaml
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass


@dataclass
class ScenarioGroup:
    """
    Represents a group of folders with shared configuration.
    
    Attributes:
        name: Group name for identification
        folders: List of folder paths to process
        config: Configuration dictionary for this group
        output_dir: Output directory for this group
    """
    name: str
    folders: List[str]
    config: Dict[str, Any]
    output_dir: str
    
    def __repr__(self):
        return f"ScenarioGroup(name={self.name}, folders={len(self.folders)}, output={self.output_dir})"


class ScenarioManager:
    """
    Manages scenario-based conversion with different configs for different folders.
    
    A scenario defines multiple groups, where each group has:
    - One or more input folders
    - A specific configuration
    - An output directory
    
    The manager automatically detects file types and applies the appropriate
    config section (excel, word_options, powerpoint_options) from the group's config.
    """
    
    def __init__(self, scenario_file: str):
        """
        Initialize scenario manager.
        
        Args:
            scenario_file: Path to scenario YAML file
        """
        self.scenario_file = scenario_file
        self.scenario_data = self._load_scenario()
        self.groups = self._parse_groups()
    
    def _load_scenario(self) -> Dict[str, Any]:
        """Load scenario configuration from YAML file."""
        if not os.path.exists(self.scenario_file):
            raise FileNotFoundError(f"Scenario file not found: {self.scenario_file}")
        
        with open(self.scenario_file, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        
        if not data:
            raise ValueError(f"Empty scenario file: {self.scenario_file}")
        
        logging.info(f"Loaded scenario from {self.scenario_file}")
        return data
    
    def _parse_groups(self) -> List[ScenarioGroup]:
        """Parse scenario groups from loaded data."""
        groups = []
        
        scenario_groups = self.scenario_data.get('groups', [])
        if not scenario_groups:
            raise ValueError("No groups defined in scenario file")
        
        for group_data in scenario_groups:
            name = group_data.get('name', 'Unnamed')
            folders = group_data.get('folders', [])
            config_path = group_data.get('config')
            output_dir = group_data.get('output', f'output_{name}')
            
            if not folders:
                logging.warning(f"Group '{name}' has no folders, skipping")
                continue
            
            # Load config for this group
            if config_path:
                config = self._load_config(config_path)
            else:
                # Use inline config if provided
                config = group_data.get('config_inline', {})
            
            # Expand folder paths to absolute
            expanded_folders = [os.path.abspath(f) for f in folders]
            output_dir = os.path.abspath(output_dir)
            
            group = ScenarioGroup(
                name=name,
                folders=expanded_folders,
                config=config,
                output_dir=output_dir
            )
            groups.append(group)
            
            logging.info(f"Parsed group: {group}")
        
        return groups
    
    def _load_config(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from file."""
        if not os.path.exists(config_path):
            logging.warning(f"Config file not found: {config_path}, using defaults")
            return {}
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        return config or {}
    
    def get_group_for_file(self, file_path: str) -> Optional[ScenarioGroup]:
        """
        Find which group a file belongs to based on its parent folder.
        
        Args:
            file_path: Path to the file
            
        Returns:
            ScenarioGroup if file is in a configured folder, None otherwise
        """
        abs_path = os.path.abspath(file_path)
        
        for group in self.groups:
            for folder in group.folders:
                # Check if file is in this folder or subfolder
                try:
                    Path(abs_path).relative_to(folder)
                    return group
                except ValueError:
                    continue
        
        return None
    
    def get_all_files(self, file_extensions: Optional[List[str]] = None) -> List[Tuple[str, ScenarioGroup]]:
        """
        Scan all folders in all groups and collect files with their groups.
        
        Args:
            file_extensions: List of extensions to filter (e.g., ['.xlsx', '.docx'])
                           If None, includes all Office file types
        
        Returns:
            List of (file_path, group) tuples
        """
        if file_extensions is None:
            # Default Office file extensions
            file_extensions = [
                '.xlsx', '.xls', '.xlsm', '.xlsb',  # Excel
                '.docx', '.doc', '.docm', '.dotx', '.dotm',  # Word
                '.pptx', '.ppt', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm'  # PowerPoint
            ]
        
        files_with_groups = []
        
        for group in self.groups:
            for folder in group.folders:
                if not os.path.exists(folder):
                    logging.warning(f"Folder not found: {folder}")
                    continue
                
                # Scan folder recursively
                for root, dirs, files in os.walk(folder):
                    # Skip hidden directories
                    dirs[:] = [d for d in dirs if not d.startswith('.')]
                    
                    for file in files:
                        # Skip temp files
                        if file.startswith('~$'):
                            continue
                        
                        # Check extension
                        ext = os.path.splitext(file)[1].lower()
                        if ext in file_extensions:
                            file_path = os.path.join(root, file)
                            files_with_groups.append((file_path, group))
        
        logging.info(f"Found {len(files_with_groups)} files across {len(self.groups)} groups")
        return files_with_groups
    
    def get_file_type(self, file_path: str) -> Optional[str]:
        """
        Detect file type from extension.
        
        Args:
            file_path: Path to file
            
        Returns:
            'excel', 'word', 'powerpoint', or None
        """
        ext = os.path.splitext(file_path)[1].lower()
        
        excel_exts = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        word_exts = ['.docx', '.doc', '.docm', '.dotx', '.dotm']
        ppt_exts = ['.pptx', '.ppt', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm']
        
        if ext in excel_exts:
            return 'excel'
        elif ext in word_exts:
            return 'word'
        elif ext in ppt_exts:
            return 'powerpoint'
        
        return None
    
    def get_output_path_for_file(self, file_path: str, group: ScenarioGroup) -> str:
        """
        Calculate output path for a file, preserving folder structure within group.
        
        Args:
            file_path: Input file path
            group: ScenarioGroup the file belongs to
            
        Returns:
            Output PDF path
        """
        abs_path = os.path.abspath(file_path)
        
        # Find which folder in the group contains this file
        source_folder = None
        for folder in group.folders:
            try:
                Path(abs_path).relative_to(folder)
                source_folder = folder
                break
            except ValueError:
                continue
        
        if source_folder is None:
            # Fallback: use group output dir directly
            filename = os.path.splitext(os.path.basename(file_path))[0] + '.pdf'
            return os.path.join(group.output_dir, filename)
        
        # Calculate relative path from source folder
        rel_path = os.path.relpath(abs_path, source_folder)
        
        # Change extension to .pdf
        pdf_filename = os.path.splitext(rel_path)[0] + '.pdf'
        
        # Construct output path
        output_path = os.path.join(group.output_dir, pdf_filename)
        
        return output_path
    
    def get_scenario_summary(self) -> str:
        """Get a summary of the scenario configuration."""
        lines = [
            f"Scenario: {self.scenario_data.get('name', 'Unnamed')}",
            f"Description: {self.scenario_data.get('description', 'No description')}",
            f"Groups: {len(self.groups)}",
            ""
        ]
        
        for i, group in enumerate(self.groups, 1):
            lines.append(f"Group {i}: {group.name}")
            lines.append(f"  Folders: {len(group.folders)}")
            for folder in group.folders:
                lines.append(f"    - {folder}")
            lines.append(f"  Output: {group.output_dir}")
            lines.append(f"  Config sections: {', '.join(group.config.keys())}")
            lines.append("")
        
        return "\n".join(lines)


def load_scenario(scenario_file: str) -> ScenarioManager:
    """
    Convenience function to load a scenario.
    
    Args:
        scenario_file: Path to scenario YAML file
        
    Returns:
        ScenarioManager instance
    """
    return ScenarioManager(scenario_file)
