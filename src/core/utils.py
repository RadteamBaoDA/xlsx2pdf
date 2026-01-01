import yaml
import os

def load_config(config_path="config.yaml"):
    """Loads configuration from a YAML file."""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found: {config_path}")
    
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)

def get_output_path(input_path, input_root, output_root, suffix="_x"):
    """
    Determines the output path for a given input file, preserving directory structure.
    """
    # Get relative path from input root
    rel_path = os.path.relpath(input_path, input_root)
    
    # Construct output path
    out_path = os.path.join(output_root, rel_path)
    
    # Change extension and add suffix
    base, _ = os.path.splitext(out_path)
    return f"{base}{suffix}.pdf"

import shutil

def ensure_dir(file_path):
    """Ensures the directory for a file exists."""
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)

def copy_to_enhanced(input_path, input_root, enhanced_root):
    """
    Copies the input file to the enhanced directory, preserving structure.
    Returns the path to the copied file.
    """
    rel_path = os.path.relpath(input_path, input_root)
    enhanced_path = os.path.join(enhanced_root, rel_path)
    
    ensure_dir(enhanced_path)
    shutil.copy2(input_path, enhanced_path)
    
    return enhanced_path
