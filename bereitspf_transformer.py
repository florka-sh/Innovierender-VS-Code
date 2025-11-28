"""
BereitsPF Excel Transformer
Uses the working BereitsPF transform_excel module
"""

import pandas as pd
from pathlib import Path
import tempfile
import shutil

# Import the working transform module
try:
    import transform_excel as bereitspf_transform
    TRANSFORM_AVAILABLE = True
except ImportError:
    TRANSFORM_AVAILABLE = False
    print("Warning: transform_excel module not found")


def transform_excel(source_path, template_path=None, defaults=None):
    """
    Transform Excel file using the BereitsPF logic
    
    Args:
        source_path: Path to source Excel
        template_path: Path to template Excel  
        defaults: Dictionary of default values
        
    Returns:
        List of dictionaries (rows)
    """
    if not TRANSFORM_AVAILABLE:
        raise ImportError("transform_excel module not available")
    
    # Apply defaults if not provided
    if defaults is None:
        defaults = {
            'SATZART': 'D',
            'FIRMA': '9241',
            'SOLL_HABEN': 'S',
            'BUCH_KREIS': 'RE',
            'BUCH_JAHR': '2025',
            'BUCH_MONAT': '11',
            'Bebuchbar': 'Ja',
            'NO_RENAME': True
        }
    
    # Create temp output file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        temp_output = Path(tmp.name)
    
    try:
        # Call the working transform function
        bereitspf_transform.transform(
            template_path=Path(template_path),
            source_path=Path(source_path),
            output_path=temp_output,
            config_path=None,
            defaults=defaults
        )
        
        # Read the generated output
        df_output = pd.read_excel(temp_output)
        
        # Convert to list of dictionaries
        results = df_output.to_dict('records')
        
        return results
        
    except Exception as e:
        print(f"Error in transform: {e}")
        raise
        
    finally:
        # Clean up temp file
        if temp_output.exists():
            temp_output.unlink()

