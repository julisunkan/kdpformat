"""
DPI Checker Service
Inspects embedded images in DOCX files and checks for minimum DPI requirements.
"""

import io
import zipfile
from PIL import Image


def check_image_dpi(docx_path, min_dpi=300):
    """
    Check all embedded images in a DOCX file for DPI compliance.
    
    Args:
        docx_path: Path to the DOCX file
        min_dpi: Minimum acceptable DPI (default 300 for print)
    
    Returns:
        List of warning dictionaries with image name and DPI info
    """
    warnings = []
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for file_name in docx_zip.namelist():
                if file_name.startswith('word/media/') and _is_image_file(file_name):
                    try:
                        image_data = docx_zip.read(file_name)
                        image = Image.open(io.BytesIO(image_data))
                        
                        dpi_x, dpi_y = image.info.get('dpi', (72, 72))
                        
                        if isinstance(dpi_x, tuple):
                            dpi_x = dpi_x[0]
                        if isinstance(dpi_y, tuple):
                            dpi_y = dpi_y[0]
                        
                        dpi_x = int(dpi_x) if dpi_x else 72
                        dpi_y = int(dpi_y) if dpi_y else 72
                        
                        avg_dpi = (dpi_x + dpi_y) // 2
                        
                        if avg_dpi < min_dpi:
                            warnings.append({
                                'image': file_name.split('/')[-1],
                                'dpi': avg_dpi,
                                'required': min_dpi,
                                'width': image.width,
                                'height': image.height,
                                'message': f"Image '{file_name.split('/')[-1]}' has {avg_dpi} DPI (minimum {min_dpi} DPI required for print)"
                            })
                        
                        image.close()
                    except Exception as e:
                        warnings.append({
                            'image': file_name.split('/')[-1],
                            'dpi': 'Unknown',
                            'required': min_dpi,
                            'message': f"Could not analyze image '{file_name.split('/')[-1]}': {str(e)}"
                        })
    except Exception as e:
        warnings.append({
            'image': 'N/A',
            'dpi': 'N/A',
            'required': min_dpi,
            'message': f"Error reading DOCX file: {str(e)}"
        })
    
    return warnings


def _is_image_file(filename):
    """Check if file is an image based on extension."""
    image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp')
    return filename.lower().endswith(image_extensions)
