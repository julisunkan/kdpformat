"""
PDF Exporter Service
Converts DOCX files to PDF using available methods.
"""

import os
import subprocess
import uuid


def convert_to_pdf(docx_path, output_dir):
    """
    Convert a DOCX file to PDF.
    
    Args:
        docx_path: Path to the DOCX file
        output_dir: Directory to save the PDF
    
    Returns:
        Tuple of (success: bool, pdf_path or error_message: str)
    """
    pdf_filename = f"PRINT_{uuid.uuid4().hex[:8]}.pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    
    success, result = _convert_with_libreoffice(docx_path, output_dir, pdf_filename)
    if success:
        return True, result
    
    return False, "PDF conversion failed. LibreOffice conversion was unsuccessful."


def _convert_with_libreoffice(docx_path, output_dir, pdf_filename):
    """
    Convert using LibreOffice headless mode.
    
    Args:
        docx_path: Path to the DOCX file
        output_dir: Directory to save the PDF
        pdf_filename: Desired PDF filename
    
    Returns:
        Tuple of (success: bool, pdf_path or error_message: str)
    """
    try:
        abs_docx_path = os.path.abspath(docx_path)
        abs_output_dir = os.path.abspath(output_dir)
        
        cmd = [
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', abs_output_dir,
            abs_docx_path
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0:
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            generated_pdf = os.path.join(abs_output_dir, f"{base_name}.pdf")
            
            if os.path.exists(generated_pdf):
                final_path = os.path.join(abs_output_dir, pdf_filename)
                os.rename(generated_pdf, final_path)
                return True, final_path
            
            for f in os.listdir(abs_output_dir):
                if f.endswith('.pdf') and base_name in f:
                    generated_pdf = os.path.join(abs_output_dir, f)
                    final_path = os.path.join(abs_output_dir, pdf_filename)
                    os.rename(generated_pdf, final_path)
                    return True, final_path
        
        return False, f"LibreOffice conversion failed: {result.stderr}"
    
    except subprocess.TimeoutExpired:
        return False, "LibreOffice conversion timed out"
    except FileNotFoundError:
        return False, "LibreOffice not found"
    except Exception as e:
        return False, f"LibreOffice error: {str(e)}"


def is_pdf_conversion_available():
    """Check if PDF conversion is available."""
    try:
        result = subprocess.run(
            ['soffice', '--version'],
            capture_output=True,
            text=True,
            timeout=10
        )
        return result.returncode == 0
    except:
        return False
