"""
PDF Exporter Service
Converts DOCX files to PDF using available methods.
Primary: docx2pdf (faster, Windows-compatible)
Fallback: LibreOffice headless CLI
"""

import os
import subprocess
import uuid


def convert_to_pdf(docx_path, output_dir):
    """
    Convert a DOCX file to PDF.
    Tries docx2pdf first, falls back to LibreOffice if unavailable.
    
    Args:
        docx_path: Path to the DOCX file
        output_dir: Directory to save the PDF
    
    Returns:
        Tuple of (success: bool, pdf_path or error_message: str)
    """
    pdf_filename = f"PRINT_{uuid.uuid4().hex[:8]}.pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    
    success, result, method = _convert_with_docx2pdf(docx_path, pdf_path)
    if success:
        return True, result
    
    docx2pdf_error = result
    
    success, result = _convert_with_libreoffice(docx_path, output_dir, pdf_filename)
    if success:
        return True, result
    
    libreoffice_error = result
    return False, f"PDF conversion failed. docx2pdf: {docx2pdf_error}. LibreOffice: {libreoffice_error}"


def _convert_with_docx2pdf(docx_path, pdf_path):
    """
    Convert using docx2pdf library.
    
    Args:
        docx_path: Path to the DOCX file
        pdf_path: Full path for output PDF
    
    Returns:
        Tuple of (success: bool, pdf_path or error_message: str, method: str)
    """
    try:
        from docx2pdf import convert
        
        abs_docx_path = os.path.abspath(docx_path)
        abs_pdf_path = os.path.abspath(pdf_path)
        
        convert(abs_docx_path, abs_pdf_path)
        
        if os.path.exists(abs_pdf_path):
            return True, abs_pdf_path, 'docx2pdf'
        
        return False, "docx2pdf completed but PDF file not found", 'docx2pdf'
    
    except ImportError:
        return False, "docx2pdf library not available", 'docx2pdf'
    except Exception as e:
        error_msg = str(e)
        if "comtypes" in error_msg.lower() or "win32" in error_msg.lower():
            return False, "docx2pdf requires Windows with Microsoft Word installed", 'docx2pdf'
        return False, f"docx2pdf error: {error_msg}", 'docx2pdf'


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
        
        return False, f"conversion failed: {result.stderr}"
    
    except subprocess.TimeoutExpired:
        return False, "conversion timed out"
    except FileNotFoundError:
        return False, "LibreOffice not found"
    except Exception as e:
        return False, f"error: {str(e)}"


def is_pdf_conversion_available():
    """Check if any PDF conversion method is available."""
    try:
        from docx2pdf import convert
        return True
    except ImportError:
        pass
    
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
