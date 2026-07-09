import io
from typing import Dict, Any, List, Tuple
import fitz  # PyMuPDF
from pdf_extractor.logging_config import logger

class PDFPage:
    def __init__(self, page_num: int, text: str, is_scanned: bool, fitz_page: fitz.Page):
        self.page_num = page_num
        self.text = text
        self.is_scanned = is_scanned
        self.fitz_page = fitz_page

    def render_to_png(self, dpi: int = 150) -> bytes:
        """Render page to PNG bytes using PyMuPDF."""
        pix = self.fitz_page.get_pixmap(dpi=dpi)
        return pix.tobytes("png")

def parse_pdf(pdf_bytes: bytes) -> List[PDFPage]:
    """Parse PDF from bytes and return a list of PDFPage objects."""
    logger.info("Opening PDF from bytes...")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    num_pages = len(doc)
    logger.info(f"PDF opened successfully. Total pages: {num_pages}")
    
    parsed_pages = []
    for i in range(num_pages):
        page = doc.load_page(i)
        text = page.get_text().strip()
        
        # Simple heuristic to determine if page is scanned or mostly empty
        is_scanned = False
        if len(text) < 100:
            is_scanned = True
            logger.info(f"Page {i+1} detected as scanned or empty (extracted char count: {len(text)})")
        
        parsed_pages.append(PDFPage(
            page_num=i + 1,
            text=text,
            is_scanned=is_scanned,
            fitz_page=page
        ))
        
    return parsed_pages

def check_pdf_validity(pdf_bytes: bytes) -> Tuple[bool, str]:
    """Check if the PDF is valid, password protected, or corrupted."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if doc.is_encrypted:
            return False, "The PDF file is password protected."
        if len(doc) == 0:
            return False, "The PDF file contains no pages."
        return True, ""
    except Exception as e:
        logger.error(f"Error validating PDF: {e}")
        return False, f"The PDF file appears corrupted or invalid: {str(e)}"
