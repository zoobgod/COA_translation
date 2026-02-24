"""
PDF text extraction module.

Uses multiple strategies to extract text from pharmaceutical COA PDFs:
1. pdfplumber (primary) - good for structured/tabular PDFs
2. PyMuPDF/fitz (fallback) - good for general text extraction
3. pytesseract OCR (last resort) - for scanned/image-based PDFs
"""

import io
import logging
from typing import Optional

import pdfplumber

try:
    import fitz  # PyMuPDF

    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

logger = logging.getLogger(__name__)


def extract_with_pdfplumber(pdf_bytes: bytes) -> tuple[Optional[str], int]:
    """
    Extract text using pdfplumber. Works well with structured PDFs
    that contain tables (common in COA documents).

    Returns (text, page_count).
    """
    try:
        text_parts = []
        page_count = 0
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page_count = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(f"--- Page {i + 1} ---\n{page_text}")

                # Also try to extract tables
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        table_text = _format_table(table)
                        if table_text:
                            text_parts.append(table_text)

        result = "\n\n".join(text_parts).strip()
        return (result if result else None), page_count
    except Exception as e:
        logger.warning(f"pdfplumber extraction failed: {e}")
        return None, 0


def extract_with_pymupdf(pdf_bytes: bytes) -> tuple[Optional[str], int]:
    """
    Extract text using PyMuPDF (fitz). Good general-purpose extraction.

    Returns (text, page_count).
    """
    if not HAS_FITZ:
        logger.warning("PyMuPDF (fitz) not available")
        return None, 0

    try:
        text_parts = []
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_count = len(doc)
        for i, page in enumerate(doc):
            page_text = page.get_text()
            if page_text.strip():
                text_parts.append(f"--- Page {i + 1} ---\n{page_text.strip()}")
        doc.close()

        result = "\n\n".join(text_parts).strip()
        return (result if result else None), page_count
    except Exception as e:
        logger.warning(f"PyMuPDF extraction failed: {e}")
        return None, 0


def extract_with_ocr(pdf_bytes: bytes) -> tuple[Optional[str], int]:
    """
    Extract text using OCR (pytesseract) for scanned/image-based PDFs.
    Converts PDF pages to images first, then runs OCR.

    Returns (text, page_count).
    """
    if not HAS_FITZ:
        logger.warning("PyMuPDF (fitz) not available; OCR requires it for rendering")
        return None, 0

    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        logger.warning("pytesseract or Pillow not installed; OCR unavailable")
        return None, 0

    try:
        text_parts = []
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_count = len(doc)
        for i, page in enumerate(doc):
            # Render page to image at 300 DPI
            mat = fitz.Matrix(300 / 72, 300 / 72)
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_bytes))

            # Run OCR - try English first, can also attempt Russian
            page_text = pytesseract.image_to_string(image, lang="eng")
            if page_text.strip():
                text_parts.append(f"--- Page {i + 1} (OCR) ---\n{page_text.strip()}")
        doc.close()

        result = "\n\n".join(text_parts).strip()
        return (result if result else None), page_count
    except Exception as e:
        logger.warning(f"OCR extraction failed: {e}")
        return None, 0


def extract_text_from_pdf(pdf_bytes: bytes) -> dict:
    """
    Main extraction function. Tries multiple strategies and returns
    the best result.

    Returns:
        dict with keys:
            - 'text': extracted text string
            - 'method': which extraction method was used
            - 'success': whether extraction succeeded
            - 'page_count': number of pages in the PDF
    """
    # Strategy 1: pdfplumber (best for tables/structured COAs)
    text, page_count = extract_with_pdfplumber(pdf_bytes)
    if text and len(text.strip()) > 50:
        return {
            "text": text,
            "method": "pdfplumber",
            "success": True,
            "page_count": page_count,
        }

    # Strategy 2: PyMuPDF
    text, pc = extract_with_pymupdf(pdf_bytes)
    if pc > 0:
        page_count = pc
    if text and len(text.strip()) > 50:
        return {
            "text": text,
            "method": "PyMuPDF",
            "success": True,
            "page_count": page_count,
        }

    # Strategy 3: OCR (for scanned documents)
    text, pc = extract_with_ocr(pdf_bytes)
    if pc > 0:
        page_count = pc
    if text and len(text.strip()) > 20:
        return {
            "text": text,
            "method": "OCR (pytesseract)",
            "success": True,
            "page_count": page_count,
        }

    # All methods failed
    return {
        "text": "",
        "method": "none",
        "success": False,
        "page_count": page_count,
    }


def _format_table(table: list) -> str:
    """Format an extracted table as readable text."""
    if not table:
        return ""

    rows = []
    for row in table:
        if row:
            cells = [str(cell).strip() if cell else "" for cell in row]
            rows.append(" | ".join(cells))

    return "\n".join(rows) if rows else ""
