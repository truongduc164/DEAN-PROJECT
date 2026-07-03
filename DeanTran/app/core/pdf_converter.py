from __future__ import annotations
import sys
from pathlib import Path
import logging

logger = logging.getLogger("DeanTran.pdf_converter")

def pdf_to_word(pdf_path: str | Path, docx_path: str | Path, log_fn=None):
    """Convert PDF file to Word (.docx) using pdf2docx."""
    from pdf2docx import Converter
    
    msg = f"Chuyển đổi PDF sang Word: {Path(pdf_path).name}"
    logger.info(msg)
    if log_fn:
        log_fn("INFO", msg)
        
    cv = Converter(str(pdf_path))
    cv.convert(str(docx_path), start=0, end=None)
    cv.close()
    
    msg_done = f"Đã chuyển đổi thành công sang Word: {Path(docx_path).name}"
    logger.info(msg_done)
    if log_fn:
        log_fn("INFO", msg_done)

def pdf_to_excel(pdf_path: str | Path, xlsx_path: str | Path, log_fn=None):
    """Convert PDF file to Excel (.xlsx) extracting tables with pdfplumber."""
    import pdfplumber
    from openpyxl import Workbook
    
    msg = f"Chuyển đổi PDF sang Excel: {Path(pdf_path).name}"
    logger.info(msg)
    if log_fn:
        log_fn("INFO", msg)
        
    wb = Workbook()
    default_sheet = wb.active
    if default_sheet is not None:
        wb.remove(default_sheet)
        
    with pdfplumber.open(pdf_path) as pdf:
        for idx, page in enumerate(pdf.pages, 1):
            sheet = wb.create_sheet(title=f"Page {idx}")
            tables = page.extract_tables()
            if tables:
                msg_tables = f"Trích xuất {len(tables)} bảng từ trang {idx}..."
                logger.info(msg_tables)
                if log_fn:
                    log_fn("INFO", msg_tables)
                row_idx = 1
                for table in tables:
                    for row in table:
                        for col_idx, val in enumerate(row, 1):
                            sheet.cell(row=row_idx, column=col_idx, value=val)
                        row_idx += 1
                    row_idx += 1  # Để trống 1 dòng giữa các bảng
            else:
                text = page.extract_text()
                if text:
                    msg_text = f"Trích xuất văn bản dòng-theo-dòng từ trang {idx}..."
                    logger.info(msg_text)
                    if log_fn:
                        log_fn("INFO", msg_text)
                    for row_idx, line in enumerate(text.split("\n"), 1):
                        sheet.cell(row=row_idx, column=1, value=line)
                        
    wb.save(xlsx_path)
    
    msg_done = f"Đã chuyển đổi thành công sang Excel: {Path(xlsx_path).name}"
    logger.info(msg_done)
    if log_fn:
        log_fn("INFO", msg_done)

def pdf_to_ppt(pdf_path: str | Path, pptx_path: str | Path, log_fn=None):
    """Convert PDF file to PowerPoint (.pptx) rendering pages as slide images with PyMuPDF."""
    import fitz
    from pptx import Presentation
    from pptx.util import Inches
    from io import BytesIO
    
    msg = f"Chuyển đổi PDF sang PowerPoint: {Path(pdf_path).name}"
    logger.info(msg)
    if log_fn:
        log_fn("INFO", msg)
        
    prs = Presentation()
    doc = fitz.open(pdf_path)
    
    for page_idx in range(len(doc)):
        msg_page = f"Đang dựng slide cho trang {page_idx + 1}/{len(doc)}..."
        logger.info(msg_page)
        if log_fn:
            log_fn("INFO", msg_page)
            
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        
        width_inch = page.rect.width / 72.0
        height_inch = page.rect.height / 72.0
        
        if page_idx == 0:
            prs.slide_width = Inches(width_inch)
            prs.slide_height = Inches(height_inch)
            
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        img_stream = BytesIO(img_data)
        slide.shapes.add_picture(img_stream, Inches(0), Inches(0), Inches(width_inch), Inches(height_inch))
        
    prs.save(pptx_path)
    
    msg_done = f"Đã chuyển đổi thành công sang PowerPoint: {Path(pptx_path).name}"
    logger.info(msg_done)
    if log_fn:
        log_fn("INFO", msg_done)
