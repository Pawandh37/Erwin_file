"""
Extract text from scanned PDF using OCR and create a clean PDF
"""
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime
import os

def create_ocr_pdf(input_pdf, output_pdf=None):
    """
    Extract text from scanned PDF using OCR and create a clean PDF.
    
    Args:
        input_pdf: Path to the original PDF file
        output_pdf: Path to save the new PDF
    """
    if not os.path.exists(input_pdf):
        print(f"Error: PDF file not found at {input_pdf}")
        return
    
    # Generate output path if not provided
    if output_pdf is None:
        base_path = os.path.splitext(input_pdf)[0]
        output_pdf = f"{base_path}_CLEAN.pdf"
    
    try:
        print(f"Reading scanned PDF: {input_pdf}")
        print("Converting PDF pages to images...")
        
        # Convert PDF to images
        images = convert_from_path(input_pdf, dpi=200)
        print(f"Extracted {len(images)} pages as images")
        
        # Extract text using OCR
        extracted_content = []
        for page_num, image in enumerate(images, 1):
            print(f"Performing OCR on page {page_num}...")
            text = pytesseract.image_to_string(image, lang='eng')
            if text.strip():
                extracted_content.append({
                    'page': page_num,
                    'text': text.strip()
                })
            else:
                extracted_content.append({
                    'page': page_num,
                    'text': '[No text detected on this page]'
                })
        
        # Create new PDF with proper formatting
        print(f"\nCreating formatted PDF: {output_pdf}")
        doc = SimpleDocTemplate(
            output_pdf,
            pagesize=letter,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=0.75*inch
        )
        
        # Get styles
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=12,
            alignment=1  # Center
        )
        
        page_header_style = ParagraphStyle(
            'PageHeader',
            parent=styles['Heading2'],
            fontSize=11,
            textColor=colors.HexColor('#2e5c8a'),
            spaceAfter=8,
            spaceBefore=8,
            borderPadding=4
        )
        
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['BodyText'],
            fontSize=10,
            leading=14,
            spaceAfter=6
        )
        
        # Build document content
        story = []
        
        # Add title
        title = Paragraph(
            f"Document: {os.path.basename(input_pdf)}<br/>Extracted via OCR & Reformatted",
            title_style
        )
        story.append(title)
        story.append(Spacer(1, 0.3*inch))
        
        # Add extraction metadata
        metadata = Paragraph(
            f"<b>Extraction Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>"
            f"<b>Original Pages:</b> {len(extracted_content)}<br/>"
            f"<b>Method:</b> OCR (Optical Character Recognition)<br/>"
            f"<b>Status:</b> Successfully reformatted and cleaned",
            body_style
        )
        story.append(metadata)
        story.append(Spacer(1, 0.3*inch))
        
        # Add content from each page
        for page_info in extracted_content:
            # Page header
            page_header = Paragraph(f"PAGE {page_info['page']}", page_header_style)
            story.append(page_header)
            
            # Content
            text = page_info['text']
            content = Paragraph(text, body_style)
            story.append(content)
            story.append(Spacer(1, 0.2*inch))
            story.append(PageBreak())
        
        # Build PDF
        doc.build(story)
        
        file_size = os.path.getsize(output_pdf)
        print(f"\n✓ Success!")
        print(f"✓ Clean PDF created: {output_pdf}")
        print(f"✓ File size: {file_size:,} bytes ({file_size/1024:.2f} KB)")
        print(f"✓ Total pages: {len(extracted_content)}")
        
    except ImportError as e:
        print(f"Error: Missing required library")
        print(f"Details: {e}")
        print("\nPlease install the required packages:")
        print("pip install pytesseract pdf2image")
        print("\nAlso install Tesseract OCR:")
        print("Download from: https://github.com/UB-Mannheim/tesseract/wiki")
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    original_pdf = r"E:\Data Model\DocScanner 30 May 2025 11-04.pdf"
    create_ocr_pdf(original_pdf)
