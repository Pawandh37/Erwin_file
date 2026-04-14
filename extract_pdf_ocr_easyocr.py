"""
Extract text from scanned PDF using EasyOCR and create clean PDF
"""
import os
from pathlib import Path
from datetime import datetime

def create_ocr_pdf_easyocr(input_pdf, output_pdf=None):
    """
    Extract text from scanned PDF using EasyOCR.
    """
    if not os.path.exists(input_pdf):
        print(f"Error: PDF file not found at {input_pdf}")
        return
    
    try:
        print("Importing required libraries...")
        from pdf2image import convert_from_path
        import easyocr
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.lib import colors
        
    except ImportError as e:
        print(f"Missing library: {e}")
        print("\nPlease install required packages:")
        print("  ./.venv/Scripts/python.exe -m pip install easyocr pdf2image")
        return
    
    if output_pdf is None:
        base_path = os.path.splitext(input_pdf)[0]
        output_pdf = f"{base_path}_CLEAN.pdf"
    
    try:
        print(f"\nReading scanned PDF: {input_pdf}")
        print("Converting PDF to images...")
        
        # Convert PDF to images
        images = convert_from_path(input_pdf, dpi=150)
        print(f"✓ Converted {len(images)} pages to images")
        
        # Initialize EasyOCR reader
        print("\nInitializing OCR reader (first run downloads model, ~100 MB)...")
        reader = easyocr.Reader(['en'], verbose=False)
        
        # Extract text from images
        extracted_content = []
        for page_num, image in enumerate(images, 1):
            print(f"Running OCR on page {page_num}/{len(images)}...", end=' ')
            
            # Perform OCR
            results = reader.readtext(image, detail=0)
            text = '\n'.join(results)
            
            if text.strip():
                extracted_content.append({
                    'page': page_num,
                    'text': text.strip()
                })
                print(f"✓ ({len(results)} lines extracted)")
            else:
                extracted_content.append({
                    'page': page_num,
                    'text': '[No text detected on this page]'
                })
                print("(No text detected)")
        
        print(f"\n✓ OCR extraction complete!")
        print(f"Creating formatted PDF: {output_pdf}")
        
        # Create new PDF
        doc = SimpleDocTemplate(
            output_pdf,
            pagesize=letter,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=0.75*inch
        )
        
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=12,
            alignment=1
        )
        
        page_header_style = ParagraphStyle(
            'PageHeader',
            parent=styles['Heading2'],
            fontSize=11,
            textColor=colors.HexColor('#2e5c8a'),
            spaceAfter=8,
            spaceBefore=8
        )
        
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['BodyText'],
            fontSize=10,
            leading=14,
            spaceAfter=6
        )
        
        # Build document
        story = []
        
        # Title
        title = Paragraph(
            f"Document: {os.path.basename(input_pdf)}<br/>Extracted via OCR & Reformatted",
            title_style
        )
        story.append(title)
        story.append(Spacer(1, 0.3*inch))
        
        # Metadata
        metadata = Paragraph(
            f"<b>Extraction Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>"
            f"<b>Pages:</b> {len(extracted_content)}<br/>"
            f"<b>Method:</b> EasyOCR (Optical Character Recognition)<br/>"
            f"<b>Status:</b> ✓ Successfully extracted and reformatted",
            body_style
        )
        story.append(metadata)
        story.append(Spacer(1, 0.4*inch))
        
        # Content
        for page_info in extracted_content:
            page_header = Paragraph(f"PAGE {page_info['page']}", page_header_style)
            story.append(page_header)
            
            text = page_info['text']
            # Escape special characters for HTML
            text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            content = Paragraph(text, body_style)
            story.append(content)
            story.append(Spacer(1, 0.3*inch))
            story.append(PageBreak())
        
        # Build PDF
        doc.build(story)
        
        file_size = os.path.getsize(output_pdf)
        print(f"\n{'='*60}")
        print(f"✓ SUCCESS!")
        print(f"{'='*60}")
        print(f"Clean PDF created: {output_pdf}")
        print(f"File size: {file_size:,} bytes ({file_size/1024:.2f} KB)")
        print(f"Total pages: {len(extracted_content)}")
        print(f"{'='*60}")
        
    except ImportError as e:
        print(f"Error: Missing required package")
        print(f"Details: {e}")
        print("\nTo use this script, install EasyOCR:")
        print("  ./.venv/Scripts/python.exe -m pip install easyocr pdf2image")
        print("\nNote: First run will download ~100 MB OCR model")
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    original_pdf = r"E:\Data Model\DocScanner 30 May 2025 11-04.pdf"
    create_ocr_pdf_easyocr(original_pdf)
