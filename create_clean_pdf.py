"""
Extract text from PDF and create a new clean, readable PDF
"""
import pdfplumber
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime
import os

def create_readable_pdf(input_pdf, output_pdf=None):
    """
    Extract text from PDF and create a new, well-formatted PDF.
    
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
        print(f"Reading original PDF: {input_pdf}")
        
        # Extract text from PDF
        extracted_content = []
        with pdfplumber.open(input_pdf) as pdf:
            print(f"Processing {len(pdf.pages)} pages...")
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    extracted_content.append({
                        'page': page_num,
                        'text': text.strip()
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
            borderPadding=4,
            borderColor=colors.HexColor('#cccccc'),
            borderWidth=0.5
        )
        
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['BodyText'],
            fontSize=10,
            leading=14,
            spaceAfter=6,
            alignment=4  # Justify
        )
        
        # Build document content
        story = []
        
        # Add title
        title = Paragraph(
            f"Document: {os.path.basename(input_pdf)}<br/>Extracted & Reformatted",
            title_style
        )
        story.append(title)
        story.append(Spacer(1, 0.3*inch))
        
        # Add extraction metadata
        metadata = Paragraph(
            f"<b>Extraction Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>"
            f"<b>Original Pages:</b> {len(extracted_content)}<br/>"
            f"<b>Status:</b> Successfully reformatted and cleaned",
            body_style
        )
        story.append(metadata)
        story.append(Spacer(1, 0.3*inch))
        
        # Add content from each page
        for page_info in extracted_content:
            # Page header
            page_header = Paragraph(
                f"PAGE {page_info['page']}",
                page_header_style
            )
            story.append(page_header)
            
            # Content
            if page_info['text']:
                # Clean up text
                text = page_info['text']
                # Replace multiple spaces/newlines
                text = ' '.join(text.split())
                
                content = Paragraph(text, body_style)
                story.append(content)
            else:
                empty_notice = Paragraph("[No text content on this page]", body_style)
                story.append(empty_notice)
            
            story.append(Spacer(1, 0.2*inch))
            story.append(PageBreak())
        
        # Build PDF
        doc.build(story)
        
        file_size = os.path.getsize(output_pdf)
        print(f"\n✓ Success!")
        print(f"✓ Clean PDF created: {output_pdf}")
        print(f"✓ File size: {file_size:,} bytes ({file_size/1024:.2f} KB)")
        print(f"✓ Total pages: {len(extracted_content)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    original_pdf = r"E:\Data Model\DocScanner 30 May 2025 11-04.pdf"
    create_readable_pdf(original_pdf)
