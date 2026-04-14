# Missing Python files recreation
write-host "Creating missing files..."

# extract_pdf_text.py
@"
import pdfplumber
import os
from pathlib import Path

def extract_pdf_text(pdf_path, output_path=None):
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found at {pdf_path}")
        return
    
    if output_path is None:
        base_path = Path(pdf_path).stem
        output_path = f"{base_path}_extracted.txt"
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"Processing PDF with {len(pdf.pages)} pages...")
            
            full_text = []
            
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"Extracting page {page_num}...")
                text = page.extract_text()
                
                if page_num > 1:
                    full_text.append("\n" + "="*80 + "\n")
                full_text.append(f"PAGE {page_num}\n")
                full_text.append("="*80 + "\n\n")
                
                if text:
                    full_text.append(text)
                else:
                    full_text.append("[No text found on this page]")
                
                full_text.append("\n\n")
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.writelines(full_text)
        
        print(f"\nSuccess! Text extracted and saved to: {output_path}")
        print(f"File size: {os.path.getsize(output_path)} bytes")
        
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")

if __name__ == "__main__":
    pdf_file = r"E:\Data Model\DocScanner 30 May 2025 11-04.pdf"
    extract_pdf_text(pdf_file)
