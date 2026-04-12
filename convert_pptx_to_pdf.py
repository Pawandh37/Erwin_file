import subprocess
import os

pptx_file = r"e:\VScode26\Data_Governance_Quality_Banking.pptx"
output_dir = r"e:\VScode26"

# Try using LibreOffice
try:
    cmd = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        pptx_file
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    print(f"✅ Conversion successful!")
    print(f"PDF saved to: {output_dir}")
    
    # Verify file exists
    pdf_file = os.path.join(output_dir, "Data_Governance_Quality_Banking.pdf")
    if os.path.exists(pdf_file):
        print(f"✅ File verified: {pdf_file}")
    else:
        print("⚠️ PDF file not found after conversion")
        
except FileNotFoundError:
    print("❌ LibreOffice not found. Trying alternative method...")
    try:
        # Try using comtypes for Windows PowerPoint
        from pptx import Presentation
        from PIL import Image
        from io import BytesIO
        print("Falling back to python-pptx method (images only - no direct PDF support)")
    except ImportError:
        print("❌ Required libraries not available")
except subprocess.CalledProcessError as e:
    print(f"❌ Conversion error: {e.stderr}")
