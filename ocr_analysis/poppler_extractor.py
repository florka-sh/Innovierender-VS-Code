"""
Invoice Extractor using run_with_poppler.bat
Extracts PDFs and displays progress in Streamlit dashboard
"""

import os
import subprocess
import json
from pathlib import Path
from datetime import datetime
import pandas as pd

class PopperExtractor:
    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.poppler_bat = os.path.join(self.script_dir, 'run_with_poppler.bat')
        
    def extract_pdf(self, pdf_path, output_folder='extracted_pages', progress_callback=None):
        """Extract PDF using run_with_poppler.bat environment"""
        
        # Create output folder
        os.makedirs(output_folder, exist_ok=True)
        
        # Get PDF info
        pdf_filename = os.path.basename(pdf_path)
        pdf_pages = self._count_pdf_pages(pdf_path)
        start_time = datetime.now()
        
        # Run extraction with poppler environment
        extraction_log = {
            'filename': pdf_filename,
            'filepath': pdf_path,
            'total_pages': pdf_pages,
            'start_time': datetime.now().isoformat(),
            'status': 'extracting',
            'pages_completed': 0,
            'errors': []
        }
        
        try:
            # Set up environment from run_with_poppler; prefer bundled InvoiceExtractor/_internal when present
            candidates = []
            # primary: this module's _internal
            candidates.append(os.path.join(self.script_dir, '_internal'))
            # sibling InvoiceExtractor/_internal (common upload location)
            candidates.append(os.path.abspath(os.path.join(self.script_dir, '..', 'InvoiceExtractor', '_internal')))
            # repo-root InvoiceExtractor/_internal (in case the script moved)
            candidates.append(os.path.abspath(os.path.join(self.script_dir, '..', '..', 'InvoiceExtractor', '_internal')))

            poppler_path = None
            poppler_home = None
            tesseract_path = None
            tessdata_path = None

            for base in candidates:
                if not base:
                    continue
                p_poppler = os.path.join(base, 'poppler', 'Library', 'bin')
                p_tess = os.path.join(base, 'Tesseract-OCR')
                if os.path.exists(p_poppler) and poppler_path is None:
                    poppler_path = p_poppler
                    poppler_home = os.path.join(base, 'poppler')
                if os.path.exists(p_tess) and tesseract_path is None:
                    tesseract_path = p_tess
                    tessdata_path = os.path.join(p_tess, 'tessdata')

            # Fallback to script-local locations if nothing found
            if poppler_path is None:
                poppler_path = os.path.join(self.script_dir, '_internal', 'poppler', 'Library', 'bin')
                poppler_home = os.path.join(self.script_dir, '_internal', 'poppler')
            if tesseract_path is None:
                tesseract_path = os.path.join(self.script_dir, '_internal', 'Tesseract-OCR')
                tessdata_path = os.path.join(tesseract_path, 'tessdata')

            # report what we are using (helpful when debugging runtime issues)
            print(f"Using poppler path: {poppler_path}")
            print(f"Using tesseract path: {tesseract_path}")
            
            # Update system environment for this process
            if poppler_path not in os.environ['PATH']:
                os.environ['PATH'] = poppler_path + os.pathsep + os.environ['PATH']
            
            if tesseract_path not in os.environ['PATH']:
                os.environ['PATH'] = tesseract_path + os.pathsep + os.environ['PATH']
                
            os.environ['POPPLER_HOME'] = poppler_home
            os.environ['TESSDATA_PREFIX'] = tessdata_path
            
            # Import with proper environment
            import pdf2image
            import pytesseract
            
            # Set Tesseract command
            tesseract_cmd = os.path.join(tesseract_path, 'tesseract.exe') if tesseract_path else 'tesseract'
            if os.path.exists(tesseract_cmd):
                try:
                    pytesseract.pytesseract_cmd = tesseract_cmd
                except Exception:
                    # in case older API or nested attribute
                    try:
                        pytesseract.pytesseract.pytesseract_cmd = tesseract_cmd
                    except Exception:
                        pass
                print(f"✓ Tesseract found at: {tesseract_cmd}")
            else:
                print(f"✗ Tesseract NOT found at: {tesseract_cmd}")
                # Try default if not found
                try:
                    pytesseract.pytesseract_cmd = 'tesseract'
                except Exception:
                    try:
                        pytesseract.pytesseract.pytesseract_cmd = 'tesseract'
                    except Exception:
                        pass
            
            # Convert PDF to images
            print(f"Converting {pdf_filename}...")
            images = pdf2image.convert_from_path(
                pdf_path,
                dpi=200,
                poppler_path=poppler_path if os.path.exists(poppler_path) else None,
                timeout=60
            )
            
            # Extract text from each page
            extracted_data = []
            for page_num, image in enumerate(images, 1):
                try:
                    # Extract text using OCR
                    text = pytesseract.image_to_string(image)
                    
                    extracted_data.append({
                        'page': page_num,
                        'text': text,  # Store full text
                        'timestamp': datetime.now().isoformat()
                    })
                    
                    extraction_log['pages_completed'] = page_num
                    print(f"✓ Page {page_num}/{pdf_pages} extracted")
                    
                    # Update progress
                    if progress_callback:
                        elapsed = (datetime.now() - start_time).total_seconds()
                        progress_callback(page_num, pdf_pages, elapsed)
                    
                except Exception as e:
                    extraction_log['errors'].append(f"Page {page_num}: {str(e)}")
                    print(f"✗ Error on page {page_num}: {str(e)}")
            
            extraction_log['status'] = 'completed'
            extraction_log['end_time'] = datetime.now().isoformat()
            
            # Save to Excel
            df = pd.DataFrame(extracted_data)
            df['Filename'] = pdf_filename
            df['Page'] = df['page']
            df['Invoice Number'] = 'To Extract'
            df['Invoice Suffix'] = ''
            df['Recipient Name'] = 'To Extract'
            df['Date'] = datetime.now().strftime('%d.%m.%Y')
            df['Description'] = df['text']
            df['Quantity'] = ''
            df['Rate'] = ''
            df['Line Total'] = ''
            df['Total Amount'] = ''
            
            # Reorder columns
            output_df = df[['Filename', 'Page', 'Invoice Number', 'Invoice Suffix', 
                          'Recipient Name', 'Date', 'Description', 'Quantity', 
                          'Rate', 'Line Total', 'Total Amount']]
            
            output_file = os.path.join(output_folder, 'new.xlsx')
            output_df.to_excel(output_file, index=False)
            
            print(f"\n✓ Extraction complete! Saved to {output_file}")
            extraction_log['output_file'] = output_file
            
            return extraction_log, output_df
            
        except Exception as e:
            extraction_log['status'] = 'failed'
            extraction_log['error'] = str(e)
            print(f"✗ Extraction failed: {str(e)}")
            return extraction_log, None
    
    def _count_pdf_pages(self, pdf_path):
        """Count total pages in PDF"""
        try:
            import PyPDF2
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                return len(reader.pages)
        except:
            # Fallback: use pdf2image
            try:
                import pdf2image
                images = pdf2image.convert_from_path(pdf_path, first_page=1, last_page=1)
                # This is just estimate, will be accurate during extraction
                return 50  # Default estimate
            except:
                return 0

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
        extractor = PopperExtractor()
        result = extractor.extract_pdf(pdf_file)
        print(f"\nResult: {result}")
    else:
        print("Usage: python poppler_extractor.py <pdf_file>")
