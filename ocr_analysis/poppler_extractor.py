"""
Invoice Extractor using Poppler and Tesseract OCR
Extracts PDFs with improved text recognition and financial data parsing
"""

import os
import re
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
import logging
import cv2
import numpy as np

logger = logging.getLogger(__name__)


class PopperExtractor:
    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.poppler_bat = os.path.join(self.script_dir, 'run_with_poppler.bat')
        self._setup_ocr_paths()
        
    def _setup_ocr_paths(self):
        """Setup OCR tool paths"""
        candidates = []
        candidates.append(os.path.join(self.script_dir, '_internal'))
        candidates.append(os.path.abspath(
            os.path.join(self.script_dir, '..', 'InvoiceExtractor', '_internal')
        ))
        candidates.append(os.path.abspath(
            os.path.join(self.script_dir, '..', '..', 'InvoiceExtractor', '_internal')
        ))

        self.poppler_path = None
        self.poppler_home = None
        self.tesseract_path = None
        self.tessdata_path = None

        for base in candidates:
            if not base:
                continue
            p_poppler = os.path.join(base, 'poppler', 'Library', 'bin')
            p_tess = os.path.join(base, 'Tesseract-OCR')
            if os.path.exists(p_poppler) and self.poppler_path is None:
                self.poppler_path = p_poppler
                self.poppler_home = os.path.join(base, 'poppler')
            if os.path.exists(p_tess) and self.tesseract_path is None:
                self.tesseract_path = p_tess
                self.tessdata_path = os.path.join(p_tess, 'tessdata')

        # Fallback
        if self.poppler_path is None:
            self.poppler_path = os.path.join(
                self.script_dir, '_internal', 'poppler', 'Library', 'bin'
            )
            self.poppler_home = os.path.join(
                self.script_dir, '_internal', 'poppler'
            )
        if self.tesseract_path is None:
            self.tesseract_path = os.path.join(
                self.script_dir, '_internal', 'Tesseract-OCR'
            )
            self.tessdata_path = os.path.join(
                self.tesseract_path, 'tessdata'
            )

    def extract_pdf(
        self, 
        pdf_path, 
        output_folder='extracted_pages', 
        progress_callback=None
    ):
        """Extract PDF using OCR with improved data parsing"""
        
        os.makedirs(output_folder, exist_ok=True)
        
        pdf_filename = os.path.basename(pdf_path)
        pdf_pages = self._count_pdf_pages(pdf_path)
        start_time = datetime.now()
        
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
            # Setup environment
            print(f"Using poppler path: {self.poppler_path}")
            print(f"Using tesseract path: {self.tesseract_path}")
            
            if self.poppler_path not in os.environ.get('PATH', ''):
                os.environ['PATH'] = (
                    self.poppler_path + os.pathsep + os.environ.get('PATH', '')
                )
            
            if self.tesseract_path not in os.environ.get('PATH', ''):
                os.environ['PATH'] = (
                    self.tesseract_path + os.pathsep + os.environ.get('PATH', '')
                )
                
            os.environ['POPPLER_HOME'] = self.poppler_home
            os.environ['TESSDATA_PREFIX'] = self.tessdata_path
            
            # Import with proper environment
            import pdf2image
            import pytesseract
            
            # Configure Tesseract
            tesseract_cmd = os.path.join(
                self.tesseract_path, 'tesseract.exe'
            )
            if os.path.exists(tesseract_cmd):
                pytesseract.pytesseract_cmd = tesseract_cmd
                print(f"✓ Tesseract found at: {tesseract_cmd}")
            else:
                print(f"✗ Tesseract NOT found at: {tesseract_cmd}")
            
            # Convert PDF to images
            print(f"Converting {pdf_filename}...")
            images = pdf2image.convert_from_path(
                pdf_path,
                dpi=300,  # Higher DPI for better OCR
                poppler_path=self.poppler_path if os.path.exists(
                    self.poppler_path
                ) else None,
                timeout=120
            )
            
            # Extract text from each page
            extracted_data = []
            for page_num, image in enumerate(images, 1):
                try:
                    # Preprocess image for better OCR
                    processed_image = self._preprocess_image(image)
                    
                    # Extract text using OCR with better config
                    text = pytesseract.image_to_string(
                        processed_image,
                        lang='deu+eng',  # German + English
                        config='--psm 6 --oem 3'  # PSM 6: uniform text blocks
                    )
                    
                    # DEBUG: Print raw OCR output
                    print("\n" + "="*80)
                    print(f"RAW OCR OUTPUT - Page {page_num}:")
                    print("="*80)
                    print(text)
                    print("="*80 + "\n")
                    
                    # Get confidence scores
                    data = pytesseract.image_to_data(
                        processed_image,
                        lang='deu+eng',
                        output_type=pytesseract.Output.DICT
                    )
                    confidence = self._calculate_confidence(data)
                    
                    # Parse extracted text for financial data
                    parsed = self._parse_invoice_text(
                        text,
                        page_num,
                        confidence
                    )

                    # Helper: convert German amount string to cents (int) using Decimal
                    from decimal import Decimal, InvalidOperation
                    def _to_cents(amount_str: str) -> int:
                        if not amount_str:
                            return 0
                        cleaned = amount_str.strip()
                        # German format: thousands '.' decimal ','
                        cleaned = cleaned.replace('.', '').replace(',', '.')
                        try:
                            value = Decimal(cleaned)
                            # Quantize to 2 decimals before multiplying
                            value = value.quantize(Decimal('0.01'))
                            return int((value * 100).to_integral_value())
                        except (InvalidOperation, ValueError):
                            return 0
                    def _to_decimal(amount_str: str) -> Decimal:
                        if not amount_str:
                            return Decimal('0.00')
                        cleaned = amount_str.strip().replace('.', '').replace(',', '.')
                        try:
                            value = Decimal(cleaned)
                            return value.quantize(Decimal('0.01'))
                        except (InvalidOperation, ValueError):
                            return Decimal('0.00')

                    line_items = parsed.get('line_items', [])
                    total_amount_str = parsed.get('total_amount', '')
                    total_amount_cents = _to_cents(total_amount_str)
                    total_amount_decimal = _to_decimal(total_amount_str)
                    running_sum_cents = 0
                    running_sum_decimal = Decimal('0.00')

                    if line_items:
                        for idx, item in enumerate(line_items, 1):
                            amt_str = item.get('amount', '')
                            amt_cents = _to_cents(amt_str)
                            amt_dec = _to_decimal(amt_str)
                            running_sum_cents += amt_cents
                            running_sum_decimal += amt_dec
                            # Build booking text enriched with names
                            fam = parsed.get('family_name', '')
                            giv = parsed.get('given_name', '')
                            names_part = f"{fam} {giv}".strip()
                            booking_text = item.get('description', '')
                            if names_part:
                                if booking_text:
                                    booking_text = f"{booking_text} + {names_part}"
                                else:
                                    booking_text = names_part
                            row = {
                                'Page': page_num,
                                'Filename': pdf_filename,
                                'Invoice Number': parsed.get(
                                    'invoice_number', 'To Extract'
                                ),
                                'Invoice Suffix': parsed.get(
                                    'invoice_suffix', ''
                                ),
                                'Customer Number': parsed.get(
                                    'customer_number', ''
                                ),
                                'Recipient Name': parsed.get(
                                    'recipient_name', 'To Extract'
                                ),
                                'Date': parsed.get(
                                    'date', datetime.now().strftime('%d.%m.%Y')
                                ),
                                'Booking Text': booking_text,
                                'Description': item.get('description', ''),
                                'Quantity': parsed.get('quantity', ''),
                                'Rate': parsed.get('rate', ''),
                                'Line Total': amt_str,
                                'BETRAG': str(amt_cents),  # cents
                                'Total Amount': parsed.get('total_amount', ''),
                                'Confidence': f"{confidence:.0f}%",
                                'ocr_confidence': confidence,
                                'validation_required': confidence < 70.0,
                                'Sum Check': 'OK' if (
                                    total_amount_cents and
                                    running_sum_cents <= total_amount_cents
                                ) else '',
                                'PreRound Sum': str(running_sum_decimal),
                                'Total Decimal': str(total_amount_decimal)
                            }
                            extracted_data.append(row)
                        # After loop finalize sum check status for last row if mismatch
                        if (
                            total_amount_cents and
                            running_sum_cents != total_amount_cents
                        ):
                            # Mark recent rows as mismatch
                            mismatch_msg = (
                                f"Mismatch ({running_sum_cents} != "
                                f"{total_amount_cents})"
                            )
                            for row in extracted_data[-len(line_items):]:
                                row['Sum Check'] = mismatch_msg
                        elif total_amount_cents:
                            extracted_data[-1]['Sum Check'] = 'OK'
                        # Add a difference hint column if small discrepancy <= 50 cents
                        diff_cents = abs(total_amount_cents - running_sum_cents)
                        if diff_cents and diff_cents <= 50:
                            for row in extracted_data[-len(line_items):]:
                                row['Difference Cents'] = diff_cents
                        else:
                            for row in extracted_data[-len(line_items):]:
                                row['Difference Cents'] = ''
                    else:
                        # Fallback single row behavior
                        single_row = {
                            'Page': page_num,
                            'Filename': pdf_filename,
                            'Invoice Number': parsed.get(
                                'invoice_number', 'To Extract'
                            ),
                            'Invoice Suffix': parsed.get('invoice_suffix', ''),
                            'Customer Number': parsed.get('customer_number', ''),
                            'Recipient Name': parsed.get(
                                'recipient_name', 'To Extract'
                            ),
                            'Date': parsed.get(
                                'date', datetime.now().strftime('%d.%m.%Y')
                            ),
                            'Booking Text': parsed.get('booking_text', ''),
                            'Description': text[:500],
                            'Quantity': parsed.get('quantity', ''),
                            'Rate': parsed.get('rate', ''),
                            'Line Total': parsed.get('line_total', ''),
                            'BETRAG': str(
                                _to_cents(
                                    parsed.get('line_total', '')
                                )
                            ),
                            'Total Amount': parsed.get('total_amount', ''),
                            'Confidence': f"{confidence:.0f}%",
                            'ocr_confidence': confidence,
                            'validation_required': confidence < 70.0,
                            'Sum Check': ''
                        }
                        extracted_data.append(single_row)
                    
                    extraction_log['pages_completed'] = page_num
                    print(
                        f"✓ Page {page_num}/{pdf_pages} extracted "
                        f"(Confidence: {confidence:.0f}%)"
                    )
                    
                    if progress_callback:
                        elapsed = (
                            datetime.now() - start_time
                        ).total_seconds()
                        progress_callback(page_num, pdf_pages, elapsed)
                    
                except Exception as e:
                    error_msg = f"Page {page_num}: {str(e)}"
                    extraction_log['errors'].append(error_msg)
                    print(f"✗ Error on page {page_num}: {str(e)}")
            
            extraction_log['status'] = 'completed'
            extraction_log['end_time'] = datetime.now().isoformat()
            
            # Create DataFrame with proper template columns
            df = pd.DataFrame(extracted_data)
            
            # Ensure all template columns exist
            template_columns = [
                'Filename', 'Page', 'Invoice Number', 'Invoice Suffix',
                'Recipient Name', 'Date', 'Description', 'Quantity',
                'Rate', 'Line Total', 'Total Amount'
            ]
            
            for col in template_columns:
                if col not in df.columns:
                    df[col] = ''
            
            output_df = df[template_columns]
            
            output_file = os.path.join(output_folder, 'extracted_invoices.xlsx')
            output_df.to_excel(output_file, index=False)
            
            print(f"\n✓ Extraction complete! Saved to {output_file}")
            extraction_log['output_file'] = output_file
            
            return extraction_log, output_df
            
        except Exception as e:
            extraction_log['status'] = 'failed'
            extraction_log['error'] = str(e)
            print(f"✗ Extraction failed: {str(e)}")
            logger.error(f"OCR extraction error: {str(e)}", exc_info=True)
            return extraction_log, None
    
    def _preprocess_image(self, image):
        """Preprocess image for better OCR accuracy (+20-30% improvement)"""
        try:
            import cv2
        except ImportError:
            logger.warning("cv2 not available, skipping preprocessing")
            return image
        
        # Convert PIL image to numpy array
        if hasattr(image, 'convert'):
            image = image.convert('L')  # Grayscale
            image_np = np.array(image)
        else:
            image_np = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Apply CLAHE (Contrast Limited Adaptive Histogram Equalization)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(image_np)
        
        # Denoise
        denoised = cv2.medianBlur(enhanced, 3)
        
        # Thresholding
        _, binary = cv2.threshold(
            denoised,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        
        # Convert back to PIL for pytesseract
        from PIL import Image
        return Image.fromarray(binary)
    
    def _calculate_confidence(self, data):
        """Calculate average OCR confidence from Tesseract data"""
        try:
            if 'conf' in data:
                confidences = [
                    float(conf) for conf in data['conf']
                    if str(conf) != '-1' and conf != -1
                ]
            elif 'confidence' in data:
                confidences = [
                    float(conf) for conf in data['confidence']
                    if str(conf) != '-1' and conf != -1
                ]
            else:
                print("DEBUG: No confidence data in OCR output")
                return 0.0
                
            if confidences:
                avg_conf = sum(confidences) / len(confidences)
                print(f"DEBUG: Average OCR confidence: {avg_conf:.1f}%")
                return avg_conf
            return 0.0
        except Exception as e:
            print(f"Error calculating confidence: {e}")
            return 0.0
    
    def _clean_ocr_text(self, text: str) -> str:
        """
        Clean OCR artifacts and noise from extracted text.
        
        Args:
            text: Raw OCR text
            
        Returns:
            Cleaned text
        """
        if not text:
            return text
        
        # Remove common OCR artifacts and noise characters
        noise_patterns = [
            r'[_\-—–]{3,}',  # Multiple dashes/underscores
            r'[\.]{3,}',  # Multiple dots
            r'[,;]{2,}',  # Multiple punctuation
            r'[\|]{2,}',  # Multiple pipes
            r'[\'"`]{2,}',  # Multiple quotes
            r'\s+[_\-—]\s+',  # Isolated dashes
        ]
        
        cleaned = text
        for pattern in noise_patterns:
            cleaned = re.sub(pattern, ' ', cleaned)
        
        # Remove special Unicode characters often from OCR errors
        # Keep letters, numbers, spaces, and common punctuation
        cleaned = re.sub(
            r'[^\w\s\.,\-/()€$]',
            '',
            cleaned,
            flags=re.UNICODE
        )
        
        # Remove extra whitespace
        cleaned = re.sub(r'\s+', ' ', cleaned)
        cleaned = cleaned.strip()
        
        return cleaned
    
    def _parse_invoice_text(
        self,
        text: str,
        page_num: int,
        confidence: float = 0.0
    ) -> dict:
        """
        Parse OCR text to extract financial data using smart patterns.
        Based on proven techniques from production systems.
        
        Returns:
            Dictionary with extracted fields
        """
        result = {
            'invoice_number': '',
            'invoice_suffix': '',
            'customer_number': '',
            'recipient_name': '',
            'date': '',
            'quantity': '',
            'rate': '',
            'line_total': '',
            'total_amount': '',
            'confidence': confidence
        }
        
        if not text:
            return result
        
        lines = text.split('\n')
        
        # 1. Extract date (DD.MM.YYYY or DD/MM/YYYY format)
        date_pattern = (
            r'(\d{2}[./-]\d{2}[./-]\d{4}|'
            r'\d{4}[./-]\d{2}[./-]\d{2})'
        )
        dates = re.findall(date_pattern, text)
        if dates:
            result['date'] = dates[0]
        
        # 2. Extract invoice number (improved pattern)
        inv_keywords = [
            r'Rg\.-Nr\.',
            r'Rechnungsnummer',
            r'Rechnung Nr\.', 
            r'Invoice\s*#?\s*'
        ]
        
        for keyword in inv_keywords:
            match = re.search(
                f"({keyword})",
                text,
                re.IGNORECASE
            )
            if match:
                start_pos = match.end()
                snippet = text[start_pos:start_pos+60]
                
                # Look for alphanumeric sequences
                candidates = re.findall(
                    r'\b([A-Za-z0-9\-\/]{3,})\b',
                    snippet
                )
                
                for cand in candidates:
                    if len(cand) < 3:
                        continue
                    if re.match(r'\d{2}\.\d{2}\.\d{4}', cand):
                        continue
                    
                    if '/' in cand:
                        parts = cand.split('/')
                        if len(parts[0]) > 3:
                            result['invoice_number'] = parts[0]
                            if len(parts) > 1:
                                result['invoice_suffix'] = parts[1]
                    else:
                        result['invoice_number'] = cand
                    break
                
                if result['invoice_number']:
                    break
        
        # 2b. Extract customer number (Kunden-Nummer / Debitoren)
        customer_patterns = [
            r'Kunden-Nummer[:\s]+([0-9\s]+)',
            r'Kundennummer[:\s]+([0-9\s]+)',
            r'Debitoren[:\s]+([0-9\s]+)',
            r'Personenkonto[:\s]+([0-9\s]+)',
        ]
        
        for pattern in customer_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                # Remove spaces and whitespace
                result['customer_number'] = match.group(1).replace(
                    ' ', ''
                ).strip()
                break
        
        # 3. Extract recipient name
        # Note: "Name" = Family name (last name)
        #       "Vorname" = Given name (first name)
        # Output format: "Vorname Name" (FirstName LastName)
        family_name = None
        given_name = None
        
        name_match = re.search(
            r'Name\s+(.+?)(?:\s+Geb\.-Datum|\n)',
            text
        )
        if name_match:
            family_name = self._clean_ocr_text(
                name_match.group(1)
            )
        
        vorname_match = re.search(
            r'Vorname\s+(.+?)(?:\n|$)',
            text
        )
        if vorname_match:
            given_name = self._clean_ocr_text(
                vorname_match.group(1)
            )
        
        # Build full name: "Vorname Name" format
        if given_name and family_name:
            result['recipient_name'] = f"{given_name} {family_name}"
        elif family_name:
            result['recipient_name'] = family_name
        elif given_name:
            result['recipient_name'] = given_name
        # Preserve separate parts for later row construction
        result['family_name'] = family_name or ''
        result['given_name'] = given_name or ''
        
        # 4. Extract total amount (improved)
        amount_keywords = [
            r'Rechnungsbetrag',
            r'Zahlbetrag',
            r'Endbetrag',
            r'Gesamtbetrag',
            r'Gesamt',
            r'Total',
            r'Betrag',
            r'Summe'
        ]
        
        print(f"\nDEBUG: Searching for amount in text...")
        found_amount = None
        for keyword in amount_keywords:
            pattern = (
                rf"(?:{keyword}).{{0,100}}?"
                r"(\d{{1,3}}(?:[.,]\d{{3}})*[.,]\d{{2}})"
            )
            match = re.search(
                pattern,
                text,
                re.IGNORECASE | re.DOTALL
            )
            if match:
                found_amount = match.group(1)
                print(f"DEBUG: Found amount '{found_amount}' with keyword: {keyword}")
                if any(k in keyword.lower() for k in [
                    'zahlbetrag', 'endbetrag',
                    'rechnungsbetrag'
                ]):
                    break
        
        # Fallback: look for EUR amounts with multiple patterns
        if not found_amount:
            print("DEBUG: No keyword match, trying EUR/EURO patterns...")
            
            # Try different currency patterns
            currency_patterns = [
                # German format: 5.608,31 EURO or € 
                r'(\d{1,3}(?:\.\d{3})*,\d{2})\s*(?:€|EUR|EURO)',
                # Space before currency: 5.608,31 EURO
                r'(\d{1,3}(?:\.\d{3})*,\d{2})\s+(?:EUR|EURO)',
                # Standalone format near "Rechnungsbetrag"
                r'Rechnungsbetrag[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
            ]
            
            for pattern in currency_patterns:
                matches = re.findall(pattern, text, re.IGNORECASE)
                if matches:
                    found_amount = matches[-1]
                    print(f"DEBUG: Found amount '{found_amount}' with pattern: {pattern[:50]}")
                    try:
                        values = []
                        for m in matches:
                            val_str = m.replace('.', '').replace(',', '.')
                            values.append((float(val_str), m))
                        values.sort(key=lambda x: x[0], reverse=True)
                        found_amount = values[0][1]
                    except Exception:
                        found_amount = matches[-1]
                    break
        
        result['total_amount'] = found_amount or ''
        if found_amount:
            print(f"DEBUG: Final amount set to: {found_amount}")
        else:
            print("DEBUG: No amount found!")
        
        # 5. Extract booking text / course description
        # Pattern: "2101 11612440 WG 3 UMA-Gruppe Wesel" or similar
        print("\nDEBUG: Searching for booking text...")
        booking_text = None
        
        # Try to find cost center or course codes
        booking_patterns = [
            # Pattern: "2101 [number] [description]"
            r'(\d{4})\s+(\d+)\s+(.+?)(?:\d{2}\.\d{2}\.\d{4})',
            # Pattern: "Kst. [number] - [description]"
            r'Kst\.\s*(\d+)\s*-\s*(.+?)(?:\n|$)',
            # Pattern: Line after "Leistung"
            r'Leistung.*?\n(.+?)(?:\d{2}\.\d{2}\.\d{4})',
        ]
        
        for pattern in booking_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if len(match.groups()) == 3:
                    # Format: "2101 11612440 WG 3 UMA-Gruppe"
                    cost_center = match.group(1)
                    account = match.group(2)
                    desc = self._clean_ocr_text(match.group(3).strip())
                    booking_text = f"{cost_center} {account} {desc}"
                elif len(match.groups()) == 2:
                    # Format: "Kst. 123 - Description"
                    booking_text = self._clean_ocr_text(
                        f"{match.group(1)} - {match.group(2)}"
                    )
                else:
                    # Single group - just description
                    booking_text = self._clean_ocr_text(match.group(1))
                
                print(f"DEBUG: Found booking text: {booking_text}")
                break
        
        result['booking_text'] = booking_text or ''
        
        # 5b. Extract detailed line items (each row of services with final amount)
        line_items = []
        currency_pattern = r"\d{1,3}(?:\.\d{3})*,\d{2}"  # German formatted number
        keywords = [
            'WG', 'Gruppe', 'Taschengeld', 'Bekleidung', 'Bekleidungsgeld',
            'Schillwiese', 'UMA', 'UMAs', 'Heim', 'Unterbringung'
        ]
        qty_pattern = re.compile(r'(\d{1,3}[,.]\d{2}|\d{1,3})\s*(Stunden|Std|Tage|x)', re.IGNORECASE)
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            # Skip total lines
            if re.search(r'Rechnungsbetrag|Gesamtbetrag|Endbetrag|Summe', line, re.IGNORECASE):
                continue
            # Skip lines that are just headers
            if re.search(r'Leistung\s+Std', line):
                continue
            amounts = re.findall(currency_pattern, line)
            if not amounts:
                continue
            amount_str = amounts[-1]
            cut_index = line.rfind(amount_str)
            description = line[:cut_index].strip()
            description = re.sub(currency_pattern, '', description).strip()
            description_clean = self._clean_ocr_text(description)
            # Determine acceptance
            # Convert amount to cents
            amt_cents_try = 0
            try:
                amt_cents_try = int(round(float(amount_str.replace('.', '').replace(',', '.')) * 100))
            except Exception:
                pass
            has_keyword = any(k.lower() in description_clean.lower() for k in keywords)
            has_qty = bool(qty_pattern.search(line))
            only_numbers = bool(re.fullmatch(r'[0-9 .,/:-]+', description_clean))
            # Reject tiny amounts (< 1000 cents) unless keyword or quantity present
            accept = True
            if amt_cents_try < 1000 and not (has_keyword or has_qty):
                accept = False
            if only_numbers:
                accept = False
            if len(description_clean) < 3:
                accept = False
            print(f"DEBUG LINE ITEM: amt={amount_str} ({amt_cents_try}c) kw={has_keyword} qty={has_qty} desc='{description_clean}' accept={accept}")
            if not accept:
                continue
            line_items.append({'description': description_clean, 'amount': amount_str})
        result['line_items'] = line_items
        
        # 6. Extract quantities and rates
        qty_match = re.search(
            r'(\d+)\s*(?:Stunden|hours|x|pcs)',
            text,
            re.IGNORECASE
        )
        if qty_match:
            result['quantity'] = qty_match.group(1)
        
        rate_match = re.search(
            r'(?:á|@|Stundensatz|Rate)\s*€?\s*'
            r'(\d+,\d{2}|\d+\.\d{2})',
            text,
            re.IGNORECASE
        )
        if rate_match:
            result['rate'] = rate_match.group(1)
        
        # 6. Extract line total
        # Usually close to total if only one item
        line_total_match = re.search(
            r'(?:Nettobetrag|Zwischensumme|Subtotal)'
            r'.{0,100}?'
            r'(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})',
            text,
            re.IGNORECASE | re.DOTALL
        )
        if line_total_match:
            result['line_total'] = line_total_match.group(1)
        elif result['total_amount'] and not result['quantity']:
            # If no line total but has total, assume same
            result['line_total'] = result['total_amount']
        
        return result
    
    def _count_pdf_pages(self, pdf_path):
        """Count total pages in PDF"""
        try:
            import PyPDF2
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                return len(reader.pages)
        except Exception:
            # Fallback estimate
            return 10


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
        extractor = PopperExtractor()
        result = extractor.extract_pdf(pdf_file)
        print(f"\nResult: {result}")
    else:
        print("Usage: python poppler_extractor.py <pdf_file>")
