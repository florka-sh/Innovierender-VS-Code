import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import re
import os
from PIL import Image
import sys

# NOTE: You might need to set the tesseract path if it's not in your PATH
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class InvoiceExtractor:
    def __init__(self, poppler_path=None, tesseract_cmd=None):
        # Try to find tesseract if not explicitly provided
        if not tesseract_cmd:
            # Determine base directory (handle PyInstaller frozen state)
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))

            # Common default locations on Windows
            possible_paths = [
                os.path.join(base_dir, 'Tesseract-OCR', 'tesseract.exe'),
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                os.path.join(os.getenv('LOCALAPPDATA', ''), 'Tesseract-OCR', 'tesseract.exe')
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    tesseract_cmd = path
                    break
        
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
            
        self.poppler_path = poppler_path

    def extract_invoices_from_pdf(self, pdf_path):
        """
        Converts PDF to images and extracts data from EACH page as a separate invoice.
        Returns a list of data dictionaries.
        """
        results = []
        try:
            images = convert_from_path(pdf_path, poppler_path=self.poppler_path)
            
            for i, image in enumerate(images):
                try:
                    # Try German + English first
                    text = pytesseract.image_to_string(image, lang='deu+eng')
                except pytesseract.TesseractError as e:
                    if "data file" in str(e) and "deu" in str(e):
                        print("Warning: German language pack not found. Falling back to English.")
                        text = pytesseract.image_to_string(image, lang='eng')
                    else:
                        raise e

                # Parse data for this specific page
                data = self.parse_invoice_data(text)
                data['page_number'] = i + 1
                results.append(data)
            
            return results
        except Exception as e:
            # Return a single error entry if the whole file fails
            return [{"error": str(e)}]

    def extract_line_items(self, text):
        """
        Extracts line items from the text.
        Assumes a table structure: Description | Quantity | Rate | Total
        """
        items = []
        
        # Split text into lines
        lines = text.split('\n')
        
        # Find the start of the table (header)
        start_idx = 0
        header_keywords = ['Leistung', 'Beschreibung', 'Artikel', 'Description', 'Pos.']
        for i, line in enumerate(lines):
            if any(k in line for k in header_keywords) and ('Betrag' in line or 'Gesamt' in line or 'Total' in line or 'Preis' in line):
                start_idx = i + 1
                break
        
        # Regex for line items
        # Pattern 1: Description ... Qty ... Rate ... Total (3 numbers)
        # 30,00 241,34 7.240,20
        pat_3_nums = re.compile(r'^(.+?)\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*$')
        
        # Pattern 2: Description ... Qty ... Total (2 numbers) - e.g. if Rate is missing or merged
        # 30,00 47,28
        pat_2_nums = re.compile(r'^(.+?)\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*$')

        # Pattern 3: Description ... Total (1 number)
        # Deutschlandticket September 2025 ... 38,00
        # Kostenzusage vom 27.8.25 120,00
        # We look for a line ending with a number, and treat the rest as description.
        pat_1_num = re.compile(r'^(.+?)\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*$')

        # Iterate from start_idx until we hit a "Total" line or end
        for i in range(start_idx, len(lines)):
            line = lines[i].strip()
            if not line: continue
            
            # Stop if we hit the footer totals
            if any(k in line.lower() for k in ['rechnungsbetrag', 'gesamtbetrag', 'nettobetrag', 'zahlbetrag', 'mwst', 'überweisen']):
                break
            
            # Try matching patterns (most specific first)
            match3 = pat_3_nums.match(line)
            if match3:
                desc, qty, rate, total = match3.groups()
                items.append({
                    "description": desc.strip(),
                    "quantity": qty,
                    "rate": rate,
                    "line_total": total
                })
                continue
                
            match2 = pat_2_nums.match(line)
            if match2:
                desc, val1, val2 = match2.groups()
                items.append({
                    "description": desc.strip(),
                    "quantity": val1,
                    "rate": "",
                    "line_total": val2
                })
                continue

            match1 = pat_1_num.match(line)
            if match1:
                desc, total = match1.groups()
                # Filter out lines that are likely just dates or noise
                # e.g. "01.09.2025 - 30.09.2025 Seite" might match if "Seite" wasn't there
                # But here we have a number at the end.
                
                # Check if description is just a date range or noise
                if len(desc) < 3: continue
                
                items.append({
                    "description": desc.strip(),
                    "quantity": "1", # Assume 1 if not specified
                    "rate": total,
                    "line_total": total
                })
                continue
                
        return items

    def parse_invoice_data(self, text):
        """
        Attempts to extract common invoice fields and line items.
        """
        data = {
            "invoice_number": None,
            "invoice_suffix": None,
            "name": None,
            "vorname": None,
            "recipient_name": None,
            "date": None,
            "total_amount": None,
            "line_items": []
        }

        # --- 1. Date Extraction ---
        date_pattern = r'(\d{2}[./-]\d{2}[./-]\d{4}|\d{4}[./-]\d{2}[./-]\d{2})'
        dates = re.findall(date_pattern, text)
        if dates:
            data["date"] = dates[0]

        # --- 2. Invoice Number Extraction ---
        # Look for "Rechnungsnummer" or "Rg.-Nr."
        # We want to capture the whole string first, then split if needed.
        inv_keywords = [r'Rg\.-Nr\.', r'Rechnungsnummer', r'Rechnung Nr\.']
        
        for keyword in inv_keywords:
            # Look for the keyword, then capture the text after it
            match = re.search(f"({keyword})", text, re.IGNORECASE)
            if match:
                # Get text after the match
                start_pos = match.end()
                # Take the next 50 chars to find the number
                snippet = text[start_pos:start_pos+50]
                
                # Look for a sequence of digits, letters, slashes, dashes
                # e.g. 126251238/4071433 or INV-2023-001
                candidates = re.findall(r'\b([A-Za-z0-9\-\/]{3,})\b', snippet)
                
                for cand in candidates:
                    # Filter out small noise
                    if len(cand) < 3: continue
                    # Ignore dates like 01.09.2025
                    if re.match(r'\d{2}\.\d{2}\.\d{4}', cand): continue
                    
                    # Check for slash
                    if '/' in cand:
                        parts = cand.split('/')
                        if len(parts[0]) > 3:
                            data["invoice_number"] = parts[0]
                            if len(parts) > 1:
                                data["invoice_suffix"] = parts[1]
                    else:
                        data["invoice_number"] = cand
                    break
                
                if data["invoice_number"]:
                    break
        
        # --- 3. Recipient Name Extraction ---
        # Name ... Geb.-Datum
        # Vorname ...
        
        # Extract Name
        # Look for "Name" followed by text until "Geb.-Datum" or newline
        name_match = re.search(r'Name\s+(.+?)(?:\s+Geb\.-Datum|\n)', text)
        if name_match:
            data["name"] = name_match.group(1).strip()
            
        # Extract Vorname
        vorname_match = re.search(r'Vorname\s+(.+?)(?:\n|$)', text)
        if vorname_match:
            data["vorname"] = vorname_match.group(1).strip()
            
        # Combine
        if data.get("name") or data.get("vorname"):
            last = data.get("name", "")
            first = data.get("vorname", "")
            data["recipient_name"] = f"{last}, {first}".strip(", ")

        # --- 4. Total Amount Extraction ---
        amount_keywords = [
            r'Gesamtbetrag', r'Gesamt', r'Total', r'Betrag', r'Summe', 
            r'Rechnungsbetrag', r'Zahlbetrag', r'Endbetrag'
        ]
        
        found_amount = None
        for keyword in amount_keywords:
            pattern = rf"(?:{keyword}).{{0,100}}?(\d{{1,3}}(?:[.,]\d{{3}})*[.,]\d{{2}})"
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                found_amount = match.group(1)
                if "zahlbetrag" in keyword.lower() or "endbetrag" in keyword.lower() or "rechnungsbetrag" in keyword.lower():
                    break
        
        if not found_amount:
            currency_pattern = r'(\d{1,3}(?:[.]\d{3})*,\d{2})\s*(?:€|EUR)'
            matches = re.findall(currency_pattern, text)
            if matches:
                try:
                    values = []
                    for m in matches:
                        val_str = m.replace('.', '').replace(',', '.')
                        values.append((float(val_str), m))
                    values.sort(key=lambda x: x[0], reverse=True)
                    found_amount = values[0][1]
                except:
                    found_amount = matches[-1]

        data["total_amount"] = found_amount
        
        # --- 5. Line Items ---
        data["line_items"] = self.extract_line_items(text)
        
        return data

    def save_to_excel(self, data_list, output_path):
        """Saves a list of dictionaries to Excel, flattening line items."""
        flat_data = []
        
        for entry in data_list:
            base_info = {
                "Filename": entry.get("filename", ""),
                "Page": entry.get("page_number", ""),
                "Invoice Number": entry.get("invoice_number", ""),
                "Invoice Suffix": entry.get("invoice_suffix", ""),
                "Recipient Name": entry.get("recipient_name", ""),
                "Date": entry.get("date", ""),
                "Total Amount": entry.get("total_amount", "")
            }
            
            if entry.get("line_items"):
                for item in entry["line_items"]:
                    row = base_info.copy()
                    row.update({
                        "Description": item.get("description", ""),
                        "Quantity": item.get("quantity", ""),
                        "Rate": item.get("rate", ""),
                        "Line Total": item.get("line_total", "")
                    })
                    flat_data.append(row)
            else:
                # No line items, just save base info
                flat_data.append(base_info)
                
        df = pd.DataFrame(flat_data)
        
        # Reorder columns nicely
        cols = ["Filename", "Page", "Invoice Number", "Invoice Suffix", "Recipient Name", "Date", "Description", "Quantity", "Rate", "Line Total", "Total Amount"]
        # Only keep columns that exist
        cols = [c for c in cols if c in df.columns]
        df = df[cols]
        
        df.to_excel(output_path, index=False)
        return output_path
