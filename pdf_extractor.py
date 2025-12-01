"""
PDF Invoice Extractor Module.
Extracts Lernförderung invoice data from PDF files.
"""

import pdfplumber
import re
from datetime import datetime
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)

# Pre-compile regex patterns for performance
INVOICE_NUMBER_PATTERN = re.compile(r'Rechnun[gq]snummer:\s*(\d+)')
CUSTOMER_NUMBER_PATTERN = re.compile(r'Kunden-Nummer:\s*([0-9\s]+)')
INVOICE_DATE_PATTERN = re.compile(
    r'Rechnun[gq]sdatum:\s*(\d{2}\.\d{2}\.\d{4})'
)
COURSE_PATTERN = re.compile(
    r'(\d{2}/\d{2})\s+([^\d]+?)\s+(\d+)\s+Zeitstunden?\s+á\s+'
    r'([\d,]+)\s*€'
)
ACCOUNT_CODE_PATTERN = re.compile(r'(\d{5}/\d{4}\s+\d+)')
STUDENT_NAME_PATTERN = re.compile(
    r'Durchführung einer Lernförderung für:\s*\n([^\n]+)'
)
SCHOOL_PATTERN = re.compile(
    r'Durchführung einer Lernförderung für:\s*\n[^\n]+\s*\n([^\n]+)'
)

# Bereitschaftspflege (care) invoice patterns
CARE_KEYWORD_PATTERN = re.compile(r'Bereitschaftspflege', re.IGNORECASE)
CARE_LINES_START_PATTERN = re.compile(
    r'Folgende Leistungen wurden erbracht', re.IGNORECASE
)
CARE_LINE_ITEM_PATTERN = re.compile(
    r'^\s*(\d+)\s+(.+?)\s+([0-9\.,]+)\s*€\s*$', re.MULTILINE
)
CARE_NETTO_PATTERN = re.compile(r'Nettobetrag\s*([0-9\.]+,\d{2})')
CARE_DEDUCTION_PATTERN = re.compile(
    r'abzüg.*?(-?[0-9\.]+,\d{2})', re.IGNORECASE
)
CARE_TOTAL_PATTERN = re.compile(r'Rechnungsbetrag\s*([0-9\.]+,\d{2})')
CARE_PAYMENT_PATTERN = re.compile(r'Zahlbetrag\s*([0-9\.]+,\d{2})')


def extract_invoices(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Extract all invoice data from a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        List of dictionaries containing extracted invoice data
    """
    invoices = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                try:
                    invoice_data = extract_page_data(page, page_num)
                    if invoice_data:
                        invoices.extend(invoice_data)
                except Exception as e:
                    logger.warning(
                        f"Error processing page {page_num}: {str(e)}"
                    )
                    continue
    except Exception as e:
        logger.error(f"Error opening PDF {pdf_path}: {str(e)}")
        raise
    
    return invoices


def extract_page_data(page, page_num: int) -> List[Dict[str, Any]]:
    """
    Extract invoice data from a single page.
    Supports both Lernförderung and Bereitschaftspflege formats.
    
    Args:
        page: pdfplumber page object
        page_num: Page number
        
    Returns:
        List of invoice line items
    """
    text = page.extract_text() or ''
    tables = page.extract_tables()
    
    # Detect invoice type
    is_care_invoice = bool(CARE_KEYWORD_PATTERN.search(text))
    metadata = _extract_metadata_from_table(tables, text)
    
    # For care invoices, extract metadata from text patterns
    if is_care_invoice:
        if not metadata.get('invoice_number'):
            inv_match = INVOICE_NUMBER_PATTERN.search(text)
            if inv_match:
                metadata['invoice_number'] = inv_match.group(1).strip()
        if not metadata.get('invoice_date'):
            date_match = INVOICE_DATE_PATTERN.search(text)
            if date_match:
                metadata['invoice_date'] = date_match.group(1).strip()
        if not metadata.get('customer_number'):
            cust_match = CUSTOMER_NUMBER_PATTERN.search(text)
            if cust_match:
                metadata['customer_number'] = (
                    cust_match.group(1).replace(' ', '').strip()
                )
    
    line_items = []
    
    if is_care_invoice:
        # Parse Bereitschaftspflege format
        lines = text.splitlines()
        collecting = False
        
        for line in lines:
            if CARE_LINES_START_PATTERN.search(line):
                collecting = True
                continue
            if not collecting:
                continue
            if CARE_NETTO_PATTERN.search(line):
                collecting = False
                continue
                
            m = CARE_LINE_ITEM_PATTERN.match(line)
            if not m:
                continue
                
            line_no, desc, amt = m.groups()
            desc_clean = desc.strip()
            
            # Remove daily rate from description (e.g., "94,02€")
            desc_clean = re.sub(r'\s*[0-9\.,]+€', '', desc_clean).strip()
            
            # Extract month/year if present
            month_year_match = re.search(
                r'\b([A-Za-z]{3}\s+\d{2})\b', desc_clean
            )
            month_year = month_year_match.group(1) if month_year_match else ''
            
            def to_float(s: str) -> float:
                if not s:
                    return 0.0
                cleaned = s.replace('€', '').replace('.', '')
                cleaned = cleaned.replace(',', '.').strip()
                try:
                    return float(cleaned)
                except ValueError:
                    return 0.0
            
            amount_val = to_float(amt)
            
            item = {
                'page_num': page_num,
                'invoice_number': metadata.get('invoice_number'),
                'invoice_date': metadata.get('invoice_date'),
                'customer_number': metadata.get('customer_number'),
                'student_name': '',
                'school': '',
                'month_year': month_year,
                'subject': desc_clean,
                'hours': '',
                'rate': '',
                'amount': amount_val,
                'account_code': '',
                'care_mode': True
            }
            line_items.append(item)
        
        # Extract totals
        def parse_amount(pattern):
            m = pattern.search(text)
            if not m:
                return 0.0
            return float(m.group(1).replace('.', '').replace(',', '.'))
        
        netto = parse_amount(CARE_NETTO_PATTERN)
        total = parse_amount(CARE_TOTAL_PATTERN)
        deduction = parse_amount(CARE_DEDUCTION_PATTERN)
        zahlbetrag = parse_amount(CARE_PAYMENT_PATTERN)
        
        if any([netto, total, zahlbetrag]):
            summary = {
                'page_num': page_num,
                'invoice_number': metadata.get('invoice_number'),
                'invoice_date': metadata.get('invoice_date'),
                'customer_number': metadata.get('customer_number'),
                'student_name': '',
                'school': '',
                'month_year': '',
                'subject': 'SUMME (Bereitschaftspflege)',
                'hours': '',
                'rate': '',
                'amount': zahlbetrag if zahlbetrag else (
                    total if total else netto
                ),
                'account_code': '',
                'care_mode': True,
                'netto': netto,
                'rechnungsbetrag': total,
                'abschlag': deduction,
                'zahlbetrag': zahlbetrag if zahlbetrag else total
            }
            line_items.append(summary)
        
        return line_items
    
    # Lernförderung parsing (existing logic)
    if tables and len(tables) > 1:
        detail_table = tables[1]
        student_info = extract_student_info(text)
        
        for row in detail_table[1:]:
            if not row or not row[0]:
                continue
            
            if row[0] and row[0].strip().isdigit():
                course_text = row[1] if len(row) > 1 and row[1] else ""
                courses = parse_courses(course_text)
                
                for course in courses:
                    item = {
                        'page_num': page_num,
                        'invoice_number': metadata.get('invoice_number'),
                        'invoice_date': metadata.get('invoice_date'),
                        'customer_number': metadata.get('customer_number'),
                        'student_name': student_info.get('student_name'),
                        'school': student_info.get('school'),
                        'month_year': course.get('month_year'),
                        'subject': course.get('subject'),
                        'hours': course.get('hours'),
                        'rate': course.get('rate'),
                        'amount': course.get('amount'),
                        'account_code': student_info.get('account_code'),
                        'care_mode': False
                    }
                    line_items.append(item)
    
    return line_items


def _extract_metadata_from_table(
    tables: List, text: str
) -> Dict[str, str]:
    """Extract invoice metadata from tables and text."""
    metadata = {}
    if tables and len(tables) > 0:
        header_table = tables[0]
        for row in header_table:
            if not row[0]:
                continue
            if 'Rechnungsnummer' in row[0]:
                metadata['invoice_number'] = (
                    row[1].strip() if row[1] else None
                )
            elif 'Rechnungsdatum' in row[0]:
                metadata['invoice_date'] = (
                    row[1].strip() if row[1] else None
                )
            elif 'Kunden-Nummer' in row[0]:
                metadata['customer_number'] = (
                    row[1].replace(' ', '') if row[1] else None
                )
    
    # Extract from text if not found in table
    if not metadata.get('customer_number'):
        match = CUSTOMER_NUMBER_PATTERN.search(text)
        if match:
            metadata['customer_number'] = match.group(1).replace(
                ' ', ''
            )
    
    if not metadata.get('invoice_number'):
        match = INVOICE_NUMBER_PATTERN.search(text)
        if match:
            metadata['invoice_number'] = match.group(1)
    
    if not metadata.get('invoice_date'):
        match = INVOICE_DATE_PATTERN.search(text)
        if match:
            metadata['invoice_date'] = match.group(1)
    
    return metadata


def extract_student_info(text: str) -> Dict[str, str]:
    """
    Extract student name, school, and account code from invoice text.
    
    Args:
        text: Full page text
        
    Returns:
        Dictionary with student information
    """
    info = {}
    
    # Extract student name
    name_match = STUDENT_NAME_PATTERN.search(text)
    if name_match:
        info['student_name'] = name_match.group(1).strip()
    
    # Extract school name
    school_match = SCHOOL_PATTERN.search(text)
    if school_match:
        school_line = school_match.group(1).strip()
        # Check if it looks like a school name (not a course line)
        if not re.match(r'\d{2}/\d{2}', school_line):
            info['school'] = school_line
    
    # Extract account code
    code_match = ACCOUNT_CODE_PATTERN.search(text)
    if code_match:
        info['account_code'] = code_match.group(1).strip()
    
    return info


def parse_courses(course_text: str) -> List[Dict[str, Any]]:
    """
    Parse course details from invoice line item text.
    
    Args:
        course_text: Text containing course information
        
    Returns:
        List of course dictionaries
    """
    courses = []
    lines = course_text.split('\n')
    
    for line in lines:
        # Match pattern: MM/YY Subject Hours Zeitstunden á Rate €
        match = COURSE_PATTERN.search(line)
        if match:
            month_year = match.group(1)
            subject = match.group(2).strip()
            hours = int(match.group(3))
            rate = float(match.group(4).replace(',', '.'))
            amount = hours * rate
            
            courses.append({
                'month_year': month_year,
                'subject': subject,
                'hours': hours,
                'rate': rate,
                'amount': amount
            })
    
    return courses


def convert_date_to_yyyymmdd(date_str: str) -> str:
    """
    Convert date from DD.MM.YYYY to YYYYMMDD format.
    
    Args:
        date_str: Date string in DD.MM.YYYY format
        
    Returns:
        Date string in YYYYMMDD format
    """
    if not date_str:
        return None
    
    try:
        date_obj = datetime.strptime(date_str, '%d.%m.%Y')
        return date_obj.strftime('%Y%m%d')
    except ValueError:
        return None


if __name__ == "__main__":
    # Test extraction
    import sys
    
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
    else:
        pdf_file = "RE_1155500316-325.pdf"
    
    print(f"Extracting data from: {pdf_file}")
    invoices = extract_invoices(pdf_file)
    
    print(f"\nFound {len(invoices)} line items:")
    for i, invoice in enumerate(invoices, 1):
        print(f"\n--- Item {i} ---")
        for key, value in invoice.items():
            print(f"  {key}: {value}")
