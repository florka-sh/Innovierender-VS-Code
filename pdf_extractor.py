"""
PDF Invoice Extractor Module
Extracts invoice data from PDF files for learning support (Lernförderung) invoices.
"""

import pdfplumber
import re
from datetime import datetime
from typing import List, Dict, Any


def extract_invoices(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Extract all invoice data from a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        List of dictionaries containing extracted invoice data
    """
    invoices = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            try:
                invoice_data = extract_page_data(page, page_num)
                if invoice_data:
                    invoices.extend(invoice_data)
            except Exception as e:
                print(f"Error processing page {page_num}: {str(e)}")
                continue
    
    return invoices


def extract_page_data(page, page_num: int) -> List[Dict[str, Any]]:
    """
    Extract invoice data from a single page.
    
    Args:
        page: pdfplumber page object
        page_num: Page number
        
    Returns:
        List of invoice line items
    """
    text = page.extract_text()
    tables = page.extract_tables()
    
    # Extract metadata from the first table (invoice header)
    metadata = {}
    if tables and len(tables) > 0:
        header_table = tables[0]
        for row in header_table:
            if row[0] and 'Rechnungsnummer' in row[0]:
                metadata['invoice_number'] = row[1].strip() if row[1] else None
            elif row[0] and 'Rechnungsdatum' in row[0]:
                metadata['invoice_date'] = row[1].strip() if row[1] else None
            elif row[0] and 'Kunden-Nummer' in row[0]:
                metadata['customer_number'] = row[1].replace(' ', '') if row[1] else None
    
    # Extract customer number from text if not found in table
    if not metadata.get('customer_number'):
        customer_match = re.search(r'Kunden-Nummer:\s*([0-9\s]+)', text)
        if customer_match:
            metadata['customer_number'] = customer_match.group(1).replace(' ', '')
    
    # Extract invoice number from text if not found
    if not metadata.get('invoice_number'):
        inv_match = re.search(r'Rechnungsnummer:\s*(\d+)', text)
        if inv_match:
            metadata['invoice_number'] = inv_match.group(1)
    
    # Extract date from text if not found
    if not metadata.get('invoice_date'):
        date_match = re.search(r'Rechnungsdatum:\s*(\d{2}\.\d{2}\.\d{4})', text)
        if date_match:
            metadata['invoice_date'] = date_match.group(1)
    
    # Extract line items from the second table (invoice details)
    line_items = []
    if tables and len(tables) > 1:
        detail_table = tables[1]
        
        # Find student name, school, and course details
        student_info = extract_student_info(text)
        
        for row in detail_table[1:]:  # Skip header row
            if not row or not row[0]:
                continue
                
            # Check if this is a line item row (has line number)
            if row[0] and row[0].strip().isdigit():
                # Extract course details from the middle column
                course_text = row[1] if len(row) > 1 and row[1] else ""
                amount_text = row[2] if len(row) > 2 and row[2] else ""
                
                # Parse course lines (e.g., "10/25 Deutsch 1 Zeitstunden á 25,00 €")
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
                    }
                    line_items.append(item)
    
    return line_items


def extract_student_info(text: str) -> Dict[str, str]:
    """
    Extract student name, school, and account code from invoice text.
    
    Args:
        text: Full page text
        
    Returns:
        Dictionary with student information
    """
    info = {}
    
    # Extract student name (line after "Durchführung einer Lernförderung für:")
    name_match = re.search(r'Durchführung einer Lernförderung für:\s*\n([^\n]+)', text)
    if name_match:
        info['student_name'] = name_match.group(1).strip()
    
    # Extract school name (typically the line after student name)
    school_match = re.search(r'Durchführung einer Lernförderung für:\s*\n[^\n]+\s*\n([^\n]+)', text)
    if school_match:
        school_line = school_match.group(1).strip()
        # Check if it looks like a school name (not a course line)
        if not re.match(r'\d{2}/\d{2}', school_line):
            info['school'] = school_line
    
    # Extract account code (pattern like "42100/0111 121512520")
    code_match = re.search(r'(\d{5}/\d{4}\s+\d+)', text)
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
    
    # Pattern: "10/25 Deutsch 1 Zeitstunden á 25,00 €" or similar
    # Also handles amounts like "25,00 €" on separate lines
    lines = course_text.split('\n')
    
    for line in lines:
        # Match pattern: MM/YY Subject Hours Zeitstunden á Rate €
        match = re.search(r'(\d{2}/\d{2})\s+([^\d]+?)\s+(\d+)\s+Zeitstunden?\s+á\s+([\d,]+)\s*€', line)
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
