"""
Excel Generator Module
Generates Excel files from extracted invoice data in the required accounting format.
"""

import pandas as pd
from datetime import datetime
from typing import List, Dict, Any


def generate_excel(invoice_data: List[Dict[str, Any]], output_path: str, config: Dict[str, Any]) -> None:
    """
    Generate Excel file from invoice data.
    
    Args:
        invoice_data: List of extracted invoice line items
        output_path: Path where Excel file should be saved
        config: Configuration dictionary with default column values
    """
    # Create list of rows for DataFrame
    rows = []
    
    for item in invoice_data:
        # Convert date to YYYYMMDD format
        invoice_date = item.get('invoice_date', '')
        if invoice_date:
            try:
                date_obj = datetime.strptime(invoice_date, '%d.%m.%Y')
                beleg_dat = date_obj.strftime('%Y%m%d')
                buch_jahr = date_obj.year
                buch_monat = date_obj.month
            except ValueError:
                beleg_dat = ''
                buch_jahr = ''
                buch_monat = ''
        else:
            beleg_dat = ''
            buch_jahr = ''
            buch_monat = ''
        
        # Convert amount to cents (multiply by 100)
        betrag = int(item.get('amount', 0) * 100) if item.get('amount') else 0
        
        # Build descriptive text from student, subject, and course info
        buch_text_parts = []
        if config.get('BUCH_TEXT_PREFIX'):
            buch_text_parts.append(config.get('BUCH_TEXT_PREFIX'))
        if item.get('student_name'):
            buch_text_parts.append(item.get('student_name'))
        if item.get('subject'):
            buch_text_parts.append(item.get('subject'))
        if item.get('school'):
            buch_text_parts.append(f"({item.get('school')})")
        
        buch_text = ' '.join(buch_text_parts) if buch_text_parts else ''
        
        # Create row with all 23 columns
        row = {
            'SATZART': config.get('SATZART', 'D'),
            'FIRMA': config.get('FIRMA', ''),
            'BELEG_NR': item.get('invoice_number', ''),
            'BELEG_DAT': beleg_dat,
            'SOLL_HABEN': config.get('SOLL_HABEN', ''),
            'BUCH_KREIS': config.get('BUCH_KREIS', ''),
            'BUCH_JAHR': buch_jahr,
            'BUCH_MONAT': buch_monat,
            'DEBI_KREDI': item.get('customer_number', ''),
            'BETRAG': betrag,
            'RECHNUNG': item.get('invoice_number', ''),
            'leer': None,
            'BUCH_TEXT': buch_text,
            'HABENKONTO': config.get('HABENKONTO', ''),
            'SOLLKONTO': None,
            'leer_1': None,
            'KOSTSTELLE': config.get('KOSTSTELLE', ''),
            'KOSTTRAGER': config.get('KOSTTRAGER', ''),
            'Kostenträgerbezeichnung': config.get('Kostenträgerbezeichnung', ''),
            'Bebuchbar': config.get('Bebuchbar', 'Ja'),
            'Debitoren.Bezeichnung': None,
            'Debitoren.Aktuelle Anschrift Anschrift-Zusatz': None,
            'AbgBenutzerdefiniert': None
        }
        
        rows.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(rows)
    
    # Save to Excel
    df.to_excel(output_path, index=False, engine='openpyxl')
    
    return df


if __name__ == "__main__":
    # Test generation
    from pdf_extractor import extract_invoices
    
    # Extract data
    invoices = extract_invoices("RE_1155500316-325.pdf")
    
    # Test config
    config = {
        'SATZART': 'D',
        'FIRMA': 9251,
        'SOLL_HABEN': 'H',
        'BUCH_KREIS': 'RA',
        'HABENKONTO': 42200,
        'KOSTSTELLE': 190,
        'KOSTTRAGER': '190111512110',
        'Kostenträgerbezeichnung': 'SPFH/HzE Siegen',
        'Bebuchbar': 'Ja',
        'BUCH_TEXT_PREFIX': '1025'
    }
    
    # Generate Excel
    output_file = "test_output.xlsx"
    df = generate_excel(invoices, output_file, config)
    
    print(f"Generated Excel file: {output_file}")
    print(f"Rows generated: {len(df)}")
    print("\nFirst row:")
    print(df.head(1).to_dict('records')[0])
