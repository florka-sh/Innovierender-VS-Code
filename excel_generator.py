"""
Excel Generator Module
Generates Excel files from extracted invoice data in the required accounting format.
"""

import pandas as pd
from datetime import datetime
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)

# Define column names as constant to avoid duplication
EXCEL_COLUMNS = [
    'SATZART', 'FIRMA', 'BELEG_NR', 'BELEG_DAT', 'SOLL_HABEN',
    'BUCH_KREIS', 'BUCH_JAHR', 'BUCH_MONAT', 'DEBI_KREDI',
    'BETRAG', 'RECHNUNG', 'leer', 'BUCH_TEXT', 'HABENKONTO',
    'SOLLKONTO', 'leer_1', 'KOSTSTELLE', 'KOSTTRAGER',
    'Kostenträgerbezeichnung', 'Bebuchbar',
    'Debitoren.Bezeichnung',
    'Debitoren.Aktuelle Anschrift Anschrift-Zusatz',
    'AbgBenutzerdefiniert'
]


def _parse_invoice_date(
    invoice_date: str
) -> tuple:
    """
    Parse and convert invoice date.
    
    Returns:
        Tuple of (beleg_dat, buch_jahr, buch_monat)
    """
    if not invoice_date:
        return '', '', ''
    
    try:
        date_obj = datetime.strptime(invoice_date, '%d.%m.%Y')
        return (
            date_obj.strftime('%Y%m%d'),
            date_obj.year,
            date_obj.month
        )
    except ValueError:
        logger.warning(
            f'Invalid date format: {invoice_date}'
        )
        return '', '', ''


def _build_booking_text(
    config: Dict[str, Any],
    item: Dict[str, Any]
) -> str:
    """Build booking text from configuration and item data."""
    parts = []
    
    if config.get('BUCH_TEXT_PREFIX'):
        parts.append(config['BUCH_TEXT_PREFIX'])
    if item.get('student_name'):
        parts.append(item['student_name'])
    if item.get('subject'):
        parts.append(item['subject'])
    if item.get('school'):
        parts.append(f"({item['school']})")
    
    return ' '.join(parts) if parts else ''


def _create_row(
    item: Dict[str, Any],
    config: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Create a single row for the Excel file.
    
    Args:
        item: Invoice line item
        config: Configuration dictionary
        
    Returns:
        Dictionary representing one Excel row
    """
    beleg_dat, buch_jahr, buch_monat = _parse_invoice_date(
        item.get('invoice_date', '')
    )
    
    # Convert amount to cents
    betrag = (
        int(item.get('amount', 0) * 100)
        if item.get('amount')
        else 0
    )
    
    buch_text = _build_booking_text(config, item)
    
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
        'Kostenträgerbezeichnung': config.get(
            'Kostenträgerbezeichnung', ''
        ),
        'Bebuchbar': config.get('Bebuchbar', 'Ja'),
        'Debitoren.Bezeichnung': None,
        'Debitoren.Aktuelle Anschrift Anschrift-Zusatz': None,
        'AbgBenutzerdefiniert': None
    }
    
    return row


def generate_excel(
    invoice_data: List[Dict[str, Any]],
    output_path: str,
    config: Dict[str, Any]
) -> pd.DataFrame:
    """
    Generate Excel file from invoice data.
    
    Args:
        invoice_data: List of extracted invoice line items
        output_path: Path where Excel file should be saved
        config: Configuration dictionary with default column values
        
    Returns:
        Generated DataFrame
    """
    try:
        rows = [_create_row(item, config) for item in invoice_data]
        
        # Create DataFrame
        df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
        
        # Save to Excel
        df.to_excel(output_path, index=False, engine='openpyxl')
        logger.info(f'Excel file generated: {output_path}')
        
        return df
    except Exception as e:
        logger.error(f'Error generating Excel: {str(e)}')
        raise


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
