"""
Simple script to generate Excel from PDF
Run this to quickly create an Excel file without using the GUI
"""

from pdf_extractor import extract_invoices
from excel_generator import generate_excel

# Configure your parameters here
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

# Input and output files
pdf_file = "RE_1155500316-325.pdf"  # Change this to your PDF file
output_file = "my_export.xlsx"       # Change this to your desired output name

print(f"📄 Extracting data from: {pdf_file}")
invoices = extract_invoices(pdf_file)
print(f"✅ Found {len(invoices)} entries")

print(f"\n📊 Generating Excel file: {output_file}")
generate_excel(invoices, output_file, config)
print(f"✅ Excel file created successfully!")
print(f"\n📁 Location: {output_file}")
