from PIL import Image, ImageDraw, ImageFont
import os
from extractor import InvoiceExtractor

def create_dummy_pdf(filename):
    # Create a white image
    img = Image.new('RGB', (800, 1000), color='white')
    d = ImageDraw.Draw(img)
    
    # Add some text that looks like an invoice
    # Note: We rely on default font, which might be small, but Tesseract should handle it.
    d.text((50, 50), "Rechnung", fill='black')
    d.text((50, 100), "Rechnungsnummer: INV-2023-001", fill='black')
    d.text((50, 150), "Datum: 23.11.2025", fill='black')
    d.text((50, 300), "Beschreibung   Menge   Preis", fill='black')
    d.text((50, 330), "Service A      1       100.00", fill='black')
    d.text((50, 360), "Service B      2       50.00", fill='black')
    d.text((50, 500), "Gesamtbetrag: 200,00 EUR", fill='black')
    
    img.save(filename)
    print(f"Created dummy PDF: {filename}")
    return filename

def test_extraction():
    pdf_file = "test_invoice.pdf"
    create_dummy_pdf(pdf_file)
    
    # Check for local Poppler
    poppler_path = None
    base_dir = os.path.dirname(os.path.abspath(__file__))
    local_poppler = os.path.join(base_dir, "poppler", "Library", "bin")
    if os.path.exists(local_poppler):
        print(f"Using local Poppler: {local_poppler}")
        poppler_path = local_poppler

    extractor = InvoiceExtractor(poppler_path=poppler_path)
    
    print("Attempting extraction...")
    # New API returns a list of dicts
    results = extractor.extract_invoices_from_pdf(pdf_file)
    
    if not results:
        print("No results returned.")
        return

    if "error" in results[0]:
        print("Extraction Failed:")
        print(results[0]["error"])
        return

    # We only have 1 page in dummy PDF
    data = results[0]
    print("Parsed Data:", data)
    
    # Basic assertions
    if data.get("invoice_number") == "INV-2023-001":
        print("SUCCESS: Invoice Number matched.")
    else:
        print(f"FAILURE: Invoice Number mismatch. Got {data.get('invoice_number')}")

    if data.get("total_amount") == "200,00":
        print("SUCCESS: Total Amount matched.")
    else:
        print(f"FAILURE: Total Amount mismatch. Got {data.get('total_amount')}")

if __name__ == "__main__":
    test_extraction()
