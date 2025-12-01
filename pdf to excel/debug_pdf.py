import os
from extractor import InvoiceExtractor

def debug_pdf():
    pdf_path = "RE September 25_stationär_1238 - 1316.pdf"
    
    if not os.path.exists(pdf_path):
        print(f"File not found: {pdf_path}")
        return

    # Setup extractor with local poppler
    poppler_path = None
    local_poppler = os.path.join(os.getcwd(), "poppler", "Library", "bin")
    if os.path.exists(local_poppler):
        poppler_path = local_poppler

    extractor = InvoiceExtractor(poppler_path=poppler_path)
    
    print(f"Extracting invoices from: {pdf_path}")
    invoices = extractor.extract_invoices_from_pdf(pdf_path)
    
    # Print first few pages to console
    for i, inv in enumerate(invoices[:5]):  # Just first 5 pages
        print(f"\nPage {i+1}:")
        print(f"  Invoice #: {inv.get('invoice_number')}")
        print(f"  Invoice Suffix: {inv.get('invoice_suffix')}")
        print(f"  Recipient Name: {inv.get('recipient_name')}")
        print(f"  Date: {inv.get('date')}")
        print(f"  Total: {inv.get('total_amount')}")
    
    print(f"\nProcessed {len(invoices)} total pages")

if __name__ == "__main__":
    debug_pdf()
