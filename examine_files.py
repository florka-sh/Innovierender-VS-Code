import pdfplumber
import openpyxl
import pandas as pd

# Examine PDF structure
print("=" * 50)
print("PDF STRUCTURE")
print("=" * 50)
with pdfplumber.open("RE_1155500316-325.pdf") as pdf:
    print(f"Number of pages: {len(pdf.pages)}")
    print("\nFirst page text (first 500 chars):")
    print(pdf.pages[0].extract_text()[:500])
    print("\nTables on first page:")
    tables = pdf.pages[0].extract_tables()
    if tables:
        for i, table in enumerate(tables):
            print(f"\nTable {i+1}:")
            print(f"Rows: {len(table)}, Columns: {len(table[0]) if table else 0}")
            if table:
                # Show first few rows
                for row in table[:3]:
                    print(row)

# Examine Excel structure
print("\n\n" + "=" * 50)
print("EXCEL STRUCTURE")
print("=" * 50)
df = pd.read_excel("9251_1025_Lernforderung Solingen Fibuübernahmepaket.xlsx")
print(f"Shape: {df.shape}")
print(f"\nColumns: {list(df.columns)}")
print(f"\nFirst 5 rows:")
print(df.head())
print(f"\nData types:")
print(df.dtypes)
