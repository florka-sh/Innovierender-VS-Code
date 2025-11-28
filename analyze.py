import pandas as pd
import pdfplumber
import sys

# Redirect output to file
output_file = open('analysis_output.txt', 'w', encoding='utf-8')
sys.stdout = output_file

# Analyze Excel
print("=" * 70)
print("EXCEL FILE STRUCTURE")
print("=" * 70)
df = pd.read_excel('9251_1025_Lernforderung Solingen Fibuübernahmepaket.xlsx')
print(f"\nShape: {df.shape[0]} rows, {df.shape[1]} columns")
print(f"\nColumn names:")
for i, col in enumerate(df.columns, 1):
    print(f"  {i}. {col}")

print(f"\nSample data (first 5 rows):")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
print(df.head(5).to_string())

# Analyze PDF
print("\n\n" + "=" * 70)
print("PDF FILE STRUCTURE")
print("=" * 70)
with pdfplumber.open('RE_1155500316-325.pdf') as pdf:
    print(f"\nTotal pages: {len(pdf.pages)}")
    
    # Check first page
    page = pdf.pages[0]
    print(f"\nFirst page text (first 2000 chars):")
    text = page.extract_text()
    print(text[:2000])
    
    # Check for tables
    tables = page.extract_tables()
    print(f"\n\nTables found on first page: {len(tables)}")
    if tables:
        for i, table in enumerate(tables):
            print(f"\n--- Table {i+1} ---")
            print(f"Rows: {len(table)}, Columns: {len(table[0]) if table else 0}")
            for r, row in enumerate(table[:8]):  # Show first 8 rows
                print(f"Row {r}: {row}")

output_file.close()
print("Analysis complete. Output saved to analysis_output.txt", file=sys.__stdout__)
