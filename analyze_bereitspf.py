import pandas as pd
import os

print("=" * 80)
print("ANALYZING BEREITSPF WORKING OUTPUTS")
print("=" * 80)

# Analyze try.xlsx
print("\n--- try.xlsx ---")
try:
    df_try = pd.read_excel('reference_outputs/try.xlsx')
    print(f"Columns: {df_try.columns.tolist()}")
    print(f"Rows: {len(df_try)}")
    print(f"\nFirst row sample:")
    print(df_try.iloc[0].to_dict())
except Exception as e:
    print(f"Error: {e}")

# Analyze output_final.xlsx
print("\n" + "=" * 80)
print("\n--- output_final.xlsx ---")
try:
    df_output = pd.read_excel('reference_outputs/output_final.xlsx')
    print(f"Columns: {df_output.columns.tolist()}")
    print(f"Rows: {len(df_output)}")
    print(f"\nFirst row sample:")
    print(df_output.iloc[0].to_dict())
except Exception as e:
    print(f"Error: {e}")

# Analyze source data
print("\n" + "=" * 80)
print("\n--- Auszahlungsbelege Pflegefamilien 11.2025.xlsx (Source) ---")
try:
    source_path = r'C:\Users\Florin Sherban\Desktop\data_analyts\BereitsPF\Auszahlungsbelege Pflegefamilien 11.2025.xlsx'
    df_source = pd.read_excel(source_path)
    print(f"Columns: {df_source.columns.tolist()}")
    print(f"Rows: {len(df_source)}")
    print(f"\nFirst few rows:")
    print(df_source.head(3))
except Exception as e:
    print(f"Error: {e}")
