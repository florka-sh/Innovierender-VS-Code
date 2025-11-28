"""
Test BereitsPF transformation to compare with working output
"""
import pandas as pd

# Read working output
output_df = pd.read_excel('reference_outputs/output_final.xlsx')
print("WORKING OUTPUT STRUCTURE:")
print(f"Columns ({len(output_df.columns)}): {output_df.columns.tolist()}")
print(f"Rows: {len(output_df)}")
print("\nSample row:")
for col in output_df.columns[:10]:
    print(f"  {col}: {output_df.iloc[0][col]}")

# Save to text file for reference
with open('working_output_structure.txt', 'w', encoding='utf-8') as f:
    f.write("WORKING OUTPUT COLUMNS:\n")
    for i, col in enumerate(output_df.columns):
        f.write(f"{i+1}. {col}\n")
    f.write(f"\nTotal columns: {len(output_df.columns)}\n")
    f.write(f"Total rows: {len(output_df)}\n")
    f.write("\nFirst row values:\n")
    for col in output_df.columns:
        f.write(f"  {col}: {output_df.iloc[0][col]}\n")

print("\n✅ Saved structure to working_output_structure.txt")
