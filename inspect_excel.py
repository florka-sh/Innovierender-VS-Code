
import pandas as pd
import os

path = [f for f in os.listdir('temp_analysis') if 'Auszahlungsbelege' in f][0]
full_path = os.path.join('temp_analysis', path)

print(f"Reading: {path}")
df = pd.read_excel(full_path, header=None)
print(df.head(15))
