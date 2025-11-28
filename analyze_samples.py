
import os
import pdfplumber
import pandas as pd

def log(msg):
    print(msg)
    with open("analysis_result.txt", "a", encoding="utf-8") as f:
        f.write(msg + "\n")

def analyze_pdf(path):
    log(f"\n--- Analyzing PDF: {os.path.basename(path)} ---")
    try:
        with pdfplumber.open(path) as pdf:
            if len(pdf.pages) > 0:
                text = pdf.pages[0].extract_text()
                if text and len(text.strip()) > 50:
                    log("✅ Text detected! Standard extraction might work.")
                    log(f"Sample text:\n{text[:200]}...")
                else:
                    log("⚠️ No text detected. OCR likely needed.")
            else:
                log("⚠️ Empty PDF.")
    except Exception as e:
        log(f"❌ Error reading PDF: {e}")

def analyze_excel(path):
    log(f"\n--- Analyzing Excel: {os.path.basename(path)} ---")
    try:
        df = pd.read_excel(path)
        log("✅ Columns found:")
        for col in df.columns:
            log(f"  - {col}")
        log(f"Rows: {len(df)}")
    except Exception as e:
        log(f"❌ Error reading Excel: {e}")

# Clear previous log
with open("analysis_result.txt", "w", encoding="utf-8") as f:
    f.write("Starting analysis...\n")

print("Starting analysis...")
files = os.listdir('temp_analysis')
for f in files:
    path = os.path.join('temp_analysis', f)
    if f.lower().endswith('.pdf'):
        analyze_pdf(path)
    elif f.lower().endswith('.xlsx'):
        analyze_excel(path)
