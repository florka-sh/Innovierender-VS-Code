# OCR Field Validation & Manual Correction Guide

## Overview
Enhanced OCR extraction now validates extracted invoice data against expected patterns from the Excel template (`9261_1025_JHV_Wesel_Fibuübernahmepaket_DEB.xlsx`) and highlights rows requiring manual review.

## Key Features

### 1. Pattern Validation
The system validates these critical fields:

| Field | Pattern | Example |
|-------|---------|---------|
| **BELEG_NR** (Invoice Number) | 9-digit numeric | `126251373` |
| **BELEG_DAT** (Date) | YYYYMMDD format | `20251104` |
| **DEBI_KREDI** (Debitor Number) | 9-digit numeric | `214211191` |
| **BETRAG** (Amount) | Numeric cents | `646722` (= 6467.22€) |
| **BUCH_TEXT** (Booking Text) | `<code> TG <loc> <name>` | `1025 TG OB Fraatz, Jannick` |
| **KOSTTRAGER** (Cost Bearer) | 12-digit numeric | `251111611410` |

### 2. Confidence Scoring
Each extracted field receives a confidence score (0-100%):
- **Green (80-100%)**: High confidence, pattern matches
- **Yellow (50-79%)**: Medium confidence, may need review
- **Red (<50%)**: Low confidence, **requires manual entry**

### 3. Visual Feedback
In the desktop app (`pdf_extractor_app.py`):
- **Yellow highlighted rows**: OCR confidence <70% or pattern mismatch
- **Normal rows**: Validation passed
- **Editable cells**: Double-click any cell to manually correct

## How to Use

### Step 1: Run OCR Extraction
1. Open `pdf_extractor_app.py`
2. Select **"Scanned PDF (OCR)"** mode
3. Click **"📁 PDF wählen"** and select your scanned invoice PDF
4. Click **"🔎 OCR Extrahieren"**
5. Wait for processing (progress bar shows status)

### Step 2: Review Validation Results
After extraction:
- **Yellow rows** = Need manual review
- Check these fields in yellow rows:
  - BELEG_NR (must be 9 digits)
  - BELEG_DAT (must be YYYYMMDD)
  - DEBI_KREDI (must be 9 digits)
  - BETRAG (must be numeric)
  - BUCH_TEXT (should follow pattern)

### Step 3: Manual Correction
For yellow/flagged rows:
1. **Double-click** the cell with incorrect data
2. **Type** the correct value following the pattern
3. **Press Enter** to confirm
4. Repeat for all flagged fields

### Step 4: Export
Once all validations pass:
1. Click **"💾 Excel speichern"**
2. Choose save location
3. File exports in correct template format

## Validation Rules

### Invoice Number (BELEG_NR)
```
✅ Valid: 126251373 (9 digits)
❌ Invalid: 12625137 (8 digits), RG-126251373 (contains letters)
```

### Date (BELEG_DAT)
```
✅ Valid: 20251104 (YYYYMMDD)
❌ Invalid: 04.11.2025 (dots), 2025-11-04 (dashes), 04112025 (wrong order)
```

### Debitor Number (DEBI_KREDI)
```
✅ Valid: 214211191 (9 digits)
❌ Invalid: 21421119 (8 digits), D-214211191 (prefix)
```

### Amount (BETRAG)
```
✅ Valid: 646722 (cents, = 6467.22€)
✅ Valid: 2301 (= 23.01€)
❌ Invalid: 6467,22 (comma), 6.467,22 (thousand separator)
```

### Booking Text (BUCH_TEXT)
```
✅ Valid: "1025 TG OB Fraatz, Jannick"
✅ Valid: "1025 TG OB Amoako Boafo, Osagie Ryan"
⚠️ Acceptable: Any text >10 chars (manual entry allowed)
```

## Common Issues & Fixes

### Issue: All rows yellow
**Cause**: OCR confidence low (<70%)  
**Fix**: Check Tesseract/Poppler installation, or manually correct flagged fields

### Issue: BELEG_NR has letters
**Cause**: OCR misread digits as letters (e.g., "0" → "O")  
**Fix**: Double-click cell, type correct 9-digit number

### Issue: BETRAG wrong format
**Cause**: OCR extracted "1.234,56" instead of cents  
**Fix**: Convert manually: `1.234,56 € → 123456` (remove dots/commas, multiply by 100)

### Issue: BUCH_TEXT empty
**Cause**: Recipient name not extracted by OCR  
**Fix**: Type manually: `<code> TG <location> <Lastname, Firstname>`

## Technical Details

### Files Modified
- **`ocr_analysis/field_validators.py`**: New validation module with pattern matching
- **`ocr_analysis/poppler_extractor.py`**: OCR extractor enhanced with confidence tracking
- **`pdf_extractor_app.py`**: UI updated to highlight validation issues and enable manual entry

### Validation Flow
1. OCR extracts text → confidence score calculated
2. Fields parsed using regex patterns
3. Each field validated against template patterns
4. Invalid/low-confidence fields flagged
5. UI highlights rows needing review
6. User corrects via double-click edit
7. Export generates clean Excel

### Extending Validation
To add new validation rules, edit `ocr_analysis/field_validators.py`:
```python
@staticmethod
def validate_custom_field(value: str) -> Tuple[bool, str]:
    if not value:
        return False, "Field is empty"
    if re.match(r'^YOUR_PATTERN$', value):
        return True, ""
    return False, "Error message"
```

## Testing
Run validation tests:
```powershell
"C:/Users/flori/OneDrive/Desktop/ib_test/Invoicereder/Innovierender-VS-Code/pdf to excel/.venv/Scripts/python.exe" -m pytest ocr_analysis/test_validators.py
```

## Support
- **Pattern mismatch?** Check `9261_1025_JHV_Wesel_Fibuübernahmepaket_DEB.xlsx` for reference format
- **OCR errors?** See `OCR_IMPROVEMENTS.md` for preprocessing tips
- **UI issues?** Check that `customtkinter` is installed: `pip show customtkinter`
