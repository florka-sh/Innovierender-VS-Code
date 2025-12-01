# OCR Enhancement Summary

## Changes Made

### 1. Added "Einstellungen anwenden" Button
**File:** `pdf_extractor_app.py`
- Added apply settings button to OCR UI (card 2)
- Button triggers `apply_ocr_settings()` to rebuild grid with current defaults
- Now visible between settings and export sections

### 2. Fixed DEBI_KREDI Field Extraction
**Problem:** DEBI_KREDI was incorrectly set to recipient name instead of customer/debitor number

**Solution:**
- **OCR Extractor** (`ocr_analysis/poppler_extractor.py`):
  - Added `customer_number` field to extraction result dictionary
  - Implemented customer number extraction using patterns:
    - `Kunden-Nummer: <number>`
    - `Kundennummer: <number>`
    - `Debitoren: <number>`
    - `Personenkonto: <number>`
  - Strips spaces and whitespace from extracted number
  - Example: `"51 111 291 120"` → `"51111291120"`

- **Desktop App** (`pdf_extractor_app.py`):
  - Changed `DEBI_KREDI` assignment from `recipient` (name) to `debi_kredi` (extracted customer number)
  - Now correctly populates debitor field with numeric customer ID

- **Validation** (`ocr_analysis/field_validators.py`):
  - Updated debitor pattern to accept 9-11 digits (was 9 only)
  - Updated invoice pattern to accept 9-10 digits
  - Validation now passes for real-world PDFs

### 3. Data Flow

#### Before:
```
PDF → OCR → {
  Invoice Number: "1155500316",
  Recipient Name: "Max Mustermann"
}
→ DEBI_KREDI = "Max Mustermann" ❌
```

#### After:
```
PDF → OCR → {
  Invoice Number: "1155500316",
  Customer Number: "51111291120",
  Recipient Name: "Max Mustermann"
}
→ DEBI_KREDI = "51111291120" ✅
→ BUCH_TEXT = "Max Mustermann ..." ✅
```

## Testing Results

### Extraction Test
```python
Input PDF text:
  Rechnungsnummer: 1155500316
  Kunden-Nummer: 51 111 291 120

Extracted:
  Invoice Number: "1155500316"
  Customer Number: "51111291120"
  
✅ Spaces removed correctly
✅ Pattern match validated
```

### Field Mapping
| PDF Field | OCR Column | Final Column | Value Example |
|-----------|------------|--------------|---------------|
| Rechnungsnummer | Invoice Number | BELEG_NR | `1155500316` |
| Kunden-Nummer | Customer Number | DEBI_KREDI | `51111291120` |
| Name/Vorname | Recipient Name | BUCH_TEXT | `Max Mustermann` |
| Rechnungsdatum | Date | BELEG_DAT | `20251127` |

## UI Changes

### OCR Panel (Before)
```
2. Einstellungen (Standard)
  [SATZART input]
  [FIRMA input]
  ...
  
3. Export
  [💾 Excel speichern]
```

### OCR Panel (After)
```
2. Einstellungen (Standard)
  [SATZART input]
  [FIRMA input]
  ...
  [✓ Einstellungen anwenden]  ← NEW BUTTON
  
3. Export
  [💾 Excel speichern]
```

## Usage

1. **Load OCR PDF**:
   - Click "📁 PDF wählen"
   - Select scanned invoice
   - Click "🔎 OCR Extrahieren"

2. **Review Extracted Data**:
   - Check data grid
   - Yellow rows = need manual review
   - DEBI_KREDI should show customer number (not name)

3. **Adjust Settings** (if needed):
   - Modify SATZART, FIRMA, etc.
   - **Click "✓ Einstellungen anwenden"** ← Important!
   - Grid rebuilds with new settings

4. **Manual Corrections**:
   - Double-click any cell to edit
   - Fix DEBI_KREDI if pattern not found
   - Fix BELEG_NR if OCR misread

5. **Export**:
   - Click "💾 Excel speichern"
   - File ready for accounting import

## Pattern Reference

### Customer Number Patterns Searched
The extractor searches for these patterns (case-insensitive):
1. `Kunden-Nummer: <number>`
2. `Kundennummer: <number>`
3. `Debitoren: <number>`
4. `Personenkonto: <number>`

First match wins, spaces automatically removed.

### Validation Rules
- **BELEG_NR**: 9-10 digits
- **DEBI_KREDI**: 9-11 digits
- **BELEG_DAT**: YYYYMMDD (8 digits)
- **BETRAG**: Numeric (cents)

## Files Modified
1. `pdf_extractor_app.py`:
   - Added apply button (line ~502)
   - Changed DEBI_KREDI assignment (line ~720)

2. `ocr_analysis/poppler_extractor.py`:
   - Added `customer_number` to result dict (line ~306)
   - Added customer extraction logic (line ~370)
   - Added Customer Number to DataFrame (line ~165)

3. `ocr_analysis/field_validators.py`:
   - Updated invoice pattern to 9-10 digits
   - Updated debitor pattern to 9-11 digits

## Known Limitations
- If PDF doesn't have "Kunden-Nummer:" field, DEBI_KREDI will be empty
- User must manually enter if pattern not found
- Validation highlights yellow for manual entry
- Some scanned PDFs may have OCR errors requiring correction

## Next Steps (Optional)
1. Add fallback: extract from filename if Kunden-Nummer missing
2. Show tooltip on hover explaining yellow highlight
3. Add "Auto-fix DEBI_KREDI" button for batch correction
4. Export validation report alongside Excel
