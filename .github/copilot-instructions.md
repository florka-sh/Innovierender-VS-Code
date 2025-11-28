<!-- Copilot instructions for contributors and AI agents -->
# Repo snapshot — quick summary for an AI coding agent

This project extracts invoice data from PDFs and converts it into accounting-style Excel exports. There are two ways to run it: a desktop GUI and a Flask web app. The main data pipeline modules are `pdf_extractor.py` (parsing PDFs), `excel_generator.py` (creating the 23-column export), and `transform_excel.py` (heuristic Excel → template transformer used by `bereitspf_transformer.py`).

Key entry points and where to look
- `pdf_extractor_app.py` — primary desktop/Tkinter GUI, houses the MultiProjectApp and PDF Reader + Excel Transformer UI.
- `app.py` — Flask web UI for upload → preview → generate workflow.
- `pdf_extractor.py` — heuristics and regex-driven PDF parsing (line items, student info, dates, currency).
- `excel_generator.py` — authoritative logic that formats rows for the 23-column accounting Excel output (amounts → cents, BUCH_TEXT construction).
- `transform_excel.py` — heuristic extractor/transformer for Bereitschaftspflege-style Excel inputs, contains CLI interactive rename logic and column-mapping rules.
- `bereitspf_transformer.py` — adapter that calls `transform_excel.transform()` and returns a list of dicts.
- `ocr_analysis/poppler_extractor.py` — provides OCR extraction using the bundled `InvoiceExtractor/_internal` (poppler + Tesseract). This is now wired into the GUI as "Scanned PDF (OCR)" (Mode 3).
- `column_mapping.json` / `data/mapping_db.xlsx` — optional mapping files used when mapping customer numbers → Kostenträger/Kostenstelle.
- `reference_outputs/` — canonical outputs for comparison; `test_output_structure.py` reads these files to assist manual verification.

Important behaviors & patterns to preserve
- Amounts: code treats money as floats in euros, then most modules multiply by 100 and store `BETRAG` as integer cents (see `pdf_extractor.py` and `excel_generator.py`).
- Dates: input dates are often DD.MM.YYYY; components convert them to YYYYMMDD (e.g. `convert_date_to_yyyymmdd` and Excel post-processing in `transform_excel.py`).
- Mapping DB: mapping overrides default Kostenträger/Kostenträgerbezeichnung if `Kundennummer` matches (look at `pdf_extractor_app.py` and `MultiProjectApp.load_stored_mapping`).
- Column shape: the project expects a fixed export schema (23 columns) used across GUI and generator code. Keep column names consistent (e.g. `SATZART`, `FIRMA`, `BELEG_NR`, `BETRAG`, `BUCH_TEXT`, `KOSTTRAGER`, `KOSTSTELLE`).
- Non-interactive vs interactive: `transform_excel` will prompt for a rename by default; when called programmatically, callers may pass `NO_RENAME` in `defaults` to skip prompts (see CLI vs `bereitspf_transformer.py`).

Developer workflows / repeated commands
- Install deps: `pip install -r requirements.txt` (core: `pdfplumber`, `pandas`, `openpyxl`, `flask`).
- Desktop GUI (local): `python pdf_extractor_app.py` (uses Tkinter). Long-running tasks use threads.
- Web server: `python app.py` → open `http://localhost:5000` (uploads saved to `uploads/`, outputs to `outputs/`).
- Transformer CLI: `python transform_excel.py <template.xlsx> <source.xlsx> --no-rename` for non-interactive runs.
- Quick structural test: `python test_output_structure.py` will produce `working_output_structure.txt` based on `reference_outputs/output_final.xlsx`.

Patterns and conventions useful for codegen
- Small, focused modules: changes that touch the export schema should update `excel_generator.py` and every UI that assembles rows (e.g. `pdf_extractor_app.py`).
- Heuristics-first: the PDF and Excel parsing is heuristic and permissive — prefer small, isolated improvements with test data rather than global rewrites.
- Keep I/O paths explicit: `uploads/`, `outputs/`, `data/` are used by apps; do not rely on implicit temp locations unless adding explicit handling.
- Use the project's conversion functions and constants rather than reinventing conversions (amounts → cents, date format conversion, column names).

When adding features/tests
- Add minimal, targeted tests that use files under `reference_outputs/` or add small synthetic files in `data/` to verify parsing/transforming logic.
- If adding new CLI flags, keep backwards compatibility. `transform_excel` CLI behavior is used by other modules — preserve `NO_RENAME` semantics.

Integration notes & gotchas
- The `bereitspf_transformer` adapter imports `transform_excel` and will fail if `transform_excel` cannot be imported — guard changes accordingly.
- Some third‑party dependencies and runtime helpers live in `InvoiceExtractor/_internal/` (bundled binaries like Tesseract / poppler). Avoid altering that folder unless updating bundled tools.
- Mode 3 (OCR) notes: `ocr_analysis/poppler_extractor.py` uses `pdf2image`, `pytesseract`, and `PyPDF2`. The repo includes a `run_with_poppler.bat` and `_internal` runtimes (Poppler/Tesseract). When running OCR in the GUI the code will set `POPPLER_HOME`/`TESSDATA_PREFIX` and try to use `InvoiceExtractor/_internal` first.
- UI code expects specific column ordering & widths in the Treeview — a column rename or reorder will affect the GUI look and downstream exports.

If anything here is unclear or you want a deeper integration guide (examples of adding a new parser, or where to add unit tests), tell me which area to expand and I'll iterate.
