# PDF Invoice Extractor - Lernförderung

Eine Desktop-Anwendung zum Extrahieren von Rechnungsdaten aus PDF-Dateien und Generieren von Excel-Dateien im Buchhaltungsformat.

## 📋 Funktionen

- **PDF-Extraktion**: Automatisches Extrahieren von Rechnungsdaten aus mehrseitigen PDFs
- **OCR für gescannte PDFs**: Mode 3 unterstützt OCR-Extraktion mit Tesseract/Poppler (bundled in `InvoiceExtractor/_internal`)
- **Excel-Transformer**: Transformiert vorhandene Excel-Dateien nach Template-Vorgabe (Bereitschaftspflege-Mode)
- **Datenvorschau**: Tabellarische Anzeige aller extrahierten Einträge mit Bearbeitungsfunktion
- **Konfigurierbar**: Eingabefelder für alle Buchhaltungsparameter
- **Excel-Export**: Generiert Excel-Dateien mit 23 Spalten im korrekten Format
- **Moderne Oberfläche**: Dunkles Theme mit benutzerfreundlichem Design

## 🚀 Installation

### Voraussetzungen
- Python 3.8 oder höher
- Die erforderlichen Bibliotheken sind bereits installiert:
  - pdfplumber
  - pandas
  - openpyxl
  - tkinter (im Lieferumfang von Python enthalten)

### Optional: Abhängigkeiten installieren
Falls Bibliotheken fehlen:
```bash
pip install pdfplumber pandas openpyxl
```

## 💻 Verwendung

### Desktop-Anwendung starten

```bash
python pdf_extractor_app.py
```

### Schritt-für-Schritt-Anleitung

**Mode 1: PDF Reader (Standard PDFs)**

1. **PDF auswählen**
   - Klicken Sie auf "Durchsuchen" und wählen Sie Ihre PDF-Datei aus
   - Klicken Sie auf "Daten extrahieren"

2. **Vorschau anzeigen**
   - Ein Fenster zeigt alle extrahierten Einträge an
   - Überprüfen Sie Rechnungsnummern, Schüler, Fächer und Beträge

3. **Parameter konfigurieren**
   - Passen Sie die Buchhaltungsparameter nach Bedarf an
   - Voreingestellte Werte basieren auf Ihrer Vorlage

4. **Excel generieren**
   - Klicken Sie auf "Excel generieren"
   - Wählen Sie den Speicherort für die Excel-Datei
   - Fertig!

**Mode 2: Excel Transformer (Bereitschaftspflege)**

1. Wählen Sie Template Excel und Quelldatei Excel
2. Konfigurieren Sie Standardwerte (SATZART, FIRMA, etc.)
3. Klicken Sie auf "Transformieren"
4. Exportieren Sie das Ergebnis

**Mode 3: Scanned PDF (OCR)**

1. Wählen Sie ein gescanntes PDF (benötigt Tesseract/Poppler unter `InvoiceExtractor/_internal`)
2. Klicken Sie auf "OCR Extrahieren" (kann einige Minuten dauern)
3. Vorschau prüfen und exportieren

## 📊 Extrahierte Daten

Die Anwendung extrahiert folgende Informationen aus PDFs:

- **Rechnungsmetadaten**: Rechnungsnummer, Datum, Kundennummer
- **Schülerinformationen**: Name, Schule
- **Kursinformationen**: Monat/Jahr, Fach, Stunden, Tarif, Betrag
- **Kontocodes**: Automatisch extrahierte Buchhaltungscodes

## 📁 Excel-Format

Die generierte Excel-Datei enthält 23 Spalten:

| Spalte | Quelle | Beispiel |
|--------|--------|----------|
| SATZART | Konfiguriert | D |
| FIRMA | Konfiguriert | 9251 |
| BELEG_NR | Aus PDF | 1155500316 |
| BELEG_DAT | Aus PDF (konvertiert) | 20251127 |
| SOLL_HABEN | Konfiguriert | H |
| BUCH_KREIS | Konfiguriert | RA |
| BUCH_JAHR | Aus Datum | 2025 |
| BUCH_MONAT | Aus Datum | 11 |
| DEBI_KREDI | Aus PDF | 51111291120 |
| BETRAG | Aus PDF (in Cent) | 2500 |
| RECHNUNG | Aus PDF | 1155500316 |
| BUCH_TEXT | Generiert | 1025 Cana Khudidah Deutsch |
| HABENKONTO | Konfiguriert | 42200 |
| KOSTSTELLE | Konfiguriert | 190 |
| KOSTTRAGER | Konfiguriert | 190111512110 |

## 🛠️ Projektstruktur

```
Lernförderung Solingen/
├── pdf_extractor_app.py      # Desktop-Anwendung (Hauptdatei)
├── pdf_extractor.py           # PDF-Extraktionsmodul
├── excel_generator.py         # Excel-Generierungsmodul
├── requirements.txt           # Python-Abhängigkeiten
├── uploads/                   # Temporäre Upload-Ordner
└── outputs/                   # Generierte Excel-Dateien
```

## ⚙️ Konfigurierbare Parameter

- **FIRMA**: Firmennummer (Standard: 9251)
- **SATZART**: Satzart (Standard: D)
- **SOLL_HABEN**: Soll/Haben-Kennzeichen (Standard: H)
- **BUCH_KREIS**: Buchungskreis (Standard: RA)
- **HABENKONTO**: Habenkonto (Standard: 42200)
- **KOSTSTELLE**: Kostenstelle (Standard: 190)
- **KOSTTRAGER**: Kostenträger (Standard: 190111512110)
- **Kostenträgerbezeichnung**: Beschreibung (Standard: SPFH/HzE Siegen)
- **Bebuchbar**: Bebuchbar-Status (Standard: Ja)
- **BUCH_TEXT_PREFIX**: Präfix für Buchungstext (Standard: 1025)

## 🔍 Fehlerbehebung

**Problem**: "Modul nicht gefunden"
- Lösung: Installieren Sie fehlende Module mit `pip install pdfplumber pandas openpyxl`

**Problem**: "Keine Daten extrahiert"
- Lösung: Stellen Sie sicher, dass die PDF-Datei Lernförderungsrechnungen im erwarteten Format enthält

**Problem**: "Excel-Datei kann nicht gespeichert werden"
- Lösung: Überprüfen Sie, ob Sie Schreibrechte für den Zielordner haben

### Scanned PDF (OCR) Mode - Fehlerbehebung

**Problem**: "Tesseract NOT found" oder "Unable to get page count"
- **Ursache**: OCR-Bibliotheken (Tesseract, Poppler) fehlen oder sind nicht konfiguriert
- **Lösung**: 
  1. Stellen Sie sicher, dass `InvoiceExtractor/_internal/` die benötigten Runtimes enthält (Poppler + Tesseract-OCR)
  2. Diese werden automatisch vom Extractor gefunden, wenn sie im Projektverzeichnis unter `InvoiceExtractor/_internal/` liegen
  3. Alternativ: Installieren Sie Tesseract global (`choco install tesseract` oder von https://github.com/UB-Mannheim/tesseract/wiki)

**Problem**: OCR ist langsam oder stürzt ab
- **Tipp**: OCR verarbeitet jede Seite einzeln mit DPI=200. Große PDFs (>50 Seiten) können mehrere Minuten dauern
- **Lösung**: Verkleinern Sie das PDF oder teilen Sie es in kleinere Dateien auf

**Hinweis**: OCR-Mode (Mode 3) nutzt die gebündelten Binaries unter `InvoiceExtractor/_internal/poppler/` und `InvoiceExtractor/_internal/Tesseract-OCR/`. Bei Problemen prüfen Sie, ob diese Ordner existieren und ausführbare Dateien enthalten.

## 📝 Beispiel

Getestet mit `RE_1155500316-325.pdf`:
- **10 Seiten** verarbeitet
- **16 Einträge** extrahiert (10 Schüler, verschiedene Fächer)
- **23 Spalten** im Excel-Export
- Identisches Format wie `9251_1025_Lernforderung Solingen Fibuübernahmepaket.xlsx`

## 🎨 Features

- ✅ Moderne dunkle Benutzeroberfläche
- ✅ Multi-Threading für reaktionsschnelle UI
- ✅ Datenvorschau vor dem Export
- ✅ Anpassbare Buchhaltungsparameter
- ✅ Fehlerbehandlung und Benutzer-Feedback
- ✅ Unterstützung für mehrseitige PDFs

## 📞 Hinweise

- Beträge werden automatisch von Euro in Cent umgerechnet (× 100)
- Datumsformat wird von DD.MM.YYYY in YYYYMMDD konvertiert
- Alle Spalten entsprechen der Vorlage
