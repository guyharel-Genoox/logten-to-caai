# Flight Logbook to CAAI

Convert **any** flight logbook into the Israeli CAAI **tofes-shaot** (flight hours summary form).

Supports Excel, CSV, PDF, and [Coradine LogTen Pro](https://coradine.com) exports. Auto-detects your column layout or use a simple mapping file.

## What It Does

1. **Imports** your flight logbook from any format (Excel, CSV, PDF, LogTen Pro) into a standardized 48-column Excel workbook
2. **Calculates** great-circle distances (NM) for all flights using a 260+ airport coordinate database
3. **Classifies** each flight per CAAI rules (role, aircraft group, day/night, XC, instrument, complex)
4. **Fills** the official CAAI tofes-shaot form (Summary, CPL, and ATPL sheets)

All 8 CAAI regulatory rules are enforced, including SIC half-credit (regulation 42(b)), safety pilot exclusions, and proper student/PIC/SIC classification.

## Prerequisites

- Python 3.8+
- A flight logbook file (Excel, CSV, PDF, or LogTen Pro export)
- A blank CAAI tofes-shaot Excel form (included in `templates/`)

## Installation

```bash
git clone https://github.com/guyharel-Genoox/logten-to-caai.git
cd logten-to-caai
pip install -r requirements.txt
cp config.example.ini config.ini
# Edit config.ini with your name and file paths
```

## Quick Start

### Option A: Any Logbook (Excel, CSV, PDF)

```bash
# Just point to your file — columns are auto-detected
python run.py --input my_logbook.xlsx --pilot "Your Name"

# CSV works too
python run.py --input flights.csv

# PDF (requires tabular format)
python run.py --input logbook.pdf
```

### Option B: LogTen Pro Export

```bash
# LogTen Pro tab-delimited export (auto-detected)
python run.py --input "Export Flights (Tab).txt" --format logten
```

### Option C: Using config.ini

Edit `config.ini`:
```ini
[pilot]
name = Your Name

[import]
input_file = ./my_logbook.xlsx
format = auto

[files]
logbook_output = ./Flight_Logbook.xlsx
caai_template = ./templates/tofes-shaot-blank.xlsx
caai_output = ./CAAI_Tofes_Shaot_Filled.xlsx
```

Then run:
```bash
python run.py
```

### Running Individual Steps

```bash
python run.py --step import          # Import and standardize logbook
python run.py --step distances       # Add distance calculations
python run.py --step caai-columns    # Add CAAI classification columns
python run.py --step fill-form       # Fill the tofes-shaot
python run.py --step analyze         # Verify CAAI compliance
```

### Review

Open your filled tofes-shaot and verify the numbers before submitting to CAAI.

## Supported Input Formats

| Format | Extensions | Notes |
|--------|-----------|-------|
| Excel | `.xlsx`, `.xls` | Most logbook apps can export to Excel |
| CSV | `.csv` | Comma-separated values |
| TSV | `.tsv`, `.txt` | Tab-separated values |
| PDF | `.pdf` | Must contain structured tables |
| LogTen Pro | `.txt` | Coradine LogTen Pro tab-delimited export |

### Column Auto-Detection

The tool recognizes common column header names from popular logbook apps (ForeFlight, Safelog, etc.) and manual spreadsheets. It also supports Hebrew headers.

If auto-detection doesn't work for your logbook, you can provide a column mapping file:

```bash
python run.py --input flights.csv --mapping my_columns.ini
```

See [docs/import-guide.md](docs/import-guide.md) for mapping file format and full column reference.

## Output

### Flight Logbook (standardized Excel)
- **Flight Log**: All flights with 48+ columns in standardized format
- **Summary & Totals**: Import summary and statistics

### CAAI Form (3 sheets filled)
- **Summary**: Table 1 (aircraft hours by type/role) + Table 2 (instrument time)
- **CPL**: PIC XC, dual, night, instrument, solo XC, complex
- **ATPL**: XC total, night PIC XC, instrument total

## CAAI Rules Implemented

| Rule | Description |
|------|-------------|
| 1 | Table 1 contains aircraft hours only (NO simulators) |
| 2 | Student = dual instruction only |
| 3 | PIC excludes safety pilot time and flights with instructor |
| 4 | PIC XC excludes safety pilot and instructor flights |
| 5 | Actual instrument on single-pilot aircraft (not dual) = PIC |
| 6 | Simulator time NOT in total flight hours |
| 7 | Simulated instrument in air during instruction = student time |
| 8 | SIC half-credit per regulation 42(b): total = PIC + SIC/2 + Student |

### Aircraft Groups
- **Group A**: Single-engine piston (C172, C150, PA28, SR22)
- **Group B**: Multi-engine piston (PA44, BE76)
- **Group C**: Multi-engine jet/turboprop (A319, A320, H25B)
- **Group D**: Single-engine turboprop

## Adding Airports

The built-in database covers ~260 airports (US, Europe, Middle East). To add your own:

1. Create a JSON file:
```json
{
    "LLSD": [32.1145, 34.7822],
    "ICAO": [latitude, longitude]
}
```

2. Set in config.ini:
```ini
custom_airports = ./my_airports.json
```

## Project Structure

```
logten-to-caai/
├── run.py                       # Pipeline runner (entry point)
├── config.example.ini           # Sample configuration
├── requirements.txt             # Dependencies
├── src/
│   ├── config.py                # Configuration loading
│   ├── column_map.py            # Excel column constants (48 base + 8 CAAI)
│   ├── column_detector.py       # Column auto-detection + alias database
│   ├── universal_importer.py    # Universal import (Excel/CSV/PDF)
│   ├── pdf_reader.py            # PDF table extraction
│   ├── logbook_creator.py       # LogTen Pro specific importer
│   ├── caai_rules.py            # CAAI classification functions
│   ├── airports.py              # Airport DB + haversine distance
│   ├── distance_calculator.py   # Add distances to logbook
│   ├── caai_columns.py          # Add CAAI classification columns
│   ├── caai_form_filler.py      # Fill tofes-shaot form
│   └── caai_analyzer.py         # Verification tool
├── templates/
│   └── tofes-shaot-blank.xlsx   # Blank CAAI form template
└── docs/
    ├── caai-rules.md            # CAAI rules documentation
    └── import-guide.md          # Import format guide + column mapping
```

## Limitations

- Airport database may not have all airports — add your own via JSON
- CAAI form template layout is hardcoded — if CAAI changes the form, the script may need updating
- Safety pilot detection relies on "safety pilot" appearing in the remarks field
- PDF import works best with well-structured tabular PDFs; scanned/image PDFs are not supported

---

## בעברית

כלי להמרת כל יומן טיסות לטופס שעות של רת"א (CAAI).

הכלי קורא יומני טיסות בכל פורמט (אקסל, CSV, PDF, או LogTen Pro), יוצר חוברת אקסל מסודרת, וממלא אוטומטית את טופס השעות (תופס-שעות) כולל:
- טבלה 1: שעות טיסה לפי סוג מטוס ותפקיד
- טבלה 2: שעות מכשירים
- דף CPL: דרישות רישיון טיס מסחרי
- דף ATPL: דרישות רישיון טיס תובלה

כל 8 כללי רת"א מיושמים, כולל חצי זיכוי SIC לפי תקנה 42(ב).

זיהוי אוטומטי של עמודות — תומך בכותרות בעברית ובאנגלית.
