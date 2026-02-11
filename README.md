# LogTen to CAAI

Convert [Coradine LogTen Pro](https://coradine.com) flight log exports into a comprehensive Excel logbook and automatically fill the Israeli CAAI **tofes-shaot** (טופס שעות - flight hours summary form).

## What It Does

1. **Parses** your LogTen Pro tab-delimited export into a 4-sheet Excel workbook (Flight Log, Summary, Aircraft, Milestones)
2. **Calculates** great-circle distances (NM) for all flights using a 260+ airport coordinate database
3. **Classifies** each flight per CAAI rules (role, aircraft group, day/night, XC, instrument, complex)
4. **Fills** the official CAAI tofes-shaot form (Summary, CPL, and ATPL sheets)

All 8 CAAI regulatory rules are enforced, including SIC half-credit (regulation 42(b)), safety pilot exclusions, and proper student/PIC/SIC classification.

## Prerequisites

- Python 3.8+
- [LogTen Pro](https://coradine.com) flight logbook (Mac/iOS)
- A blank CAAI tofes-shaot Excel form (place in `templates/`)

## Installation

```bash
git clone https://github.com/your-username/logten-to-caai.git
cd logten-to-caai
pip install -r requirements.txt
cp config.example.ini config.ini
# Edit config.ini with your name and file paths
```

## Quick Start

### 1. Export from LogTen Pro

In LogTen Pro, go to **File > Export > Tab-delimited** and save the file.

### 2. Configure

Edit `config.ini`:
```ini
[pilot]
name = Your Name

[files]
logten_export = ./Export Flights (Tab).txt
logbook_output = ./Flight_Logbook.xlsx
caai_template = ./templates/tofes-shaot-blank.xlsx
caai_output = ./CAAI_Tofes_Shaot_Filled.xlsx
```

### 3. Run

```bash
# Full pipeline
python run.py

# Or run individual steps
python run.py --step logbook         # Create Excel logbook
python run.py --step distances       # Add distance calculations
python run.py --step caai-columns    # Add CAAI columns
python run.py --step fill-form       # Fill the tofes-shaot
python run.py --step analyze         # Verify CAAI compliance
```

### 4. Review

Open your filled tofes-shaot and verify the numbers before submitting to CAAI.

## Output

### Flight Logbook (4 sheets)
- **Flight Log**: All flights with 56 columns (48 base + 8 CAAI)
- **Summary & Totals**: Career totals, yearly breakdown, aircraft types, airports
- **Aircraft**: All unique aircraft flown
- **Milestones**: Checkrides, first solo, type ratings

### CAAI Form (3 sheets filled)
- **Summary (סיכום ניסיון תעופתי)**: Table 1 (aircraft hours by type/role) + Table 2 (instrument time)
- **CPL (רישיון טיס מסחרי)**: PIC XC, dual, night, instrument, solo XC, complex
- **ATPL (רישיון טיס תובלה)**: XC total, night PIC XC, instrument total

## CAAI Rules Implemented

| Rule | Description |
|------|-------------|
| 1 | Table 1 contains aircraft hours only (NO simulators) |
| 2 | Student (מתלמד) = dual instruction only |
| 3 | PIC excludes safety pilot time and flights with instructor |
| 4 | PIC XC excludes safety pilot and instructor flights |
| 5 | Actual instrument on single-pilot aircraft (not dual) = PIC |
| 6 | Simulator time NOT in total flight hours |
| 7 | Simulated instrument in air during instruction = student time |
| 8 | SIC half-credit per regulation 42(b): total = PIC + SIC/2 + Student |

### Aircraft Groups
- **Group A (א)**: Single-engine piston (C172, C150, PA28, SR22)
- **Group B (ב)**: Multi-engine piston (PA44, BE76)
- **Group C (ג)**: Multi-engine jet/turboprop (A319, A320, H25B)
- **Group D (ד)**: Single-engine turboprop

## Adding Airports

The built-in database covers ~260 airports (US, Europe, Middle East). To add your own:

1. Create a JSON file:
```json
{
    "ICAO": [latitude, longitude],
    "LLSD": [32.1145, 34.7822]
}
```

2. Set in config.ini:
```ini
custom_airports = ./my_airports.json
```

## Project Structure

```
logten-to-caai/
├── run.py                    # Pipeline runner
├── config.example.ini        # Sample configuration
├── requirements.txt          # openpyxl
├── src/
│   ├── config.py             # Configuration loading
│   ├── column_map.py         # Excel column constants
│   ├── caai_rules.py         # CAAI classification functions
│   ├── airports.py           # Airport DB + haversine
│   ├── logbook_creator.py    # LogTen export -> Excel
│   ├── distance_calculator.py # Add distances
│   ├── caai_columns.py       # Add CAAI columns
│   ├── caai_form_filler.py   # Fill tofes-shaot
│   └── caai_analyzer.py      # Verification tool
├── templates/
│   └── tofes-shaot-blank.xlsx
└── docs/
    ├── caai-rules.md
    └── logten-export-guide.md
```

## Limitations

- Designed for LogTen Pro tab-delimited exports (specific field names required)
- Airport database may not have all airports - add your own via JSON
- CAAI form template layout is hardcoded - if CAAI changes the form, the script may need updating
- Safety pilot detection relies on "safety pilot" appearing in the remarks field

---

## בעברית

כלי להמרת יומן טיסות מ-LogTen Pro לטופס שעות של רת"א (CAAI).

הכלי קורא את קובץ הייצוא מ-LogTen Pro, יוצר חוברת אקסל מסודרת, ומלא אוטומטית את טופס השעות (תופס-שעות) כולל:
- טבלה 1: שעות טיסה לפי סוג מטוס ותפקיד
- טבלה 2: שעות מכשירים
- דף CPL: דרישות רישיון טיס מסחרי
- דף ATPL: דרישות רישיון טיס תובלה

כל 8 כללי רת"א מיושמים, כולל חצי זיכוי SIC לפי תקנה 42(ב).
