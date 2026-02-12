# Import Guide

This guide explains how to import your flight logbook from any format.

## Supported Formats

### Excel (.xlsx, .xls)

Most logbook apps (ForeFlight, Safelog, etc.) can export to Excel. Just point the tool to your file:

```bash
python run.py --input my_logbook.xlsx
```

The tool reads the first sheet (or a sheet named "Flight Log" if it exists). The first row must contain column headers.

### CSV (.csv)

```bash
python run.py --input flights.csv
```

Standard comma-separated format. UTF-8 encoding recommended.

### TSV (.tsv, .txt)

```bash
python run.py --input flights.tsv
```

Tab-separated format. This is what many apps export by default.

### PDF (.pdf)

```bash
python run.py --input logbook.pdf
```

The PDF must contain structured tables (not scanned images). The tool uses pdfplumber to extract tabular data.

**Tips for PDF import:**
- Well-formatted tables with clear borders work best
- Multi-page tables are handled automatically (repeated headers are merged)
- Summary/total rows are automatically skipped
- If PDF extraction fails, try exporting from your app as Excel or CSV instead

### LogTen Pro

For Coradine LogTen Pro tab-delimited exports:

```bash
python run.py --input "Export Flights (Tab).txt" --format logten
```

The LogTen format is auto-detected if the file contains LogTen field names. See the LogTen-specific export instructions below.

## Column Auto-Detection

The tool automatically matches your column headers to the standardized format. It recognizes common names from popular logbook apps:

| Our Column | Recognized Headers |
|-----------|-------------------|
| Date | date, flight date, flt date |
| From | from, departure, dep, origin |
| To | to, arrival, arr, destination |
| Registration | registration, reg, tail, tail number, ident |
| Aircraft Type | aircraft type, type, type code, a/c type |
| Total Time | total time, total, duration, flight time, block time |
| PIC | pic, pilot in command, p1 |
| SIC | sic, second in command, co-pilot, p2 |
| Night | night, night time |
| Cross Country | cross country, xc, x-country |
| Actual Instrument | actual instrument, actual inst, actual ifr, imc |
| Simulated Instrument | simulated instrument, sim inst, hood |
| Dual Received | dual received, dual recv, dual |
| Dual Given | dual given, instruction given, cfi time |
| Solo | solo, solo time |
| Day Landings | day landings, day ldg |
| Night Landings | night landings, night ldg |
| Instructor | instructor, cfi name |
| Remarks | remarks, comments, notes |

Hebrew headers are also supported (e.g., "תאריך", "דגם כלי טיס", "הערות").

## Column Mapping File

If auto-detection doesn't match your columns correctly, create a mapping file:

### Format

```ini
[columns]
# Map: Our Column Name = Your Column Header
Date = Flight Date
Aircraft Type = A/C Type
Total Time = Block Hours
PIC = P1 Time
SIC = P2 Time
Night = Night Hours
Cross Country = XC Time
Dual Received = Training Received
Instructor = CFI Name
Remarks = Notes
```

### Using a mapping file

```bash
python run.py --input flights.csv --mapping my_columns.ini
```

Or in config.ini:
```ini
[import]
input_file = ./flights.csv
column_mapping = ./my_columns.ini
```

### Column index mapping

You can also map by column index (0-based):

```ini
[columns]
Date = 0
Aircraft Type = 3
Total Time = 5
PIC = 6
```

## Required Columns

For CAAI form filling, these columns are **required** (marked with *):

| Column | Required | Notes |
|--------|----------|-------|
| Date * | Yes | Flight date |
| From * | Yes | Departure airport |
| To * | Yes | Arrival airport |
| Registration * | Yes | Aircraft registration |
| Aircraft Type * | Yes | Type code (C172, A319, etc.) |
| Total Time * | Yes | Total flight time (decimal hours) |
| PIC * | Yes | PIC time |
| SIC * | Yes | SIC time |
| Night * | Yes | Night time |
| Cross Country * | Yes | XC time |
| Actual Instrument * | Yes | Actual instrument time |
| Simulated Instrument * | Yes | Simulated instrument time |
| Dual Received * | Yes | Dual received time |
| Dual Given * | Yes | Dual given (instructor) time |
| Solo * | Yes | Solo time |
| Multi-Pilot * | Yes | Multi-pilot time |
| Day Landings * | Yes | Day landing count |
| Night Landings * | Yes | Night landing count |
| Instructor * | Yes | Instructor name (for role detection) |
| Remarks * | Yes | Remarks (for safety pilot detection) |
| Engine Type | Recommended | Helps with aircraft classification |
| Class | Recommended | Aircraft class |
| Distance | Recommended | Auto-calculated if airports are known |

## Time Format

Time values can be in any of these formats:
- Decimal hours: `1.5`, `2.33`
- H:MM: `1:30`, `2:20`
- European decimal (comma): `1,5`

All are automatically normalized to decimal hours.

## LogTen Pro Export Instructions

### LogTen Pro for Mac

1. Open LogTen Pro
2. Go to **File > Export**
3. Select **Tab-delimited** as the format
4. Select **All Flights** (or the date range you need)
5. Save the file (default name: `Export Flights (Tab) - YYYY-MM-DD HH-MM-SS.txt`)

### LogTen Pro for iOS/iPad

1. Open LogTen Pro
2. Tap the **More** menu
3. Tap **Export**
4. Select **Tab-delimited**
5. Share/save the file to your computer

**Tips:**
- Make sure **all flights** are exported, including simulator sessions
- Airport codes should be in ICAO format (4-letter codes)
- The tool handles multiline remarks automatically

## Troubleshooting

### "Required column missing" warnings
Your logbook doesn't have all the columns needed for CAAI processing. Options:
1. Add the missing columns to your source file
2. Create a column mapping file if the columns exist under different names
3. The tool will still work but the CAAI form may be incomplete

### PDF extraction issues
If the PDF produces garbled data:
- Try exporting from your logbook app as Excel or CSV instead
- Ensure the PDF contains actual text tables (not scanned images)
- Check that table borders are clear and consistent

### Column detection mismatches
Run with just the import step to see the mapping report:
```bash
python run.py --step import --input my_logbook.xlsx
```
Review the "COLUMN MAPPING REPORT" output, then create a mapping file for any incorrect matches.
