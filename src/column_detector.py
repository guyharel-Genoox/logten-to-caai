"""
Auto-detection and mapping of flight logbook column headers.

Maps common header names (from ForeFlight, Safelog, manual spreadsheets,
Hebrew logbooks, etc.) to our standardized 48-column layout defined in
column_map.py.

Supports:
- Auto-detection via fuzzy header matching
- Explicit mapping via INI config file
- Validation of required columns for CAAI processing
"""

import configparser
import os
import re

from .column_map import (
    COL_ROW_NUM, COL_DATE, COL_FROM, COL_TO, COL_ROUTE,
    COL_REGISTRATION, COL_AIRCRAFT_TYPE, COL_MAKE, COL_MODEL,
    COL_ENGINE_TYPE, COL_CATEGORY, COL_CLASS,
    COL_TOTAL_TIME, COL_PIC, COL_SIC, COL_NIGHT,
    COL_CROSS_COUNTRY, COL_ACTUAL_INSTRUMENT, COL_SIMULATED_INSTRUMENT,
    COL_DUAL_RECEIVED, COL_DUAL_GIVEN, COL_SOLO, COL_SIMULATOR,
    COL_MULTI_PILOT, COL_DAY_LANDINGS, COL_DAY_TAKEOFFS,
    COL_NIGHT_LANDINGS, COL_NIGHT_TAKEOFFS,
    COL_APPROACH_1, COL_APPROACH_2, COL_APPROACH_3, COL_APPROACH_4,
    COL_HOLDS, COL_IFR, COL_GO_AROUNDS,
    COL_PIC_NAME, COL_INSTRUCTOR, COL_STUDENT, COL_OBSERVER, COL_DPE,
    COL_DISTANCE, COL_REMARKS,
    COL_COMPLEX, COL_HIGH_PERF, COL_EFIS, COL_RETRACTABLE,
    COL_PRESSURIZED, COL_REVIEW_IPC,
)


# ============ Header Alias Database ============
# Maps our column index (from column_map.py) to a list of known aliases.
# Aliases are checked case-insensitively. The FIRST match wins.
# More specific aliases should come before generic ones.

HEADER_ALIASES = {
    COL_DATE: [
        'date', 'flight date', 'flt date', 'flight_date',
        'dep date', 'departure date',
        # Hebrew
        'תאריך',
    ],
    COL_FROM: [
        'from', 'departure', 'dep', 'origin', 'route from',
        'dep airport', 'departure airport', 'depart',
        # Hebrew
        'מ-', 'ממקום',
    ],
    COL_TO: [
        'to', 'arrival', 'arr', 'dest', 'destination', 'route to',
        'arr airport', 'arrival airport',
        # Hebrew
        'ל-', 'למקום',
    ],
    COL_ROUTE: [
        'route', 'via', 'route of flight',
    ],
    COL_REGISTRATION: [
        'registration', 'reg', 'tail', 'tail number', 'tail no',
        'aircraft id', 'ident', 'aircraft ident', 'a/c reg',
        'tail #', 'n-number',
        # Hebrew
        'רישום', 'סימן קריאה',
    ],
    COL_AIRCRAFT_TYPE: [
        'aircraft type', 'type', 'type code', 'a/c type',
        'make/model', 'aircraft', 'ac type', 'airplane type',
        # Hebrew
        'דגם כלי טיס', 'דגם', 'סוג מטוס',
    ],
    COL_MAKE: [
        'make', 'manufacturer', 'aircraft make',
        # Hebrew
        'יצרן',
    ],
    COL_MODEL: [
        'model', 'aircraft model', 'a/c model',
        # Hebrew
        'דגם מלא',
    ],
    COL_ENGINE_TYPE: [
        'engine type', 'engine', 'eng type', 'powerplant',
        # Hebrew
        'סוג מנוע',
    ],
    COL_CATEGORY: [
        'category', 'cat', 'aircraft category',
        # Hebrew
        'קטגוריה',
    ],
    COL_CLASS: [
        'class', 'aircraft class', 'a/c class',
        # Hebrew
        'סיווג',
    ],
    COL_TOTAL_TIME: [
        'total time', 'total', 'total flight time', 'duration',
        'flight time', 'block time', 'total duration', 'ttl time',
        'total hrs', 'flight hours',
        # Hebrew
        'סה"כ זמן', 'זמן טיסה', 'סה"כ',
    ],
    COL_PIC: [
        'pic', 'pilot in command', 'p1', 'pic time',
        'pic hours', 'command',
        # Hebrew
        'טייס אחראי', 'מפקד',
    ],
    COL_SIC: [
        'sic', 'second in command', 'co-pilot', 'copilot',
        'p2', 'sic time', 'sic hours', 'first officer',
        # Hebrew
        'טייס משנה',
    ],
    COL_NIGHT: [
        'night', 'night time', 'night hours', 'nite',
        # Hebrew
        'לילה',
    ],
    COL_CROSS_COUNTRY: [
        'cross country', 'xc', 'x-country', 'cc',
        'cross-country', 'xcountry', 'xc time',
        # Hebrew
        'חוצה ארץ',
    ],
    COL_ACTUAL_INSTRUMENT: [
        'actual instrument', 'actual inst', 'actual ifr',
        'act inst', 'actual imc', 'imc',
        # Hebrew
        'מכשירים בפועל',
    ],
    COL_SIMULATED_INSTRUMENT: [
        'simulated instrument', 'sim inst', 'hood',
        'sim ifr', 'simulated inst', 'sim instrument',
        # Hebrew
        'מכשירים מדומה',
    ],
    COL_DUAL_RECEIVED: [
        'dual received', 'dual recv', 'dual',
        'instruction received', 'dual rcvd', 'training received',
        # Hebrew
        'הדרכה שהתקבלה',
    ],
    COL_DUAL_GIVEN: [
        'dual given', 'instruction given', 'cfi time',
        'instructor time', 'dual gvn', 'training given',
        # Hebrew
        'הדרכה שניתנה',
    ],
    COL_SOLO: [
        'solo', 'solo time', 'solo hours',
        # Hebrew
        'סולו',
    ],
    COL_SIMULATOR: [
        'simulator', 'sim', 'ftd', 'ffs', 'sim time',
        'training device', 'flight sim',
        # Hebrew
        'סימולטור',
    ],
    COL_MULTI_PILOT: [
        'multi-pilot', 'multi pilot', 'multipilot', 'multi crew',
        'multi-crew', 'multicrew', 'mp',
        # Hebrew
        'רב טייס',
    ],
    COL_DAY_LANDINGS: [
        'day landings', 'day ldg', 'ldg day',
        'day land', 'landings day', 'day ldgs',
        # Hebrew
        'נחיתות יום',
    ],
    COL_DAY_TAKEOFFS: [
        'day takeoffs', 'day t/o', 'day to',
        'takeoffs day', 'day tkoffs',
        # Hebrew
        'המראות יום',
    ],
    COL_NIGHT_LANDINGS: [
        'night landings', 'night ldg', 'ldg night',
        'night land', 'landings night', 'night ldgs',
        # Hebrew
        'נחיתות לילה',
    ],
    COL_NIGHT_TAKEOFFS: [
        'night takeoffs', 'night t/o', 'night to',
        'takeoffs night', 'night tkoffs',
        # Hebrew
        'המראות לילה',
    ],
    COL_APPROACH_1: [
        'approach 1', 'approach1', 'apch 1', 'approach',
        'approaches', 'instrument approach',
    ],
    COL_APPROACH_2: [
        'approach 2', 'approach2', 'apch 2',
    ],
    COL_APPROACH_3: [
        'approach 3', 'approach3', 'apch 3',
    ],
    COL_APPROACH_4: [
        'approach 4', 'approach4', 'apch 4',
    ],
    COL_HOLDS: [
        'holds', 'holding', 'hold',
    ],
    COL_IFR: [
        'ifr', 'ifr time', 'instrument flight rules',
    ],
    COL_GO_AROUNDS: [
        'go arounds', 'go-arounds', 'go around',
        'missed approach', 'missed approaches',
    ],
    COL_PIC_NAME: [
        'pic name', 'pilot name', 'commander name',
        'captain name', 'captain',
    ],
    COL_INSTRUCTOR: [
        'instructor', 'cfi name', 'instructor name',
        'flight instructor',
        # Hebrew
        'מדריך',
    ],
    COL_STUDENT: [
        'student', 'student name', 'trainee',
        # Hebrew
        'תלמיד',
    ],
    COL_OBSERVER: [
        'observer', 'safety pilot', 'observer/safety pilot',
        # Hebrew
        'משקיף',
    ],
    COL_DPE: [
        'dpe', 'examiner', 'dpe/examiner', 'check pilot',
        # Hebrew
        'בוחן',
    ],
    COL_DISTANCE: [
        'distance', 'distance (nm)', 'dist', 'nm',
        'distance nm', 'nautical miles',
        # Hebrew
        'מרחק',
    ],
    COL_REMARKS: [
        'remarks', 'comments', 'notes', 'remark',
        # Hebrew
        'הערות',
    ],
    COL_COMPLEX: [
        'complex', 'complex aircraft',
    ],
    COL_HIGH_PERF: [
        'high perf', 'high performance', 'high perf.',
        'hp', 'high power',
    ],
    COL_EFIS: [
        'efis', 'glass cockpit', 'taa',
    ],
    COL_RETRACTABLE: [
        'retractable', 'retract', 'rg',
        'retractable gear',
    ],
    COL_PRESSURIZED: [
        'pressurized', 'press', 'pressurised',
    ],
    COL_REVIEW_IPC: [
        'review/ipc', 'review', 'ipc', 'bfr',
        'flight review', 'biennial flight review',
    ],
}

# Columns required for CAAI processing (from caai_form_filler.read_logbook)
REQUIRED_COLUMNS = {
    COL_DATE, COL_FROM, COL_TO, COL_REGISTRATION, COL_AIRCRAFT_TYPE,
    COL_TOTAL_TIME, COL_PIC, COL_SIC, COL_NIGHT, COL_CROSS_COUNTRY,
    COL_ACTUAL_INSTRUMENT, COL_SIMULATED_INSTRUMENT,
    COL_DUAL_RECEIVED, COL_DUAL_GIVEN, COL_SOLO, COL_MULTI_PILOT,
    COL_DAY_LANDINGS, COL_NIGHT_LANDINGS,
    COL_INSTRUCTOR, COL_REMARKS,
}

# Additional columns that are useful but not strictly required
RECOMMENDED_COLUMNS = {
    COL_ENGINE_TYPE, COL_CLASS, COL_DISTANCE,
}

# Human-readable names for column indices (for reports)
COLUMN_NAMES = {
    COL_ROW_NUM: 'Row Number',
    COL_DATE: 'Date',
    COL_FROM: 'From Airport',
    COL_TO: 'To Airport',
    COL_ROUTE: 'Route',
    COL_REGISTRATION: 'Registration',
    COL_AIRCRAFT_TYPE: 'Aircraft Type',
    COL_MAKE: 'Make',
    COL_MODEL: 'Model',
    COL_ENGINE_TYPE: 'Engine Type',
    COL_CATEGORY: 'Category',
    COL_CLASS: 'Class',
    COL_TOTAL_TIME: 'Total Time',
    COL_PIC: 'PIC',
    COL_SIC: 'SIC',
    COL_NIGHT: 'Night',
    COL_CROSS_COUNTRY: 'Cross Country',
    COL_ACTUAL_INSTRUMENT: 'Actual Instrument',
    COL_SIMULATED_INSTRUMENT: 'Simulated Instrument',
    COL_DUAL_RECEIVED: 'Dual Received',
    COL_DUAL_GIVEN: 'Dual Given',
    COL_SOLO: 'Solo',
    COL_SIMULATOR: 'Simulator',
    COL_MULTI_PILOT: 'Multi-Pilot',
    COL_DAY_LANDINGS: 'Day Landings',
    COL_DAY_TAKEOFFS: 'Day Takeoffs',
    COL_NIGHT_LANDINGS: 'Night Landings',
    COL_NIGHT_TAKEOFFS: 'Night Takeoffs',
    COL_APPROACH_1: 'Approach 1',
    COL_APPROACH_2: 'Approach 2',
    COL_APPROACH_3: 'Approach 3',
    COL_APPROACH_4: 'Approach 4',
    COL_HOLDS: 'Holds',
    COL_IFR: 'IFR',
    COL_GO_AROUNDS: 'Go Arounds',
    COL_PIC_NAME: 'PIC Name',
    COL_INSTRUCTOR: 'Instructor',
    COL_STUDENT: 'Student',
    COL_OBSERVER: 'Observer/Safety Pilot',
    COL_DPE: 'DPE/Examiner',
    COL_DISTANCE: 'Distance (NM)',
    COL_REMARKS: 'Remarks',
    COL_COMPLEX: 'Complex',
    COL_HIGH_PERF: 'High Performance',
    COL_EFIS: 'EFIS',
    COL_RETRACTABLE: 'Retractable',
    COL_PRESSURIZED: 'Pressurized',
    COL_REVIEW_IPC: 'Review/IPC',
}


def _normalize_header(header):
    """Normalize a header string for matching.

    Strips whitespace, lowercases, removes special chars except letters/numbers/spaces.
    """
    if not header:
        return ''
    h = str(header).strip().lower()
    # Keep Hebrew chars, Latin chars, digits, spaces
    h = re.sub(r'[^\w\s\u0590-\u05FF]', ' ', h)
    h = re.sub(r'\s+', ' ', h).strip()
    return h


def detect_columns(headers):
    """Auto-detect column mapping from header names.

    Uses the HEADER_ALIASES database to match user's column headers
    to our standardized column indices.

    Args:
        headers: List of raw header strings from the source file.

    Returns:
        Dict mapping our column index → source column index (0-based).
        Only includes detected columns.
    """
    mapping = {}
    used_source_cols = set()
    normalized_headers = [_normalize_header(h) for h in headers]

    # First pass: exact matches (highest confidence)
    for our_col, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            norm_alias = _normalize_header(alias)
            for src_idx, norm_header in enumerate(normalized_headers):
                if src_idx in used_source_cols:
                    continue
                if norm_header == norm_alias:
                    mapping[our_col] = src_idx
                    used_source_cols.add(src_idx)
                    break
            if our_col in mapping:
                break

    # Second pass: substring matches for unmapped columns
    for our_col, aliases in HEADER_ALIASES.items():
        if our_col in mapping:
            continue
        for alias in aliases:
            norm_alias = _normalize_header(alias)
            if len(norm_alias) < 3:
                continue  # Skip very short aliases for substring matching
            for src_idx, norm_header in enumerate(normalized_headers):
                if src_idx in used_source_cols:
                    continue
                if norm_alias in norm_header or norm_header in norm_alias:
                    mapping[our_col] = src_idx
                    used_source_cols.add(src_idx)
                    break
            if our_col in mapping:
                break

    return mapping


def load_column_mapping(mapping_file):
    """Load explicit column mapping from an INI file.

    The mapping file maps our standard column names to source column
    names or 0-based indices.

    Format:
        [columns]
        Date = Flight Date
        Aircraft Type = A/C Type
        Total Time = Block Hours
        PIC = P1 Time
        # etc.

    Or by index:
        [columns]
        Date = 0
        Aircraft Type = 3
        Total Time = 5

    Args:
        mapping_file: Path to the mapping INI file.

    Returns:
        Dict mapping our column index → source column name or index.
        The caller must resolve names against actual headers.
    """
    parser = configparser.ConfigParser()
    parser.read(mapping_file, encoding='utf-8')

    if not parser.has_section('columns'):
        raise ValueError(f"Mapping file {mapping_file} must have a [columns] section")

    # Build reverse lookup: our column name → our column index
    name_to_col = {}
    for col_idx, name in COLUMN_NAMES.items():
        name_to_col[_normalize_header(name)] = col_idx

    mapping = {}
    for our_name, source_val in parser.items('columns'):
        our_norm = _normalize_header(our_name)
        our_col = name_to_col.get(our_norm)
        if our_col is None:
            # Try matching against aliases too
            for col_idx, aliases in HEADER_ALIASES.items():
                if our_norm in [_normalize_header(a) for a in aliases]:
                    our_col = col_idx
                    break

        if our_col is None:
            print(f"  WARNING: Unknown column name in mapping: '{our_name}'")
            continue

        # Source value can be a column name (string) or index (int)
        source_val = source_val.strip()
        if source_val.isdigit():
            mapping[our_col] = int(source_val)
        else:
            mapping[our_col] = source_val

    return mapping


def resolve_mapping_names(mapping, headers):
    """Resolve string source column names to indices.

    Args:
        mapping: Dict from load_column_mapping (may contain string names).
        headers: List of actual header strings from the source.

    Returns:
        Dict mapping our column index → source column index (int).
    """
    normalized_headers = [_normalize_header(h) for h in headers]
    resolved = {}

    for our_col, source_val in mapping.items():
        if isinstance(source_val, int):
            resolved[our_col] = source_val
        else:
            norm_source = _normalize_header(source_val)
            found = False
            for idx, norm_h in enumerate(normalized_headers):
                if norm_h == norm_source or norm_source in norm_h:
                    resolved[our_col] = idx
                    found = True
                    break
            if not found:
                print(f"  WARNING: Source column '{source_val}' not found in headers")

    return resolved


def get_required_columns():
    """Return the set of column indices required for CAAI processing."""
    return REQUIRED_COLUMNS.copy()


def validate_mapping(mapping, headers=None):
    """Validate a column mapping and return warnings.

    Args:
        mapping: Dict mapping our column index → source column index.
        headers: Optional list of source headers (for display).

    Returns:
        Tuple of (errors: list[str], warnings: list[str]).
        Errors = missing required columns.
        Warnings = missing recommended columns.
    """
    errors = []
    warnings = []

    for col in REQUIRED_COLUMNS:
        if col not in mapping:
            name = COLUMN_NAMES.get(col, f'Column {col}')
            errors.append(f"Required column missing: {name}")

    for col in RECOMMENDED_COLUMNS:
        if col not in mapping:
            name = COLUMN_NAMES.get(col, f'Column {col}')
            warnings.append(f"Recommended column missing: {name} (will use defaults)")

    return errors, warnings


def print_mapping_report(mapping, headers):
    """Print a human-readable report of the column mapping.

    Args:
        mapping: Dict mapping our column index → source column index.
        headers: List of source header strings.
    """
    print("\n" + "=" * 70)
    print("COLUMN MAPPING REPORT")
    print("=" * 70)

    # Detected columns
    print(f"\n  Detected {len(mapping)} columns:")
    for our_col in sorted(mapping.keys()):
        our_name = COLUMN_NAMES.get(our_col, f'Column {our_col}')
        src_idx = mapping[our_col]
        src_name = headers[src_idx] if src_idx < len(headers) else f'Index {src_idx}'
        required = '*' if our_col in REQUIRED_COLUMNS else ' '
        print(f"  {required} {our_name:<25} <- '{src_name}' (col {src_idx})")

    # Missing required
    errors, warnings = validate_mapping(mapping)
    if errors:
        print(f"\n  MISSING REQUIRED COLUMNS ({len(errors)}):")
        for err in errors:
            print(f"    ! {err}")

    if warnings:
        print(f"\n  Missing optional columns ({len(warnings)}):")
        for warn in warnings:
            print(f"    - {warn}")

    # Unmapped source columns
    mapped_src = set(mapping.values())
    unmapped = [i for i in range(len(headers)) if i not in mapped_src and headers[i].strip()]
    if unmapped:
        print(f"\n  Unmapped source columns ({len(unmapped)}):")
        for idx in unmapped[:20]:
            print(f"    ? '{headers[idx]}' (col {idx})")
        if len(unmapped) > 20:
            print(f"    ... and {len(unmapped) - 20} more")

    print()
