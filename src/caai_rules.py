"""
Shared CAAI (Israeli Civil Aviation Authority) classification functions.

These functions implement the aircraft categorization and role assignment
rules required by the CAAI tofes-shaot (flight hours summary form).

CAAI Aircraft Groups:
    Group A (א) - Single-Engine Piston (בוכנה חד מנועי)
    Group B (ב) - Multi-Engine Piston (בוכנה רב מנועי)
    Group C (ג) - Multi-Engine Jet/Turboprop (סילון/טורבו פרופ רב מנועי)
    Group D (ד) - Single-Engine Turboprop (טורבו פרופ חד מנועי)
"""

# Maps internal category code to Hebrew group letter
CAAI_GROUP_MAP = {
    'C': 'א',   # SE Piston (קבוצה א')
    'D': 'ד',   # SE Turboprop (קבוצה ד')
    'E': 'ב',   # ME Piston (קבוצה ב')
    'F': 'ג',   # ME Jet/Turboprop (קבוצה ג')
}

# Known multi-engine aircraft types
MULTI_ENGINE_TYPES = {'A319', 'A320', 'H25B', 'PA44', 'BE76'}

# Known complex aircraft (retractable gear + variable pitch propeller)
COMPLEX_TYPES = {'PA44', 'BE76'}

# Type normalization map for variant names
TYPE_NORMALIZATION = {
    'C172R': 'C172',
    'C172K': 'C172',
    'P28A-161': 'PA28',
    'P28A-181': 'PA28',
}


def is_simulator(aircraft_type, registration):
    """Check if this entry represents a simulator/training device.

    Args:
        aircraft_type: Aircraft type code (e.g. 'C172', 'A320 FFS')
        registration: Aircraft registration (e.g. 'N12345', 'FRASCA 142')

    Returns:
        True if this is a simulator/training device.
    """
    atype = aircraft_type.upper()
    reg = registration.upper()
    parts = reg.split()
    return ('SIM' in atype or 'FTD' in atype or 'FFS' in atype or
            'FRASCA' in reg or 'FLIGHT SAFETY' in reg or 'CAE' in reg or
            (bool(parts) and parts[0] == 'ATP'))


def get_caai_category(aircraft_type):
    """Return CAAI category code for an aircraft type.

    Category codes:
        'C' = SE Piston (Group A / קבוצה א')
        'D' = SE Turboprop (Group D / קבוצה ד')
        'E' = ME Piston (Group B / קבוצה ב')
        'F' = ME Jet/Turboprop (Group C / קבוצה ג')

    Args:
        aircraft_type: Aircraft type code (e.g. 'C172', 'A319')

    Returns:
        Category code string ('C', 'D', 'E', or 'F').
    """
    atype = aircraft_type.upper()
    if 'A319' in atype or 'A320' in atype:
        return 'F'
    if 'H25B' in atype:
        return 'F'
    if 'PA44' in atype or 'BE76' in atype:
        return 'E'
    return 'C'


def is_single_engine(aircraft_type):
    """Check if aircraft is single-engine (no SIC concept per CAAI).

    Args:
        aircraft_type: Aircraft type code

    Returns:
        True if single-engine aircraft.
    """
    atype = aircraft_type.upper()
    return atype not in MULTI_ENGINE_TYPES and 'SIM' not in atype


def normalize_type(aircraft_type):
    """Normalize aircraft type code to standard form.

    Handles variant names like 'C172R' -> 'C172', 'P28A-161' -> 'PA28'.

    Args:
        aircraft_type: Raw aircraft type code

    Returns:
        Normalized type code string.
    """
    atype = aircraft_type.upper().strip()
    return TYPE_NORMALIZATION.get(atype, atype)


def is_complex_aircraft(aircraft_type):
    """Check if aircraft is complex (retractable gear + variable pitch prop).

    Args:
        aircraft_type: Aircraft type code (will be normalized)

    Returns:
        True if complex aircraft.
    """
    return normalize_type(aircraft_type) in COMPLEX_TYPES


def get_caai_group_letter(aircraft_type):
    """Get the Hebrew group letter for an aircraft type.

    Args:
        aircraft_type: Aircraft type code

    Returns:
        Hebrew letter string ('א', 'ב', 'ג', or 'ד').
    """
    category = get_caai_category(aircraft_type)
    return CAAI_GROUP_MAP.get(category, 'א')
