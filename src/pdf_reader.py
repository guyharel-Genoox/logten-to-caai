"""
PDF flight logbook table extraction.

Uses pdfplumber to extract tabular data from PDF logbooks.
Handles multi-page tables, repeated headers, and summary rows.

Requires: pip install pdfplumber
"""

import re


def _is_summary_row(row):
    """Check if a row is a summary/total row that should be skipped."""
    if not row:
        return False
    text = ' '.join(str(cell or '') for cell in row).lower()
    return any(kw in text for kw in [
        'total', 'subtotal', 'sub-total', 'grand total',
        'page total', 'totals', 'סה"כ', 'סיכום',
    ])


def _is_empty_row(row):
    """Check if a row is effectively empty."""
    if not row:
        return True
    return all(not str(cell or '').strip() for cell in row)


def _rows_similar(row1, row2, threshold=0.7):
    """Check if two rows have similar content (for header detection).

    Used to detect repeated headers across pages.
    """
    if not row1 or not row2:
        return False
    if len(row1) != len(row2):
        return False
    matches = sum(
        1 for a, b in zip(row1, row2)
        if str(a or '').strip().lower() == str(b or '').strip().lower()
    )
    return matches / len(row1) >= threshold


def _clean_cell(cell):
    """Clean a cell value from PDF extraction artifacts."""
    if cell is None:
        return ''
    s = str(cell).strip()
    # Remove common PDF artifacts
    s = re.sub(r'\s+', ' ', s)
    return s


def read_pdf_tables(pdf_path):
    """Extract tabular flight data from a PDF logbook.

    Tries pdfplumber table extraction first. Merges tables across pages,
    detects and removes repeated headers and summary rows.

    Args:
        pdf_path: Path to the PDF file.

    Returns:
        Tuple of (headers, data_rows) where:
            headers: List of header strings
            data_rows: List of rows, each row is a list of cell strings

    Raises:
        ImportError: If pdfplumber is not installed.
        ValueError: If no tables found in the PDF.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "PDF import requires pdfplumber. Install it with:\n"
            "  pip install pdfplumber"
        )

    headers = None
    data_rows = []
    expected_col_count = None

    with pdfplumber.open(pdf_path) as pdf:
        print(f"  Reading PDF: {len(pdf.pages)} pages")

        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            if not tables:
                # Try with different settings
                tables = page.extract_tables({
                    'vertical_strategy': 'text',
                    'horizontal_strategy': 'text',
                })

            if not tables:
                continue

            # Use the largest table on each page
            table = max(tables, key=lambda t: len(t))

            for row_idx, row in enumerate(table):
                if _is_empty_row(row):
                    continue

                # Clean all cells
                cleaned = [_clean_cell(c) for c in row]

                # First non-empty row with text = potential header
                if headers is None:
                    # Check if this looks like a header (has text, not numbers)
                    text_cells = sum(1 for c in cleaned if c and not c.replace('.', '').replace(',', '').replace(':', '').isdigit())
                    if text_cells >= len(cleaned) * 0.4:
                        headers = cleaned
                        expected_col_count = len(headers)
                        print(f"  Page {page_num}: Found headers ({expected_col_count} columns)")
                        continue

                if headers is None:
                    continue

                # Skip repeated headers on subsequent pages
                if _rows_similar(cleaned, headers):
                    continue

                # Skip summary rows
                if _is_summary_row(cleaned):
                    continue

                # Normalize row length to match headers
                if len(cleaned) < expected_col_count:
                    cleaned.extend([''] * (expected_col_count - len(cleaned)))
                elif len(cleaned) > expected_col_count:
                    cleaned = cleaned[:expected_col_count]

                data_rows.append(cleaned)

    if headers is None:
        raise ValueError(
            f"No tables found in {pdf_path}. "
            f"The PDF may not contain structured tabular data. "
            f"Try exporting your logbook as Excel or CSV instead."
        )

    print(f"  Extracted {len(data_rows)} flight records from PDF")
    return headers, data_rows
