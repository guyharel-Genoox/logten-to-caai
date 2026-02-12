#!/usr/bin/env python3
"""
Flight logbook to Israeli CAAI tofes-shaot pipeline runner.

Converts any flight logbook (Excel, CSV, PDF, or LogTen Pro export) into:
1. A standardized Excel flight logbook (48 columns)
2. Distance calculations for all flights
3. CAAI classification columns
4. A filled CAAI tofes-shaot form (Summary, CPL, ATPL sheets)

Usage:
    python run.py --input flights.xlsx                    # Any Excel logbook
    python run.py --input flights.csv                     # CSV logbook
    python run.py --input flights.pdf                     # PDF logbook
    python run.py --input "Export Flights (Tab).txt" --format logten  # LogTen Pro
    python run.py                                         # Uses config.ini
    python run.py --step fill-form                        # Run single step
    python run.py --step analyze                          # Verify CAAI compliance
"""

import argparse
import sys
import os

from src.config import Config


STEPS = ['import', 'distances', 'caai-columns', 'fill-form', 'analyze']
# Keep 'logbook' as hidden alias for backward compatibility
STEP_ALIASES = {'logbook': 'import'}


def run_import(config):
    """Step 1: Import flight data into standardized Excel logbook.

    Auto-detects the source format and column layout, or uses explicit
    settings from config/CLI.
    """
    source, fmt = config.get_import_source()

    print("\n" + "=" * 70)
    print("STEP 1: Importing flight data")
    print("=" * 70)
    print(f"  Source: {source}")
    print(f"  Format: {fmt}")

    config.validate('import')

    if fmt == 'logten':
        # Use the original LogTen Pro importer
        from src.logbook_creator import create_logbook
        print("  Using LogTen Pro importer")
        create_logbook(source, config.logbook_output, config.pilot_name)
    else:
        # Use universal importer
        from src.universal_importer import create_standardized_logbook
        print("  Using universal importer")
        create_standardized_logbook(
            input_file=source,
            output_file=config.logbook_output,
            pilot_name=config.pilot_name,
            fmt=fmt,
            column_mapping=config.column_mapping or None,
        )


def run_distances(config):
    """Step 2: Add distance calculations to logbook."""
    from src.distance_calculator import add_distances
    print("\n" + "=" * 70)
    print("STEP 2: Calculating flight distances")
    print("=" * 70)
    config.validate('distances')
    custom = config.custom_airports if config.custom_airports else None
    add_distances(config.logbook_output, custom)


def run_caai_columns(config):
    """Step 3: Add CAAI classification columns."""
    from src.caai_columns import add_caai_columns
    print("\n" + "=" * 70)
    print("STEP 3: Adding CAAI classification columns")
    print("=" * 70)
    config.validate('caai-columns')
    add_caai_columns(config.logbook_output)


def run_fill_form(config):
    """Step 4: Fill the CAAI tofes-shaot form."""
    from src.caai_form_filler import fill_caai_form
    print("\n" + "=" * 70)
    print("STEP 4: Filling CAAI tofes-shaot form")
    print("=" * 70)
    config.validate('fill-form')
    fill_caai_form(config.logbook_output, config.caai_template, config.caai_output)


def run_analyze(config):
    """Step 5 (optional): Analyze and verify CAAI categorization."""
    from src.caai_analyzer import analyze_caai
    print("\n" + "=" * 70)
    print("STEP 5: CAAI compliance analysis")
    print("=" * 70)
    config.validate('analyze')
    analyze_caai(config.logbook_output)


STEP_FUNCTIONS = {
    'import': run_import,
    'logbook': run_import,  # backward compat alias
    'distances': run_distances,
    'caai-columns': run_caai_columns,
    'fill-form': run_fill_form,
    'analyze': run_analyze,
}


def main():
    parser = argparse.ArgumentParser(
        description='Flight logbook to Israeli CAAI tofes-shaot pipeline',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Steps (run in order, or individually with --step):
  import         Import flight data from any format (Excel, CSV, PDF, LogTen)
  distances      Add great-circle distances to logbook
  caai-columns   Add CAAI classification columns to logbook
  fill-form      Fill CAAI tofes-shaot form from logbook
  analyze        Verify CAAI categorization (optional)

Supported input formats:
  .xlsx / .xls   Excel spreadsheet
  .csv           Comma-separated values
  .tsv / .txt    Tab-separated values
  .pdf           PDF with tabular data
  LogTen Pro     Tab-delimited export (auto-detected or --format logten)

Examples:
  python run.py --input my_logbook.xlsx                   # Excel logbook
  python run.py --input flights.csv                       # CSV logbook
  python run.py --input logbook.pdf                       # PDF logbook
  python run.py --input flights.csv --mapping columns.ini # With column mapping
  python run.py --input export.txt --format logten        # LogTen Pro export
  python run.py                                           # Use config.ini
  python run.py --step fill-form                          # Run only form filling
        """,
    )
    parser.add_argument('--config', '-c', default='config.ini',
                        help='Config file path (default: config.ini)')
    parser.add_argument('--step', '-s', choices=STEPS + ['logbook'],
                        help='Run only this step')
    parser.add_argument('--input', '-i', default=None,
                        help='Input file (any format: Excel, CSV, PDF, LogTen)')
    parser.add_argument('--format', '-f', default=None,
                        choices=['auto', 'excel', 'csv', 'tsv', 'pdf', 'logten'],
                        help='Input format (default: auto-detect)')
    parser.add_argument('--mapping', '-m', default=None,
                        help='Column mapping INI file (default: auto-detect)')
    parser.add_argument('--logbook', '-l', default=None,
                        help='Override logbook output file path')
    parser.add_argument('--pilot', '-p', default=None,
                        help='Override pilot name')
    parser.add_argument('--template', '-t', default=None,
                        help='Override CAAI form template path')
    parser.add_argument('--output', '-o', default=None,
                        help='Override CAAI form output path')
    # Keep --export for backward compatibility
    parser.add_argument('--export', '-e', default=None,
                        help=argparse.SUPPRESS)  # hidden, backward compat

    args = parser.parse_args()

    # Resolve step aliases
    step = args.step
    if step in STEP_ALIASES:
        step = STEP_ALIASES[step]

    # Load config
    config = Config.from_file(args.config)

    # Apply CLI overrides
    if args.input:
        config.override(input_file=args.input)
    elif args.export:
        # Backward compat: --export sets logten format
        config.override(input_file=args.export, input_format='logten')

    config.override(
        input_format=args.format,
        column_mapping=args.mapping,
        logbook_output=args.logbook,
        pilot_name=args.pilot,
        caai_template=args.template,
        caai_output=args.output,
    )

    # If no input_file but logten_export is set, use it (backward compat)
    if not config.input_file and config.logten_export:
        config.input_file = config.logten_export
        if config.input_format == 'auto':
            config.input_format = 'logten'

    source, fmt = config.get_import_source()

    print("Flight Logbook -> CAAI tofes-shaot Pipeline")
    print("=" * 70)
    print(f"Config: {os.path.abspath(args.config)}")
    if config.pilot_name:
        print(f"Pilot: {config.pilot_name}")
    print(f"Input: {source or '(not set)'} [{fmt or 'auto'}]")
    print(f"Logbook: {config.logbook_output}")
    print(f"CAAI template: {config.caai_template}")
    print(f"CAAI output: {config.caai_output}")

    try:
        if step:
            # Run single step
            STEP_FUNCTIONS[step](config)
        else:
            # Run full pipeline (steps 1-4, skip analyze)
            for s in ['import', 'distances', 'caai-columns', 'fill-form']:
                STEP_FUNCTIONS[s](config)

        print("\n" + "=" * 70)
        print("PIPELINE COMPLETE")
        print("=" * 70)
        if not step or step == 'fill-form':
            print(f"\nOutput files:")
            print(f"  Logbook: {config.logbook_output}")
            print(f"  CAAI form: {config.caai_output}")

    except FileNotFoundError as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        raise


if __name__ == '__main__':
    main()
