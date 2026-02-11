#!/usr/bin/env python3
"""
LogTen Pro to Israeli CAAI tofes-shaot pipeline runner.

Converts a Coradine LogTen Pro flight log export into:
1. A comprehensive Excel flight logbook (4 sheets)
2. Distance calculations for all flights
3. CAAI classification columns
4. A filled CAAI tofes-shaot form (Summary, CPL, ATPL sheets)

Usage:
    python run.py                              # Full pipeline using config.ini
    python run.py --export flights.txt         # Override input file
    python run.py --step logbook               # Run single step
    python run.py --step distances
    python run.py --step caai-columns
    python run.py --step fill-form
    python run.py --step analyze               # Optional verification
"""

import argparse
import sys
import os

from src.config import Config


STEPS = ['logbook', 'distances', 'caai-columns', 'fill-form', 'analyze']


def run_logbook(config):
    """Step 1: Create Excel logbook from LogTen Pro export."""
    from src.logbook_creator import create_logbook
    print("\n" + "=" * 70)
    print("STEP 1: Creating flight logbook from LogTen Pro export")
    print("=" * 70)
    config.validate('logbook')
    create_logbook(config.logten_export, config.logbook_output, config.pilot_name)


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
    'logbook': run_logbook,
    'distances': run_distances,
    'caai-columns': run_caai_columns,
    'fill-form': run_fill_form,
    'analyze': run_analyze,
}


def main():
    parser = argparse.ArgumentParser(
        description='LogTen Pro to Israeli CAAI tofes-shaot pipeline',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Steps (run in order, or individually with --step):
  logbook        Create Excel logbook from LogTen Pro export
  distances      Add great-circle distances to logbook
  caai-columns   Add CAAI classification columns to logbook
  fill-form      Fill CAAI tofes-shaot form from logbook
  analyze        Verify CAAI categorization (optional)

Examples:
  python run.py                              # Run full pipeline
  python run.py --step fill-form             # Run only form filling
  python run.py --export "my_flights.txt"    # Override input file
        """,
    )
    parser.add_argument('--config', '-c', default='config.ini',
                        help='Config file path (default: config.ini)')
    parser.add_argument('--step', '-s', choices=STEPS,
                        help='Run only this step')
    parser.add_argument('--export', '-e', default=None,
                        help='Override LogTen export file path')
    parser.add_argument('--logbook', '-l', default=None,
                        help='Override logbook file path')
    parser.add_argument('--pilot', '-p', default=None,
                        help='Override pilot name')
    parser.add_argument('--template', '-t', default=None,
                        help='Override CAAI form template path')
    parser.add_argument('--output', '-o', default=None,
                        help='Override CAAI form output path')

    args = parser.parse_args()

    # Load config
    config = Config.from_file(args.config)

    # Apply CLI overrides
    config.override(
        logten_export=args.export,
        logbook_output=args.logbook,
        pilot_name=args.pilot,
        caai_template=args.template,
        caai_output=args.output,
    )

    print("LogTen Pro -> CAAI tofes-shaot Pipeline")
    print("=" * 70)
    print(f"Config: {os.path.abspath(args.config)}")
    if config.pilot_name:
        print(f"Pilot: {config.pilot_name}")
    print(f"LogTen export: {config.logten_export}")
    print(f"Logbook: {config.logbook_output}")
    print(f"CAAI template: {config.caai_template}")
    print(f"CAAI output: {config.caai_output}")

    try:
        if args.step:
            # Run single step
            STEP_FUNCTIONS[args.step](config)
        else:
            # Run full pipeline (steps 1-4, skip analyze)
            for step in ['logbook', 'distances', 'caai-columns', 'fill-form']:
                STEP_FUNCTIONS[step](config)

        print("\n" + "=" * 70)
        print("PIPELINE COMPLETE")
        print("=" * 70)
        if not args.step or args.step == 'fill-form':
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
