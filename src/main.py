#!/usr/bin/env python3
"""
QSP Quarterly Reporting CLI

Usage:
    python -m src.main generate --year 2025 --quarter 3 --bdo-file ./inputs/Cijfers_QSP.xlsx ...
    python -m src.main validate --input-dir ./inputs
    python -m src.main generate --config config/config.yaml
"""

import argparse
import sys
from pathlib import Path
from loguru import logger

from .orchestrator import QuarterlyReportOrchestrator, QuarterlyReportConfig, run_quarterly_report


def setup_logging(log_file: str = None, verbose: bool = False):
    """Configure logging."""
    logger.remove()
    
    level = "DEBUG" if verbose else "INFO"
    
    # Console output with colors
    logger.add(
        sys.stderr,
        format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | {message}",
        level=level,
        colorize=True
    )
    
    # File output
    if log_file:
        logger.add(
            log_file,
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} | {message}",
            level="DEBUG",
            rotation="10 MB"
        )


def cmd_generate(args):
    """Generate quarterly report."""
    logger.info("Starting quarterly report generation")
    
    # Check if config file is provided
    if args.config:
        logger.info(f"Loading configuration from {args.config}")
        results = run_quarterly_report(args.config)
        # Note: dry_run not supported with config file yet
    else:
        # Use individual file arguments
        config = QuarterlyReportConfig(
            year=args.year,
            quarter=args.quarter,
            bdo_file=Path(args.bdo_file),
            previous_management_accounts=Path(args.prev_ma),
            rent_roll_file=Path(args.rent_roll),
            sales_tracker_file=Path(args.sales_tracker),
            previous_compliance_calc=Path(args.prev_compliance),
            word_template=Path(args.word_template),
            output_dir=Path(args.output_dir)
        )
        
        orchestrator = QuarterlyReportOrchestrator(config)
        results = orchestrator.run(dry_run=args.dry_run)
    
    # Display results
    if results['status'] == 'success':
        logger.success("✓ Quarterly report generated successfully")
        if results.get('output_files'):
            logger.info("Generated files:")
            for output_file in results['output_files']:
                logger.info(f"  → {output_file}")
        
        if results.get('warnings'):
            logger.warning(f"⚠ {len(results['warnings'])} warning(s) found:")
            for warning in results['warnings'][:10]:  # Show first 10
                logger.warning(f"  • {warning}")
            if len(results['warnings']) > 10:
                logger.warning(f"  ... and {len(results['warnings']) - 10} more")
    else:
        logger.error("✗ Report generation failed")
        if results.get('errors'):
            for error in results['errors']:
                logger.error(f"  → {error}")
        sys.exit(1)


def cmd_validate(args):
    """Validate input files."""
    from .parsers.bdo_parser import BDOParser
    
    logger.info(f"Validating input files in {args.input_dir}")
    
    input_dir = Path(args.input_dir)
    if not input_dir.exists():
        logger.error(f"Input directory does not exist: {input_dir}")
        sys.exit(1)
    
    issues = []
    
    # Check for expected files
    expected_patterns = [
        ("Cijfers*.xlsx", "BDO quarterly financials"),
        ("Management*Accounts*.xlsx", "Previous Management Accounts"),
        ("*huurlijst*.xlsx", "Rent roll"),
        ("Unit*Sales*.xlsx", "Sales tracker"),
        ("Compliance*.xlsx", "Compliance calculator"),
        ("Quarterly*.docx", "Word template"),
    ]
    
    for pattern, description in expected_patterns:
        matches = list(input_dir.glob(pattern))
        if not matches:
            issues.append(f"Missing: {description} ({pattern})")
        else:
            logger.success(f"✓ Found {description}: {matches[0].name}")
    
    # Try to parse BDO file if found
    bdo_files = list(input_dir.glob("Cijfers*.xlsx"))
    if bdo_files:
        try:
            parser = BDOParser()
            result = parser.parse(str(bdo_files[0]))
            logger.success(f"✓ BDO file parsed successfully: {len(result.accounts)} accounts")
            if result.warnings:
                logger.warning(f"⚠ BDO file has {len(result.warnings)} warning(s)")
        except Exception as e:
            issues.append(f"BDO file parsing failed: {e}")
            logger.error(f"✗ BDO file parsing failed: {e}")
    
    if issues:
        logger.error("Validation issues found:")
        for issue in issues:
            logger.error(f"  → {issue}")
        sys.exit(1)
    else:
        logger.success("✓ All input files validated successfully")


def main():
    parser = argparse.ArgumentParser(
        description="QSP Quarterly Reporting Automation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Generate report with config file
  python -m src.main generate --config config/config.yaml
  
  # Generate report with individual files
  python -m src.main generate --year 2025 --quarter 3 \\
      --bdo-file ./inputs/Cijfers_QSP.xlsx \\
      --prev-ma ./inputs/Management_Accounts_Q2.xlsx \\
      --rent-roll ./inputs/QSP_huurlijst.xlsx \\
      --sales-tracker ./inputs/Unit_Sales_tracker.xlsx \\
      --prev-compliance ./inputs/Compliance_Q2.xlsx \\
      --word-template ./inputs/Quarterly_QSP_Q2.docx \\
      --output-dir ./outputs
  
  # Validate input files
  python -m src.main validate --input-dir ./inputs
  
  # Dry run (validate without generating)
  python -m src.main generate --config config/config.yaml --dry-run
        """
    )
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose logging')
    parser.add_argument('--log-file', help='Log file path')
    
    subparsers = parser.add_subparsers(dest='command', help='Command to run', required=True)
    
    # Generate command
    gen_parser = subparsers.add_parser('generate', help='Generate quarterly report')
    gen_group = gen_parser.add_mutually_exclusive_group(required=True)
    gen_group.add_argument('--config', help='Path to config YAML file')
    gen_group.add_argument('--year', type=int, help='Report year (required if not using --config)')
    
    gen_parser.add_argument('--quarter', type=int, choices=[1, 2, 3, 4], 
                           help='Quarter number (required if not using --config)')
    gen_parser.add_argument('--bdo-file', help='BDO Excel file path (required if not using --config)')
    gen_parser.add_argument('--prev-ma', help='Previous Management Accounts file (required if not using --config)')
    gen_parser.add_argument('--rent-roll', help='Rent roll Excel file (required if not using --config)')
    gen_parser.add_argument('--sales-tracker', help='Sales tracker Excel file (required if not using --config)')
    gen_parser.add_argument('--prev-compliance', help='Previous compliance calculator (required if not using --config)')
    gen_parser.add_argument('--word-template', help='Word template file (required if not using --config)')
    gen_parser.add_argument('--output-dir', help='Output directory (required if not using --config)')
    gen_parser.add_argument('--dry-run', action='store_true', help='Validate only, do not generate files')
    gen_parser.set_defaults(func=cmd_generate)
    
    # Validate command
    val_parser = subparsers.add_parser('validate', help='Validate input files')
    val_parser.add_argument('--input-dir', required=True, help='Input directory to validate')
    val_parser.set_defaults(func=cmd_validate)
    
    args = parser.parse_args()
    
    setup_logging(args.log_file, args.verbose)
    
    # Validate required arguments for generate command when not using config
    if args.command == 'generate' and not args.config:
        required_args = ['year', 'quarter', 'bdo_file', 'prev_ma', 'rent_roll', 
                        'sales_tracker', 'prev_compliance', 'word_template', 'output_dir']
        missing = [arg for arg in required_args if not getattr(args, arg, None)]
        if missing:
            parser.error(f"Missing required arguments when not using --config: {', '.join(missing)}")
    
    args.func(args)


if __name__ == '__main__':
    main()

