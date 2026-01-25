"""
Main Orchestrator

Coordinates the entire quarterly reporting workflow.
"""

from dataclasses import dataclass
from typing import Optional
from pathlib import Path
from datetime import datetime
import yaml
import glob
from loguru import logger

from .parsers.bdo_parser import BDOParser
from .parsers.rent_roll_parser import RentRollParser
from .parsers.sales_tracker_parser import SalesTrackerParser
from .transformers.management_accounts import ManagementAccountsBuilder, ManagementAccountsConfig
from .transformers.compliance_builder import ComplianceBuilder, ComplianceConfig
from .transformers.compliance_calc import ComplianceCalculator
from .generators.word_updater import WordTemplateUpdater, ReportValues
from .generators.pdf_assembler import PDFAssembler, PDFSource
from .validators.balance_validator import BalanceValidator
from .validators.covenant_validator import CovenantValidator
from .validators.reconciliation import ReconciliationValidator


@dataclass
class QuarterlyReportConfig:
    """Configuration for quarterly report generation."""
    year: int
    quarter: int
    
    # Input file paths
    bdo_file: Path
    previous_management_accounts: Path
    rent_roll_file: Path
    sales_tracker_file: Path
    previous_compliance_calc: Path
    word_template: Path
    
    # Output paths
    output_dir: Path
    
    @property
    def quarter_str(self) -> str:
        return f"Q{self.quarter} {self.year}"
    
    @property
    def previous_quarter_str(self) -> str:
        prev_q = self.quarter - 1
        prev_y = self.year
        if prev_q == 0:
            prev_q = 4
            prev_y -= 1
        return f"Q{prev_q} {prev_y}"
    
    @property
    def period_end(self) -> datetime:
        month = self.quarter * 3
        if month == 3:
            day = 31
        elif month == 6:
            day = 30
        elif month == 9:
            day = 30
        else:
            day = 31
        return datetime(self.year, month, day)


@dataclass
class WordReportConfig:
    """Configuration for Word report generation (Phase 2)."""
    year: int
    quarter: int
    
    # Input file paths
    management_accounts_file: Path  # Generated MA from Phase 1
    compliance_certificate_file: Path  # Edited compliance cert from client
    rent_roll_file: Path
    sales_tracker_file: Path
    word_template: Path
    
    # Output paths
    output_dir: Path
    
    @property
    def quarter_str(self) -> str:
        return f"Q{self.quarter} {self.year}"
    
    @property
    def previous_quarter_str(self) -> str:
        prev_q = self.quarter - 1
        prev_y = self.year
        if prev_q == 0:
            prev_q = 4
            prev_y -= 1
        return f"Q{prev_q} {prev_y}"
    
    @property
    def period_end(self) -> datetime:
        month = self.quarter * 3
        if month == 3:
            day = 31
        elif month == 6:
            day = 30
        elif month == 9:
            day = 30
        else:
            day = 31
        return datetime(self.year, month, day)


@dataclass
class FinalPDFConfig:
    """Configuration for final PDF generation (Phase 3)."""
    year: int
    quarter: int
    
    # Input file paths
    word_report_file: Path  # Final Word report (potentially edited)
    management_accounts_file: Path  # Final MA
    compliance_certificate_file: Path  # Final compliance cert (client-approved)
    rent_roll_file: Path
    sales_tracker_file: Path
    
    # Output paths
    output_dir: Path
    
    @property
    def quarter_str(self) -> str:
        return f"Q{self.quarter} {self.year}"


class QuarterlyReportOrchestrator:
    """
    Orchestrates the complete quarterly reporting workflow.
    
    Workflow Steps:
    1. Parse BDO quarterly financials
    2. Parse rent roll
    3. Parse sales tracker
    4. Build Management Accounts
    5. Calculate compliance metrics
    6. Validate all data
    7. Update Word template
    8. Assemble final PDF
    """
    
    def __init__(self, config: QuarterlyReportConfig):
        self.config = config
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialize parsers and processors
        self.bdo_parser = BDOParser()
        self.rent_roll_parser = RentRollParser()
        self.sales_parser = SalesTrackerParser()
        self.compliance_calc = ComplianceCalculator()
        self.balance_validator = BalanceValidator()
        self.covenant_validator = CovenantValidator()
        self.reconciliation_validator = ReconciliationValidator()
        
    def run(self, dry_run: bool = False) -> dict:
        """
        Execute the complete quarterly reporting workflow.
        
        Args:
            dry_run: If True, validate inputs but don't generate outputs
            
        Returns:
            Dict with workflow results and any warnings
        """
        logger.info(f"Starting quarterly report generation for {self.config.quarter_str}")
        
        results = {
            'quarter': self.config.quarter_str,
            'started_at': datetime.now().isoformat(),
            'steps': {},
            'warnings': [],
            'errors': [],
            'output_files': []
        }
        
        try:
            # Step 1: Parse BDO data
            logger.info("Step 1: Parsing BDO quarterly financials")
            bdo_result = self.bdo_parser.parse(str(self.config.bdo_file))
            results['steps']['bdo_parse'] = {
                'status': 'success',
                'accounts_parsed': len(bdo_result.accounts),
                'warnings': bdo_result.warnings
            }
            if bdo_result.warnings:
                results['warnings'].extend(bdo_result.warnings)
            
            # Step 2: Parse rent roll
            logger.info("Step 2: Parsing rent roll")
            rent_roll = self.rent_roll_parser.parse(str(self.config.rent_roll_file))
            results['steps']['rent_roll_parse'] = {
                'status': 'success',
                'units_count': rent_roll.total_units,
                'annual_rent': rent_roll.total_annual_rent
            }
            if rent_roll.warnings:
                results['warnings'].extend(rent_roll.warnings)
            
            # Step 3: Parse sales tracker
            logger.info("Step 3: Parsing sales tracker")
            sales_data = self.sales_parser.parse(
                str(self.config.sales_tracker_file),
                period_end=self.config.period_end
            )
            results['steps']['sales_parse'] = {
                'status': 'success',
                'total_sales': sales_data.total_units_sold,
                'quarter_sales': sales_data.quarter_units_sold,
                'quarter_proceeds': sales_data.quarter_proceeds
            }
            if sales_data.warnings:
                results['warnings'].extend(sales_data.warnings)
            
            if dry_run:
                logger.info("Dry run - stopping before file generation")
                results['dry_run'] = True
                results['completed_at'] = datetime.now().isoformat()
                results['status'] = 'success'
                return results
            
            # Step 4: Build Management Accounts
            logger.info("Step 4: Building Management Accounts")
            ma_output = self.config.output_dir / f"Management Accounts {self.config.quarter_str} - Draft 1.xlsx"
            ma_config = ManagementAccountsConfig(
                quarter=self.config.quarter_str,
                period_end=self.config.period_end,
                bdo_sheet_name=f"BDO - Q{self.config.quarter}-{str(self.config.year)[-2:]}",
                summary_sheet_name=f"Management Cijfers - {self.config.quarter_str}",
                cash_proceeds_sale=sales_data.quarter_proceeds  # Row 50: from sales tracker
            )
            ma_builder = ManagementAccountsBuilder(
                str(self.config.previous_management_accounts),
                str(ma_output),
                ma_config,
                bdo_source_path=str(self.config.bdo_file)  # Pass BDO file for direct copy
            )
            ma_builder.build(bdo_result)
            results['steps']['management_accounts'] = {
                'status': 'success',
                'output': str(ma_output)
            }
            results['output_files'].append(str(ma_output))
            
            # Step 4b: Build Compliance Certificate
            # The Q{n} Management Accounts sheet should contain CURRENT quarter's LTM data.
            # We pass the newly generated Management Accounts file (ma_output).
            logger.info("Step 4b: Building Compliance Certificate")
            compliance_output = self.config.output_dir / f"Compliance Certificate Berekening QSP - {self.config.quarter_str}_updated.xlsx"
            compliance_config = ComplianceConfig(
                year=self.config.year,
                quarter=self.config.quarter
            )
            compliance_builder = ComplianceBuilder(
                str(self.config.previous_compliance_calc),
                str(compliance_output),
                compliance_config
            )
            # Pass the newly generated Management Accounts file for Q{n} MA sheet data (LTM Q{n})
            compliance_builder.build(bdo_result, str(ma_output))
            results['steps']['compliance_certificate'] = {
                'status': 'success',
                'output': str(compliance_output)
            }
            results['output_files'].append(str(compliance_output))
            
            # Step 5: Calculate compliance
            logger.info("Step 5: Calculating compliance metrics")
            totals = self.bdo_parser.calculate_totals(bdo_result)
            
            # Calculate NOI (simplified: 85% of annual rent)
            net_operating_income = rent_roll.total_annual_rent * 0.85
            
            # Calculate debt service from BDO (interest expense + principal)
            # This is a simplified calculation - in practice, this should come from loan schedule
            debt_service = 2_500_000  # TODO: Calculate from BDO or loan schedule
            
            # Calculate net debt (liabilities - current assets)
            net_debt = totals['total_liabilities'] - totals['current_assets']
            
            # Property value (real estate net)
            property_value = totals['real_estate_net']
            
            compliance_result = self.compliance_calc.calculate(
                net_operating_income=net_operating_income,
                debt_service=debt_service,
                net_debt=net_debt,
                property_value=property_value,
                quarter=self.config.quarter_str
            )
            
            # Validate covenants
            covenant_validation = self.covenant_validator.validate(compliance_result)
            
            results['steps']['compliance'] = {
                'status': 'success',
                'overall': compliance_result.overall_status.value,
                'covenants': {k: v.status.value for k, v in compliance_result.covenants.items()},
                'failed_count': len(covenant_validation.failed_covenants),
                'warning_count': len(covenant_validation.warning_covenants)
            }
            if compliance_result.warnings:
                results['warnings'].extend(compliance_result.warnings)
            if covenant_validation.failed_covenants:
                for issue in covenant_validation.failed_covenants:
                    results['errors'].append(issue.message)
            
            # Step 6: Validate data
            logger.info("Step 6: Validating calculations")
            validation = self.balance_validator.validate(bdo_result, totals)
            results['steps']['validation'] = {
                'status': 'success' if validation.passed else 'warning',
                'passed': validation.passed,
                'checks': [
                    {
                        'name': check.name,
                        'passed': check.passed,
                        'message': check.message
                    }
                    for check in validation.checks
                ]
            }
            if validation.warnings:
                results['warnings'].extend(validation.warnings)
            if validation.errors:
                results['errors'].extend(validation.errors)
            
            # Cross-file reconciliation
            executive_summary = {
                'rent_roll_annual': rent_roll.total_annual_rent / 1000,  # In thousands
                'units_sold': sales_data.quarter_units_sold
            }
            reconciliation = self.reconciliation_validator.validate(
                rent_roll=rent_roll,
                sales_tracker=sales_data,
                bdo_result=bdo_result,
                executive_summary=executive_summary,
                management_accounts_cash=totals.get('current_assets', 0.0)  # Simplified
            )
            if not reconciliation.passed:
                results['warnings'].extend(reconciliation.warnings)
            
            # Step 7: Update Word template
            logger.info("Step 7: Updating Word template")
            word_output = self.config.output_dir / f"Quarterly QSP - {self.config.quarter_str} - Draft.docx"
            
            # Extract values from the generated Management Accounts Excel file
            # This ensures we use the same calculated values that appear in the Excel output
            from .generators.word_updater import extract_values_from_management_accounts
            
            ma_values = extract_values_from_management_accounts(
                str(ma_output),
                self.config.quarter,
                self.config.year
            )
            
            # Use extracted values, with fallbacks to BDO data
            gtri_amount = ma_values.get('gtri', 0.0)
            gtri_ltm = ma_values.get('gtri_ltm', 0.0)
            vacancy_amount = ma_values.get('vacancy_amount', 0.0)
            gross_rental_income = ma_values.get('gross_rental', 0.0)
            gross_rental_ltm = ma_values.get('gross_rental_ltm', 0.0)
            vacancy_pct = ma_values.get('vacancy_pct', 0.0)
            maintenance_amount = ma_values.get('maintenance', 0.0)
            maintenance_ltm = ma_values.get('maintenance_ltm', 0.0)
            
            # Fallback: Extract from BDO data if MA extraction failed
            if gtri_amount == 0:
                for code, entry in bdo_result.accounts.items():
                    if code == '8000003':
                        gtri_amount = abs(entry.closing_balance - entry.opening_balance) / 1000
                    elif code == '8000004':
                        vacancy_amount = abs(entry.closing_balance - entry.opening_balance) / 1000
                    elif code.startswith('4100'):
                        maintenance_amount += abs(entry.closing_balance - entry.opening_balance) / 1000
                
                gross_rental_income = gtri_amount - vacancy_amount
                vacancy_pct = (vacancy_amount / gtri_amount * 100) if gtri_amount > 0 else 0.0
                # Estimate LTM as 4x quarterly (approximation)
                gtri_ltm = gtri_amount * 4
                gross_rental_ltm = gross_rental_income * 4
                maintenance_ltm = maintenance_amount * 4
            
            # Calculate unit sales proceeds from sales tracker (in thousands)
            unit_sales_proceeds = ma_values.get('unit_sales_proceeds', 0.0)
            if unit_sales_proceeds == 0 and sales_data.quarter_proceeds:
                unit_sales_proceeds = sales_data.quarter_proceeds / 1000
            
            # Rent roll total (annual, in thousands)
            rent_roll_k = rent_roll.total_annual_rent / 1000
            
            logger.info(f"Page 4 KPIs extracted from Management Accounts:")
            logger.info(f"  GTRI (Q3): €{gtri_amount:,.1f}k")
            logger.info(f"  GTRI (LTM): €{gtri_ltm:,.1f}k")
            logger.info(f"  Gross rental income: €{gross_rental_income:,.1f}k")
            logger.info(f"  Financial vacancy: {vacancy_pct:.1f}% (€{vacancy_amount:,.1f}k)")
            logger.info(f"  Rent roll (annual): €{rent_roll_k:,.1f}k")
            logger.info(f"  Maintenance: €{maintenance_amount:,.1f}k")
            logger.info(f"  Unit sales proceeds: €{unit_sales_proceeds:,.1f}k")
            logger.info(f"  Unit sales: {sales_data.quarter_units_sold} units, €{unit_sales_proceeds:,.0f}k")
            
            report_values = ReportValues(
                report_date=datetime.now().strftime("%d %B %Y"),
                # Page 4 - from BDO/Management Accounts
                gtri=gtri_amount,  # Q3 GTRI in thousands
                # gtri_ltm is used for "rent roll yields" value in Word document
                # Use rent_roll_k (annual rent from rent roll file) as this matches the expected document
                gtri_ltm=rent_roll_k,  # Rent roll annual (appears as "yields €X,XXXk" in Word)
                gross_rental_income=gross_rental_income,  # Q3 gross rental in thousands
                gross_rental_income_ltm=gross_rental_ltm,
                financial_vacancy_pct=vacancy_pct,
                financial_vacancy_amount=vacancy_amount,
                # Page 4 - from Rent Roll
                rent_roll_annual=rent_roll_k,
                rent_roll_units=rent_roll.total_units,
                # Page 4 - from Sales Tracker
                units_sold_quarter=sales_data.quarter_units_sold,
                unit_sales_proceeds=unit_sales_proceeds,
                # Page 4/6 - Maintenance and CAPEX
                maintenance_amount=maintenance_amount,
                maintenance_ltm=maintenance_ltm,
                capex_amount=0,  # CAPEX is entered manually per mainPlan.pdf
                capex_ltm=0,
                # Narratives
                unit_sales_narrative=f"{sales_data.quarter_units_sold} unit(s) sold for €{unit_sales_proceeds:,.0f}k" if sales_data.quarter_units_sold > 0 else "No unit sales this quarter.",
                maintenance_detail="",
                sustainability_detail="",
                # Previous quarter values (Q2 2025 values from template document)
                # Used for finding and replacing old values with new values
                prev_gtri=3200.6,  # Q2 GTRI
                prev_gtri_ltm=12940.2,  # Q2 LTM GTRI (shown as "rent roll yields")
                prev_gross_rental_income=3067.5,  # Q2 gross rental income
                prev_vacancy_pct=4.1,  # Q2 vacancy percentage
                prev_vacancy_amount=106.3,  # Q2 vacancy amount
                prev_rent_roll=0,
                prev_maintenance=297.0,  # Q2 maintenance
                prev_unit_sales_proceeds=762.5,  # Q2 unit sales proceeds
            )
            word_updater = WordTemplateUpdater(str(self.config.word_template), str(word_output))
            word_updater.update_with_python_docx(
                report_values,
                self.config.previous_quarter_str,
                self.config.quarter_str
            )
            results['steps']['word_update'] = {
                'status': 'success',
                'output': str(word_output)
            }
            results['output_files'].append(str(word_output))
            
            # Step 8: Assemble PDF (optional - requires LibreOffice)
            # Per Issue 8: Each section introduction should be immediately followed by its attachment
            # The reference PDF structure has sections interleaved with attachments, not bulk at end
            logger.info("Step 8: Assembling final PDF")
            try:
                pdf_output = self.config.output_dir / f"Quarterly QSP - {self.config.quarter_str} - ASX.pdf"
                
                # PDF Assembly Order (per reference document structure):
                # 1. Main Report (Word) - includes section intros
                # 2. Rent Roll (Excel) - Attachment 1
                # 3. Management Accounts (Excel) - Attachment 2
                # 4. Sales Tracker (Excel) - Attachment 3
                # 5. Compliance Certificate (Excel sheets) - Attachment 4
                #    - SFA CC (main compliance)
                #    - Q{quarter} Management Accounts (MA data in CC)
                #    - Suppl. Calc
                #    - Impact Unit Sales
                pdf_sources = [
                    # Section 1: Main Report
                    PDFSource("Main Report", word_output, 'docx'),
                    
                    # Attachment 1: Rent Roll
                    PDFSource("Rent Roll", self.config.rent_roll_file, 'xlsx'),
                    
                    # Attachment 2: Management Accounts
                    PDFSource(
                        "Management Accounts",
                        ma_output,
                        'xlsx',
                        ma_config.summary_sheet_name
                    ),
                    
                    # Attachment 3: Sales Tracker
                    PDFSource("Sales Tracker", self.config.sales_tracker_file, 'xlsx'),
                    
                    # Attachment 4: Compliance Certificate (all sheets in order)
                    PDFSource("Compliance SFA CC", compliance_output, 'xlsx', "SFA CC"),
                    PDFSource(
                        f"Q{self.config.quarter} Management Accounts",
                        compliance_output,
                        'xlsx',
                        f"Q{self.config.quarter} Management Accounts"
                    ),
                    PDFSource("Compliance Suppl Calc", compliance_output, 'xlsx', "Suppl. Calc"),
                    PDFSource("Compliance Impact Sales", compliance_output, 'xlsx', "Impact Unit Sales"),
                ]
                assembler = PDFAssembler(str(pdf_output))
                assembler.assemble(pdf_sources)
                results['steps']['pdf_assembly'] = {
                    'status': 'success',
                    'output': str(pdf_output)
                }
                results['output_files'].append(str(pdf_output))
            except FileNotFoundError as e:
                # LibreOffice not installed - skip PDF assembly
                logger.warning(f"PDF assembly skipped: {e}")
                logger.warning("Install LibreOffice to enable PDF generation: brew install --cask libreoffice")
                results['steps']['pdf_assembly'] = {
                    'status': 'skipped',
                    'reason': 'LibreOffice not installed'
                }
                results['warnings'].append("PDF assembly skipped - LibreOffice not installed")
            
            results['completed_at'] = datetime.now().isoformat()
            results['status'] = 'success'
            
        except Exception as e:
            logger.exception(f"Workflow failed: {e}")
            results['errors'].append(str(e))
            results['status'] = 'failed'
            results['completed_at'] = datetime.now().isoformat()
        
        return results

    def run_accounts_only(self, dry_run: bool = False) -> dict:
        """
        Execute Phase 1: Generate Management Accounts and Compliance Certificate only.
        
        This is the first phase of the split workflow. Outputs are:
        - Management Accounts Excel
        - Compliance Certificate Excel (Draft)
        
        Args:
            dry_run: If True, validate inputs but don't generate outputs
            
        Returns:
            Dict with workflow results and any warnings
        """
        logger.info(f"[Phase 1] Starting accounts generation for {self.config.quarter_str}")
        
        results = {
            'phase': 1,
            'quarter': self.config.quarter_str,
            'started_at': datetime.now().isoformat(),
            'steps': {},
            'warnings': [],
            'errors': [],
            'output_files': []
        }
        
        try:
            # Step 1: Parse BDO data
            logger.info("Step 1: Parsing BDO quarterly financials")
            bdo_result = self.bdo_parser.parse(str(self.config.bdo_file))
            results['steps']['bdo_parse'] = {
                'status': 'success',
                'accounts_parsed': len(bdo_result.accounts),
                'warnings': bdo_result.warnings
            }
            if bdo_result.warnings:
                results['warnings'].extend(bdo_result.warnings)
            
            # Step 2: Parse rent roll
            logger.info("Step 2: Parsing rent roll")
            rent_roll = self.rent_roll_parser.parse(str(self.config.rent_roll_file))
            results['steps']['rent_roll_parse'] = {
                'status': 'success',
                'units_count': rent_roll.total_units,
                'annual_rent': rent_roll.total_annual_rent
            }
            if rent_roll.warnings:
                results['warnings'].extend(rent_roll.warnings)
            
            # Step 3: Parse sales tracker
            logger.info("Step 3: Parsing sales tracker")
            sales_data = self.sales_parser.parse(
                str(self.config.sales_tracker_file),
                period_end=self.config.period_end
            )
            results['steps']['sales_parse'] = {
                'status': 'success',
                'total_sales': sales_data.total_units_sold,
                'quarter_sales': sales_data.quarter_units_sold,
                'quarter_proceeds': sales_data.quarter_proceeds
            }
            if sales_data.warnings:
                results['warnings'].extend(sales_data.warnings)
            
            if dry_run:
                logger.info("Dry run - stopping before file generation")
                results['dry_run'] = True
                results['completed_at'] = datetime.now().isoformat()
                results['status'] = 'success'
                return results
            
            # Step 4: Build Management Accounts
            logger.info("Step 4: Building Management Accounts")
            ma_output = self.config.output_dir / f"Management Accounts {self.config.quarter_str} - Draft 1.xlsx"
            ma_config = ManagementAccountsConfig(
                quarter=self.config.quarter_str,
                period_end=self.config.period_end,
                bdo_sheet_name=f"BDO - Q{self.config.quarter}-{str(self.config.year)[-2:]}",
                summary_sheet_name=f"Management Cijfers - {self.config.quarter_str}",
                cash_proceeds_sale=sales_data.quarter_proceeds
            )
            ma_builder = ManagementAccountsBuilder(
                str(self.config.previous_management_accounts),
                str(ma_output),
                ma_config,
                bdo_source_path=str(self.config.bdo_file)
            )
            ma_builder.build(bdo_result)
            results['steps']['management_accounts'] = {
                'status': 'success',
                'output': str(ma_output)
            }
            results['output_files'].append(str(ma_output))
            
            # Step 4b: Build Compliance Certificate
            logger.info("Step 4b: Building Compliance Certificate")
            compliance_output = self.config.output_dir / f"Compliance Certificate Berekening QSP - {self.config.quarter_str}_updated.xlsx"
            compliance_config = ComplianceConfig(
                year=self.config.year,
                quarter=self.config.quarter
            )
            compliance_builder = ComplianceBuilder(
                str(self.config.previous_compliance_calc),
                str(compliance_output),
                compliance_config
            )
            compliance_builder.build(bdo_result, str(ma_output))
            results['steps']['compliance_certificate'] = {
                'status': 'success',
                'output': str(compliance_output)
            }
            results['output_files'].append(str(compliance_output))
            
            results['completed_at'] = datetime.now().isoformat()
            results['status'] = 'success'
            
        except Exception as e:
            logger.exception(f"Phase 1 workflow failed: {e}")
            results['errors'].append(str(e))
            results['status'] = 'failed'
            results['completed_at'] = datetime.now().isoformat()
        
        return results


class WordReportOrchestrator:
    """
    Orchestrates Phase 2: Word report generation.
    
    Uses Management Accounts and Compliance Certificate from Phase 1
    (potentially edited by client) to generate the Word report.
    """
    
    def __init__(self, config: WordReportConfig):
        self.config = config
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialize parsers
        self.rent_roll_parser = RentRollParser()
        self.sales_parser = SalesTrackerParser()
    
    def run(self) -> dict:
        """
        Execute Phase 2: Generate Word report.
        
        Returns:
            Dict with workflow results
        """
        logger.info(f"[Phase 2] Starting Word report generation for {self.config.quarter_str}")
        
        results = {
            'phase': 2,
            'quarter': self.config.quarter_str,
            'started_at': datetime.now().isoformat(),
            'steps': {},
            'warnings': [],
            'errors': [],
            'output_files': []
        }
        
        try:
            # Parse rent roll for report data
            logger.info("Parsing rent roll")
            rent_roll = self.rent_roll_parser.parse(str(self.config.rent_roll_file))
            if rent_roll.warnings:
                results['warnings'].extend(rent_roll.warnings)
            
            # Parse sales tracker for report data
            logger.info("Parsing sales tracker")
            sales_data = self.sales_parser.parse(
                str(self.config.sales_tracker_file),
                period_end=self.config.period_end
            )
            if sales_data.warnings:
                results['warnings'].extend(sales_data.warnings)
            
            # Generate Word report
            logger.info("Generating Word report")
            word_output = self.config.output_dir / f"Quarterly QSP - {self.config.quarter_str} - Draft.docx"
            
            # Extract values from the Management Accounts Excel file
            from .generators.word_updater import extract_values_from_management_accounts, WordTemplateUpdater, ReportValues
            
            ma_values = extract_values_from_management_accounts(
                str(self.config.management_accounts_file),
                self.config.quarter,
                self.config.year
            )
            
            # Use extracted values, with fallbacks
            gtri_amount = ma_values.get('gtri', 0.0)
            gtri_ltm = ma_values.get('gtri_ltm', 0.0)
            vacancy_amount = ma_values.get('vacancy_amount', 0.0)
            gross_rental_income = ma_values.get('gross_rental', 0.0)
            gross_rental_ltm = ma_values.get('gross_rental_ltm', 0.0)
            vacancy_pct = ma_values.get('vacancy_pct', 0.0)
            maintenance_amount = ma_values.get('maintenance', 0.0)
            maintenance_ltm = ma_values.get('maintenance_ltm', 0.0)
            unit_sales_proceeds = ma_values.get('unit_sales_proceeds', 0.0)
            
            if unit_sales_proceeds == 0 and sales_data.quarter_proceeds:
                unit_sales_proceeds = sales_data.quarter_proceeds / 1000
            
            # Rent roll total (annual, in thousands)
            rent_roll_k = rent_roll.total_annual_rent / 1000
            
            logger.info(f"Word report KPIs extracted from Management Accounts:")
            logger.info(f"  GTRI: €{gtri_amount:,.1f}k")
            logger.info(f"  Gross rental income: €{gross_rental_income:,.1f}k")
            logger.info(f"  Financial vacancy: {vacancy_pct:.1f}%")
            logger.info(f"  Rent roll (annual): €{rent_roll_k:,.1f}k")
            
            report_values = ReportValues(
                report_date=datetime.now().strftime("%d %B %Y"),
                gtri=gtri_amount,
                gtri_ltm=rent_roll_k,
                gross_rental_income=gross_rental_income,
                gross_rental_income_ltm=gross_rental_ltm,
                financial_vacancy_pct=vacancy_pct,
                financial_vacancy_amount=vacancy_amount,
                rent_roll_annual=rent_roll_k,
                rent_roll_units=rent_roll.total_units,
                units_sold_quarter=sales_data.quarter_units_sold,
                unit_sales_proceeds=unit_sales_proceeds,
                maintenance_amount=maintenance_amount,
                maintenance_ltm=maintenance_ltm,
                capex_amount=0,
                capex_ltm=0,
                unit_sales_narrative=f"{sales_data.quarter_units_sold} unit(s) sold for €{unit_sales_proceeds:,.0f}k" if sales_data.quarter_units_sold > 0 else "No unit sales this quarter.",
                maintenance_detail="",
                sustainability_detail="",
                prev_gtri=3200.6,
                prev_gtri_ltm=12940.2,
                prev_gross_rental_income=3067.5,
                prev_vacancy_pct=4.1,
                prev_vacancy_amount=106.3,
                prev_rent_roll=0,
                prev_maintenance=297.0,
                prev_unit_sales_proceeds=762.5,
            )
            
            word_updater = WordTemplateUpdater(str(self.config.word_template), str(word_output))
            word_updater.update_with_python_docx(
                report_values,
                self.config.previous_quarter_str,
                self.config.quarter_str
            )
            
            results['steps']['word_update'] = {
                'status': 'success',
                'output': str(word_output)
            }
            results['output_files'].append(str(word_output))
            
            results['completed_at'] = datetime.now().isoformat()
            results['status'] = 'success'
            
        except Exception as e:
            logger.exception(f"Phase 2 workflow failed: {e}")
            results['errors'].append(str(e))
            results['status'] = 'failed'
            results['completed_at'] = datetime.now().isoformat()
        
        return results


class FinalPDFOrchestrator:
    """
    Orchestrates Phase 3: Final PDF assembly.
    
    Takes finalized Word report and Compliance Certificate
    (client-approved) and assembles the final PDF.
    """
    
    def __init__(self, config: FinalPDFConfig):
        self.config = config
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
    
    def run(self) -> dict:
        """
        Execute Phase 3: Generate final PDF.
        
        Returns:
            Dict with workflow results
        """
        logger.info(f"[Phase 3] Starting final PDF assembly for {self.config.quarter_str}")
        
        results = {
            'phase': 3,
            'quarter': self.config.quarter_str,
            'started_at': datetime.now().isoformat(),
            'steps': {},
            'warnings': [],
            'errors': [],
            'output_files': []
        }
        
        try:
            # Assemble PDF
            logger.info("Assembling final PDF")
            pdf_output = self.config.output_dir / f"Quarterly QSP - {self.config.quarter_str} - ASX.pdf"
            
            # Determine summary sheet name
            summary_sheet_name = f"Management Cijfers - {self.config.quarter_str}"
            
            pdf_sources = [
                # Section 1: Main Report
                PDFSource("Main Report", self.config.word_report_file, 'docx'),
                
                # Attachment 1: Rent Roll
                PDFSource("Rent Roll", self.config.rent_roll_file, 'xlsx'),
                
                # Attachment 2: Management Accounts
                PDFSource(
                    "Management Accounts",
                    self.config.management_accounts_file,
                    'xlsx',
                    summary_sheet_name
                ),
                
                # Attachment 3: Sales Tracker
                PDFSource("Sales Tracker", self.config.sales_tracker_file, 'xlsx'),
                
                # Attachment 4: Compliance Certificate (all sheets in order)
                PDFSource("Compliance SFA CC", self.config.compliance_certificate_file, 'xlsx', "SFA CC"),
                PDFSource(
                    f"Q{self.config.quarter} Management Accounts",
                    self.config.compliance_certificate_file,
                    'xlsx',
                    f"Q{self.config.quarter} Management Accounts"
                ),
                PDFSource("Compliance Suppl Calc", self.config.compliance_certificate_file, 'xlsx', "Suppl. Calc"),
                PDFSource("Compliance Impact Sales", self.config.compliance_certificate_file, 'xlsx', "Impact Unit Sales"),
            ]
            
            assembler = PDFAssembler(str(pdf_output))
            assembler.assemble(pdf_sources)
            
            results['steps']['pdf_assembly'] = {
                'status': 'success',
                'output': str(pdf_output)
            }
            results['output_files'].append(str(pdf_output))
            
            results['completed_at'] = datetime.now().isoformat()
            results['status'] = 'success'
            
        except FileNotFoundError as e:
            logger.warning(f"PDF assembly failed: {e}")
            logger.warning("Install LibreOffice to enable PDF generation: brew install --cask libreoffice")
            results['steps']['pdf_assembly'] = {
                'status': 'failed',
                'reason': str(e)
            }
            results['errors'].append(f"PDF assembly failed - {str(e)}")
            results['status'] = 'failed'
            results['completed_at'] = datetime.now().isoformat()
            
        except Exception as e:
            logger.exception(f"Phase 3 workflow failed: {e}")
            results['errors'].append(str(e))
            results['status'] = 'failed'
            results['completed_at'] = datetime.now().isoformat()
        
        return results


def run_quarterly_report(config_path: str = "config/config.yaml") -> dict:
    """
    Convenience function to run quarterly report from config file.
    
    Args:
        config_path: Path to config YAML file
        
    Returns:
        Dict with workflow results
    """
    with open(config_path, 'r', encoding='utf-8') as f:
        config_data = yaml.safe_load(f)
    
    # Build config from YAML
    current = config_data['current_quarter']
    paths = config_data['paths']
    file_patterns = config_data.get('file_patterns', {})
    
    # Find input files using patterns
    input_dir = Path(paths['input_folder'])
    
    # Find BDO file
    bdo_files = list(input_dir.glob(file_patterns.get('bdo_financials', 'Cijfers*QSP*.xlsx')))
    if not bdo_files:
        raise FileNotFoundError(f"BDO file not found in {input_dir}")
    bdo_file = bdo_files[0]
    
    # Find Management Accounts
    ma_files = list(input_dir.glob(file_patterns.get('management_accounts', 'Management*Accounts*.xlsx')))
    if not ma_files:
        raise FileNotFoundError(f"Management Accounts file not found in {input_dir}")
    previous_ma = ma_files[0]
    
    # Find rent roll
    rent_roll_files = list(input_dir.glob(file_patterns.get('rent_roll', 'QSP*huurlijst*.xlsx')))
    if not rent_roll_files:
        raise FileNotFoundError(f"Rent roll file not found in {input_dir}")
    rent_roll_file = rent_roll_files[0]
    
    # Find sales tracker
    sales_files = list(input_dir.glob(file_patterns.get('sales_tracker', 'Unit*Sales*tracker*.xlsx')))
    if not sales_files:
        raise FileNotFoundError(f"Sales tracker file not found in {input_dir}")
    sales_tracker_file = sales_files[0]
    
    # Find compliance calc (optional)
    compliance_files = list(input_dir.glob(file_patterns.get('compliance_calc', 'Compliance*Certificate*.xlsx')))
    previous_compliance = compliance_files[0] if compliance_files else Path(paths['input_folder']) / "compliance.xlsx"
    
    # Find Word template
    word_files = list(input_dir.glob(file_patterns.get('word_template', 'Quarterly*QSP*.docx')))
    if not word_files:
        raise FileNotFoundError(f"Word template not found in {input_dir}")
    word_template = word_files[0]
    
    config = QuarterlyReportConfig(
        year=current['year'],
        quarter=current['quarter'],
        bdo_file=bdo_file,
        previous_management_accounts=previous_ma,
        rent_roll_file=rent_roll_file,
        sales_tracker_file=sales_tracker_file,
        previous_compliance_calc=previous_compliance,
        word_template=word_template,
        output_dir=Path(paths['output_folder'])
    )
    
    orchestrator = QuarterlyReportOrchestrator(config)
    return orchestrator.run()

