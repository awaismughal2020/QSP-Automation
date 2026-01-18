"""
Word Template Updater

Updates the quarterly report Word template with:
- Global find/replace for quarter references
- Specific value updates on designated pages
- Date updates
- PAGE 4: Numeric KPI population (GTRI, gross rental income, vacancy, maintenance, etc.)
- PAGE 5-6: Actions and CAPEX values

NUMERIC POPULATION (Page 4):
The Word document contains financial KPIs embedded in narrative text like:
"GTRI of the portfolio amounted to €3,200k"
"Financial vacancy was 5.4%"
etc.

These values must be extracted from the generated Excel outputs and injected
into the Word document, replacing the previous quarter's values.

VALUE SOURCES (from Management Cijfers sheet):
- Column AA: Current quarter values (Q3 2025 = column 27)
- Column AB: LTM values (Q3 2025 LTM = column 28)
- Dynamic column calculation: base_column + (quarter_total - base_quarter_total)
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path
import subprocess
import re
import shutil
import tempfile
from loguru import logger


@dataclass
class ReportValues:
    """Values to insert into the quarterly report."""
    # Page 1
    report_date: str  # e.g., "14 October 2025"
    
    # Page 4 - Executive Summary (from Management Accounts / BDO)
    gtri: float  # Gross Theoretical Rental Income (€k) - quarterly
    gtri_ltm: float = 0.0  # GTRI LTM (€k) - for annual references / rent roll yields
    gross_rental_income: float = 0.0  # Actual rental income (€k) - quarterly
    gross_rental_income_ltm: float = 0.0  # Gross rental income LTM (€k)
    financial_vacancy_pct: float = 0.0  # Vacancy percentage
    financial_vacancy_amount: float = 0.0  # Vacancy in €k
    
    # Page 4 - From Rent Roll
    rent_roll_annual: float = 0.0  # Annual rent roll total (€k)
    rent_roll_units: int = 0  # Number of units in rent roll
    
    # Page 4 - From Sales Tracker
    units_sold_quarter: int = 0  # Units sold this quarter
    unit_sales_proceeds: float = 0.0  # Proceeds from sales (€k)
    
    # Page 4/6 - Maintenance and CAPEX
    maintenance_amount: float = 0.0  # Maintenance spend (€k) - quarterly
    maintenance_ltm: float = 0.0  # Maintenance LTM (€k)
    capex_amount: float = 0.0  # CAPEX spend (€k) - quarterly
    capex_ltm: float = 0.0  # CAPEX LTM (€k)
    
    # Page 5 - Unit Sales narrative
    unit_sales_narrative: str = ""  # Free text about unit sales actions
    
    # Page 6
    maintenance_detail: str = ""  # Maintenance description
    sustainability_detail: str = ""  # Sustainability/CAPEX description
    
    # Previous quarter values (for finding/replacing)
    # These are extracted from the template document (Q2 values)
    prev_gtri: float = 0.0
    prev_gtri_ltm: float = 0.0  # Previous LTM GTRI (for rent roll yields)
    prev_gross_rental_income: float = 0.0
    prev_vacancy_pct: float = 0.0
    prev_vacancy_amount: float = 0.0
    prev_rent_roll: float = 0.0
    prev_maintenance: float = 0.0
    prev_unit_sales_proceeds: float = 0.0


def extract_values_from_management_accounts(ma_path: str, quarter: int, year: int) -> dict:
    """
    Extract financial values from Management Accounts Excel file.
    
    The Management Cijfers sheet contains:
    - Column AA (index 27) for Q3 2025 quarterly values
    - Column AB (index 28) for Q3 2025 LTM values
    - Each subsequent quarter advances by 1 column
    
    Args:
        ma_path: Path to Management Accounts Excel file
        quarter: Quarter number (1-4)
        year: Year (e.g., 2025)
        
    Returns:
        Dictionary with extracted values in €k
    """
    import openpyxl
    
    try:
        # Load with data_only=True to get calculated values
        wb = openpyxl.load_workbook(ma_path, data_only=True)
        
        # Find Management Cijfers sheet
        cijfers_sheet = None
        for name in wb.sheetnames:
            if 'Management Cijfers' in name:
                cijfers_sheet = wb[name]
                break
        
        if cijfers_sheet is None:
            logger.warning("Management Cijfers sheet not found")
            wb.close()
            return {}
        
        # Calculate column indices dynamically
        # Base: Q3 2025 = Column AA (index 27)
        # Formula: column = 27 + ((year * 4 + quarter) - (2025 * 4 + 3))
        base_quarter_total = 2025 * 4 + 3  # Q3 2025 = 8103
        target_quarter_total = year * 4 + quarter
        quarterly_column = 27 + (target_quarter_total - base_quarter_total)
        ltm_column = quarterly_column + 1
        
        logger.info(f"Extracting values from columns {quarterly_column} (quarterly) and {ltm_column} (LTM)")
        
        # Find BDO sheet for direct value extraction
        bdo_sheet = None
        expected_bdo_name = f"BDO - Q{quarter}-{str(year)[-2:]}"
        for name in wb.sheetnames:
            if expected_bdo_name in name:
                bdo_sheet = wb[name]
                break
        
        values = {}
        
        # Extract from BDO sheet (more reliable for calculated values)
        if bdo_sheet:
            # Row 76: GTRI (stored as negative, take absolute)
            gtri_q = bdo_sheet.cell(row=76, column=7).value  # Column G
            gtri_ltm = bdo_sheet.cell(row=76, column=8).value  # Column H
            values['gtri'] = abs(gtri_q) / 1000 if gtri_q else 0
            values['gtri_ltm'] = abs(gtri_ltm) / 1000 if gtri_ltm else 0
            
            # Row 77: Financial vacancy (stored as positive)
            vac_q = bdo_sheet.cell(row=77, column=7).value
            vac_ltm = bdo_sheet.cell(row=77, column=8).value
            values['vacancy_amount'] = abs(vac_q) / 1000 if vac_q else 0
            values['vacancy_amount_ltm'] = abs(vac_ltm) / 1000 if vac_ltm else 0
            
            # Row 88: Maintenance (stored as positive)
            maint_q = bdo_sheet.cell(row=88, column=7).value
            maint_ltm = bdo_sheet.cell(row=88, column=8).value
            values['maintenance'] = abs(maint_q) / 1000 if maint_q else 0
            values['maintenance_ltm'] = abs(maint_ltm) / 1000 if maint_ltm else 0
            
            logger.info(f"Extracted from BDO: GTRI={values['gtri']:.1f}k, Vacancy={values['vacancy_amount']:.1f}k")
        
        # Calculate derived values
        values['gross_rental'] = values.get('gtri', 0) - values.get('vacancy_amount', 0)
        values['gross_rental_ltm'] = values.get('gtri_ltm', 0) - values.get('vacancy_amount_ltm', 0)
        
        # Vacancy percentage
        if values.get('gtri', 0) > 0:
            values['vacancy_pct'] = (values.get('vacancy_amount', 0) / values['gtri']) * 100
        else:
            values['vacancy_pct'] = 0
        
        # Extract cash proceeds from Management Cijfers row 50
        cash_proceeds = cijfers_sheet.cell(row=50, column=quarterly_column).value
        values['unit_sales_proceeds'] = abs(cash_proceeds) / 1000 if cash_proceeds else 0
        
        wb.close()
        return values
        
    except Exception as e:
        logger.error(f"Error extracting values from Management Accounts: {e}")
        return {}


@dataclass 
class NumericValueMapping:
    """Maps a KPI to its context pattern and value formatter."""
    kpi_name: str
    # Regex pattern to find the value in context (must have a capture group for the number)
    context_pattern: str
    # How to format the new value
    formatter: str  # 'euro_k', 'euro_m', 'percent', 'integer'
    # Which ReportValues attribute to use
    value_attr: str
    # Optional: previous value attribute for better matching
    prev_value_attr: Optional[str] = None


class WordTemplateUpdater:
    """
    Updates Word template for quarterly report.
    
    Uses pandoc for conversion to/from markdown for text operations,
    and python-docx for structural operations.
    """
    
    # Placeholder patterns in the template
    PLACEHOLDERS = {
        '{{GTRI}}': 'gtri',
        '{{GROSS_RENTAL_INCOME}}': 'gross_rental_income',
        '{{RENT_ROLL}}': 'rent_roll_annual',
        '{{VACANCY}}': 'financial_vacancy_pct',
        '{{UNITS_SOLD}}': 'units_sold_quarter',
        '{{MAINTENANCE}}': 'maintenance_amount',
        '{{CAPEX}}': 'capex_amount',
    }
    
    def __init__(self, template_path: str, output_path: str):
        self.template_path = Path(template_path)
        self.output_path = Path(output_path)
        # Use system temp directory
        self.working_dir = Path(tempfile.gettempdir()) / "word_update"
        self.working_dir.mkdir(parents=True, exist_ok=True)
        
    def update(self, 
               values: ReportValues, 
               previous_quarter: str,
               current_quarter: str) -> Path:
        """
        Update template with new values using pandoc.
        Simpler method but may lose some formatting.
        
        Args:
            values: ReportValues with all data to insert
            previous_quarter: Previous quarter string (e.g., "Q2 2025")
            current_quarter: Current quarter string (e.g., "Q3 2025")
            
        Returns:
            Path to updated Word document
        """
        logger.info(f"Updating Word template for {current_quarter} (pandoc method)")
        
        # Step 1: Convert to markdown for text operations
        md_path = self.working_dir / "template.md"
        self._docx_to_markdown(self.template_path, md_path)
        
        # Step 2: Read markdown content
        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Step 3: Global quarter replacement
        content = self._replace_quarter(content, previous_quarter, current_quarter)
        
        # Step 4: Replace placeholders with values
        content = self._replace_placeholders(content, values)
        
        # Step 5: Write updated markdown
        updated_md = self.working_dir / "updated.md"
        with open(updated_md, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # Step 6: Convert back to docx (preserving reference doc styling)
        self._markdown_to_docx(updated_md, self.output_path, self.template_path)
        
        logger.info(f"Generated updated report: {self.output_path}")
        return self.output_path
    
    def _docx_to_markdown(self, docx_path: Path, md_path: Path):
        """Convert Word document to markdown using pandoc."""
        cmd = [
            'pandoc',
            str(docx_path),
            '-o', str(md_path),
            '--wrap=none'  # Preserve line structure
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"Pandoc conversion failed: {result.stderr}")
    
    def _markdown_to_docx(self, md_path: Path, docx_path: Path, reference_doc: Path):
        """Convert markdown back to Word using pandoc with reference doc."""
        cmd = [
            'pandoc',
            str(md_path),
            '-o', str(docx_path),
            f'--reference-doc={reference_doc}'  # Use original as style reference
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"Pandoc conversion failed: {result.stderr}")
    
    def _replace_quarter(self, content: str, old_quarter: str, new_quarter: str) -> str:
        """
        Replace ALL quarter references comprehensively.
        
        This addresses Issue 7.1 from feedback: mixed quarter references.
        Handles all possible formats and variations.
        """
        import re
        
        # Extract quarter number and year
        old_parts = old_quarter.split()  # e.g., ["Q2", "2025"]
        new_parts = new_quarter.split()
        
        if len(old_parts) != 2 or len(new_parts) != 2:
            logger.warning(f"Unexpected quarter format: {old_quarter} -> {new_quarter}")
            return content
        
        old_q_num = old_parts[0][1]  # "2" from "Q2"
        old_year = old_parts[1]  # "2025"
        new_q_num = new_parts[0][1]  # "3" from "Q3"
        new_year = new_parts[1]  # "2025"
        
        # Handle ordinal words (second quarter → third quarter, etc.)
        ordinal_map = {
            '1': ('first', 'eerste'),
            '2': ('second', 'tweede'),
            '3': ('third', 'derde'),
            '4': ('fourth', 'vierde'),
        }
        if old_q_num in ordinal_map and new_q_num in ordinal_map:
            old_ordinal_en, old_ordinal_nl = ordinal_map[old_q_num]
            new_ordinal_en, new_ordinal_nl = ordinal_map[new_q_num]
            
            # English ordinal patterns - case variations
            content = re.sub(rf'\b{old_ordinal_en}\s+quarter\b', f'{new_ordinal_en} quarter', content, flags=re.IGNORECASE)
            content = re.sub(rf'\bthe\s+{old_ordinal_en}\b', f'the {new_ordinal_en}', content, flags=re.IGNORECASE)
            
            # Dutch ordinal patterns
            content = re.sub(rf'\b{old_ordinal_nl}\s+kwartaal\b', f'{new_ordinal_nl} kwartaal', content, flags=re.IGNORECASE)
        
        # Comprehensive replacement patterns
        replacements = [
            # Standard formats
            (old_quarter, new_quarter),  # Q2 2025 -> Q3 2025
            (old_quarter.lower(), new_quarter.lower()),  # q2 2025 -> q3 2025
            (old_quarter.replace(' ', '-'), new_quarter.replace(' ', '-')),  # Q2-2025
            (old_quarter.replace(' ', '/'), new_quarter.replace(' ', '/')),  # Q2/2025
            (f"Q{old_q_num} {old_year[-2:]}", f"Q{new_q_num} {new_year[-2:]}"),  # Q2 25
            (f"Q{old_q_num}-{old_year[-2:]}", f"Q{new_q_num}-{new_year[-2:]}"),  # Q2-25
            (f"{old_year[-2:]}Q{old_q_num}", f"{new_year[-2:]}Q{new_q_num}"),  # 25Q2
            
            # Year-quarter formats with different separators
            (f"{old_year} – Q{old_q_num}", f"{new_year} – Q{new_q_num}"),  # 2025 – Q2
            (f"{old_year} - Q{old_q_num}", f"{new_year} - Q{new_q_num}"),  # 2025 - Q2
            
            # Title and header formats (as seen in Sample.pdf)
            (f"Quarterly Report {old_quarter}", f"Quarterly Report {new_quarter}"),
            (f"Quarterly QSP - {old_quarter}", f"Quarterly QSP - {new_quarter}"),
            (f"in {old_quarter}", f"in {new_quarter}"),
            
            # Text variations
            (f"Actions {old_quarter}", f"Actions {new_quarter}"),
            (f"Portfolio highlights – {old_quarter}", f"Portfolio highlights – {new_quarter}"),
            (f"Portfolio highlights - {old_quarter}", f"Portfolio highlights - {new_quarter}"),
            (f"Report {old_quarter}", f"Report {new_quarter}"),
            (f"Quarter {old_q_num} {old_year}", f"Quarter {new_q_num} {new_year}"),
            
            # Compliance certificate references
            (f"Interest Period: {self._get_period_date(old_q_num, old_year)}", 
             f"Interest Period: {self._get_period_date(new_q_num, new_year)}"),
            
            # Dutch variations
            (f"Kwartaal {old_q_num} {old_year}", f"Kwartaal {new_q_num} {new_year}"),
        ]
        
        for old_pattern, new_pattern in replacements:
            content = content.replace(old_pattern, new_pattern)
            # Also try case-insensitive for text variations
            if not old_pattern[0].isdigit():
                content = re.sub(re.escape(old_pattern), new_pattern, content, flags=re.IGNORECASE)
        
        # Handle month references
        quarter_to_month = {
            'Q1': ('March', '31', 'maart'),
            'Q2': ('June', '30', 'juni'),
            'Q3': ('September', '30', 'september'),
            'Q4': ('December', '31', 'december'),
        }
        
        old_q = old_parts[0]
        new_q = new_parts[0]
        
        if old_q in quarter_to_month and new_q in quarter_to_month:
            old_month, old_day, old_month_nl = quarter_to_month[old_q]
            new_month, new_day, new_month_nl = quarter_to_month[new_q]
            
            # Replace English months
            content = content.replace(old_month, new_month)
            content = content.replace(old_month.lower(), new_month.lower())
            
            # Replace Dutch months
            content = content.replace(old_month_nl, new_month_nl)
            content = content.replace(old_month_nl.capitalize(), new_month_nl.capitalize())
            
            # Replace date formats like "30 June 2025" and "30-6-2025"
            for year in [old_year, str(int(old_year) - 1), str(int(old_year) + 1)]:
                # English format: 30 June 2025
                old_date = f"{old_day} {old_month} {year}"
                new_date = f"{new_day} {new_month} {new_year}"
                content = content.replace(old_date, new_date)
                
                # Dutch format: 30 juni 2025
                old_date_nl = f"{old_day} {old_month_nl} {year}"
                new_date_nl = f"{new_day} {new_month_nl} {new_year}"
                content = content.replace(old_date_nl, new_date_nl)
                
                # Numeric formats: 30-6-2025, 30/6/2025
                old_month_num = {'March': 3, 'June': 6, 'September': 9, 'December': 12}[old_month]
                new_month_num = {'March': 3, 'June': 6, 'September': 9, 'December': 12}[new_month]
                
                content = content.replace(f"{old_day}-{old_month_num}-{year}", f"{new_day}-{new_month_num}-{new_year}")
                content = content.replace(f"{old_day}/{old_month_num}/{year}", f"{new_day}/{new_month_num}/{new_year}")
                
                # Rent roll date format: 1-10-2025 (1st of next month after quarter)
                next_month_old = old_month_num + 1 if old_month_num < 12 else 1
                next_year_old = year if old_month_num < 12 else str(int(year) + 1)
                next_month_new = new_month_num + 1 if new_month_num < 12 else 1
                next_year_new = new_year if new_month_num < 12 else str(int(new_year) + 1)
                
                content = content.replace(f"1-{next_month_old}-{next_year_old}", f"1-{next_month_new}-{next_year_new}")
        
        return content
    
    def _get_period_date(self, q_num: str, year: str) -> str:
        """Get period end date string for a quarter."""
        month_days = {'1': ('3', '31'), '2': ('6', '30'), '3': ('9', '30'), '4': ('12', '31')}
        if q_num in month_days:
            month, day = month_days[q_num]
            return f"{day}-{month}-{year}"
        return ""
    
    def _replace_placeholders(self, content: str, values: ReportValues) -> str:
        """Replace placeholder values in content."""
        replacements = {
            '{{GTRI}}': f"€{values.gtri:,.1f}k",
            '{{GROSS_RENTAL_INCOME}}': f"€{values.gross_rental_income:,.1f}k",
            '{{RENT_ROLL}}': f"€{values.rent_roll_annual:,.1f}k",
            '{{VACANCY}}': f"{values.financial_vacancy_pct:.1f}%",
            '{{UNITS_SOLD}}': str(values.units_sold_quarter),
            '{{MAINTENANCE}}': f"€{values.maintenance_amount:,.0f}k",
            '{{CAPEX}}': f"€{values.capex_amount:,.0f}k",
            '{{REPORT_DATE}}': values.report_date,
        }
        
        for placeholder, value in replacements.items():
            content = content.replace(placeholder, value)
        
        # Handle narrative sections
        if values.unit_sales_narrative:
            content = content.replace('{{UNIT_SALES_NARRATIVE}}', values.unit_sales_narrative)
        
        if values.maintenance_detail:
            content = content.replace('{{MAINTENANCE_DETAIL}}', values.maintenance_detail)
        
        if values.sustainability_detail:
            content = content.replace('{{SUSTAINABILITY_DETAIL}}', values.sustainability_detail)
        
        return content
    
    def update_with_python_docx(self, values: ReportValues, 
                                 previous_quarter: str, 
                                 current_quarter: str) -> Path:
        """
        Update Word document including text boxes using direct XML manipulation.
        
        This method handles:
        - Regular paragraphs (via python-docx)
        - Text boxes and shapes (via direct XML manipulation)
        - Headers and footers
        
        Args:
            values: ReportValues with all data to insert
            previous_quarter: Previous quarter string (e.g., "Q2 2025")
            current_quarter: Current quarter string (e.g., "Q3 2025")
            
        Returns:
            Path to updated Word document
        """
        from docx import Document
        import zipfile
        import os
        
        logger.info(f"Updating Word template for {current_quarter} (python-docx + XML method)")
        
        # Copy template to output location
        shutil.copy2(self.template_path, self.output_path)
        
        # Step 1: Update regular paragraphs/tables via python-docx (basic text replacements)
        doc = Document(self.output_path)
        
        old_date_patterns = self._get_old_date_patterns(previous_quarter)
        
        # Process paragraphs
        for para in doc.paragraphs:
            self._process_paragraph(para, values, previous_quarter, current_quarter)
            self._update_cover_page_date(para, old_date_patterns, values.report_date)
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._process_paragraph(para, values, previous_quarter, current_quarter)
                        self._update_cover_page_date(para, old_date_patterns, values.report_date)
        
        # Process headers/footers
        for section in doc.sections:
            for para in section.header.paragraphs:
                self._process_paragraph(para, values, previous_quarter, current_quarter)
            for para in section.footer.paragraphs:
                self._process_paragraph(para, values, previous_quarter, current_quarter)
        
        doc.save(self.output_path)
        
        # Step 2: After python-docx saves, apply direct XML manipulation for:
        # - Text boxes (which python-docx can't access)
        # - Fragmented text replacements
        # - Context-aware numeric replacements (€ values split across multiple XML elements)
        # This must happen AFTER python-docx saves, otherwise doc.save() overwrites our changes
        text_replacements = self._build_replacement_map(values, previous_quarter, current_quarter)
        self._update_text_boxes_xml(text_replacements, values)
        
        logger.info(f"Generated updated report: {self.output_path}")
        return self.output_path
    
    def _build_replacement_map(self, values: ReportValues, old_quarter: str, new_quarter: str) -> List[Tuple[str, str]]:
        """
        Build comprehensive replacement map for all text that needs updating.
        
        Returns list of (old_text, new_text) tuples, sorted longest-first.
        
        INCLUDES: 
        - Quarter text replacements (Q2 -> Q3)
        - Date replacements
        - NUMERIC VALUE REPLACEMENTS (Page 4 KPIs)
        """
        replacements = []
        
        # Extract quarter components
        old_parts = old_quarter.split()
        new_parts = new_quarter.split()
        old_q_num = old_parts[0][1] if len(old_parts) == 2 else ''
        old_year = old_parts[1] if len(old_parts) == 2 else ''
        new_q_num = new_parts[0][1] if len(new_parts) == 2 else ''
        new_year = new_parts[1] if len(new_parts) == 2 else ''
        
        # === NUMERIC VALUE REPLACEMENTS (Page 4 KPIs) ===
        # These replace last quarter's numeric values with new quarter's values
        numeric_replacements = self._build_numeric_replacements(values)
        replacements.extend(numeric_replacements)
        
        # Quarter text replacements
        replacements.extend([
            (old_quarter, new_quarter),
            (old_quarter.replace(' ', '-'), new_quarter.replace(' ', '-')),
            (old_quarter.replace(' ', '/'), new_quarter.replace(' ', '/')),
            (f"Q{old_q_num} {old_year[-2:]}", f"Q{new_q_num} {new_year[-2:]}"),
            (f"Q{old_q_num}-{old_year[-2:]}", f"Q{new_q_num}-{new_year[-2:]}"),
            (f"{old_year[-2:]}Q{old_q_num}", f"{new_year[-2:]}Q{new_q_num}"),
            (f"{old_year} – Q{old_q_num}", f"{new_year} – Q{new_q_num}"),
            (f"{old_year} - Q{old_q_num}", f"{new_year} - Q{new_q_num}"),
        ])
        
        # Ordinal replacements
        ordinal_map = {
            '1': ('first', 'eerste'),
            '2': ('second', 'tweede'),
            '3': ('third', 'derde'),
            '4': ('fourth', 'vierde'),
        }
        if old_q_num in ordinal_map and new_q_num in ordinal_map:
            old_en, old_nl = ordinal_map[old_q_num]
            new_en, new_nl = ordinal_map[new_q_num]
            replacements.extend([
                (f"{old_en} quarter", f"{new_en} quarter"),
                (f"the {old_en}", f"the {new_en}"),
                (f"{old_nl} kwartaal", f"{new_nl} kwartaal"),
            ])
        
        # Date replacements
        quarter_to_month = {
            '1': ('March', 'maart', '31', '3'),
            '2': ('June', 'juni', '30', '6'),
            '3': ('September', 'september', '30', '9'),
            '4': ('December', 'december', '31', '12'),
        }
        
        if old_q_num in quarter_to_month and new_q_num in quarter_to_month:
            old_month, old_month_nl, old_day, old_month_num = quarter_to_month[old_q_num]
            new_month, new_month_nl, new_day, new_month_num = quarter_to_month[new_q_num]
            
            # Month names
            replacements.extend([
                (old_month, new_month),
                (old_month.lower(), new_month.lower()),
                (old_month_nl, new_month_nl),
            ])
            
            # Date formats
            replacements.extend([
                (f"{old_day} {old_month} {old_year}", f"{new_day} {new_month} {new_year}"),
                (f"{old_day} {old_month_nl} {old_year}", f"{new_day} {new_month_nl} {new_year}"),
                (f"{old_day}-{old_month_num}-{old_year}", f"{new_day}-{new_month_num}-{new_year}"),
                # Rent roll date (1st of next month)
                (f"1-{int(old_month_num)+1}-{old_year}", f"1-{int(new_month_num)+1}-{new_year}"),
            ])
        
        # Cover page date - get patterns and replace with current report date
        old_date_patterns = self._get_old_date_patterns(old_quarter)
        for old_date in old_date_patterns:
            replacements.append((old_date, values.report_date))
        
        # Sort by length (longest first) to prevent partial matches
        replacements.sort(key=lambda x: len(x[0]), reverse=True)
        
        return replacements
    
    def _build_numeric_replacements(self, values: ReportValues) -> List[Tuple[str, str]]:
        """
        Build replacement tuples for numeric values in the Word document.
        
        The Word document contains financial KPIs embedded in narrative text.
        We need to find patterns like "€3,200k" or "5.4%" and replace them
        with the new quarter's values.
        
        This method generates replacement pairs for common numeric formats
        that might appear in the document.
        """
        replacements = []
        
        # Helper functions for formatting
        def format_euro_k(value: float, decimals: int = 0) -> List[str]:
            """Generate multiple format variations for Euro amounts in thousands."""
            formats = []
            if value == 0:
                return formats
            
            abs_val = abs(value)
            
            # Various formatting styles that might appear in the document
            if decimals == 0:
                formats.extend([
                    f"€{abs_val:,.0f}k",
                    f"€{abs_val:,.0f} k",
                    f"€ {abs_val:,.0f}k",
                    f"€{abs_val:,.0f}",  # Without k suffix
                    f"€{int(abs_val):,}k",
                    f"€{int(abs_val):,}",
                ])
            else:
                formats.extend([
                    f"€{abs_val:,.{decimals}f}k",
                    f"€{abs_val:,.{decimals}f} k", 
                    f"€ {abs_val:,.{decimals}f}k",
                    f"€{abs_val:,.{decimals}f}",
                ])
            
            # European format with dot as thousand separator
            if abs_val >= 1000:
                euro_fmt = f"{abs_val:,.0f}".replace(',', '.')
                formats.extend([
                    f"€{euro_fmt}k",
                    f"€{euro_fmt}",
                ])
            
            return formats
        
        def format_euro_m(value: float) -> List[str]:
            """Generate format variations for Euro amounts in millions."""
            formats = []
            if value == 0:
                return formats
            
            abs_val = abs(value)
            formats.extend([
                f"€{abs_val:.1f}m",
                f"€{abs_val:.1f} m",
                f"€ {abs_val:.1f}m",
                f"€{abs_val:.2f}m",
            ])
            return formats
        
        def format_percent(value: float) -> List[str]:
            """Generate format variations for percentages."""
            formats = []
            abs_val = abs(value)
            formats.extend([
                f"{abs_val:.1f}%",
                f"{abs_val:.1f} %",
                f"{abs_val:.2f}%",
                f"{int(abs_val)}%",
            ])
            return formats
        
        # === PAGE 4 NUMERIC REPLACEMENTS ===
        # These are the key KPIs that need to be updated
        
        # 1. GTRI (Gross Theoretical Rental Income) - quarterly value
        # Format: typically "€3,200k" or "€3.2m"
        if values.gtri > 0:
            gtri_k = values.gtri  # Already in thousands
            
            # If we have previous GTRI value, create direct replacement
            if values.prev_gtri > 0:
                for old_fmt in format_euro_k(values.prev_gtri, 0):
                    for new_fmt in format_euro_k(gtri_k, 0)[:1]:  # Just use first format
                        replacements.append((old_fmt, new_fmt))
        
        # 2. Gross Rental Income - quarterly value
        if values.gross_rental_income > 0 and values.prev_gross_rental_income > 0:
            for old_fmt in format_euro_k(values.prev_gross_rental_income, 0):
                for new_fmt in format_euro_k(values.gross_rental_income, 0)[:1]:
                    replacements.append((old_fmt, new_fmt))
        
        # 3. Financial Vacancy percentage
        if values.prev_vacancy_pct > 0:
            for old_fmt in format_percent(values.prev_vacancy_pct):
                for new_fmt in format_percent(values.financial_vacancy_pct)[:1]:
                    replacements.append((old_fmt, new_fmt))
        
        # 4. Rent Roll total
        if values.prev_rent_roll > 0 and values.rent_roll_annual > 0:
            for old_fmt in format_euro_k(values.prev_rent_roll, 0):
                for new_fmt in format_euro_k(values.rent_roll_annual, 0)[:1]:
                    replacements.append((old_fmt, new_fmt))
        
        # 5. Maintenance amount
        if values.prev_maintenance > 0 and values.maintenance_amount > 0:
            for old_fmt in format_euro_k(values.prev_maintenance, 0):
                for new_fmt in format_euro_k(values.maintenance_amount, 0)[:1]:
                    replacements.append((old_fmt, new_fmt))
        
        # 6. Unit sales proceeds
        if values.prev_unit_sales_proceeds > 0 and values.unit_sales_proceeds > 0:
            for old_fmt in format_euro_k(values.prev_unit_sales_proceeds, 0):
                for new_fmt in format_euro_k(values.unit_sales_proceeds, 0)[:1]:
                    replacements.append((old_fmt, new_fmt))
        
        logger.info(f"Built {len(replacements)} numeric value replacements")
        return replacements
    
    def _update_text_boxes_xml(self, replacements: List[Tuple[str, str]], values: ReportValues = None):
        """
        Update text in text boxes by directly manipulating the docx XML.
        
        Text boxes are stored as drawing elements in the document.xml,
        and python-docx can't access them directly.
        
        IMPORTANT: Word XML often splits text across multiple <w:t> elements,
        e.g., "Q2 2025" might be stored as ["Q2", " 2025"]. We handle this
        by using regex patterns that account for XML tags between text fragments.
        
        Also applies context-aware numeric value replacements for Page 4 KPIs.
        """
        import zipfile
        import os
        
        # Read the docx as a zip file
        temp_dir = self.working_dir / "docx_xml"
        temp_dir.mkdir(parents=True, exist_ok=True)
        
        # Extract all files
        with zipfile.ZipFile(self.output_path, 'r') as zin:
            zin.extractall(temp_dir)
        
        # Process document.xml and any other relevant XML files
        xml_files = [
            temp_dir / "word" / "document.xml",
        ]
        
        # Also process header/footer files if they exist
        word_dir = temp_dir / "word"
        for f in word_dir.iterdir():
            if f.suffix == '.xml' and ('header' in f.name or 'footer' in f.name):
                xml_files.append(f)
        
        updated_count = 0
        numeric_updates = 0
        for xml_file in xml_files:
            if xml_file.exists():
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                original_content = content
                
                # Build fragmented patterns that account for XML tags between text pieces
                # E.g., "Q2 2025" might be: <w:t>Q2</w:t></w:r><w:r><w:t> 2025</w:t>
                content = self._apply_fragmented_replacements(content, replacements)
                
                # Apply context-aware numeric replacements for Page 4 KPIs
                if values:
                    content, num_count = self._apply_context_numeric_replacements(content, values)
                    numeric_updates += num_count
                    
                    # Apply fragmented euro value replacements
                    content, frag_count = self._apply_fragmented_euro_replacements(content, values)
                    numeric_updates += frag_count
                    
                    # Apply fragmented date replacements (e.g., "1- 7 -202 5" -> "1- 10 -202 5")
                    content = self._apply_fragmented_date_replacements(content, replacements)
                
                if content != original_content:
                    with open(xml_file, 'w', encoding='utf-8') as f:
                        f.write(content)
                    updated_count += 1
        
        if numeric_updates > 0:
            logger.info(f"Applied {numeric_updates} context-aware numeric replacements")
        
        logger.info(f"Updated {updated_count} XML files in docx")
        
        # Repack the docx
        with zipfile.ZipFile(self.output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zout.write(file_path, arcname)
        
        # Cleanup temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    def _apply_context_numeric_replacements(self, content: str, values: ReportValues) -> Tuple[str, int]:
        """
        Apply context-aware numeric value replacements in the Word XML.
        
        This method finds numeric values (€ amounts, percentages) that appear near
        specific KPI context words (GTRI, rental income, vacancy, etc.) and replaces
        them with the new quarter's values.
        
        Word XML fragments numbers across multiple <w:t> elements, e.g.:
        "€3,200k" might be: <w:t>€</w:t><w:t>3,</w:t><w:t>200</w:t><w:t>k</w:t>
        
        The document has specific Q2 values that need to be replaced with Q3 values:
        - GTRI: €3,200.6k → €3,273.6k
        - Financial vacancy: €106.3k → €177.5k (amount), 4.1% → 5.4% (percentage)
        - Gross rental income: €3,067.5k → €3,095.9k
        - Rent roll yields: €12,940.2k → €13,317.9k (this is LTM GTRI)
        - Unit sales proceeds: €762.5k → €203.5k
        - Maintenance: €297k → €244k
        
        Returns:
            (updated_content, count_of_replacements)
        """
        updates_made = 0
        
        # Define specific value replacements: (old_value, new_value, value_format, context_keywords)
        # Using exact Q2 values found in the template document
        specific_replacements = [
            # GTRI quarterly: €3,200.6k → new GTRI
            (3200.6, values.gtri, 'euro_k_decimal', ['GTRI', 'Gross Theoretical rental income']),
            (3200, values.gtri, 'euro_k', ['GTRI', 'Gross Theoretical rental income']),
            
            # Financial vacancy amount: €106.3k → new vacancy amount  
            (106.3, values.financial_vacancy_amount, 'euro_k_decimal', ['vacancy']),
            (128.5, values.financial_vacancy_amount, 'euro_k_decimal', ['vacancy']),  # Alternative Q2 value
            
            # Financial vacancy percentage: 4.1% → new percentage
            (4.1, values.financial_vacancy_pct, 'percent', ['vacancy']),
            
            # Gross rental income: €3,067.5k → new gross rental
            (3067.5, values.gross_rental_income, 'euro_k_decimal', ['gross rental income']),
            (3067, values.gross_rental_income, 'euro_k', ['gross rental income']),
            
            # Rent roll yields (this is LTM GTRI): €12,940.2k → new LTM GTRI
            (12940.2, values.gtri_ltm, 'euro_k_decimal', ['rent roll', 'yields']),
            (12940, values.gtri_ltm, 'euro_k', ['rent roll', 'yields']),
            
            # Unit sales proceeds: €762.5k → new proceeds
            (762.5, values.unit_sales_proceeds, 'euro_k_decimal', ['Unit sale proceeds', 'proceeds']),
            (762, values.unit_sales_proceeds, 'euro_k', ['sale proceeds']),
            
            # Maintenance: €297k → new maintenance
            (297, values.maintenance_amount, 'euro_k', ['Maintenance']),
        ]
        
        # Apply specific replacements
        for old_val, new_val, fmt, keywords in specific_replacements:
            if new_val <= 0:
                continue
            
            # Generate old and new formatted strings
            if fmt == 'euro_k_decimal':
                old_patterns = [
                    f'€{old_val:,.1f}k', f'€ {old_val:,.1f}k', f'€{old_val:,.1f} k',
                    f'€{old_val:.1f}k', f'€ {old_val:.1f}k',
                ]
                new_formatted = f'€{new_val:,.1f}k'
            elif fmt == 'euro_k':
                old_patterns = [
                    f'€{old_val:,.0f}k', f'€ {old_val:,.0f}k', f'€{old_val:,.0f} k',
                    f'€{int(old_val):,}k', f'€ {int(old_val):,}k',
                ]
                new_formatted = f'€{new_val:,.0f}k'
            elif fmt == 'percent':
                old_patterns = [f'{old_val:.1f}%', f'{old_val:.1f} %', f'{old_val}%']
                new_formatted = f'{new_val:.1f}%'
            else:
                continue
            
            # Try direct replacement first
            for old_pattern in old_patterns:
                if old_pattern in content:
                    content = content.replace(old_pattern, new_formatted)
                    updates_made += 1
                    logger.debug(f"Replaced '{old_pattern}' with '{new_formatted}'")
        
        # Define KPI contexts and their new values for context-aware replacement
        # Format: (context_keywords, new_value, value_format)
        kpi_mappings = [
            # GTRI - appears as "GTRI of the portfolio amounted to €X,XXXk"
            (['GTRI', 'Gross Theoretical Rental Income'], values.gtri, 'euro_k'),
            
            # Gross rental income - "Gross rental income was €X,XXXk"
            (['Gross rental income', 'gross rental income'], values.gross_rental_income, 'euro_k'),
            
            # Financial vacancy - "Financial vacancy was X.X%" or "vacancy of X.X%"
            (['Financial vacancy', 'vacancy'], values.financial_vacancy_pct, 'percent'),
            
            # Rent roll yields - uses LTM GTRI value
            (['rent roll', 'Rent roll', 'yields'], values.gtri_ltm, 'euro_k'),
            
            # Maintenance - "Maintenance expenses of €XXXk"
            (['Maintenance', 'maintenance', 'repair costs'], values.maintenance_amount, 'euro_k'),
            
            # Unit sales proceeds - "Unit sales proceeds of €XXXk"
            (['Unit sales proceeds', 'sale proceeds', 'proceeds'], values.unit_sales_proceeds, 'euro_k'),
            
            # CAPEX - "CAPEX of €XXXk"
            (['CAPEX', 'capex', 'capital expenditure'], values.capex_amount, 'euro_k'),
        ]
        
        for keywords, new_value, value_format in kpi_mappings:
            if new_value <= 0:
                continue
                
            # Find context regions containing the keywords
            for keyword in keywords:
                # Find all occurrences of the keyword in XML content
                keyword_positions = []
                pos = 0
                while True:
                    pos = content.find(keyword, pos)
                    if pos == -1:
                        break
                    keyword_positions.append(pos)
                    pos += 1
                
                for kw_pos in keyword_positions:
                    # Look for Euro amounts or percentages - need ~1200 chars as numbers can be highly fragmented
                    search_region = content[kw_pos:kw_pos + 1500]
                    
                    if value_format == 'euro_k':
                        # Pattern to find Euro amounts (handles XML fragmentation)
                        # Matches: €X,XXX or €X.XXX or just the number parts
                        
                        # First try to find the complete value in one <w:t> tag
                        euro_pattern = r'>€([\d,\.]+)([km]?)</w:t>'
                        match = re.search(euro_pattern, search_region)
                        if match:
                            old_val = match.group(0)
                            suffix = match.group(2) or 'k'
                            new_formatted = f'>€{new_value:,.0f}{suffix}</w:t>'
                            
                            # Only replace if this is in the original content at the right location
                            full_pos = kw_pos + match.start()
                            if content[full_pos:full_pos + len(old_val)] == old_val:
                                content = content[:full_pos] + new_formatted + content[full_pos + len(old_val):]
                                updates_made += 1
                                logger.debug(f"Replaced '{old_val}' with '{new_formatted}' near '{keyword}'")
                                break
                        
                        # Handle fragmented case: €</w:t>...<w:t>X,XXX</w:t>
                        # Word often splits numbers like "3,200.6" into: "3," | "20" | "0" | "." | "6"
                        # Strategy: Find the full region, replace in middle section, then swap back
                        
                        # Find the Euro sign and everything up to the 'k' or 'm' suffix
                        full_value_pattern = r'(>€</w:t>)(.*?)(>k</w:t>|>m</w:t>)'
                        match = re.search(full_value_pattern, search_region, re.DOTALL | re.IGNORECASE)
                        
                        if match:
                            middle_section = match.group(2)  # Everything between € and k
                            
                            # Find all number/dot fragments in the middle section
                            frag_pattern = r'>([\d,\.]+)</w:t>'
                            fragments = list(re.finditer(frag_pattern, middle_section))
                            
                            if fragments:
                                new_num = f'{new_value:,.0f}'
                                
                                # Rebuild middle section: first fragment gets new value, rest become empty
                                new_middle = middle_section
                                for i, frag in enumerate(fragments):
                                    old_text = f'>{frag.group(1)}</w:t>'
                                    if i == 0:
                                        # First fragment gets the new value
                                        new_middle = new_middle.replace(old_text, f'>{new_num}</w:t>', 1)
                                        logger.debug(f"Replaced fragment '{frag.group(1)}' with '{new_num}' near '{keyword}'")
                                    else:
                                        # Subsequent fragments become empty
                                        new_middle = new_middle.replace(old_text, '></w:t>', 1)
                                        logger.debug(f"Cleared fragment '{frag.group(1)}' near '{keyword}'")
                                
                                # Calculate absolute position and replace the middle section
                                abs_start = kw_pos + match.start(2)
                                abs_end = kw_pos + match.end(2)
                                content = content[:abs_start] + new_middle + content[abs_end:]
                                updates_made += 1
                                break
                    
                    elif value_format == 'percent':
                        # Pattern to find percentages: X.X% or X%
                        # Handle both complete (>5.4%</w:t>) and fragmented (>5.</w:t>...<w:t>4</w:t>...<w:t>%</w:t>)
                        
                        # First try complete pattern
                        pct_pattern = r'>([\d,\.]+)\s*(%)</w:t>'
                        match = re.search(pct_pattern, search_region)
                        if match:
                            old_val = match.group(0)
                            new_formatted = f'>{new_value:.1f}%</w:t>'
                            
                            full_pos = kw_pos + match.start()
                            if content[full_pos:full_pos + len(old_val)] == old_val:
                                content = content[:full_pos] + new_formatted + content[full_pos + len(old_val):]
                                updates_made += 1
                                logger.debug(f"Replaced percentage '{old_val}' with '{new_formatted}' near '{keyword}'")
                                break
                        
                        # Try fragmented pattern: >X.</w:t>...<w:t>Y</w:t>
                        # where X.Y is the percentage value (the % might be in same or separate element)
                        # Pattern: ">4.</w:t>...<w:t>1</w:t>" 
                        frag_pct_pattern = r'(>)(\d+)\.</w:t>.*?<w:t[^>]*>(\d+)</w:t>'
                        match = re.search(frag_pct_pattern, search_region, re.DOTALL)
                        if match:
                            old_int = match.group(2)
                            old_decimal = match.group(3)
                            
                            new_int = str(int(new_value))
                            new_decimal = str(int(round((new_value % 1) * 10)))  # Get first decimal digit
                            
                            logger.debug(f"Found fragmented percentage {old_int}.{old_decimal}% near '{keyword}', replacing with {new_int}.{new_decimal}%")
                            
                            # Replace integer part first (e.g., >4.</w:t> -> >4.</w:t>)
                            # match.start(1) gives position of the '>' which precedes the integer
                            int_pos = kw_pos + match.start(1)
                            old_int_full = f'>{old_int}.</w:t>'
                            new_int_full = f'>{new_int}.</w:t>'
                            
                            logger.debug(f"Looking for '{old_int_full}' at position {int_pos}, found: '{content[int_pos:int_pos + len(old_int_full)]}'")
                            
                            if content[int_pos:int_pos + len(old_int_full)] == old_int_full:
                                content = content[:int_pos] + new_int_full + content[int_pos + len(old_int_full):]
                                updates_made += 1
                                logger.debug(f"Replaced fragmented percentage integer '{old_int}.' with '{new_int}.' near '{keyword}'")
                                
                                # Now find and replace the decimal part
                                # Need to re-search after replacement
                                search_region_new = content[kw_pos:kw_pos + 800]
                                frag_pct_pattern2 = r'(>)(\d+)\.</w:t>.*?<w:t[^>]*>(\d+)</w:t>'
                                match_new = re.search(frag_pct_pattern2, search_region_new, re.DOTALL)
                                
                                if match_new:
                                    # Position of the decimal in the new search region
                                    dec_pos = kw_pos + match_new.start(3)
                                    old_dec_full = f'>{old_decimal}</w:t>'
                                    new_dec_full = f'>{new_decimal}</w:t>'
                                    
                                    # Need to find the actual position - the decimal is after the integer part
                                    # Search for the decimal value pattern after the integer
                                    dec_search_start = int_pos + len(new_int_full)
                                    dec_search_region = content[dec_search_start:dec_search_start + 300]
                                    dec_pattern = rf'<w:t[^>]*>({old_decimal})</w:t>'
                                    dec_match = re.search(dec_pattern, dec_search_region)
                                    
                                    if dec_match:
                                        actual_dec_pos = dec_search_start + dec_match.start(1)
                                        if content[actual_dec_pos:actual_dec_pos + len(old_decimal)] == old_decimal:
                                            content = content[:actual_dec_pos] + new_decimal + content[actual_dec_pos + len(old_decimal):]
                                            updates_made += 1
                                            logger.debug(f"Replaced fragmented percentage decimal '{old_decimal}' with '{new_decimal}' near '{keyword}'")
                            break
        
        return content, updates_made
    
    def _apply_fragmented_replacements(self, content: str, replacements: List[Tuple[str, str]]) -> str:
        """
        Apply replacements handling fragmented text in Word XML.
        
        Word often splits text across multiple <w:t> elements. This method handles:
        1. Direct text replacements (when text is not fragmented)
        2. XML-aware pattern replacements for common fragments
        """
        import re
        
        # First, apply direct replacements
        for old_text, new_text in replacements:
            if old_text in content:
                content = content.replace(old_text, new_text)
        
        # Apply common fragmented patterns using XML-aware regex
        # These handle cases where Word splits text across <w:t> elements
        content = self._apply_xml_aware_patterns(content, replacements)
        
        return content
    
    def _apply_fragmented_euro_replacements(self, content: str, values: ReportValues) -> Tuple[str, int]:
        """
        Apply replacements for fragmented euro values in Word XML.
        
        Word XML often splits numbers like "€3,200.6k" into multiple elements:
        - <w:t>€</w:t><w:t>3,0</w:t><w:t>6</w:t><w:t>7</w:t><w:t>.</w:t><w:t>5</w:t><w:t> k</w:t>
        
        This method finds specific old values and replaces them with new values,
        handling the XML fragmentation.
        
        Returns:
            (updated_content, count_of_replacements)
        """
        import re
        updates_made = 0
        
        # Define specific value replacements
        # Format: (old_first_frag, old_pattern_regex, new_value, description)
        # The pattern should capture the first numeric fragment that will hold the new value
        
        value_replacements = [
            # Gross rental income: "€ 3,0 6 7 . 5 k" -> new value
            # Pattern: >3,0</w:t>...<w:t>6</w:t>...<w:t>7</w:t>...<w:t>.</w:t>...<w:t>5</w:t>
            (r'>(3,0)</w:t>(.*?<w:t[^>]*>)(6)(</w:t>.*?<w:t[^>]*>)(7)(</w:t>.*?<w:t[^>]*>)(\.)(<w:t[^>]*>|</w:t>.*?<w:t[^>]*>)(5)</w:t>',
             values.gross_rental_income, 'gross_rental'),
            
            # Rent roll yields: " yields €12, 940 . 2 k" -> new LTM value  
            # The structure is: " yields €12," + "940" + "." + "2"
            # Pattern: > yields €12,</w:t>...<w:t>940</w:t>...<w:t>.</w:t>...<w:t>2</w:t>
            (r'>( yields €12,)</w:t>(.*?<w:t[^>]*>)(940)(</w:t>.*?<w:t[^>]*>)(\.)(<w:t[^>]*>|</w:t>.*?<w:t[^>]*>)(2)</w:t>',
             values.gtri_ltm, 'rent_roll_yields'),
            
            # Unit sale proceeds: "€ 762 . 5 k" -> new value
            # Pattern: >762</w:t>...<w:t>.</w:t>...<w:t>5</w:t>
            (r'>(762)</w:t>(.*?<w:t[^>]*>)(\.)(<w:t[^>]*>|</w:t>.*?<w:t[^>]*>)(5)</w:t>',
             values.unit_sales_proceeds, 'unit_sales'),
        ]
        
        for pattern, new_value, desc in value_replacements:
            if new_value <= 0:
                continue
            
            match = re.search(pattern, content, re.DOTALL)
            if match:
                # Format new value
                new_formatted = f'{new_value:,.1f}'
                
                old_full = match.group(0)
                first_group = match.group(1)
                
                # Handle rent_roll_yields specially - transform structure to match expected
                # Uploaded: " yields €12," + "940" + "." + "2" + "k"
                # Expected: " yields €" + "13,317" + "." + "9" + "k"
                if desc == 'rent_roll_yields':
                    int_part = int(new_value)
                    dec_part = int(round((new_value - int_part) * 10))
                    
                    # Replace " yields €12," with " yields €"
                    new_full = old_full.replace(f'>{first_group}</w:t>', '> yields €</w:t>', 1)
                    # Replace "940" with the formatted integer part
                    new_full = re.sub(r'>(940)</w:t>', f'>{int_part:,}</w:t>', new_full)
                    # Keep "." as is
                    # Replace "2" with the decimal digit
                    new_full = re.sub(r'>(2)</w:t>', f'>{dec_part}</w:t>', new_full)
                else:
                    # Replace first numeric group with new value
                    new_full = old_full.replace(f'>{first_group}</w:t>', f'>{new_formatted}</w:t>', 1)
                    
                    # Clear subsequent numeric fragments (replace digits with empty)
                    first_close = new_full.find('</w:t>')
                    if first_close > 0:
                        rest = new_full[first_close:]
                        # Clear numeric fragments
                        rest = re.sub(r'>(\d+)</w:t>', '></w:t>', rest)
                        rest = re.sub(r'>( ?\d+)</w:t>', '></w:t>', rest)
                        rest = re.sub(r'>(\.)</w:t>', '></w:t>', rest)
                        new_full = new_full[:first_close] + rest
                
                if new_full != old_full:
                    content = content.replace(old_full, new_full, 1)
                    updates_made += 1
                    logger.debug(f"Replaced fragmented {desc}: {new_formatted}")
        
        return content, updates_made
    
    def _apply_fragmented_date_replacements(self, content: str, replacements: List[Tuple[str, str]]) -> str:
        """
        Apply replacements for fragmented dates in Word XML.
        
        Dates like "1-7-2025" are often fragmented as "1- 7 -202 5" across multiple
        <w:t> elements. This method handles these patterns for rent roll dates.
        
        Args:
            content: The XML content to process
            replacements: List of (old, new) text replacement tuples
            
        Returns:
            Updated content with date replacements applied
        """
        import re
        
        # Extract old and new quarter info from replacements
        old_q_num = None
        new_q_num = None
        for old_text, new_text in replacements:
            old_match = re.match(r'Q(\d)\s+(\d{4})', old_text)
            new_match = re.match(r'Q(\d)\s+(\d{4})', new_text)
            if old_match and new_match:
                old_q_num = int(old_match.group(1))
                new_q_num = int(new_match.group(1))
                break
        
        if not old_q_num:
            return content
        
        # Calculate rent roll date months (first day of month after quarter end)
        # Q2 ends June -> rent roll 1 July (7)
        # Q3 ends September -> rent roll 1 October (10)
        old_month = (old_q_num * 3) + 1
        if old_month > 12:
            old_month = 1
        new_month = (new_q_num * 3) + 1
        if new_month > 12:
            new_month = 1
        
        logger.debug(f"Replacing rent roll date month: {old_month} -> {new_month}")
        
        # Find "per 1-" context followed by month number
        # Pattern matches: >per </w:t>...<w:t>1-</w:t>...<w:t> 7 </w:t>
        # We need to replace the " 7 " with " 10 "
        
        # First, find "1-" followed by month in the XML
        # The pattern is: >1-</w:t></w:r><w:r...><w:t> 7 </w:t>
        date_pattern = rf'(>1-</w:t>)(.*?)(<w:t[^>]*>)(\s*){old_month}(\s*)(</w:t>)'
        
        def replace_month(match):
            before = match.group(1)
            middle = match.group(2)
            tag_open = match.group(3)
            space_before = match.group(4)
            space_after = match.group(5)
            tag_close = match.group(6)
            return f'{before}{middle}{tag_open}{space_before}{new_month}{space_after}{tag_close}'
        
        content = re.sub(date_pattern, replace_month, content, flags=re.DOTALL)
        
        return content
    
    def _apply_xml_aware_patterns(self, content: str, replacements: List[Tuple[str, str]]) -> str:
        """
        Apply XML-aware patterns for commonly fragmented text.
        
        Word frequently fragments text like "Q2 2025" into separate runs:
        - <w:t>Q2</w:t>...<w:t> 2025</w:t>
        - <w:t>Q</w:t>...<w:t>2</w:t>...<w:t> 2025</w:t>
        - <w:t>21</w:t>...<w:t> juli</w:t>...<w:t> 2025</w:t>
        - <w:t>in Q</w:t>...<w:t>2</w:t>...<w:t> 2025</w:t>
        
        This function handles these cases with targeted regex patterns.
        """
        import re
        
        # Extract quarter numbers from replacements for dynamic pattern building
        old_q_num = None
        new_q_num = None
        old_year = None
        new_year = None
        
        for old_text, new_text in replacements:
            # Find "Q2 2025" -> "Q3 2025" style replacements
            old_match = re.match(r'Q(\d)\s+(\d{4})', old_text)
            new_match = re.match(r'Q(\d)\s+(\d{4})', new_text)
            if old_match and new_match:
                old_q_num = old_match.group(1)
                old_year = old_match.group(2)
                new_q_num = new_match.group(1)
                new_year = new_match.group(2)
                break
        
        if not old_q_num:
            return content
        
        # CRITICAL: Handle fragmented pattern where "in Q" is followed by "2" in next run
        # Pattern: <w:t>in Q</w:t></w:r>...<w:t>2</w:t> → replace 2 with 3
        #
        # IMPORTANT: Use a restrictive pattern that only matches the IMMEDIATELY NEXT <w:t> element
        # to prevent accidentally matching "2" that's part of year "2025" (which can be stored as "20"|"2"|"5")
        #
        # KEY FIX: Use (?:(?!</w:rPr>).)* instead of .*? in <w:rPr> to prevent backtracking
        # from consuming multiple </w:rPr> closings (which would skip over multiple runs)
        
        # Pattern for matching immediately next <w:t> after closing </w:r>
        # The (?:(?!</w:rPr>).)* is a negative lookahead that prevents matching </w:rPr>
        next_wt_pattern = r'</w:t></w:r><w:r[^>]*>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:t[^>]*>'
        
        # Match: 'in Q</w:t>' followed by IMMEDIATE next <w:t> containing the quarter number
        fragmented_pattern = rf'(in Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Also handle '% in Q</w:t>' pattern
        fragmented_pattern2 = rf'(% in Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern2, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle 'In Q</w:t>' (capital I)
        fragmented_pattern3 = rf'(In Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern3, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle ' Q</w:t>' followed by number (space before Q)
        fragmented_pattern4 = rf'( Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern4, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle 'Actions Q</w:t>' followed by number
        fragmented_pattern5 = rf'(Actions Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern5, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle '– Q</w:t>' followed by number (for Portfolio highlights)
        fragmented_pattern6 = rf'(– Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern6, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle '- Q</w:t>' followed by number (dash variant)
        fragmented_pattern7 = rf'(- Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern7, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle 's Q</w:t>' followed by number (for 'highlights Q2')
        fragmented_pattern8 = rf'(s Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern8, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle single 'Q</w:t>' followed by number in separate element
        # This catches cases where "Q" and "2" are completely split
        fragmented_pattern9 = rf'(>Q{next_wt_pattern}){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern9, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # XML tag pattern that may appear between text fragments
        xml_gap = r'</w:t></w:r>(?:<w:r[^>]*>)?(?:<w:rPr[^>]*>.*?</w:rPr>)?<w:t[^>]*>'
        
        # Define fragmented pattern replacements
        xml_patterns = [
            # "Q2 2025" fragmented as Q + 2 + " 2025" or "Q2" + " 2025"
            (rf'>Q{old_q_num}</w:t>', f'>Q{new_q_num}</w:t>'),
            (rf'>Q{old_q_num}<', f'>Q{new_q_num}<'),
            (rf'>Q{old_q_num} {old_year}<', f'>Q{new_q_num} {new_year}<'),
            
            # "in Q2" patterns (common in narrative text)
            (rf'> in Q{old_q_num}<', f'> in Q{new_q_num}<'),
            (rf'>in Q{old_q_num}<', f'>in Q{new_q_num}<'),
            (rf' in Q{old_q_num} {old_year}', f' in Q{new_q_num} {new_year}'),
            (rf'in Q{old_q_num} {old_year}', f'in Q{new_q_num} {new_year}'),
            
            # "Actions Q2 2025" patterns
            (rf'>Actions Q{old_q_num} {old_year}<', f'>Actions Q{new_q_num} {new_year}<'),
            (rf'Actions Q{old_q_num} {old_year}', f'Actions Q{new_q_num} {new_year}'),
            (rf'>Actions Q{old_q_num}<', f'>Actions Q{new_q_num}<'),
            
            # "In Q2 2025" patterns
            (rf'>In Q{old_q_num} {old_year}<', f'>In Q{new_q_num} {new_year}<'),
            (rf'In Q{old_q_num} {old_year}', f'In Q{new_q_num} {new_year}'),
            
            # Unit/amount references like "units in Q2 2025"
            (rf' Q{old_q_num} {old_year}', f' Q{new_q_num} {new_year}'),
            
            # "2025 – Q2" patterns
            (rf'>{old_year} – Q{old_q_num}<', f'>{new_year} – Q{new_q_num}<'),
            (rf'>{old_year} - Q{old_q_num}<', f'>{new_year} - Q{new_q_num}<'),
            (rf'>{old_year[-2:]} – Q{old_q_num}<', f'>{new_year[-2:]} – Q{new_q_num}<'),
            (rf'5 – Q{old_q_num}<', f'5 – Q{new_q_num}<'),  # For "2025 – Q2" split
            (rf'{old_year} – Q{old_q_num}', f'{new_year} – Q{new_q_num}'),
            
            # Ordinal patterns - dynamically determined by quarter
            # TODO: Make these dynamic based on quarter numbers
            (r'>second quarter<', '>third quarter<'),
            (r'>second<', '>third<'),  # For split "second quarter"
            (r'>the second<', '>the third<'),
            (r'second quarter', 'third quarter'),
            (r'the second quarter', 'the third quarter'),
            
            # NOTE: Cover page date is handled by _build_replacement_map with values.report_date
            # Do NOT add hardcoded date patterns here that would override the proper replacement
            
            # Quarter-end date patterns (these are data dates, not report date)
            (rf'>30-6-{old_year}<', f'>30-9-{new_year}<'),
            (rf'>1-7-{old_year}<', f'>1-10-{new_year}<'),
            
            # Month patterns for quarter-end references (e.g., "as of June 2025")
            # NOT for report/cover page dates
            (r'>June<', '>September<'),
            (r'>juni<', '>september<'),
        ]
        
        for pattern, replacement in xml_patterns:
            try:
                if re.search(pattern, content):
                    content = re.sub(pattern, replacement, content)
                    logger.debug(f"XML pattern: {pattern} -> {replacement}")
            except re.error as e:
                logger.debug(f"Regex error for pattern {pattern}: {e}")
        
        return content
    
    def _get_old_date_patterns(self, previous_quarter: str) -> List[str]:
        """
        Generate date patterns that might appear on cover page from previous quarter.
        
        IMPORTANT: Patterns are sorted longest-first to prevent substring matching issues.
        e.g., "21 juli 2025" must be matched before "1 juli 2025"
        """
        patterns = []
        
        # Extract year from quarter
        parts = previous_quarter.split()
        if len(parts) == 2:
            year = int(parts[1])
            q_num = int(parts[0][1])
            
            # Calculate previous quarter end month
            month = q_num * 3
            
            # Generate various date formats for several months around report date
            # Reports are typically dated 2-3 weeks after quarter end
            for m in range(month, min(month + 3, 13)):
                for d in range(1, 32):
                    try:
                        from datetime import datetime
                        date = datetime(year, m, d)
                        # Common date formats
                        patterns.extend([
                            date.strftime('%d %B %Y'),  # 21 July 2025
                            date.strftime('%d %b %Y'),  # 21 Jul 2025
                            # Dutch formats
                            self._format_dutch_date(date),  # 21 juli 2025
                        ])
                    except ValueError:
                        continue
        
        # Remove duplicates and empty, then SORT BY LENGTH (longest first)
        # This prevents "1 juli 2025" from matching inside "21 juli 2025"
        unique_patterns = list(set(p for p in patterns if p))
        return sorted(unique_patterns, key=len, reverse=True)
    
    def _format_dutch_date(self, date) -> str:
        """Format date in Dutch style."""
        dutch_months = {
            1: 'januari', 2: 'februari', 3: 'maart', 4: 'april',
            5: 'mei', 6: 'juni', 7: 'juli', 8: 'augustus',
            9: 'september', 10: 'oktober', 11: 'november', 12: 'december'
        }
        return f"{date.day} {dutch_months[date.month]} {date.year}"
    
    def _update_cover_page_date(self, para, old_patterns: List[str], new_date: str):
        """Update cover page date to current run date (Issue 7.2)."""
        text = para.text
        for old_pattern in old_patterns:
            if old_pattern in text:
                self._replace_in_runs(para, old_pattern, new_date)
                return  # Only replace once per paragraph
    
    def _process_paragraph(self, para, values: ReportValues, 
                          old_quarter: str, new_quarter: str):
        """Process a single paragraph for replacements."""
        # Check if paragraph contains text to replace
        text = para.text
        
        # Extract quarter components for dynamic patterns
        old_parts = old_quarter.split()
        new_parts = new_quarter.split()
        old_q_num = old_parts[0][1] if len(old_parts) == 2 else ''
        old_year = old_parts[1] if len(old_parts) == 2 else ''
        new_q_num = new_parts[0][1] if len(new_parts) == 2 else ''
        new_year = new_parts[1] if len(new_parts) == 2 else ''
        
        # Quarter replacement
        if old_quarter in text:
            self._replace_in_runs(para, old_quarter, new_quarter)
        
        # Handle various quarter formats dynamically
        quarter_patterns = [
            (old_quarter.replace(' ', '-'), new_quarter.replace(' ', '-')),  # Q2-2025
            (old_quarter.replace(' ', '/'), new_quarter.replace(' ', '/')),  # Q2/2025
            (f"Q{old_q_num} {old_year[-2:]}", f"Q{new_q_num} {new_year[-2:]}"),  # Q2 25
            (f"Q{old_q_num}-{old_year[-2:]}", f"Q{new_q_num}-{new_year[-2:]}"),  # Q2-25
            (f"{old_year[-2:]}Q{old_q_num}", f"{new_year[-2:]}Q{new_q_num}"),  # 25Q2
        ]
        for old_pattern, new_pattern in quarter_patterns:
            if old_pattern in text:
                self._replace_in_runs(para, old_pattern, new_pattern)
        
        # Handle ordinal quarter references (second quarter, third quarter, etc.)
        ordinal_map = {
            '1': ('first', 'eerste'),
            '2': ('second', 'tweede'),
            '3': ('third', 'derde'),
            '4': ('fourth', 'vierde'),
        }
        if old_q_num in ordinal_map and new_q_num in ordinal_map:
            old_ordinal_en, old_ordinal_nl = ordinal_map[old_q_num]
            new_ordinal_en, new_ordinal_nl = ordinal_map[new_q_num]
            
            # English ordinal patterns
            ordinal_patterns = [
                (f"{old_ordinal_en} quarter", f"{new_ordinal_en} quarter"),
                (f"{old_ordinal_en.capitalize()} quarter", f"{new_ordinal_en.capitalize()} quarter"),
                (f"{old_ordinal_en} Quarter", f"{new_ordinal_en} Quarter"),
                # Dutch ordinal patterns
                (f"{old_ordinal_nl} kwartaal", f"{new_ordinal_nl} kwartaal"),
                (f"{old_ordinal_nl} Kwartaal", f"{new_ordinal_nl} Kwartaal"),
            ]
            for old_ord, new_ord in ordinal_patterns:
                if old_ord in text:
                    self._replace_in_runs(para, old_ord, new_ord)
        
        # Handle month/date replacements dynamically
        quarter_to_month = {
            '1': ('March', 'maart', '31', '3'),
            '2': ('June', 'juni', '30', '6'),
            '3': ('September', 'september', '30', '9'),
            '4': ('December', 'december', '31', '12'),
        }
        
        if old_q_num in quarter_to_month and new_q_num in quarter_to_month:
            old_month, old_month_nl, old_day, old_month_num = quarter_to_month[old_q_num]
            new_month, new_month_nl, new_day, new_month_num = quarter_to_month[new_q_num]
            
            # Replace month names (only if they appear as standalone words related to quarter)
            # Be careful not to replace months in unrelated contexts
            if old_month in text:
                self._replace_in_runs(para, old_month, new_month)
            if old_month_nl in text:
                self._replace_in_runs(para, old_month_nl, new_month_nl)
            
            # Replace date patterns - sort by length to prevent partial matches
            date_patterns = [
                # Longer patterns first
                (f"{old_day} {old_month} {old_year}", f"{new_day} {new_month} {new_year}"),
                (f"{old_day} {old_month_nl} {old_year}", f"{new_day} {new_month_nl} {new_year}"),
                (f"{old_day}-{old_month_num}-{old_year}", f"{new_day}-{new_month_num}-{new_year}"),
                # Rent roll date: 1-7-2025 (first of month after quarter)
                (f"1-{int(old_month_num)+1}-{old_year}", f"1-{int(new_month_num)+1}-{new_year}"),
            ]
            for old_date, new_date in date_patterns:
                if old_date in text:
                    self._replace_in_runs(para, old_date, new_date)
        
        # Placeholder replacements
        for placeholder, attr_name in self.PLACEHOLDERS.items():
            if placeholder in text:
                value = getattr(values, attr_name, '')
                
                # Format value appropriately
                if isinstance(value, float):
                    if 'pct' in attr_name or 'vacancy' in attr_name.lower():
                        formatted = f"{value:.1f}%"
                    elif 'amount' in attr_name.lower() or 'gtri' in attr_name.lower() or 'rent' in attr_name.lower():
                        formatted = f"€{value:,.1f}k"
                    else:
                        formatted = f"{value:,.1f}"
                elif isinstance(value, int):
                    formatted = str(value)
                else:
                    formatted = str(value)
                
                self._replace_in_runs(para, placeholder, formatted)
        
        # Handle REPORT_DATE placeholder
        if '{{REPORT_DATE}}' in text:
            self._replace_in_runs(para, '{{REPORT_DATE}}', values.report_date)
        
        # Handle narrative placeholders
        if '{{UNIT_SALES_NARRATIVE}}' in text and values.unit_sales_narrative:
            self._replace_in_runs(para, '{{UNIT_SALES_NARRATIVE}}', values.unit_sales_narrative)
        
        if '{{MAINTENANCE_DETAIL}}' in text and values.maintenance_detail:
            self._replace_in_runs(para, '{{MAINTENANCE_DETAIL}}', values.maintenance_detail)
        
        if '{{SUSTAINABILITY_DETAIL}}' in text and values.sustainability_detail:
            self._replace_in_runs(para, '{{SUSTAINABILITY_DETAIL}}', values.sustainability_detail)
    
    def _replace_in_runs(self, para, old_text: str, new_text: str):
        """Replace text while preserving run formatting."""
        # Simple case: text is in a single run
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                return
        
        # Complex case: text spans multiple runs
        # Reconstruct and replace
        full_text = para.text
        if old_text in full_text:
            new_full = full_text.replace(old_text, new_text)
            
            # Clear all runs and add new text with first run's formatting
            if para.runs:
                first_run = para.runs[0]
                for run in para.runs:
                    run.text = ''
                first_run.text = new_full

