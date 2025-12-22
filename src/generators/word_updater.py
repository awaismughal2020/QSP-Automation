"""
Word Template Updater

Updates the quarterly report Word template with:
- Global find/replace for quarter references
- Specific value updates on designated pages
- Date updates
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
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
    
    # Page 4 - Executive Summary
    gtri: float  # Gross Theoretical Rental Income (€k)
    gross_rental_income: float  # Actual rental income (€k)
    rent_roll_annual: float  # Annual rent roll (€k)
    financial_vacancy_pct: float  # Vacancy percentage
    units_sold_quarter: int  # Units sold this quarter
    maintenance_amount: float  # Maintenance spend (€k)
    capex_amount: float  # CAPEX spend (€k)
    
    # Page 5 - Unit Sales narrative
    unit_sales_narrative: str  # Free text about unit sales actions
    
    # Page 6
    maintenance_detail: str  # Maintenance description
    sustainability_detail: str  # Sustainability/CAPEX description


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
            
            # Text variations
            (f"Actions {old_quarter}", f"Actions {new_quarter}"),
            (f"Portfolio highlights – {old_quarter}", f"Portfolio highlights – {new_quarter}"),
            (f"Portfolio highlights - {old_quarter}", f"Portfolio highlights - {new_quarter}"),
            (f"Report {old_quarter}", f"Report {new_quarter}"),
            (f"Quarter {old_q_num} {old_year}", f"Quarter {new_q_num} {new_year}"),
            
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
            
            # Replace date formats like "30 June 2025"
            for year in [old_year, str(int(old_year) - 1), str(int(old_year) + 1)]:
                old_date = f"{old_day} {old_month} {year}"
                new_date = f"{new_day} {new_month} {new_year}"
                content = content.replace(old_date, new_date)
                
                # Dutch format: 30 juni 2025
                old_date_nl = f"{old_day} {old_month_nl} {year}"
                new_date_nl = f"{new_day} {new_month_nl} {new_year}"
                content = content.replace(old_date_nl, new_date_nl)
        
        return content
    
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
        Alternative update method using python-docx directly.
        Better for preserving complex formatting.
        
        Args:
            values: ReportValues with all data to insert
            previous_quarter: Previous quarter string (e.g., "Q2 2025")
            current_quarter: Current quarter string (e.g., "Q3 2025")
            
        Returns:
            Path to updated Word document
        """
        from docx import Document
        
        logger.info(f"Updating Word template for {current_quarter} (python-docx method)")
        
        # Copy template to output location
        shutil.copy2(self.template_path, self.output_path)
        
        # Open and modify
        doc = Document(self.output_path)
        
        # Build list of old dates to replace (for cover page dynamic date - Issue 7.2)
        # The cover page date should be the current run date, not copied from template
        old_date_patterns = self._get_old_date_patterns(previous_quarter)
        
        # Process paragraphs
        for para in doc.paragraphs:
            self._process_paragraph(para, values, previous_quarter, current_quarter)
            # Also update any old dates on cover page to current run date
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
        logger.info(f"Generated updated report: {self.output_path}")
        return self.output_path
    
    def _get_old_date_patterns(self, previous_quarter: str) -> List[str]:
        """Generate date patterns that might appear on cover page from previous quarter."""
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
                            date.strftime('%-d %B %Y'),  # 21 July 2025 (no leading zero)
                            # Dutch formats
                            self._format_dutch_date(date),  # 21 juli 2025
                        ])
                    except ValueError:
                        continue
        
        # Remove duplicates and empty
        return list(set(p for p in patterns if p))
    
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
        
        # Quarter replacement
        if old_quarter in text:
            self._replace_in_runs(para, old_quarter, new_quarter)
        
        # Handle various quarter formats
        quarter_patterns = [
            (old_quarter.replace(' ', '-'), new_quarter.replace(' ', '-')),
            (old_quarter.replace(' ', '/'), new_quarter.replace(' ', '/')),
        ]
        for old_pattern, new_pattern in quarter_patterns:
            if old_pattern in text:
                self._replace_in_runs(para, old_pattern, new_pattern)
        
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

