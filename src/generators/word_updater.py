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
        
        # Build replacement map for text boxes
        text_replacements = self._build_replacement_map(values, previous_quarter, current_quarter)
        
        # Step 1: Update text boxes via direct XML manipulation
        self._update_text_boxes_xml(text_replacements)
        
        # Step 2: Update regular paragraphs/tables via python-docx
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
        logger.info(f"Generated updated report: {self.output_path}")
        return self.output_path
    
    def _build_replacement_map(self, values: ReportValues, old_quarter: str, new_quarter: str) -> List[Tuple[str, str]]:
        """
        Build comprehensive replacement map for all text that needs updating.
        
        Returns list of (old_text, new_text) tuples, sorted longest-first.
        """
        replacements = []
        
        # Extract quarter components
        old_parts = old_quarter.split()
        new_parts = new_quarter.split()
        old_q_num = old_parts[0][1] if len(old_parts) == 2 else ''
        old_year = old_parts[1] if len(old_parts) == 2 else ''
        new_q_num = new_parts[0][1] if len(new_parts) == 2 else ''
        new_year = new_parts[1] if len(new_parts) == 2 else ''
        
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
    
    def _update_text_boxes_xml(self, replacements: List[Tuple[str, str]]):
        """
        Update text in text boxes by directly manipulating the docx XML.
        
        Text boxes are stored as drawing elements in the document.xml,
        and python-docx can't access them directly.
        
        IMPORTANT: Word XML often splits text across multiple <w:t> elements,
        e.g., "Q2 2025" might be stored as ["Q2", " 2025"]. We handle this
        by using regex patterns that account for XML tags between text fragments.
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
        for xml_file in xml_files:
            if xml_file.exists():
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                original_content = content
                
                # Build fragmented patterns that account for XML tags between text pieces
                # E.g., "Q2 2025" might be: <w:t>Q2</w:t></w:r><w:r><w:t> 2025</w:t>
                content = self._apply_fragmented_replacements(content, replacements)
                
                if content != original_content:
                    with open(xml_file, 'w', encoding='utf-8') as f:
                        f.write(content)
                    updated_count += 1
        
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
        # This is a common fragmentation in Word
        def replace_fragmented_q_num(match):
            """Replace the quarter number in fragmented 'in Q' + '2' pattern."""
            return match.group(0).replace(f'>{old_q_num}</w:t>', f'>{new_q_num}</w:t>')
        
        # Match: 'in Q</w:t>' followed eventually by '>2</w:t>'
        fragmented_pattern = rf'(in Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Also handle '% in Q</w:t>' pattern
        fragmented_pattern2 = rf'(% in Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern2, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle 'In Q</w:t>' (capital I)
        fragmented_pattern3 = rf'(In Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern3, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle ' Q</w:t>' followed by number (space before Q)
        fragmented_pattern4 = rf'( Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern4, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle 'Actions Q</w:t>' followed by number
        fragmented_pattern5 = rf'(Actions Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern5, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle '– Q</w:t>' followed by number (for Portfolio highlights)
        fragmented_pattern6 = rf'(– Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern6, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle '- Q</w:t>' followed by number (dash variant)
        fragmented_pattern7 = rf'(- Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern7, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle 's Q</w:t>' followed by number (for 'highlights Q2')
        fragmented_pattern8 = rf'(s Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
        content = re.sub(fragmented_pattern8, rf'\g<1>{new_q_num}\g<2>', content, flags=re.DOTALL)
        
        # Handle single 'Q</w:t>' followed by number in separate element
        # This catches cases where "Q" and "2" are completely split
        fragmented_pattern9 = rf'(>Q</w:t></w:r>.*?<w:t[^>]*>){old_q_num}(</w:t>)'
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
            
            # Ordinal patterns
            (r'>second quarter<', '>third quarter<'),
            (r'>second<', '>third<'),  # For split "second quarter"
            (r'>the second<', '>the third<'),
            (r'second quarter', 'third quarter'),
            (r'the second quarter', 'the third quarter'),
            
            # Date patterns - juli to oktober for Q2->Q3
            (rf'>21 juli {old_year}<', f'>21 oktober {new_year}<'),
            (rf'> juli<', '> oktober<'),
            (rf'>juli<', '>oktober<'),
            (rf'>1-7-{old_year}<', f'>1-10-{new_year}<'),
            (rf'>30-6-{old_year}<', f'>30-9-{new_year}<'),
            (rf'1-7-{old_year}', f'1-10-{new_year}'),
            
            # Month patterns
            (r'>June<', '>September<'),
            (r'>juni<', '>september<'),
            (r'>July<', '>October<'),
            (r'>juli<', '>oktober<'),
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

