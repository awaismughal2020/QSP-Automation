"""
Management Accounts Transformer

Builds the Management Accounts Excel file by:
1. Copying previous quarter structure
2. Cloning the previous BDO sheet and adding new column for current quarter
3. Updating Management Cijfers summary with new column
4. Updating all date references

CRITICAL: The BDO sheet structure must be preserved exactly because formulas
in Management Cijfers reference specific row numbers. Cloning the sheet and
inserting columns (not creating new rows) ensures formula integrity.
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, PatternFill, Side
from openpyxl.styles.fills import FILL_SOLID
from copy import copy
import yaml
from datetime import datetime
from loguru import logger

from ..parsers.bdo_parser import BDOParseResult, AccountEntry


@dataclass
class ManagementAccountsConfig:
    """Configuration for Management Accounts generation."""
    quarter: str  # e.g., "Q3 2025"
    period_end: datetime
    bdo_sheet_name: str  # e.g., "BDO - Q3-25"
    summary_sheet_name: str  # e.g., "Management Cijfers - Q3 2025"


class ManagementAccountsBuilder:
    """
    Builds updated Management Accounts workbook.
    
    The Management Accounts file contains:
    - Historical BDO sheets (one per quarter from Q1 2019 onwards)
    - Summary "Management Cijfers" sheet with calculated columns
    
    CRITICAL APPROACH (per feedback):
    - Clone the previous BDO sheet to preserve exact row structure
    - Insert a new column for the current quarter's mutations
    - Update values based on account codes (semantic mapping)
    - This ensures all formula references remain valid
    
    Line item mappings are loaded from config/line_item_mappings.yaml
    """
    
    def __init__(self, previous_file_path: str, output_path: str, config: ManagementAccountsConfig,
                 mappings_path: str = "config/line_item_mappings.yaml"):
        self.previous_path = Path(previous_file_path)
        self.output_path = Path(output_path)
        self.config = config
        self.workbook = None
        self.mappings_path = Path(mappings_path)
        self.line_item_mappings = self._load_mappings()
        self.equity_prefixes = []
        self.income_prefixes = []
        self.expense_prefixes = []
        self._load_validation_config()
        
        # Account code to BDO row mapping (built during BDO sheet processing)
        self._account_row_map = {}
    
    def _load_mappings(self) -> Dict[str, Tuple[List[str], str]]:
        """Load line item mappings from YAML config file."""
        mappings = {}
        
        if not self.mappings_path.exists():
            logger.warning(f"Mappings file not found: {self.mappings_path}, using defaults")
            return self._get_default_mappings()
        
        with open(self.mappings_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        # Process balance sheet mappings
        if 'balance_sheet' in config:
            for label, item_config in config['balance_sheet'].items():
                if item_config:
                    accounts = item_config.get('accounts', [])
                    calc_type = item_config.get('calc_type', 'sum')
                    mappings[label] = (accounts, calc_type)
        
        # Process P&L mappings
        if 'profit_loss' in config:
            for label, item_config in config['profit_loss'].items():
                if item_config:
                    accounts = item_config.get('accounts', [])
                    calc_type = item_config.get('calc_type', 'sum')
                    mappings[label] = (accounts, calc_type)
        
        logger.info(f"Loaded {len(mappings)} line item mappings from {self.mappings_path}")
        return mappings
    
    def _load_validation_config(self):
        """Load validation account prefixes from config."""
        if not self.mappings_path.exists():
            self.equity_prefixes = ['10', '11']
            self.income_prefixes = ['8']
            self.expense_prefixes = ['4']
            return
        
        with open(self.mappings_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        if 'equity_accounts' in config:
            self.equity_prefixes = config['equity_accounts'].get('prefixes', ['10', '11'])
        else:
            self.equity_prefixes = ['10', '11']
            
        if 'income_accounts' in config:
            self.income_prefixes = config['income_accounts'].get('prefixes', ['8'])
        else:
            self.income_prefixes = ['8']
            
        if 'expense_accounts' in config:
            self.expense_prefixes = config['expense_accounts'].get('prefixes', ['4'])
        else:
            self.expense_prefixes = ['4']
    
    def _get_default_mappings(self) -> Dict[str, Tuple[List[str], str]]:
        """Return default mappings if config file not found."""
        return {
            "Deferred Tax Asset": (["1790002"], "direct"),
            "Real estate": (["1600000", "1600200", "1601000", "1610803", "1611003"], "sum"),
            "Financial fixed assets": (["1760*", "1790000"], "sum"),
            "Accounts receivable": (["2000000", "2000100", "2000300"], "sum"),
            "Cash": (["2400*"], "sum"),
            "Equity": (["1000*", "1100000", "1160000"], "sum"),
            "Bank loan": (["1930001", "1930101"], "sum"),
        }
        
    def build(self, bdo_result: BDOParseResult) -> Path:
        """
        Build new Management Accounts file.
        
        Args:
            bdo_result: Parsed BDO data for current quarter
            
        Returns:
            Path to generated file
        """
        logger.info(f"Building Management Accounts for {self.config.quarter}")
        
        # Load previous quarter file
        self.workbook = openpyxl.load_workbook(self.previous_path)
        logger.info(f"Loaded previous file with {len(self.workbook.sheetnames)} sheets")
        
        # Step 1: Clone and update BDO sheet (preserves row structure!)
        self._clone_and_update_bdo_sheet(bdo_result)
        
        # Step 2: Update summary sheet with new columns
        self._update_summary_sheet(bdo_result)
        
        # Step 3: Update all date references throughout the workbook
        self._update_date_references()
        
        # Step 4: Validate calculations
        self._validate_calculations(bdo_result)
        
        # Save output
        self.workbook.save(self.output_path)
        logger.info(f"Saved to {self.output_path}")
        
        return self.output_path
    
    def _clone_and_update_bdo_sheet(self, bdo_result: BDOParseResult):
        """
        Clone the previous BDO sheet and add new quarter column.
        
        CRITICAL: This approach preserves the exact row structure of the BDO sheet,
        ensuring all formula references in Management Cijfers remain valid.
        
        The previous BDO sheet has structure:
        - Row 1: Headers
        - Row 2: Empty
        - Row 3+: Category headers (bold) and account rows
        - Columns: A=Code, B=Name, C=Opening Balance, D-F=Previous quarters, G=Last quarter, H=Saldi
        
        For the new sheet:
        - Copy entire structure
        - Column G becomes previous quarter data
        - Insert new column G for current quarter mutations
        - Column H (now I) becomes Saldi = SUM(C:H)
        """
        sheet_name = self.config.bdo_sheet_name
        
        # Find the most recent BDO sheet to clone
        bdo_sheets = [name for name in self.workbook.sheetnames if name.startswith('BDO')]
        if not bdo_sheets:
            logger.error("No existing BDO sheet found to clone!")
            raise ValueError("No BDO sheet found in previous Management Accounts file")
        
        prev_bdo_name = bdo_sheets[-1]  # Most recent
        prev_bdo = self.workbook[prev_bdo_name]
        logger.info(f"Cloning BDO sheet from: {prev_bdo_name}")
        
        # Check if new sheet already exists
        if sheet_name in self.workbook.sheetnames:
            logger.warning(f"Sheet {sheet_name} already exists, will remove and recreate")
            del self.workbook[sheet_name]
        
        # Copy the sheet
        new_sheet = self.workbook.copy_worksheet(prev_bdo)
        new_sheet.title = sheet_name
        
        # Find position to move the sheet (after the last BDO sheet)
        bdo_sheets = [name for name in self.workbook.sheetnames if name.startswith('BDO')]
        last_bdo_idx = self.workbook.sheetnames.index(bdo_sheets[-1])
        # Move is done automatically by copy_worksheet
        
        # Build account code to row mapping from the cloned sheet
        self._build_account_row_map(new_sheet)
        
        # Determine column structure
        # Find the "Saldi" column (last data column with values)
        saldi_col = self._find_saldi_column(new_sheet)
        mutations_col = saldi_col - 1  # Last mutations column
        
        logger.info(f"Current structure: Mutations col={mutations_col}, Saldi col={saldi_col}")
        
        # Insert new column for Q3 mutations (before Saldi column)
        new_sheet.insert_cols(saldi_col)
        new_mutations_col = saldi_col
        new_saldi_col = saldi_col + 1
        
        # Update header for new mutations column
        period_str = self.config.period_end.strftime('%d-%m-%Y')
        new_sheet.cell(row=1, column=new_mutations_col).value = f"Mutaties {self.config.quarter}"
        new_sheet.cell(row=1, column=new_mutations_col).font = Font(bold=True)
        new_sheet.cell(row=1, column=new_mutations_col).alignment = Alignment(horizontal='center')
        
        # Update Saldi column header
        new_sheet.cell(row=1, column=new_saldi_col).value = f"Saldi per {period_str}"
        new_sheet.cell(row=1, column=new_saldi_col).font = Font(bold=True)
        new_sheet.cell(row=1, column=new_saldi_col).alignment = Alignment(horizontal='center')
        
        # Copy column width from previous mutations column
        prev_mutations_letter = get_column_letter(mutations_col)
        new_mutations_letter = get_column_letter(new_mutations_col)
        if prev_mutations_letter in new_sheet.column_dimensions:
            new_sheet.column_dimensions[new_mutations_letter].width = \
                new_sheet.column_dimensions[prev_mutations_letter].width
        
        # Populate the new mutations column with BDO data
        accounts_updated = 0
        for row_idx in range(2, new_sheet.max_row + 1):
            code_cell = new_sheet.cell(row=row_idx, column=1)
            code = str(code_cell.value).strip() if code_cell.value else ''
            
            # Skip if not an account code (category headers, empty rows)
            if not code or not code[0].isdigit():
                # For non-account rows, check if there's a formula in Saldi column
                # and update it to include the new column
                self._update_saldi_formula(new_sheet, row_idx, new_saldi_col, new_mutations_col)
                continue
            
            # Find the account in BDO result
            if code in bdo_result.accounts:
                entry = bdo_result.accounts[code]
                
                # Calculate mutations for this quarter
                mutations = entry.closing_balance - entry.opening_balance
                
                # Set mutations value
                new_sheet.cell(row=row_idx, column=new_mutations_col).value = mutations
                new_sheet.cell(row=row_idx, column=new_mutations_col).number_format = '#,##0.00'
                
                # Update Saldi column to be closing balance (or update formula)
                saldi_cell = new_sheet.cell(row=row_idx, column=new_saldi_col)
                if saldi_cell.value and isinstance(saldi_cell.value, str) and saldi_cell.value.startswith('='):
                    # Update formula to include new column
                    self._update_saldi_formula(new_sheet, row_idx, new_saldi_col, new_mutations_col)
                else:
                    # Direct value
                    saldi_cell.value = entry.closing_balance
                    saldi_cell.number_format = '#,##0.00'
                
                accounts_updated += 1
            else:
                # Account not in BDO result - set mutations to 0
                new_sheet.cell(row=row_idx, column=new_mutations_col).value = 0
                new_sheet.cell(row=row_idx, column=new_mutations_col).number_format = '#,##0.00'
                
                # Update Saldi formula if present
                self._update_saldi_formula(new_sheet, row_idx, new_saldi_col, new_mutations_col)
        
        logger.info(f"Created BDO sheet '{sheet_name}' with {accounts_updated} accounts updated")
        logger.info(f"New mutations column: {get_column_letter(new_mutations_col)}, Saldi column: {get_column_letter(new_saldi_col)}")
    
    def _build_account_row_map(self, sheet):
        """Build mapping of account codes to row numbers."""
        self._account_row_map = {}
        for row_idx in range(1, sheet.max_row + 1):
            code = sheet.cell(row=row_idx, column=1).value
            if code and isinstance(code, str) and code.strip() and code.strip()[0].isdigit():
                self._account_row_map[code.strip()] = row_idx
        logger.info(f"Built account row map with {len(self._account_row_map)} accounts")
    
    def _find_saldi_column(self, sheet) -> int:
        """Find the Saldi (closing balance) column in BDO sheet."""
        # Look for "Saldi" in row 1
        for col_idx in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col_idx).value
            if cell_value and 'Saldi' in str(cell_value):
                return col_idx
        
        # Fallback: return last column with data
        return sheet.max_column
    
    def _update_saldi_formula(self, sheet, row_idx: int, saldi_col: int, new_mutations_col: int):
        """Update Saldi formula to include the new mutations column."""
        saldi_cell = sheet.cell(row=row_idx, column=saldi_col)
        
        if not saldi_cell.value:
            return
            
        if isinstance(saldi_cell.value, str) and saldi_cell.value.startswith('='):
            # It's a formula - need to update it
            old_formula = saldi_cell.value
            
            # The formula typically references a range like C6:G6
            # We need to extend it to include the new column
            import re
            
            # Pattern to match SUM ranges like SUM(C6:G6) or SUM(C6:H6)
            sum_match = re.search(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', old_formula, re.IGNORECASE)
            if sum_match:
                start_col_letter, start_row, end_col_letter, end_row = sum_match.groups()
                # Extend the end column to include the new mutations column
                new_end_letter = get_column_letter(new_mutations_col)
                new_formula = f"=SUM({start_col_letter}{start_row}:{new_end_letter}{end_row})"
                saldi_cell.value = new_formula
    
    def _update_summary_sheet(self, bdo_result: BDOParseResult):
        """
        Update Management Cijfers summary sheet with new quarter column.
        
        Expected structure for Q3 2025:
        - Column Z: Q2 2025 (existing)
        - Column AA: Q3 2025 (NEW - formulas reference BDO column G/mutations)
        - Column AB: LTM Q3 2025 (NEW - highlighted blue, formulas reference BDO column H/saldi)
        - Column AC: Empty spacer
        - Column AD: Commentary
        
        The LTM column should be highlighted blue, not the quarterly column.
        """
        # Find the summary sheet
        summary_sheets = [name for name in self.workbook.sheetnames 
                         if 'Management Cijfers' in name]
        
        if not summary_sheets:
            summary_sheets = [name for name in self.workbook.sheetnames 
                             if 'Cijfers' in name and 'QSP' in name]
        
        if not summary_sheets:
            logger.error("No summary sheet found!")
            raise ValueError("Management Cijfers sheet not found")
        
        old_summary_name = summary_sheets[-1]
        summary_sheet = self.workbook[old_summary_name]
        
        # Rename sheet to new quarter
        new_summary_name = self.config.summary_sheet_name
        if old_summary_name != new_summary_name:
            summary_sheet.title = new_summary_name
        
        logger.info(f"Updating summary sheet: {new_summary_name}")
        
        # Find the current LTM column (row 22 has P&L headers)
        ltm_col = self._find_ltm_column(summary_sheet)
        if not ltm_col:
            logger.warning("LTM column not found, using last data column")
            ltm_col = self._find_last_data_column(summary_sheet) + 1
        
        logger.info(f"Found LTM column at: {ltm_col} ({get_column_letter(ltm_col)})")
        
        # The current LTM column will become the new Q3 column
        # We insert a new column after it for the new LTM
        q3_data_col = ltm_col
        new_ltm_col = ltm_col + 1
        
        # Find Commentary column to know if we need to shift
        commentary_col = None
        for col_idx in range(ltm_col + 1, min(ltm_col + 5, summary_sheet.max_column + 5)):
            cell = summary_sheet.cell(row=22, column=col_idx)
            if cell.value and 'Commentary' in str(cell.value):
                commentary_col = col_idx
                break
            cell2 = summary_sheet.cell(row=2, column=col_idx)
            if cell2.value and 'Commentary' in str(cell2.value):
                commentary_col = col_idx
                break
        
        # Insert a new column for LTM Q3
        logger.info(f"Inserting new column at position {new_ltm_col} for LTM")
        summary_sheet.insert_cols(new_ltm_col)
        
        # Copy formatting from previous LTM column to new LTM column
        self._copy_column_formatting(summary_sheet, q3_data_col, new_ltm_col)
        
        # Get the new BDO sheet name and column positions
        bdo_sheet_name = self.config.bdo_sheet_name
        
        # Find mutations and saldi columns in the new BDO sheet
        bdo_sheet = self.workbook[bdo_sheet_name]
        bdo_saldi_col = self._find_saldi_column(bdo_sheet)
        bdo_mutations_col = bdo_saldi_col - 1
        
        bdo_mutations_letter = get_column_letter(bdo_mutations_col)
        bdo_saldi_letter = get_column_letter(bdo_saldi_col)
        
        logger.info(f"BDO sheet columns: Mutations={bdo_mutations_letter}, Saldi={bdo_saldi_letter}")
        
        # Update Q3 column (was LTM Q2) - formulas reference BDO mutations column
        self._update_column_formulas(summary_sheet, q3_data_col, bdo_sheet_name, bdo_mutations_letter)
        
        # Update new LTM column - formulas reference BDO saldi column
        self._update_column_formulas(summary_sheet, new_ltm_col, bdo_sheet_name, bdo_saldi_letter)
        
        # Update headers
        # Row 22: P&L headers
        summary_sheet.cell(row=22, column=q3_data_col).value = self.config.quarter
        summary_sheet.cell(row=22, column=q3_data_col).font = Font(bold=True)
        # Remove blue highlighting from Q3 column
        summary_sheet.cell(row=22, column=q3_data_col).fill = PatternFill()  # Clear fill
        
        summary_sheet.cell(row=22, column=new_ltm_col).value = f"LTM {self.config.quarter}"
        summary_sheet.cell(row=22, column=new_ltm_col).font = Font(bold=True)
        # Apply blue highlighting to LTM column
        blue_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        summary_sheet.cell(row=22, column=new_ltm_col).fill = blue_fill
        
        # Row 2: Balance sheet date header
        summary_sheet.cell(row=2, column=q3_data_col).value = self.config.period_end
        summary_sheet.cell(row=2, column=q3_data_col).number_format = 'YYYY-MM-DD'
        summary_sheet.cell(row=2, column=q3_data_col).font = Font(bold=True)
        
        # LTM column in row 2 might not have a header, but add date if needed
        summary_sheet.cell(row=2, column=new_ltm_col).value = self.config.period_end
        summary_sheet.cell(row=2, column=new_ltm_col).number_format = 'YYYY-MM-DD'
        summary_sheet.cell(row=2, column=new_ltm_col).font = Font(bold=True)
        
        # Apply blue fill to entire LTM column (for data rows)
        for row_idx in range(3, summary_sheet.max_row + 1):
            cell = summary_sheet.cell(row=row_idx, column=new_ltm_col)
            if cell.value is not None:
                cell.fill = blue_fill
        
        # Update title
        title_cell = summary_sheet.cell(row=1, column=1)
        title_cell.value = f"Management Accounts QSP ESS B.V. - {self.config.quarter}"
        
        logger.info(f"Updated summary sheet: Q3 col={get_column_letter(q3_data_col)}, LTM col={get_column_letter(new_ltm_col)}")
    
    def _update_column_formulas(self, sheet, col_idx: int, bdo_sheet_name: str, bdo_col_letter: str):
        """
        Update formulas in a column to reference the new BDO sheet.
        
        Preserves the row references (semantic mapping through row structure)
        but updates the sheet name and column letter.
        """
        import re
        
        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            
            if not cell.value:
                continue
                
            if isinstance(cell.value, str) and cell.value.startswith('='):
                old_formula = cell.value
                
                # Pattern to match BDO sheet references like 'BDO - Q2-25'!G73
                # Replace sheet name and column letter, keep row number
                pattern = r"'BDO[^']*'!([A-Z]+)(\d+)"
                
                def replace_ref(match):
                    old_col = match.group(1)
                    row_num = match.group(2)
                    return f"'{bdo_sheet_name}'!{bdo_col_letter}{row_num}"
                
                new_formula = re.sub(pattern, replace_ref, old_formula)
                
                if new_formula != old_formula:
                    cell.value = new_formula
    
    def _find_ltm_column(self, sheet) -> int:
        """Find the LAST LTM (Last Twelve Months) column in the P&L section."""
        ltm_col = None
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=22, column=col_idx)
            if cell.value and 'LTM' in str(cell.value):
                ltm_col = col_idx  # Keep updating to get the last one
        
        if ltm_col:
            cell_value = sheet.cell(row=22, column=ltm_col).value
            logger.info(f"Found last LTM column at {ltm_col}: {cell_value}")
        return ltm_col
    
    def _find_last_data_column(self, sheet) -> int:
        """Find the last DATE column in row 2 (header row)."""
        last_date_col = 1
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=2, column=col_idx)
            if cell.value is not None:
                if isinstance(cell.value, datetime):
                    last_date_col = col_idx
                elif isinstance(cell.value, str):
                    if any(skip in cell.value.lower() for skip in ['balance', 'commentary', 'sheet']):
                        continue
        
        return last_date_col
    
    def _copy_column_formatting(self, sheet, source_col: int, target_col: int):
        """Copy column formatting from source to target."""
        source_letter = get_column_letter(source_col)
        target_letter = get_column_letter(target_col)
        
        # Copy column width
        if source_letter in sheet.column_dimensions:
            sheet.column_dimensions[target_letter].width = sheet.column_dimensions[source_letter].width
        
        # Copy cell formatting for all rows with data
        for row_idx in range(1, sheet.max_row + 1):
            source_cell = sheet.cell(row=row_idx, column=source_col)
            target_cell = sheet.cell(row=row_idx, column=target_col)
            
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.border = copy(source_cell.border)
                # Don't copy fill for now - we'll apply blue fill to LTM later
    
    def _update_date_references(self):
        """
        Update all date references throughout the workbook.
        
        This addresses Issue 4 from feedback: dates like "Per 30-6-2025" 
        should be updated to "Per 30-9-2025" for Q3.
        """
        old_date_patterns = self._generate_date_patterns(self._get_previous_period_end())
        new_date_str = self.config.period_end.strftime('%d-%m-%Y')
        new_date_patterns = self._generate_date_patterns(self.config.period_end)
        
        # Update the summary sheet
        summary_sheets = [name for name in self.workbook.sheetnames 
                         if 'Management Cijfers' in name or 'Cijfers' in name]
        
        for sheet_name in summary_sheets:
            sheet = self.workbook[sheet_name]
            
            for row_idx in range(1, sheet.max_row + 1):
                for col_idx in range(1, min(sheet.max_column + 1, 10)):  # First 10 columns
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    if cell.value and isinstance(cell.value, str):
                        original_value = cell.value
                        new_value = original_value
                        
                        # Replace date patterns
                        for old_pattern, new_pattern in zip(old_date_patterns, new_date_patterns):
                            if old_pattern in new_value:
                                new_value = new_value.replace(old_pattern, new_pattern)
                        
                        if new_value != original_value:
                            cell.value = new_value
                            logger.debug(f"Updated date in {sheet_name}!{get_column_letter(col_idx)}{row_idx}: {original_value} -> {new_value}")
    
    def _get_previous_period_end(self) -> datetime:
        """Get the previous quarter's period end date."""
        month = self.config.period_end.month
        year = self.config.period_end.year
        
        # Go back 3 months
        if month <= 3:
            prev_month = 12
            prev_year = year - 1
        else:
            prev_month = month - 3
            prev_year = year
        
        # Determine day
        if prev_month in [1, 3, 5, 7, 8, 10, 12]:
            day = 31
        elif prev_month == 2:
            day = 28
        else:
            day = 30
        
        return datetime(prev_year, prev_month, day)
    
    def _generate_date_patterns(self, date: datetime) -> List[str]:
        """Generate various date patterns for replacement."""
        return [
            date.strftime('%d-%m-%Y'),      # 30-09-2025
            date.strftime('%d-%m-%y'),       # 30-09-25
            date.strftime('%d/%m/%Y'),       # 30/09/2025
            date.strftime('%-d-%-m-%Y'),     # 30-9-2025 (no leading zeros)
            f"Per {date.strftime('%d-%m-%Y')}",  # Per 30-09-2025
            f"Per {date.strftime('%-d-%-m-%Y')}", # Per 30-9-2025
            date.strftime('%d %B %Y'),       # 30 September 2025
            date.strftime('%B %d, %Y'),      # September 30, 2025
        ]
    
    def _validate_calculations(self, bdo_result: BDOParseResult):
        """
        Validate that calculations are correct.
        
        Key validation (per Issue 3.2): Total (Equity Movement) = Direct Result
        """
        # Calculate equity movement using configurable equity prefixes
        equity_start = 0.0
        equity_end = 0.0
        
        for code, entry in bdo_result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            if first_two in self.equity_prefixes:
                equity_start += entry.opening_balance
                equity_end += entry.closing_balance
        
        equity_movement = equity_end - equity_start
        
        # Calculate direct result from P&L (mutations during the period)
        # Direct result = Income - Expenses (for the quarter)
        income_mutations = 0.0
        expense_mutations = 0.0
        result_mutations = 0.0
        
        for code, entry in bdo_result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            first_digit = code[0] if code else ''
            
            # Check if it's a result account (95xxxxx)
            if first_two == '95':
                result_mutations += (entry.closing_balance - entry.opening_balance)
            # Check income prefixes
            elif any(first_digit == prefix[0] for prefix in self.income_prefixes):
                income_mutations += (entry.closing_balance - entry.opening_balance)
            # Check expense prefixes
            elif any(first_digit == prefix[0] for prefix in self.expense_prefixes):
                expense_mutations += (entry.closing_balance - entry.opening_balance)
        
        # Direct result is the sum of P&L mutations
        direct_result = result_mutations if result_mutations != 0 else (income_mutations + expense_mutations)
        
        # Check within tolerance
        tolerance = 1000.00  # EUR 1000 tolerance for rounding
        if abs(equity_movement - direct_result) > tolerance:
            logger.warning(
                f"VALIDATION WARNING: Equity movement ({equity_movement:,.2f}) != Direct Result ({direct_result:,.2f})"
            )
            logger.warning(f"  Difference: {abs(equity_movement - direct_result):,.2f}")
        else:
            logger.info(f"✓ Equity movement validation passed: {equity_movement:,.2f} ≈ {direct_result:,.2f}")

