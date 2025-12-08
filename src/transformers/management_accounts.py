"""
Management Accounts Transformer

Builds the Management Accounts Excel file by:
1. Copying previous quarter structure
2. Adding new BDO sheet for current quarter
3. Updating Management Cijfers summary with new column
4. Recalculating formulas programmatically
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, PatternFill
from copy import copy
import pandas as pd
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
    
    This builder:
    1. Copies the previous quarter file
    2. Adds new BDO sheet with current quarter data
    3. Adds new column to summary sheet
    4. Updates formulas to reference new BDO sheet
    
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
                accounts = item_config.get('accounts', [])
                calc_type = item_config.get('calc_type', 'sum')
                mappings[label] = (accounts, calc_type)
        
        # Process P&L mappings
        if 'profit_loss' in config:
            for label, item_config in config['profit_loss'].items():
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
        
        # Step 1: Add new BDO sheet
        self._add_bdo_sheet(bdo_result)
        
        # Step 2: Update summary sheet
        self._update_summary_sheet(bdo_result)
        
        # Step 3: Validate calculations
        self._validate_calculations(bdo_result)
        
        # Save output
        self.workbook.save(self.output_path)
        logger.info(f"Saved to {self.output_path}")
        
        return self.output_path
    
    def _add_bdo_sheet(self, bdo_result: BDOParseResult):
        """Add new BDO sheet with current quarter data."""
        sheet_name = self.config.bdo_sheet_name
        
        # Check if sheet already exists (shouldn't, but handle gracefully)
        if sheet_name in self.workbook.sheetnames:
            logger.warning(f"Sheet {sheet_name} already exists, will overwrite")
            del self.workbook[sheet_name]
        
        # Find position for new sheet (after last BDO sheet, before summary)
        bdo_sheets = [name for name in self.workbook.sheetnames if name.startswith('BDO')]
        if bdo_sheets:
            last_bdo_idx = self.workbook.sheetnames.index(bdo_sheets[-1])
            insert_idx = last_bdo_idx + 1
        else:
            insert_idx = len(self.workbook.sheetnames) - 1  # Before last sheet
        
        # Create new sheet
        new_sheet = self.workbook.create_sheet(sheet_name, insert_idx)
        
        # Copy formatting from previous BDO sheet
        if bdo_sheets:
            self._copy_sheet_formatting(self.workbook[bdo_sheets[-1]], new_sheet)
        
        # Populate with BDO data
        self._populate_bdo_sheet(new_sheet, bdo_result)
        
        logger.info(f"Added sheet '{sheet_name}' at position {insert_idx}")
    
    def _copy_sheet_formatting(self, source_sheet, target_sheet):
        """Copy column widths and basic formatting from source to target."""
        for col_idx, col_dim in source_sheet.column_dimensions.items():
            target_sheet.column_dimensions[col_idx].width = col_dim.width
            
        # Copy row heights for header rows
        for row_idx in range(1, 6):
            if row_idx in source_sheet.row_dimensions:
                target_sheet.row_dimensions[row_idx].height = source_sheet.row_dimensions[row_idx].height
    
    def _populate_bdo_sheet(self, sheet, bdo_result: BDOParseResult):
        """
        Populate BDO sheet with parsed account data.
        
        CRITICAL: The structure MUST match the original BDO sheets exactly:
        - Column A (1): Account Code
        - Column B (2): Account Name  
        - Column C (3): Opening Balance (Eindbalans previous period)
        - Columns D-F (4-6): Previous quarter mutations (can be empty)
        - Column G (7): Current Quarter Mutations
        - Column H (8): Closing Balance (Saldi = SUM of C:G)
        
        This is required because Management Cijfers formulas reference:
        - Column G for quarter data: ='BDO - Q3-25'!G73
        - Column H for LTM data: ='BDO - Q3-25'!H73
        """
        # Build period end date string
        period_end = bdo_result.period_end
        if hasattr(period_end, 'strftime'):
            period_str = period_end.strftime('%d-%m-%Y')
        else:
            period_str = str(period_end)
        
        # Header row - match original BDO sheet structure
        # Note: self.config.quarter is already "Q3 2025" format
        headers = {
            1: None,  # Account Code (no header in original)
            2: None,  # Account Name (no header in original)  
            3: f"Eindbalans per {period_str}",  # Opening balance
            4: None,  # Previous Q mutations (empty for new sheets)
            5: None,  # Previous Q mutations (empty for new sheets)
            6: None,  # Previous Q mutations (empty for new sheets)
            7: f"Mutaties {self.config.quarter}",  # Current quarter (e.g., "Mutaties Q3 2025")
            8: f"Saldi per {period_str}",  # Closing balance
        }
        
        for col_idx, header in headers.items():
            if header:
                cell = sheet.cell(row=1, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
        
        # Data rows - sorted by account code
        row_idx = 2
        for code in sorted(bdo_result.accounts.keys()):
            entry = bdo_result.accounts[code]
            
            # Column A: Account Code
            sheet.cell(row=row_idx, column=1, value=entry.code)
            
            # Column B: Account Name
            sheet.cell(row=row_idx, column=2, value=entry.name)
            
            # Column C: Opening Balance
            sheet.cell(row=row_idx, column=3, value=entry.opening_balance)
            
            # Columns D-F: Empty (previous quarter mutations not available)
            # These would need historical BDO data to populate
            
            # Column G: Current Quarter Mutations (closing - opening)
            mutations = entry.closing_balance - entry.opening_balance
            sheet.cell(row=row_idx, column=7, value=mutations)
            
            # Column H: Closing Balance (or use formula =SUM(C:G))
            # Using the actual value for reliability
            sheet.cell(row=row_idx, column=8, value=entry.closing_balance)
            
            # Number formatting for financial columns
            for col in [3, 7, 8]:
                cell = sheet.cell(row=row_idx, column=col)
                if cell.value is not None:
                    cell.number_format = '#,##0.00'
            
            row_idx += 1
        
        # Store row mapping for formula generation
        self._bdo_row_mapping = {}
        row_idx = 2
        for code in sorted(bdo_result.accounts.keys()):
            self._bdo_row_mapping[code] = row_idx
            row_idx += 1
        
        logger.info(f"Populated {row_idx - 2} account rows in 8-column BDO format")
    
    def _update_summary_sheet(self, bdo_result: BDOParseResult):
        """
        Update Management Cijfers summary sheet with new quarter column.
        
        The correct process (based on manual Q3 file analysis):
        1. Find the LTM column (e.g., "LTM Q2 2025" at col 27)
        2. The LTM column BECOMES the Q3 data column - update formulas to reference BDO - Q3-25!G
        3. Add a NEW LTM column after it (col 28) with formulas referencing BDO - Q3-25!H
        4. Update Balance Sheet row date header
        5. Update P&L section headers
        
        CRITICAL: Formulas must reference the correct BDO sheet and row numbers!
        """
        # Find or update the summary sheet
        summary_sheets = [name for name in self.workbook.sheetnames 
                         if 'Management Cijfers' in name]
        
        if not summary_sheets:
            summary_sheets = [name for name in self.workbook.sheetnames 
                             if 'Cijfers' in name and 'QSP' in name]
        
        if not summary_sheets:
            logger.error("No summary sheet found!")
            raise ValueError("Management Cijfers sheet not found")
        
        old_summary_name = summary_sheets[-1]
        old_summary = self.workbook[old_summary_name]
        
        # Rename sheet to new quarter
        new_summary_name = self.config.summary_sheet_name
        if old_summary_name != new_summary_name:
            old_summary.title = new_summary_name
        
        summary_sheet = self.workbook[new_summary_name]
        
        # Find the LTM column (contains "LTM" in header, usually row 22 for P&L section)
        ltm_col = self._find_ltm_column(summary_sheet)
        if not ltm_col:
            logger.warning("LTM column not found, falling back to last date column method")
            ltm_col = self._find_last_data_column(summary_sheet) + 1
        
        # The Q3 data goes into the current LTM column position
        q3_data_col = ltm_col
        # The new LTM goes into the next column
        new_ltm_col = ltm_col + 1
        
        logger.info(f"Q3 data column: {q3_data_col}, New LTM column: {new_ltm_col}")
        
        # Check if we need to shift Commentary
        commentary_col = None
        for col_idx in range(ltm_col + 1, summary_sheet.max_column + 5):
            cell = summary_sheet.cell(row=2, column=col_idx)
            if cell.value and 'Commentary' in str(cell.value):
                commentary_col = col_idx
                break
        
        # If new LTM column would overwrite Commentary, insert column
        if commentary_col and new_ltm_col >= commentary_col:
            logger.info(f"Inserting column at {new_ltm_col} to make room for LTM")
            summary_sheet.insert_cols(new_ltm_col)
        
        # Build BDO row mapping for formula generation
        bdo_sheet_name = self.config.bdo_sheet_name
        bdo_row_map = self._build_bdo_row_mapping(bdo_result)
        
        # Update Q3 data column (was LTM Q2) - formulas reference BDO column G
        self._update_column_with_formulas(summary_sheet, q3_data_col, bdo_sheet_name, 
                                          bdo_row_map, col_letter='G')
        
        # Update new LTM column - formulas reference BDO column H
        self._update_column_with_formulas(summary_sheet, new_ltm_col, bdo_sheet_name,
                                          bdo_row_map, col_letter='H')
        
        # Update headers
        # Balance sheet header (row 2) - date for Q3
        date_col = self._find_balance_sheet_date_column(summary_sheet)
        if date_col:
            # The date column for balance sheet might be separate from P&L columns
            # Find the column that had the previous date and update to Q3
            for col_idx in range(date_col, min(date_col + 5, summary_sheet.max_column + 1)):
                cell = summary_sheet.cell(row=2, column=col_idx)
                if isinstance(cell.value, datetime):
                    # This is Q2 date, need to add Q3 date
                    break
        
        # Add Q3 date in balance sheet row
        summary_sheet.cell(row=2, column=q3_data_col + 1).value = self.config.period_end
        summary_sheet.cell(row=2, column=q3_data_col + 1).number_format = 'YYYY-MM-DD'
        summary_sheet.cell(row=2, column=q3_data_col + 1).font = Font(bold=True)
        
        # Update P&L headers (row 22)
        # self.config.quarter is already "Q3 2025", so don't add extra "Q" prefix
        summary_sheet.cell(row=22, column=q3_data_col).value = self.config.quarter
        summary_sheet.cell(row=22, column=new_ltm_col).value = f"LTM {self.config.quarter}"
        
        logger.info(f"Set headers: Col {q3_data_col}='{self.config.quarter}', Col {new_ltm_col}='LTM {self.config.quarter}'")
        
        # Update title
        title_cell = summary_sheet.cell(row=1, column=1)
        title_cell.value = f"Management Accounts QSP ESS B.V. - {self.config.quarter}"
        
        logger.info(f"Updated summary sheet with Q3 formulas: {new_summary_name}")
    
    def _find_ltm_column(self, sheet) -> int:
        """Find the LAST LTM (Last Twelve Months) column in the P&L section."""
        # P&L headers are typically in row 22
        # We need the LAST LTM column (most recent quarter), not the first
        ltm_col = None
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=22, column=col_idx)
            if cell.value and 'LTM' in str(cell.value):
                ltm_col = col_idx  # Keep updating to get the last one
        
        if ltm_col:
            cell_value = sheet.cell(row=22, column=ltm_col).value
            logger.info(f"Found last LTM column at {ltm_col}: {cell_value}")
        return ltm_col
    
    def _find_balance_sheet_date_column(self, sheet) -> int:
        """Find the last date column in the Balance Sheet header (row 2)."""
        last_date_col = None
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=2, column=col_idx)
            if isinstance(cell.value, datetime):
                last_date_col = col_idx
        return last_date_col
    
    def _build_bdo_row_mapping(self, bdo_result: BDOParseResult) -> dict:
        """Build mapping of account codes to BDO sheet row numbers."""
        row_map = {}
        row_idx = 2  # Data starts at row 2 (row 1 is headers)
        for code in sorted(bdo_result.accounts.keys()):
            row_map[code] = row_idx
            row_idx += 1
        return row_map
    
    def _update_column_with_formulas(self, sheet, col_idx: int, bdo_sheet_name: str,
                                     bdo_row_map: dict, col_letter: str):
        """
        Update a column with formulas referencing the BDO sheet.
        
        The formulas follow the pattern: =-'BDO - Q3-25'!G76
        - Negative sign because BDO uses opposite sign convention
        - Column G for quarter data, Column H for LTM data
        """
        # Map Management Cijfers rows to account codes
        row_to_account = self._get_row_account_mapping(sheet)
        
        for row_idx, account_codes in row_to_account.items():
            if not account_codes:
                continue
            
            # Build formula for this row
            # If multiple accounts, sum them
            if len(account_codes) == 1:
                code = account_codes[0]
                if code in bdo_row_map:
                    bdo_row = bdo_row_map[code]
                    formula = f"=-'{bdo_sheet_name}'!{col_letter}{bdo_row}"
                    sheet.cell(row=row_idx, column=col_idx).value = formula
            else:
                # Multiple accounts - create SUM formula
                parts = []
                for code in account_codes:
                    if code in bdo_row_map:
                        bdo_row = bdo_row_map[code]
                        parts.append(f"'{bdo_sheet_name}'!{col_letter}{bdo_row}")
                if parts:
                    formula = f"=-({'+'.join(parts)})"
                    sheet.cell(row=row_idx, column=col_idx).value = formula
    
    def _get_row_account_mapping(self, sheet) -> dict:
        """
        Get mapping of Management Cijfers rows to account codes.
        
        This is derived from the line_item_mappings config.
        """
        row_to_accounts = {}
        
        # Build row label to row index mapping
        row_labels = {}
        for row_idx in range(1, sheet.max_row + 1):
            label = sheet.cell(row=row_idx, column=1).value
            if label:
                row_labels[str(label).strip()] = row_idx
        
        # Map labels to account codes
        for label, (account_codes, calc_type) in self.line_item_mappings.items():
            row_idx = row_labels.get(label)
            if row_idx:
                # Expand wildcard patterns
                expanded_codes = []
                for code_pattern in account_codes:
                    if code_pattern.endswith('*'):
                        prefix = code_pattern[:-1]
                        # Find matching codes from BDO result
                        # For now, just use the pattern as-is
                        expanded_codes.append(prefix)
                    else:
                        expanded_codes.append(code_pattern)
                row_to_accounts[row_idx] = expanded_codes
        
        return row_to_accounts
    
    def _find_last_data_column(self, sheet) -> int:
        """
        Find the last DATE column in row 2 (header row).
        
        The summary sheet has date headers (datetime) for each quarter,
        followed by a 'Commentary' column. We need to find the last date column
        to know where to insert the new quarter data.
        """
        last_date_col = 1
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=2, column=col_idx)
            if cell.value is not None:
                # Check if it's a date (datetime object or date-like string)
                if isinstance(cell.value, datetime):
                    last_date_col = col_idx
                elif isinstance(cell.value, str):
                    # Skip text columns like "Balance sheet", "Commentary"
                    if any(skip in cell.value.lower() for skip in ['balance', 'commentary', 'sheet']):
                        continue
                    # Try to parse as date
                    try:
                        from dateutil import parser
                        parser.parse(cell.value)
                        last_date_col = col_idx
                    except:
                        pass
        
        logger.debug(f"Last date column found: {last_date_col}")
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
                target_cell.fill = copy(source_cell.fill)
    
    def _populate_summary_values(self, sheet, col_idx: int, bdo_result: BDOParseResult):
        """Populate summary sheet column with calculated values."""
        # First, identify row labels in column A
        row_labels = {}
        for row_idx in range(1, sheet.max_row + 1):
            label = sheet.cell(row=row_idx, column=1).value
            if label:
                label_clean = str(label).strip()
                row_labels[label_clean] = row_idx
        
        # Map and populate values
        for label, (account_codes, calc_type) in self.line_item_mappings.items():
            if label not in row_labels:
                # Try fuzzy match
                matched_row = self._fuzzy_match_row(label, row_labels)
                if matched_row is None:
                    logger.warning(f"Row label not found: {label}")
                    continue
                row_idx = matched_row
            else:
                row_idx = row_labels[label]
            
            # Calculate value from BDO accounts
            value = self._calculate_value(bdo_result, account_codes, calc_type)
            
            # Set cell value
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.number_format = '#,##0.00'
            
            logger.debug(f"Set {label} = {value:,.2f}")
    
    def _fuzzy_match_row(self, target: str, row_labels: Dict[str, int]) -> Optional[int]:
        """Find closest matching row label."""
        target_lower = target.lower()
        for label, row_idx in row_labels.items():
            if target_lower in label.lower() or label.lower() in target_lower:
                return row_idx
        return None
    
    def _calculate_value(self, bdo_result: BDOParseResult, 
                         account_codes: List[str], calc_type: str) -> float:
        """Calculate value from BDO accounts."""
        total = 0.0
        
        for pattern in account_codes:
            for code, entry in bdo_result.accounts.items():
                if pattern.endswith('*'):
                    if code.startswith(pattern[:-1]):
                        total += entry.closing_balance
                elif code == pattern:
                    total += entry.closing_balance
        
        return total
    
    def _validate_calculations(self, bdo_result: BDOParseResult):
        """Validate that calculations are correct."""
        # Key validation: Total (Equity Movement) = Direct Result
        # This is the critical check from the manual process
        
        # Calculate equity movement using configurable equity prefixes
        equity_start = 0.0
        equity_end = 0.0
        
        for code, entry in bdo_result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            if first_two in self.equity_prefixes:
                equity_start += entry.opening_balance
                equity_end += entry.closing_balance
        
        equity_movement = equity_end - equity_start
        
        # Calculate direct result from P&L using configurable prefixes
        direct_result = 0.0
        for code, entry in bdo_result.accounts.items():
            first_digit = code[0] if code else ''
            # Check income prefixes
            if any(first_digit == prefix[0] for prefix in self.income_prefixes):
                direct_result += entry.closing_balance
            # Check expense prefixes
            elif any(first_digit == prefix[0] for prefix in self.expense_prefixes):
                direct_result += entry.closing_balance
        
        # Check within tolerance
        tolerance = 100.00  # EUR 100 tolerance for rounding
        if abs(equity_movement - direct_result) > tolerance:
            logger.warning(
                f"Equity movement ({equity_movement:,.2f}) != Direct Result ({direct_result:,.2f})"
            )
        else:
            logger.info("✓ Equity movement validation passed")

