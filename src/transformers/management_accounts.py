"""
Management Accounts Transformer

Builds the Management Accounts Excel file by:
1. Copying ALL data from the BDO file (Cijfers_QSP_) into a new sheet
2. Inserting a new column in Management Cijfers before the LTM column
3. Copying style and formula syntax from the previous quarter column
4. Updating formulas to reference the new BDO sheet with correct row numbers
5. Ensuring Management Cijfers sheet is always at the end

WORKFLOW:
- Copy all data from BDO file → Create new sheet "BDO - Q{quarter}-{YY}"
- Build account code to row mapping from new BDO sheet
- In Management Cijfers sheet, insert new column before LTM column
- Copy column Z style/formulas to new column
- Update formulas to reference new BDO sheet with validated row numbers
- Update SUM formulas in shifted columns
- Move Management Cijfers sheet to be the last sheet
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Font, Alignment, Border, PatternFill, Side
from openpyxl.styles.fills import FILL_SOLID
from copy import copy
import yaml
import re
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
    
    WORKFLOW:
    1. Copy ALL data from BDO file into new sheet named "BDO - Q{quarter}-{YY}"
    2. Build account code to row mapping from the new BDO sheet
    3. In Management Cijfers, insert new column before LTM column
    4. Copy previous quarter column (Z) style/formulas to new column
    5. Update formulas with correct BDO sheet name and validate row references
    6. Update SUM formulas in shifted LTM column
    7. Move Management Cijfers sheet to be last
    """
    
    def __init__(self, previous_file_path: str, output_path: str, config: ManagementAccountsConfig,
                 bdo_source_path: str = None, mappings_path: str = "config/line_item_mappings.yaml"):
        self.previous_path = Path(previous_file_path)
        self.output_path = Path(output_path)
        self.config = config
        self.bdo_source_path = Path(bdo_source_path) if bdo_source_path else None
        self.workbook = None
        self.mappings_path = Path(mappings_path)
        self.line_item_mappings = self._load_mappings()
        self.equity_prefixes = []
        self.income_prefixes = []
        self.expense_prefixes = []
        self._load_validation_config()
        
        # Track the previous BDO sheet name for formula updates
        self._prev_bdo_sheet_name = None
        
        # Account code to row mapping in the NEW BDO sheet
        self._new_bdo_row_map = {}
        
        # Account code to row mapping in the PREVIOUS BDO sheet (for row offset calculation)
        self._prev_bdo_row_map = {}
    
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
        
        # Get previous BDO sheet name and build row map
        bdo_sheets = [name for name in self.workbook.sheetnames if name.startswith('BDO')]
        if bdo_sheets:
            self._prev_bdo_sheet_name = bdo_sheets[-1]  # Most recent
            logger.info(f"Previous BDO sheet: {self._prev_bdo_sheet_name}")
            # Build row map from previous BDO sheet
            self._build_row_map(self.workbook[self._prev_bdo_sheet_name], self._prev_bdo_row_map)
        
        # Step 1: Copy ALL data from BDO file into new sheet
        self._copy_bdo_data_to_new_sheet(bdo_result)
        
        # Step 2: Build row map from the NEW BDO sheet
        new_bdo_sheet = self.workbook[self.config.bdo_sheet_name]
        self._build_row_map(new_bdo_sheet, self._new_bdo_row_map)
        logger.info(f"Built row map for new BDO sheet: {len(self._new_bdo_row_map)} accounts")
        
        # Step 3: Update summary sheet - insert new column, copy style/formulas
        self._update_summary_sheet()
        
        # Step 4: Update all date references throughout the workbook
        self._update_date_references()
        
        # Step 5: Move BDO sheet to second-to-last position
        self._move_bdo_sheet_to_second_last()
        
        # Step 6: Move Management Cijfers sheet to the end (after BDO)
        self._move_summary_sheet_to_end()
        
        # Step 7: Validate calculations
        self._validate_calculations(bdo_result)
        
        # Save output
        self.workbook.save(self.output_path)
        logger.info(f"Saved to {self.output_path}")
        
        return self.output_path
    
    def _build_row_map(self, sheet, row_map: dict):
        """
        Build mapping of account codes to row numbers in a BDO sheet.
        This is critical for validating that formula row references are correct.
        """
        row_map.clear()
        for row_idx in range(1, sheet.max_row + 1):
            code = sheet.cell(row=row_idx, column=1).value
            if code and isinstance(code, (str, int, float)):
                code_str = str(code).strip()
                if code_str and code_str[0].isdigit():
                    row_map[code_str] = row_idx
    
    def _copy_bdo_data_to_new_sheet(self, bdo_result: BDOParseResult):
        """
        Copy ALL data from the BDO file into a new sheet in Management Accounts.
        The new sheet is inserted BEFORE the Management Cijfers sheet.
        """
        new_sheet_name = self.config.bdo_sheet_name
        
        # Check if new sheet already exists
        if new_sheet_name in self.workbook.sheetnames:
            logger.warning(f"Sheet {new_sheet_name} already exists, will remove and recreate")
            del self.workbook[new_sheet_name]
        
        # Get source data from the BDO file
        if self.bdo_source_path and self.bdo_source_path.exists():
            bdo_wb = openpyxl.load_workbook(self.bdo_source_path)
            source_sheet = bdo_wb.active
            logger.info(f"Copying from BDO file: {self.bdo_source_path}")
        else:
            logger.warning("BDO source file not provided, will clone previous BDO sheet")
            self._clone_and_update_bdo_sheet(bdo_result)
            return
        
        # Find where to insert - before Management Cijfers sheet
        insert_index = len(self.workbook.sheetnames)  # Default to end
        for i, name in enumerate(self.workbook.sheetnames):
            if 'Management Cijfers' in name or 'Cijfers' in name:
                insert_index = i
                break
        
        # Create new sheet
        new_sheet = self.workbook.create_sheet(title=new_sheet_name, index=insert_index)
        
        # Copy ALL cells including values, formulas, and formatting
        for row in range(1, source_sheet.max_row + 1):
            for col in range(1, source_sheet.max_column + 1):
                source_cell = source_sheet.cell(row=row, column=col)
                target_cell = new_sheet.cell(row=row, column=col)
                
                # Copy value (including formulas)
                target_cell.value = source_cell.value
                
                # Copy formatting
                if source_cell.has_style:
                    target_cell.font = copy(source_cell.font)
                    target_cell.alignment = copy(source_cell.alignment)
                    target_cell.number_format = source_cell.number_format
                    target_cell.border = copy(source_cell.border)
                    target_cell.fill = copy(source_cell.fill)
        
        # Copy column widths
        for col_letter, dim in source_sheet.column_dimensions.items():
            new_sheet.column_dimensions[col_letter].width = dim.width
        
        # Copy row heights
        for row_idx, dim in source_sheet.row_dimensions.items():
            new_sheet.row_dimensions[row_idx].height = dim.height
        
        # Copy merged cells
        for merged_range in source_sheet.merged_cells.ranges:
            new_sheet.merge_cells(str(merged_range))
        
        logger.info(f"Copied BDO data to new sheet: {new_sheet_name}")
        logger.info(f"  Rows: {source_sheet.max_row}, Columns: {source_sheet.max_column}")
        
        bdo_wb.close()
    
    def _clone_and_update_bdo_sheet(self, bdo_result: BDOParseResult):
        """Fallback: Clone the previous BDO sheet and update with new data."""
        sheet_name = self.config.bdo_sheet_name
        
        bdo_sheets = [name for name in self.workbook.sheetnames if name.startswith('BDO')]
        if not bdo_sheets:
            raise ValueError("No BDO sheet found in previous Management Accounts file")
        
        prev_bdo_name = bdo_sheets[-1]
        prev_bdo = self.workbook[prev_bdo_name]
        logger.info(f"Cloning BDO sheet from: {prev_bdo_name}")
        
        if sheet_name in self.workbook.sheetnames:
            del self.workbook[sheet_name]
        
        # Find insert position (before Management Cijfers)
        insert_index = len(self.workbook.sheetnames)
        for i, name in enumerate(self.workbook.sheetnames):
            if 'Management Cijfers' in name:
                insert_index = i
                break
        
        new_sheet = self.workbook.copy_worksheet(prev_bdo)
        new_sheet.title = sheet_name
        
        # Move to correct position
        self.workbook.move_sheet(new_sheet, offset=insert_index - self.workbook.sheetnames.index(sheet_name))
        
        # Find Saldi column and insert new mutations column
        saldi_col = self._find_saldi_column(new_sheet)
        new_sheet.insert_cols(saldi_col)
        new_mutations_col = saldi_col
        new_saldi_col = saldi_col + 1
        
        # Update headers
        period_str = self.config.period_end.strftime('%d-%m-%Y')
        new_sheet.cell(row=1, column=new_mutations_col).value = f"Mutaties {self.config.quarter}"
        new_sheet.cell(row=1, column=new_mutations_col).font = Font(bold=True)
        new_sheet.cell(row=1, column=new_saldi_col).value = f"Saldi per {period_str}"
        new_sheet.cell(row=1, column=new_saldi_col).font = Font(bold=True)
        
        # Populate values
        for row_idx in range(2, new_sheet.max_row + 1):
            code = str(new_sheet.cell(row=row_idx, column=1).value or '').strip()
            if not code or not code[0].isdigit():
                continue
            
            if code in bdo_result.accounts:
                entry = bdo_result.accounts[code]
                mutations = entry.closing_balance - entry.opening_balance
                new_sheet.cell(row=row_idx, column=new_mutations_col).value = mutations
                new_sheet.cell(row=row_idx, column=new_mutations_col).number_format = '#,##0.00'
                new_sheet.cell(row=row_idx, column=new_saldi_col).value = entry.closing_balance
                new_sheet.cell(row=row_idx, column=new_saldi_col).number_format = '#,##0.00'
            else:
                new_sheet.cell(row=row_idx, column=new_mutations_col).value = 0
    
    def _find_saldi_column(self, sheet) -> int:
        """Find the Saldi (closing balance) column in BDO sheet."""
        for col_idx in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col_idx).value
            if cell_value and 'Saldi' in str(cell_value):
                return col_idx
        return sheet.max_column
    
    def _move_bdo_sheet_to_second_last(self):
        """Move the newly created BDO sheet to be the second-to-last sheet."""
        bdo_name = self.config.bdo_sheet_name
        if bdo_name in self.workbook.sheetnames:
            current_index = self.workbook.sheetnames.index(bdo_name)
            # Target is second-to-last (before Management Cijfers which will be moved to last)
            target_index = len(self.workbook.sheetnames) - 1
            if current_index != target_index:
                offset = target_index - current_index
                self.workbook.move_sheet(bdo_name, offset=offset)
                logger.info(f"Moved {bdo_name} to second-to-last (position {target_index + 1})")
    
    def _move_summary_sheet_to_end(self):
        """Move the Management Cijfers sheet to be the last sheet in the workbook."""
        summary_name = self.config.summary_sheet_name
        if summary_name in self.workbook.sheetnames:
            current_index = self.workbook.sheetnames.index(summary_name)
            target_index = len(self.workbook.sheetnames) - 1
            if current_index != target_index:
                self.workbook.move_sheet(summary_name, offset=target_index - current_index)
                logger.info(f"Moved {summary_name} to end (position {target_index + 1})")
    
    def _update_summary_sheet(self):
        """
        Update Management Cijfers summary sheet.
        
        IMPORTANT: Balance Sheet (rows 1-20) and P&L (rows 22+) have different structures:
        - Balance Sheet: Data only in sparse columns (C, E, ..., AA for LTM)
        - P&L: Data in consecutive quarterly columns (T, U, V, W, X, Y, Z for quarters, AA for LTM)
        
        Steps:
        1. Insert new column at LTM position (this shifts LTM to the right)
        2. For P&L section: Copy from previous quarter column (Z) to new column
        3. For Balance Sheet section: Copy from LTM column (now shifted right) to new column
        4. Update all formulas to reference new BDO sheet with correct row numbers
        5. Update column headers
        """
        # Find the summary sheet
        summary_sheets = [name for name in self.workbook.sheetnames 
                         if 'Management Cijfers' in name]
        
        if not summary_sheets:
            summary_sheets = [name for name in self.workbook.sheetnames 
                             if 'Cijfers' in name and 'QSP' in name]
        
        if not summary_sheets:
            raise ValueError("Management Cijfers sheet not found")
        
        old_summary_name = summary_sheets[-1]
        summary_sheet = self.workbook[old_summary_name]
        
        # Rename sheet to new quarter
        new_summary_name = self.config.summary_sheet_name
        if old_summary_name != new_summary_name:
            summary_sheet.title = new_summary_name
        
        logger.info(f"Updating summary sheet: {new_summary_name}")
        
        # Find the LTM column (this contains data for BOTH Balance Sheet and P&L)
        ltm_col = self._find_ltm_column(summary_sheet)
        if not ltm_col:
            ltm_col = self._find_last_data_column(summary_sheet) + 1
        
        prev_quarter_col = ltm_col - 1  # Previous quarter column (for P&L section)
        insert_position = ltm_col
        
        logger.info(f"LTM column: {get_column_letter(ltm_col)}, Previous quarter: {get_column_letter(prev_quarter_col)}")
        
        # Step 1: Insert new column at LTM position
        # This shifts the old LTM column to the right by 1
        summary_sheet.insert_cols(insert_position)
        
        new_quarter_col = insert_position  # The newly inserted blank column
        new_ltm_col = insert_position + 1  # Where old LTM data now is
        
        logger.info(f"After insert: New quarter col={get_column_letter(new_quarter_col)}, LTM col={get_column_letter(new_ltm_col)}")
        
        # Step 2: Populate the new quarter column with data
        # For P&L rows (22+): Copy from previous quarter column (prev_quarter_col)
        # For Balance Sheet rows (1-20): Copy from LTM column (new_ltm_col, which has the old LTM formulas)
        
        # First, copy Balance Sheet section from LTM (which has the data)
        self._copy_section_formulas(summary_sheet, new_ltm_col, new_quarter_col, 
                                     start_row=1, end_row=20, section_name="Balance Sheet")
        
        # Then, copy P&L section from previous quarter column
        self._copy_section_formulas(summary_sheet, prev_quarter_col, new_quarter_col, 
                                     start_row=21, end_row=summary_sheet.max_row, section_name="P&L")
        
        # Step 3: Update formulas in the new quarter column to reference new BDO sheet
        self._update_formulas_with_row_validation(summary_sheet, new_quarter_col)
        
        # Step 4: Update SUM formulas in the shifted LTM column
        self._update_internal_sum_formulas(summary_sheet, new_ltm_col, new_quarter_col)
        
        # Step 5: Update formulas in LTM column to reference new BDO sheet
        self._update_formulas_with_row_validation(summary_sheet, new_ltm_col)
        
        # Step 6: Update headers
        # Row 22: P&L header for new quarter column
        header_cell = summary_sheet.cell(row=22, column=new_quarter_col)
        header_cell.value = self.config.quarter
        header_cell.font = Font(bold=True)
        header_cell.fill = PatternFill()  # Clear blue fill
        
        # Row 2: Balance sheet date header
        date_cell = summary_sheet.cell(row=2, column=new_quarter_col)
        date_cell.value = self.config.period_end
        date_cell.number_format = 'YYYY-MM-DD'
        date_cell.font = Font(bold=True)
        
        # LTM column headers (both row 2 and row 22)
        blue_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        ltm_header = summary_sheet.cell(row=22, column=new_ltm_col)
        ltm_header.value = f"LTM {self.config.quarter}"
        ltm_header.font = Font(bold=True)
        ltm_header.fill = blue_fill
        
        # Update title
        summary_sheet.cell(row=1, column=1).value = f"Management Accounts QSP ESS B.V. - {self.config.quarter}"
        
        logger.info(f"Updated summary sheet columns: {get_column_letter(new_quarter_col)} and {get_column_letter(new_ltm_col)}")
    
    def _copy_section_formulas(self, sheet, source_col: int, target_col: int, 
                                start_row: int, end_row: int, section_name: str):
        """
        Copy formulas and formatting from source column to target column for a specific row range.
        """
        source_letter = get_column_letter(source_col)
        target_letter = get_column_letter(target_col)
        copied_count = 0
        
        for row_idx in range(start_row, end_row + 1):
            source_cell = sheet.cell(row=row_idx, column=source_col)
            target_cell = sheet.cell(row=row_idx, column=target_col)
            
            if source_cell.value is not None:
                value = source_cell.value
                
                # If it's a formula, update internal column references
                if isinstance(value, str) and value.startswith('='):
                    value = self._update_internal_column_refs(value, source_letter, target_letter)
                    copied_count += 1
                
                target_cell.value = value
            
            # Copy formatting
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.border = copy(source_cell.border)
        
        logger.info(f"Copied {section_name} section (rows {start_row}-{end_row}): {copied_count} formulas from {source_letter} to {target_letter}")
    
    def _find_ltm_column(self, sheet) -> int:
        """Find the LTM column in row 22."""
        for col_idx in range(sheet.max_column, 0, -1):
            cell = sheet.cell(row=22, column=col_idx)
            if cell.value and 'LTM' in str(cell.value):
                return col_idx
        return None
    
    def _find_last_data_column(self, sheet) -> int:
        """Find the last date column in row 2."""
        last_date_col = 1
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=2, column=col_idx)
            if isinstance(cell.value, datetime):
                last_date_col = col_idx
        return last_date_col
    
    def _copy_column_with_formulas(self, sheet, source_col: int, target_col: int):
        """
        Copy column including formulas and all formatting.
        Also updates internal column references from source to target column.
        """
        source_letter = get_column_letter(source_col)
        target_letter = get_column_letter(target_col)
        
        if source_letter in sheet.column_dimensions:
            sheet.column_dimensions[target_letter].width = sheet.column_dimensions[source_letter].width
        
        for row_idx in range(1, sheet.max_row + 1):
            source_cell = sheet.cell(row=row_idx, column=source_col)
            target_cell = sheet.cell(row=row_idx, column=target_col)
            
            if source_cell.value is not None:
                value = source_cell.value
                
                # If it's a formula, update internal column references
                if isinstance(value, str) and value.startswith('='):
                    value = self._update_internal_column_refs(value, source_letter, target_letter)
                
                target_cell.value = value
            
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.border = copy(source_cell.border)
        
        logger.info(f"Copied column {source_letter} to {target_letter} (with column reference updates)")
    
    def _update_internal_column_refs(self, formula: str, source_col: str, target_col: str) -> str:
        """
        Update internal column references in a formula.
        Changes references from source column to target column.
        
        Examples:
            =SUM(Z23:Z24) with source=Z, target=AA -> =SUM(AA23:AA24)
            =Z25/Z23-1 with source=Z, target=AA -> =AA25/AA23-1
        
        Does NOT update references to other sheets (like BDO sheets).
        """
        # Pattern to match column references that are NOT sheet references
        # Match: letter+number but not preceded by '!' (sheet reference)
        # and not preceded by another letter (multi-letter column)
        
        # First, protect sheet references by temporarily replacing them
        sheet_refs = []
        def protect_sheet_ref(match):
            idx = len(sheet_refs)
            sheet_refs.append(match.group(0))
            return f"__SHEET_REF_{idx}__"
        
        # Protect sheet references like 'Sheet Name'!A1 or Sheet!A1
        protected_formula = re.sub(r"'[^']+'![A-Z]+\$?\d+", protect_sheet_ref, formula)
        protected_formula = re.sub(r"[A-Za-z_][A-Za-z0-9_]*![A-Z]+\$?\d+", protect_sheet_ref, protected_formula)
        
        # Now update column references in the main sheet
        # Pattern: column letter followed by optional $ and row number
        # Only match the source column letter when it appears as a column reference
        pattern = rf'(\$?){source_col}(\$?\d+)'
        replacement = rf'\1{target_col}\2'
        updated_formula = re.sub(pattern, replacement, protected_formula)
        
        # Also handle column references in range expressions like SUM(Z3:Z18)
        # The above pattern should handle this, but let's also handle bare column refs
        
        # Restore sheet references
        for idx, ref in enumerate(sheet_refs):
            updated_formula = updated_formula.replace(f"__SHEET_REF_{idx}__", ref)
        
        return updated_formula
    
    def _update_formulas_with_row_validation(self, sheet, col_idx: int):
        """
        Update formulas to reference the new BDO sheet.
        Also validate that row references point to rows with actual data.
        If a row has no data, search nearby rows (+1, +2, +3) to find valid data.
        """
        new_bdo_name = self.config.bdo_sheet_name
        new_bdo_sheet = self.workbook[new_bdo_name]
        updated_count = 0
        adjusted_rows = 0
        
        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            
            if not cell.value or not isinstance(cell.value, str):
                continue
            
            if not cell.value.startswith('='):
                continue
            
            old_formula = cell.value
            new_formula = old_formula
            
            # Find all BDO references in the formula
            # Pattern: 'BDO - Q2-25'!G73 or 'BDO - Q2-25'!$G$73
            pattern = r"'(BDO[^']+)'!(\$?)([A-Z]+)(\$?)(\d+)"
            
            def replace_bdo_ref(match):
                nonlocal adjusted_rows
                old_sheet = match.group(1)
                dollar1 = match.group(2)
                col_letter = match.group(3)
                dollar2 = match.group(4)
                row_num = int(match.group(5))
                
                # Validate that the row has data in the new BDO sheet
                validated_row = self._find_valid_row(new_bdo_sheet, col_letter, row_num)
                
                if validated_row != row_num:
                    adjusted_rows += 1
                    logger.debug(f"Adjusted row reference: {col_letter}{row_num} -> {col_letter}{validated_row}")
                
                return f"'{new_bdo_name}'!{dollar1}{col_letter}{dollar2}{validated_row}"
            
            new_formula = re.sub(pattern, replace_bdo_ref, new_formula)
            
            if new_formula != old_formula:
                cell.value = new_formula
                updated_count += 1
        
        logger.info(f"Updated {updated_count} formulas in column {get_column_letter(col_idx)}, adjusted {adjusted_rows} row references")
    
    def _find_valid_row(self, bdo_sheet, col_letter: str, original_row: int, search_range: int = 5) -> int:
        """
        Find a valid row that has data in the specified column.
        
        Logic:
        1. If original row has ANY value (including 0), keep it - don't adjust
        2. If original row is empty/None, search nearby rows for data
        
        This ensures we only adjust row references when truly necessary.
        """
        col_idx = column_index_from_string(col_letter)
        
        # Check original row first - if it has ANY value (including 0), keep it
        original_value = bdo_sheet.cell(row=original_row, column=col_idx).value
        if original_value is not None and original_value != '':
            # Original row has data (could be 0, which is valid) - keep it
            return original_row
        
        # Original row is empty/None - search for valid data nearby
        # Search forward first (original+1, original+2, etc.)
        for offset in range(1, search_range + 1):
            new_row = original_row + offset
            if new_row <= bdo_sheet.max_row:
                value = bdo_sheet.cell(row=new_row, column=col_idx).value
                if value is not None and value != '':
                    logger.debug(f"Row {original_row} was empty, found data at row {new_row}")
                    return new_row
        
        # Search backward (original-1, original-2, etc.)
        for offset in range(1, search_range + 1):
            new_row = original_row - offset
            if new_row >= 1:
                value = bdo_sheet.cell(row=new_row, column=col_idx).value
                if value is not None and value != '':
                    logger.debug(f"Row {original_row} was empty, found data at row {new_row}")
                    return new_row
        
        # No valid row found - keep original (formula will show 0 or error)
        logger.warning(f"No valid data found near row {original_row} in column {col_letter}")
        return original_row
    
    def _update_internal_sum_formulas(self, sheet, target_col: int, new_quarter_col: int):
        """
        Update internal SUM formulas that reference columns within the sheet.
        When we insert a column, formulas like =SUM(AA3:AA18) need to be updated
        to =SUM(AB3:AB18) if the column shifted.
        
        Also handles formulas that need to include the new quarter column.
        """
        target_letter = get_column_letter(target_col)
        new_quarter_letter = get_column_letter(new_quarter_col)
        updated_count = 0
        
        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=target_col)
            
            if not cell.value or not isinstance(cell.value, str):
                continue
            
            if not cell.value.startswith('='):
                continue
            
            old_formula = cell.value
            new_formula = old_formula
            
            # Pattern to find SUM ranges within the same column
            # e.g., =SUM(AA3:AA18) or =SUM($AA$3:$AA$18)
            sum_pattern = r'SUM\((\$?)([A-Z]+)(\$?)(\d+):(\$?)([A-Z]+)(\$?)(\d+)\)'
            
            def fix_sum_range(match):
                d1, start_col, d2, start_row, d3, end_col, d4, end_row = match.groups()
                
                # If this SUM references the old LTM column letter (before shift),
                # update it to the new target column letter
                # The target_col now contains what was previously in ltm_col before the shift
                
                return f"SUM({d1}{target_letter}{d2}{start_row}:{d3}{target_letter}{d4}{end_row})"
            
            # Check if formula references the target column's own cells (self-referencing SUM)
            if f'SUM(' in old_formula.upper():
                # For SUM formulas, we need to ensure they reference the correct column
                # After column insertion, the column letter in the formula should match target_letter
                new_formula = re.sub(sum_pattern, fix_sum_range, new_formula, flags=re.IGNORECASE)
            
            # Also update any BDO references in LTM column
            if 'BDO' in new_formula:
                bdo_pattern = r"'(BDO[^']+)'"
                new_formula = re.sub(bdo_pattern, f"'{self.config.bdo_sheet_name}'", new_formula)
            
            if new_formula != old_formula:
                cell.value = new_formula
                updated_count += 1
        
        if updated_count > 0:
            logger.info(f"Updated {updated_count} internal formulas in column {target_letter}")
    
    def _update_date_references(self):
        """Update all date references throughout the workbook."""
        old_date_patterns = self._generate_date_patterns(self._get_previous_period_end())
        new_date_patterns = self._generate_date_patterns(self.config.period_end)
        
        summary_sheets = [name for name in self.workbook.sheetnames 
                         if 'Management Cijfers' in name or 'Cijfers' in name]
        
        for sheet_name in summary_sheets:
            sheet = self.workbook[sheet_name]
            
            for row_idx in range(1, sheet.max_row + 1):
                for col_idx in range(1, min(sheet.max_column + 1, 10)):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    if cell.value and isinstance(cell.value, str):
                        original = cell.value
                        new_value = original
                        
                        for old_pattern, new_pattern in zip(old_date_patterns, new_date_patterns):
                            if old_pattern in new_value:
                                new_value = new_value.replace(old_pattern, new_pattern)
                        
                        if new_value != original:
                            cell.value = new_value
    
    def _get_previous_period_end(self) -> datetime:
        """Get the previous quarter's period end date."""
        month = self.config.period_end.month
        year = self.config.period_end.year
        
        if month <= 3:
            prev_month = 12
            prev_year = year - 1
        else:
            prev_month = month - 3
            prev_year = year
        
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
            date.strftime('%d-%m-%Y'),
            date.strftime('%d-%m-%y'),
            date.strftime('%d/%m/%Y'),
            date.strftime('%-d-%-m-%Y'),
            f"Per {date.strftime('%d-%m-%Y')}",
            f"Per {date.strftime('%-d-%-m-%Y')}",
            date.strftime('%d %B %Y'),
            date.strftime('%B %d, %Y'),
        ]
    
    def _validate_calculations(self, bdo_result: BDOParseResult):
        """Validate that calculations are correct."""
        equity_start = 0.0
        equity_end = 0.0
        
        for code, entry in bdo_result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            if first_two in self.equity_prefixes:
                equity_start += entry.opening_balance
                equity_end += entry.closing_balance
        
        equity_movement = equity_end - equity_start
        
        income_mutations = 0.0
        expense_mutations = 0.0
        result_mutations = 0.0
        
        for code, entry in bdo_result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            first_digit = code[0] if code else ''
            
            if first_two == '95':
                result_mutations += (entry.closing_balance - entry.opening_balance)
            elif any(first_digit == prefix[0] for prefix in self.income_prefixes):
                income_mutations += (entry.closing_balance - entry.opening_balance)
            elif any(first_digit == prefix[0] for prefix in self.expense_prefixes):
                expense_mutations += (entry.closing_balance - entry.opening_balance)
        
        direct_result = result_mutations if result_mutations != 0 else (income_mutations + expense_mutations)
        
        tolerance = 1000.00
        if abs(equity_movement - direct_result) > tolerance:
            logger.warning(
                f"VALIDATION WARNING: Equity movement ({equity_movement:,.2f}) != Direct Result ({direct_result:,.2f})"
            )
            logger.warning(f"  Difference: {abs(equity_movement - direct_result):,.2f}")
        else:
            logger.info(f"✓ Equity movement validation passed: {equity_movement:,.2f} ≈ {direct_result:,.2f}")

