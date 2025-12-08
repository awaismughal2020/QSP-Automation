"""
Compliance Certificate Builder

Builds the Compliance Certificate Excel file by:
1. Copying previous quarter structure
2. Updating Q Management Accounts sheet with new data
3. Adding new quarterly column to Suppl. Calc and Impact Unit Sales
4. Updating formulas in SFA CC sheet
"""

from dataclasses import dataclass
from typing import Dict, Optional
from pathlib import Path
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, PatternFill
from copy import copy
from datetime import datetime
from loguru import logger

from ..parsers.bdo_parser import BDOParseResult


@dataclass
class ComplianceConfig:
    """Configuration for Compliance Certificate generation."""
    year: int
    quarter: int
    
    @property
    def quarter_str(self) -> str:
        return f"Q{self.quarter} {self.year}"
    
    @property
    def short_quarter(self) -> str:
        """Short format: 25Q3"""
        return f"{str(self.year)[-2:]}Q{self.quarter}"
    
    @property
    def prev_quarter_str(self) -> str:
        prev_q = self.quarter - 1
        prev_y = self.year
        if prev_q == 0:
            prev_q = 4
            prev_y -= 1
        return f"Q{prev_q} {prev_y}"


class ComplianceBuilder:
    """
    Builds updated Compliance Certificate workbook.
    
    The Compliance Certificate file contains:
    - SFA CC: Main compliance certificate with covenant calculations
    - Suppl. Calc: Supplementary calculations by quarter
    - Impact Unit Sales: Unit sales impact calculations
    - Qx Management Accounts: Management accounts data
    
    This builder:
    1. Copies the previous quarter file
    2. Updates Management Accounts sheet with new data
    3. Adds new column to Suppl. Calc and Impact Unit Sales
    4. Updates formulas in SFA CC
    """
    
    def __init__(self, previous_file_path: str, output_path: str, config: ComplianceConfig):
        self.previous_path = Path(previous_file_path)
        self.output_path = Path(output_path)
        self.config = config
        self.workbook = None
        
    def build(self, bdo_result: BDOParseResult, management_accounts_path: Optional[str] = None) -> Path:
        """
        Build new Compliance Certificate file.
        
        Args:
            bdo_result: Parsed BDO data for current quarter
            management_accounts_path: Path to Management Accounts file for data copy
            
        Returns:
            Path to generated file
        """
        logger.info(f"Building Compliance Certificate for {self.config.quarter_str}")
        
        # Load previous quarter file
        self.workbook = openpyxl.load_workbook(self.previous_path)
        logger.info(f"Loaded previous file with sheets: {self.workbook.sheetnames}")
        
        # Step 1: Update Management Accounts sheet
        self._update_management_accounts_sheet(bdo_result, management_accounts_path)
        
        # Step 2: Add new column to Suppl. Calc
        self._update_suppl_calc_sheet()
        
        # Step 3: Add new column to Impact Unit Sales
        self._update_impact_unit_sales_sheet()
        
        # Step 4: Update SFA CC formulas (if needed)
        self._update_sfa_cc_sheet()
        
        # Save output
        self.workbook.save(self.output_path)
        logger.info(f"Saved to {self.output_path}")
        
        return self.output_path
    
    def _update_management_accounts_sheet(self, bdo_result: BDOParseResult, 
                                          management_accounts_path: Optional[str]):
        """Update the Management Accounts sheet with new quarter data."""
        # The new sheet name format should match what the formulas expect: "Q3 Management Accounts"
        # NOT "Q3 2025 Management Accounts" as the formulas don't include the year
        new_ma_sheet_name = f"Q{self.config.quarter} Management Accounts"
        
        # Handle different naming patterns
        ma_sheets = [name for name in self.workbook.sheetnames if 'Management Accounts' in name]
        
        if ma_sheets:
            old_sheet_name = ma_sheets[0]  # Use the first one found
            logger.info(f"Found Management Accounts sheet: {old_sheet_name}")
            
            # Rename to new quarter format
            old_sheet = self.workbook[old_sheet_name]
            old_sheet.title = new_ma_sheet_name
            logger.info(f"Renamed sheet to: {new_ma_sheet_name}")
            
            # If we have a management accounts file, copy data from it
            if management_accounts_path:
                self._copy_ma_data(new_ma_sheet_name, management_accounts_path)
        else:
            logger.warning("No Management Accounts sheet found to update")
    
    def _copy_ma_data(self, target_sheet_name: str, source_path: str):
        """
        Copy Management Accounts data from external file.
        
        Creates a simplified structure for the Compliance Certificate:
        - Column A: Row labels (Balance sheet items, P&L items)
        - Column C: Q3 2025 values (the new quarter data)
        """
        try:
            source_wb = openpyxl.load_workbook(source_path, data_only=True)
            
            # Find the Management Cijfers sheet
            source_sheet_name = None
            for name in source_wb.sheetnames:
                if 'Management Cijfers' in name and f'Q{self.config.quarter}' in name:
                    source_sheet_name = name
                    break
            
            if not source_sheet_name:
                # Try without quarter in name
                for name in source_wb.sheetnames:
                    if 'Management Cijfers' in name:
                        source_sheet_name = name
                        break
            
            if not source_sheet_name:
                logger.warning(f"No Management Cijfers sheet found in {source_path}")
                return
            
            source_sheet = source_wb[source_sheet_name]
            target_sheet = self.workbook[target_sheet_name]
            
            # Find the Q3 2025 column in the source (should be the last date column)
            q3_col = None
            from datetime import datetime
            for col in range(1, source_sheet.max_column + 1):
                val = source_sheet.cell(row=2, column=col).value
                if isinstance(val, datetime) and val.year == 2025 and val.month == 9:
                    q3_col = col
                    break
            
            if not q3_col:
                # Find the last date column as fallback
                for col in range(source_sheet.max_column, 0, -1):
                    val = source_sheet.cell(row=2, column=col).value
                    if isinstance(val, datetime):
                        q3_col = col
                        break
            
            if not q3_col:
                logger.warning("Could not find Q3 2025 column in source")
                return
            
            logger.info(f"Found Q3 data in column {q3_col}")
            
            # Clear the target sheet first
            for row in range(1, target_sheet.max_row + 1):
                for col in range(1, target_sheet.max_column + 1):
                    target_sheet.cell(row=row, column=col).value = None
            
            # Create the simplified structure
            # The target structure should match the SFA CC formula references
            # SFA CC formulas reference rows like C22 for Gross Theoretical rental income
            # Target row = Source row - 1 (to account for header offset)
            
            # Row 1: Header with date
            target_sheet.cell(row=1, column=1).value = "Balance sheet"
            target_sheet.cell(row=1, column=3).value = datetime(2025, 9, 30)
            target_sheet.cell(row=1, column=3).number_format = 'YYYY-MM-DD'
            
            # Define P&L rows where signs need to be negated
            # These are income items that come from BDO as negative (credit) but should be positive
            # Row numbers refer to TARGET rows
            income_rows = {22, 24, 26, 27, 29}  # Gross rental income, service charges, etc.
            
            # Copy row labels (column A) and Q3 values (to column C)
            # Source has title at row 1, header at row 2, data starts at row 3
            # Target has header at row 1, data starts at row 2
            # So target row = source row - 1
            for source_row in range(2, min(source_sheet.max_row + 1, 130)):
                target_row = source_row - 1  # Shift up by 1
                
                # Copy label from column A
                label = source_sheet.cell(row=source_row, column=1).value
                if label:
                    # Skip the "Balance sheet" header row (we already have it at row 1)
                    if source_row == 2 and str(label).lower() == "balance sheet":
                        continue
                    
                    target_sheet.cell(row=target_row, column=1).value = label
                    
                    # Copy Q3 value to column C
                    value = source_sheet.cell(row=source_row, column=q3_col).value
                    if value is not None:
                        # For income items, negate the sign (BDO shows credits as negative)
                        if isinstance(value, (int, float)) and target_row in income_rows:
                            value = -value
                        
                        target_sheet.cell(row=target_row, column=3).value = value
                        if isinstance(value, (int, float)):
                            target_sheet.cell(row=target_row, column=3).number_format = '#,##0.00'
            
            logger.info(f"Created simplified Q3 Management Accounts from {source_sheet_name}")
            
        except Exception as e:
            logger.warning(f"Error copying MA data: {e}")
            import traceback
            logger.warning(traceback.format_exc())
    
    def _update_suppl_calc_sheet(self):
        """
        Update Suppl. Calc sheet by adding NEXT forecast quarter.
        
        The Q2 file already has 25Q3 forecasts. When building Q3:
        - We need to add 26Q3 (the next forecast quarter)
        - This shifts NTM column from S to T
        - This allows SFA CC formulas to reference the new NTM column
        """
        if 'Suppl. Calc' not in self.workbook.sheetnames:
            logger.warning("Suppl. Calc sheet not found")
            return
        
        sheet = self.workbook['Suppl. Calc']
        
        # Find the NTM (Next Twelve Months) column - this is where we insert
        ntm_col = None
        for col in range(1, sheet.max_column + 5):
            cell = sheet.cell(row=2, column=col)
            if cell.value and str(cell.value).upper() == 'NTM':
                ntm_col = col
                break
        
        if ntm_col is None:
            logger.warning("NTM column not found in Suppl. Calc")
            return
        
        # Calculate the next forecast quarter (current quarter + 4)
        # E.g., for Q3 2025, the next forecast to add is Q3 2026 (26Q3)
        next_forecast_quarter = f"{str(self.config.year + 1)[-2:]}Q{self.config.quarter}"
        
        # Check if the next forecast quarter already exists
        for col in range(1, sheet.max_column + 1):
            if sheet.cell(row=2, column=col).value == next_forecast_quarter:
                logger.info(f"Column {next_forecast_quarter} already exists")
                return
        
        # Insert a new column at the NTM position (shifts NTM to the right)
        logger.info(f"Inserting column for {next_forecast_quarter} at position {ntm_col} (NTM column)")
        sheet.insert_cols(ntm_col)
        
        # Add the new forecast quarter header
        header_cell = sheet.cell(row=2, column=ntm_col)
        header_cell.value = next_forecast_quarter
        header_cell.font = Font(bold=True)
        header_cell.alignment = Alignment(horizontal='center')
        
        # Add "Forecast" label
        type_cell = sheet.cell(row=3, column=ntm_col)
        type_cell.value = "Forecast"
        type_cell.alignment = Alignment(horizontal='center')
        
        # Copy formatting from previous column
        self._copy_column_formatting(sheet, ntm_col - 1, ntm_col)
        
        # Store the new NTM column position (shifted by 1)
        self._new_ntm_col = ntm_col + 1
        self._old_ntm_col = ntm_col  # This was the old NTM position
        
        logger.info(f"Added {next_forecast_quarter} column, NTM shifted to column {self._new_ntm_col}")
    
    def _update_impact_unit_sales_sheet(self):
        """
        Update Impact Unit Sales sheet similarly to Suppl. Calc.
        Add next forecast quarter and shift subsequent columns.
        """
        if 'Impact Unit Sales' not in self.workbook.sheetnames:
            logger.warning("Impact Unit Sales sheet not found")
            return
        
        sheet = self.workbook['Impact Unit Sales']
        
        # Find the last quarterly column (not NTM or totals)
        last_quarter_col = 3
        for col in range(4, sheet.max_column + 5):
            cell = sheet.cell(row=2, column=col)
            val = str(cell.value) if cell.value else ''
            # Look for quarter patterns like 25Q2, 26Q1, etc.
            if 'Q' in val and len(val) <= 5:
                last_quarter_col = col
        
        # Calculate next forecast quarter
        next_forecast_quarter = f"{str(self.config.year + 1)[-2:]}Q{self.config.quarter}"
        
        # Check if already exists
        for col in range(1, sheet.max_column + 1):
            if sheet.cell(row=2, column=col).value == next_forecast_quarter:
                logger.info(f"Column {next_forecast_quarter} already exists in Impact Unit Sales")
                return
        
        # Insert at position after last quarter
        new_col = last_quarter_col + 1
        sheet.insert_cols(new_col)
        
        # Add new quarter header
        header_cell = sheet.cell(row=2, column=new_col)
        header_cell.value = next_forecast_quarter
        header_cell.font = Font(bold=True)
        header_cell.alignment = Alignment(horizontal='center')
        
        # Copy column formatting
        self._copy_column_formatting(sheet, last_quarter_col, new_col)
        
        logger.info(f"Added {next_forecast_quarter} column to Impact Unit Sales at position {new_col}")
    
    def _update_sfa_cc_sheet(self):
        """
        Update SFA CC sheet with current quarter references.
        
        This includes:
        1. Updating sheet name references in formulas (Q2 Management Accounts → Q3 Management Accounts)
        2. Updating Suppl. Calc column references for NTM (column S → T after insert)
        3. Updating text references to quarter
        """
        if 'SFA CC' not in self.workbook.sheetnames:
            logger.warning("SFA CC sheet not found")
            return
        
        sheet = self.workbook['SFA CC']
        
        # Get the NTM column shift info
        # After inserting 26Q3, NTM shifted from S (19) to T (20)
        old_ntm_col_letter = None
        new_ntm_col_letter = None
        
        if 'Suppl. Calc' in self.workbook.sheetnames:
            suppl = self.workbook['Suppl. Calc']
            for col in range(1, suppl.max_column + 5):
                val = suppl.cell(row=2, column=col).value
                if val and str(val).upper() == 'NTM':
                    new_ntm_col_letter = get_column_letter(col)
                    # The old NTM was one column to the left (before we inserted)
                    old_ntm_col_letter = get_column_letter(col - 1)
                    break
            
            if old_ntm_col_letter and new_ntm_col_letter:
                logger.info(f"NTM column shift: {old_ntm_col_letter} → {new_ntm_col_letter}")
        
        # Build replacement patterns for Management Accounts sheet name
        prev_q = self.config.quarter - 1
        prev_y = self.config.year
        if prev_q == 0:
            prev_q = 4
            prev_y -= 1
        
        old_ma_patterns = [
            f"'Q{prev_q} {prev_y} Management Accounts'",
            f"'Q{prev_q} Management Accounts'",
            f"Q{prev_q} Management Accounts",
        ]
        new_ma_name = f"'Q{self.config.quarter} Management Accounts'"
        
        updated_count = 0
        formula_count = 0
        
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value and isinstance(cell.value, str):
                    original_value = cell.value
                    new_value = original_value
                    
                    # Check if it's a formula
                    if new_value.startswith('='):
                        # Update Management Accounts sheet references
                        for old_pattern in old_ma_patterns:
                            if old_pattern in new_value:
                                new_value = new_value.replace(old_pattern, new_ma_name)
                        
                        # Update Suppl. Calc NTM column references
                        if old_ntm_col_letter and new_ntm_col_letter and "'Suppl. Calc'" in new_value:
                            # Replace column letter references for NTM (e.g., S5 → T5)
                            # With $ sign
                            pattern = rf"'Suppl\. Calc'!\${old_ntm_col_letter}(\$?\d+)"
                            replacement = f"'Suppl. Calc'!${new_ntm_col_letter}\\1"
                            new_value = re.sub(pattern, replacement, new_value)
                            
                            # Without $ sign
                            pattern = rf"'Suppl\. Calc'!{old_ntm_col_letter}(\d+)"
                            replacement = f"'Suppl. Calc'!{new_ntm_col_letter}\\1"
                            new_value = re.sub(pattern, replacement, new_value)
                        
                        # Update Impact Unit Sales column references similarly
                        if old_ntm_col_letter and new_ntm_col_letter and "'Impact Unit Sales'" in new_value:
                            pattern = rf"'Impact Unit Sales'!\${old_ntm_col_letter}(\$?\d+)"
                            replacement = f"'Impact Unit Sales'!${new_ntm_col_letter}\\1"
                            new_value = re.sub(pattern, replacement, new_value)
                            
                            pattern = rf"'Impact Unit Sales'!{old_ntm_col_letter}(\d+)"
                            replacement = f"'Impact Unit Sales'!{new_ntm_col_letter}\\1"
                            new_value = re.sub(pattern, replacement, new_value)
                        
                        if new_value != original_value:
                            formula_count += 1
                    
                    # Update text references to quarter (non-formula)
                    else:
                        if self.config.prev_quarter_str in new_value:
                            new_value = new_value.replace(self.config.prev_quarter_str, self.config.quarter_str)
                    
                    if new_value != original_value:
                        cell.value = new_value
                        updated_count += 1
        
        logger.info(f"Updated {updated_count} cells in SFA CC ({formula_count} formulas)")
    
    def _copy_column_formatting(self, sheet, source_col: int, target_col: int):
        """Copy column formatting from source to target."""
        source_letter = get_column_letter(source_col)
        target_letter = get_column_letter(target_col)
        
        # Copy column width
        if source_letter in sheet.column_dimensions:
            sheet.column_dimensions[target_letter].width = sheet.column_dimensions[source_letter].width
        
        # Copy cell formatting for data rows
        for row in range(1, sheet.max_row + 1):
            source_cell = sheet.cell(row=row, column=source_col)
            target_cell = sheet.cell(row=row, column=target_col)
            
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.border = copy(source_cell.border)
                target_cell.fill = copy(source_cell.fill)

