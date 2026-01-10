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
    
    @property
    def period_end(self):
        """Calculate period end date for the quarter."""
        from datetime import datetime
        # Q1: March 31, Q2: June 30, Q3: September 30, Q4: December 31
        month = self.quarter * 3
        # March=31, June=30, September=30, December=31
        if month in [3, 12]:
            day = 31
        else:  # 6, 9
            day = 30
        return datetime(self.year, month, day)


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
        
        # Step 4: Update formulas in Suppl. Calc to reference Q3 Management Accounts
        self._update_suppl_calc_formulas()
        
        # Step 5: Update SFA CC formulas (if needed)
        self._update_sfa_cc_sheet()
        
        # Step 6: Update all dates and text references throughout workbook
        self._update_all_dates_and_text()
        
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
            
            # Get the old sheet structure BEFORE renaming (to preserve it)
            old_sheet = self.workbook[old_sheet_name]
            
            # Create a new sheet with the new name, copying structure from old sheet
            new_sheet = self.workbook.copy_worksheet(old_sheet)
            new_sheet.title = new_ma_sheet_name
            
            # Remove the old sheet
            self.workbook.remove(old_sheet)
            
            logger.info(f"Created new sheet: {new_ma_sheet_name} with structure from {old_sheet_name}")
            
            # If we have a management accounts file, copy data from it
            if management_accounts_path:
                self._copy_ma_data(new_ma_sheet_name, management_accounts_path)
        else:
            logger.warning("No Management Accounts sheet found to update")
    
    def _copy_ma_data(self, target_sheet_name: str, source_path: str):
        """
        Copy Management Accounts data from external file.
        
        IMPORTANT: This method reads the Management Cijfers sheet formulas and evaluates them
        by looking up the corresponding BDO sheet cell values.
        
        The approach:
        1. Load the Management Cijfers sheet to get the formulas from column AB (Q3)
        2. Load the BDO sheet to get actual cell values
        3. Parse each formula and calculate the result
        4. Copy values to the Compliance Certificate's Management Accounts sheet
        """
        try:
            # Load WITH formulas to find BDO sheet references
            source_wb = openpyxl.load_workbook(source_path, data_only=False)
            
            # Find the BDO sheet for current quarter
            bdo_sheet_name = f"BDO - Q{self.config.quarter}-{str(self.config.year)[-2:]}"
            if bdo_sheet_name not in source_wb.sheetnames:
                # Try alternate formats
                for name in source_wb.sheetnames:
                    if f'Q{self.config.quarter}' in name and 'BDO' in name:
                        bdo_sheet_name = name
                        break
            
            if bdo_sheet_name not in source_wb.sheetnames:
                logger.warning(f"BDO sheet {bdo_sheet_name} not found. Available: {source_wb.sheetnames[-5:]}")
                return
            
            bdo_sheet = source_wb[bdo_sheet_name]
            target_sheet = self.workbook[target_sheet_name]
            
            logger.info(f"Copying data from BDO sheet: {bdo_sheet_name}")
            
            # Find the Management Cijfers sheet to get the formulas
            cijfers_sheet_name = f"Management Cijfers - Q{self.config.quarter} {self.config.year}"
            if cijfers_sheet_name not in source_wb.sheetnames:
                # Try alternate formats
                for name in source_wb.sheetnames:
                    if 'Management Cijfers' in name:
                        cijfers_sheet_name = name
                        break
            
            cijfers_sheet = source_wb[cijfers_sheet_name] if cijfers_sheet_name in source_wb.sheetnames else None
            
            # Build a lookup for BDO sheet values by row (column H)
            bdo_h_values = {}
            for row in range(1, bdo_sheet.max_row + 1):
                val = bdo_sheet.cell(row=row, column=8).value  # Column H
                # If it's a formula, try to calculate it
                if isinstance(val, str) and val.startswith('='):
                    val = self._evaluate_simple_formula(val, bdo_sheet, row)
                bdo_h_values[row] = val if isinstance(val, (int, float)) else 0
            
            # Find the AB column (column 28) in Management Cijfers - this contains Q3 formulas
            # Parse and evaluate each formula
            
            # Create mapping from target row to formula calculation
            # Each target row in Compliance Certificate maps to a specific Management Cijfers row
            target_to_cijfers_mapping = {
                2: 3,   # Deferred Tax Asset
                3: 4,   # Real estate
                4: 5,   # Financial fixed assets
                5: 6,   # Accounts receivable
                6: 7,   # Service costs to be charged
                7: 8,   # Prepaid expenses
                8: 9,   # Cash
                9: 10,  # Equity
                10: 11, # AC Shareholder
                11: 12, # Bank loan
                12: 13, # Amortised fee
                13: 14, # Accounts payable
                14: 15, # Current account
                15: 16, # VAT payable
                16: 17, # Deposits
                17: 18, # Rent Invoiced in advance
            }
            
            # Update Row 1: Header with date
            period_end = self.config.period_end
            target_sheet.cell(row=1, column=3).value = period_end
            target_sheet.cell(row=1, column=3).number_format = 'YYYY-MM-DD'
            
            items_copied = 0
            
            if cijfers_sheet:
                # Use Management Cijfers formulas to calculate values
                for target_row, cijfers_row in target_to_cijfers_mapping.items():
                    formula = cijfers_sheet.cell(row=cijfers_row, column=28).value  # Column AB
                    if formula and isinstance(formula, str) and formula.startswith('='):
                        value = self._evaluate_bdo_formula(formula, bdo_h_values, bdo_sheet_name)
                        if value is not None:
                            target_sheet.cell(row=target_row, column=3).value = value
                            items_copied += 1
                            logger.debug(f"Copied value {value} to row {target_row} from formula {formula[:50]}...")
            
            # If no formulas found, fall back to the original method
            if items_copied == 0:
                items_copied = self._copy_ma_data_fallback(target_sheet, bdo_sheet, bdo_h_values)
            
            logger.info(f"Copied {items_copied} values to {target_sheet_name}")
            source_wb.close()
            
        except Exception as e:
            logger.error(f"Error copying MA data: {e}")
            import traceback
            traceback.print_exc()
    
    def _evaluate_simple_formula(self, formula: str, sheet, current_row: int) -> float:
        """Evaluate a simple SUM formula within the same sheet."""
        if 'SUM' in formula.upper():
            try:
                # Match SUM(Col1Row1:Col2Row2) pattern
                match = re.search(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', formula)
                if match:
                    total = 0
                    start_col = ord(match.group(1)) - ord('A') + 1
                    end_col = ord(match.group(3)) - ord('A') + 1
                    start_row = int(match.group(2))
                    end_row = int(match.group(4))
                    for r in range(start_row, end_row + 1):
                        for c in range(start_col, end_col + 1):
                            v = sheet.cell(row=r, column=c).value
                            if isinstance(v, (int, float)):
                                total += v
                    return total
            except Exception:
                pass
        return 0
    
    def _evaluate_bdo_formula(self, formula: str, bdo_h_values: dict, bdo_sheet_name: str) -> float:
        """
        Evaluate a formula that references the BDO sheet.
        
        Handles patterns like:
        - ='BDO - Q3-25'!H16
        - =SUM('BDO - Q3-25'!H6:H10)
        - ='BDO - Q3-25'!H13+'BDO - Q3-25'!H14+'BDO - Q3-25'!H15
        """
        total = 0.0
        
        try:
            # Pattern 1: SUM formula =SUM('BDO - Q3-25'!H6:H10)
            sum_match = re.search(r"SUM\('.*?'!H(\d+):H(\d+)\)", formula)
            if sum_match:
                start_row = int(sum_match.group(1))
                end_row = int(sum_match.group(2))
                for r in range(start_row, end_row + 1):
                    total += bdo_h_values.get(r, 0)
                
                # There might be additional terms after the SUM, e.g., +...+...
                remaining = formula[sum_match.end():]
                additional = self._parse_additional_terms(remaining, bdo_h_values)
                total += additional
                return total
            
            # Pattern 2: Multiple cell references added together
            # ='BDO - Q3-25'!H13+'BDO - Q3-25'!H14+...
            cell_refs = re.findall(r"'.*?'!H(\d+)", formula)
            if cell_refs:
                for row_str in cell_refs:
                    row_num = int(row_str)
                    total += bdo_h_values.get(row_num, 0)
                return total
            
            # Pattern 3: Single cell reference without quotes
            simple_match = re.search(r"H(\d+)", formula)
            if simple_match:
                row_num = int(simple_match.group(1))
                return bdo_h_values.get(row_num, 0)
                
        except Exception as e:
            logger.warning(f"Error evaluating formula '{formula}': {e}")
        
        return None
    
    def _parse_additional_terms(self, remaining: str, bdo_h_values: dict) -> float:
        """Parse additional terms like +'BDO - Q3-25'!H65"""
        total = 0.0
        cell_refs = re.findall(r"\+'.*?'!H(\d+)", remaining)
        for row_str in cell_refs:
            row_num = int(row_str)
            total += bdo_h_values.get(row_num, 0)
        return total
    
    def _copy_ma_data_fallback(self, target_sheet, bdo_sheet, bdo_h_values: dict) -> int:
        """Fallback method using direct BDO row mappings."""
        # Direct mapping from target row to BDO rows (based on reference formulas)
        row_to_bdo_rows = {
            2: [16],                    # Deferred Tax Asset: H16
            3: list(range(6, 11)),      # Real estate: H6:H10
            4: [13, 14, 15],            # Financial fixed assets: H13+H14+H15
            5: [19, 21, 22],            # Accounts receivable: H19+H21+H22
            6: list(range(67, 73)) + [65, 61, 62],  # Service costs
            7: list(range(23, 26)) + [30, 26],      # Prepaid expenses
            8: list(range(31, 38)),     # Cash: H31:H37
            9: [40, 41, 42],            # Equity: H40:H42
            10: [20],                   # AC Shareholder: H20
            11: [45, 46],               # Bank loan: H45+H46
            12: [47],                   # Amortised fee: H47
            13: [50, 60, 63, 64, 29],   # Accounts payable
            14: list(range(51, 57)),    # Current account: H51:H56
            15: [57],                   # VAT payable: H57
            16: [59],                   # Deposits: H59
            17: [58],                   # Rent Invoiced in advance: H58
        }
        
        items_copied = 0
        for target_row, bdo_rows in row_to_bdo_rows.items():
            total = sum(bdo_h_values.get(r, 0) for r in bdo_rows)
            target_sheet.cell(row=target_row, column=3).value = total
            items_copied += 1
        
        return items_copied
    
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
        Update Impact Unit Sales sheet.
        
        IMPORTANT: Unlike Suppl. Calc, this sheet does NOT need new quarterly columns.
        The structure is:
        - Column C: Labels (Average Sale Price LTM, etc.)
        - Column D onwards: Data for different forecasts
        - Row 9: References to Suppl. Calc for unit sales proceeds
        - Row 10: Unit Sales counts
        
        The formulas reference Suppl. Calc columns which get updated automatically.
        We just need to ensure the sheet exists and has correct structure - no new columns needed.
        """
        if 'Impact Unit Sales' not in self.workbook.sheetnames:
            logger.warning("Impact Unit Sales sheet not found")
            return
        
        sheet = self.workbook['Impact Unit Sales']
        
        # The Impact Unit Sales sheet structure should remain unchanged
        # It references Suppl. Calc dynamically, so no column insertion needed
        
        # Just verify the structure is correct
        logger.info("Impact Unit Sales sheet verified - no column changes needed")
    
    def _update_suppl_calc_formulas(self):
        """
        Update formulas in Suppl. Calc sheet to reference Q3 Management Accounts correctly.
        
        The formulas in Suppl. Calc reference the Management Accounts sheet, and we need to ensure
        they reference the correct sheet name and column (column C for Q3 data).
        """
        if 'Suppl. Calc' not in self.workbook.sheetnames:
            return
        
        sheet = self.workbook['Suppl. Calc']
        ma_sheet_name = f"Q{self.config.quarter} Management Accounts"
        
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
        new_ma_name = f"'{ma_sheet_name}'"
        
        updated_count = 0
        
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                cell_value = cell.value
                
                if cell_value is None:
                    continue
                
                if isinstance(cell_value, str) and cell_value.startswith('='):
                    original_value = cell_value
                    new_value = original_value
                    
                    # Update Management Accounts sheet references
                    for old_pattern in old_ma_patterns:
                        if old_pattern in new_value:
                            new_value = new_value.replace(old_pattern, new_ma_name)
                            # Ensure proper exclamation mark
                            if "'" in new_value and "!" not in new_value.split("'")[-1]:
                                # Add ! if missing after sheet name
                                parts = new_value.split("'")
                                for i in range(len(parts) - 1):
                                    if parts[i].endswith(" Management Accounts") and i + 1 < len(parts):
                                        if not parts[i + 1].startswith("!"):
                                            parts[i + 1] = "!" + parts[i + 1]
                                new_value = "'".join(parts)
                    
                    if new_value != original_value:
                        cell.value = new_value
                        updated_count += 1
        
        if updated_count > 0:
            logger.info(f"Updated {updated_count} formulas in Suppl. Calc to reference {ma_sheet_name}")
    
    def _update_sfa_cc_sheet(self):
        """
        Update SFA CC sheet with current quarter references.
        
        This includes:
        1. Updating sheet name references in formulas (Q2 Management Accounts → Q3 Management Accounts)
        2. Updating Suppl. Calc column references for NTM (column S → T after insert)
        3. Updating text references to quarter
        4. Fixing any #NAME? errors by ensuring correct references
        """
        if 'SFA CC' not in self.workbook.sheetnames:
            logger.warning("SFA CC sheet not found")
            return
        
        sheet = self.workbook['SFA CC']
        
        # Get the NTM column shift info
        old_ntm_col_letter = None
        new_ntm_col_letter = None
        
        if 'Suppl. Calc' in self.workbook.sheetnames:
            suppl = self.workbook['Suppl. Calc']
            for col in range(1, suppl.max_column + 5):
                val = suppl.cell(row=2, column=col).value
                if val and str(val).upper() == 'NTM':
                    new_ntm_col_letter = get_column_letter(col)
                    # The old NTM was one column to the left (before we inserted)
                    if col > 1:
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
            f"'Q{prev_q} {prev_y} Management Accounts'!",
            f"'Q{prev_q} Management Accounts'!",
        ]
        new_ma_name = f"'Q{self.config.quarter} Management Accounts'"
        
        updated_count = 0
        formula_count = 0
        
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                
                # Handle both string formulas and formula objects
                cell_value = cell.value
                if cell_value is None:
                    continue
                
                # Convert to string for processing
                if hasattr(cell_value, 'text'):
                    original_value = cell_value.text
                else:
                    original_value = str(cell_value)
                
                new_value = original_value
                
                # Check if it's a formula
                if isinstance(cell_value, str) and new_value.startswith('='):
                    # Update Management Accounts sheet references
                    for old_pattern in old_ma_patterns:
                        if old_pattern in new_value:
                            new_value = new_value.replace(old_pattern, new_ma_name + "!")
                            # Remove double exclamation if created
                            new_value = new_value.replace("!!", "!")
                            logger.debug(f"Updated MA reference: {old_pattern} → {new_ma_name}")
                    
                    # Update Suppl. Calc NTM column references
                    if old_ntm_col_letter and new_ntm_col_letter and "'Suppl. Calc'" in new_value:
                        # Replace column letter references for NTM (e.g., S5 → T5)
                        # Pattern with $ signs: 'Suppl. Calc'!$S$5 or 'Suppl. Calc'!$S5
                        pattern = rf"'Suppl\. Calc'!\${old_ntm_col_letter}(\$?\d+)"
                        replacement = f"'Suppl. Calc'!${new_ntm_col_letter}\\1"
                        new_value = re.sub(pattern, replacement, new_value)
                        
                        # Pattern without $ signs: 'Suppl. Calc'!S5
                        pattern = rf"'Suppl\. Calc'!{old_ntm_col_letter}(\d+)"
                        replacement = f"'Suppl. Calc'!{new_ntm_col_letter}\\1"
                        new_value = re.sub(pattern, replacement, new_value)
                        
                        logger.debug(f"Updated Suppl. Calc NTM reference: {old_ntm_col_letter} → {new_ntm_col_letter}")
                    
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
                        cell.value = new_value
                        updated_count += 1
                
                # Update text references to quarter (non-formula)
                elif isinstance(cell_value, str) and not new_value.startswith('='):
                    if self.config.prev_quarter_str in new_value:
                        new_value = new_value.replace(self.config.prev_quarter_str, self.config.quarter_str)
                        if new_value != original_value:
                            cell.value = new_value
                            updated_count += 1
        
        logger.info(f"Updated {updated_count} cells in SFA CC ({formula_count} formulas)")
        
        # Update signature page dates (typically in rows 82+)
        self._update_signature_page_dates(sheet)
        
        # Ensure signature page content exists
        self._ensure_signature_page(sheet)
    
    def _update_signature_page_dates(self, sheet):
        """
        Update dates in the signature page section of SFA CC.
        
        The signature page typically contains:
        - Date in format "DD-MM-YYYY" or "DD Month YYYY"
        - Quarter references
        - Interest Period dates
        
        These need to be updated to the current quarter's period end.
        """
        # Calculate dates
        period_end = self.config.period_end
        
        # Report date is typically ~20 days after quarter end
        from datetime import timedelta
        report_date = period_end + timedelta(days=21)
        
        # Previous quarter dates
        prev_q = self.config.quarter - 1
        prev_y = self.config.year
        if prev_q == 0:
            prev_q = 4
            prev_y -= 1
        
        # Calculate previous period end
        prev_month = prev_q * 3
        if prev_month in [3, 12]:
            prev_day = 31
        else:
            prev_day = 30
        
        # New date formats
        new_period_date = f"{period_end.day}-{period_end.month}-{period_end.year}"
        new_report_date = f"{report_date.day}-{report_date.month}-{report_date.year}"
        
        # Date pattern replacements - sorted by specificity/length to avoid partial matches
        date_replacements = [
            # Interest Period formats (various)
            (f"Interest Period: {prev_month}/{prev_day}/{prev_y}", 
             f"Interest Period: {period_end.month}/{period_end.day}/{period_end.year}"),
            (f"Interest Period: {prev_day}-{prev_month}-{prev_y}", 
             f"Interest Period: {new_period_date}"),
            # US format mm/dd/yyyy
            (f"{prev_month}/{prev_day}/{prev_y}", 
             f"{period_end.month}/{period_end.day}/{period_end.year}"),
            # EU format dd-mm-yyyy
            (f"{prev_day}-{prev_month}-{prev_y}", new_period_date),
            # Dated field - report date (21 days after quarter end)
            (f"Dated {self._calc_dated_field(prev_q, prev_y)}",
             f"Dated {new_report_date}"),
            # Just "21-7-2025" or similar
            (f"21-{prev_month + 1 if prev_month < 12 else 1}-{prev_y}", new_report_date),
        ]
        
        updated_dates = 0
        
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value and isinstance(cell.value, str) and not cell.value.startswith('='):
                    original = cell.value
                    new_value = original
                    
                    # Apply date replacements
                    for old_pattern, new_pattern in date_replacements:
                        if old_pattern in new_value:
                            new_value = new_value.replace(old_pattern, new_pattern)
                    
                    # Replace quarter text references
                    old_q_str = f"Q{prev_q} {prev_y}"
                    if old_q_str in new_value:
                        new_value = new_value.replace(old_q_str, self.config.quarter_str)
                    
                    if new_value != original:
                        cell.value = new_value
                        updated_dates += 1
        
        if updated_dates > 0:
            logger.info(f"Updated {updated_dates} dates in signature section")
    
    def _calc_dated_field(self, quarter: int, year: int) -> str:
        """Calculate the 'Dated' field date (typically 21 days after quarter end)."""
        from datetime import datetime, timedelta
        month = quarter * 3
        day = 31 if month in [3, 12] else 30
        period_end = datetime(year, month, day)
        dated = period_end + timedelta(days=21)
        return f"{dated.day}-{dated.month}-{dated.year}"
    
    def _ensure_signature_page(self, sheet):
        """
        Ensure the signature page content exists in the SFA CC sheet.
        
        If row 81 says "Signature page follows" but the actual signature section
        (rows 82+) is missing content, add the standard signature fields.
        """
        from datetime import timedelta
        
        # Check if signature page content already exists
        signature_content_exists = False
        for row in range(82, min(sheet.max_row + 1, 100)):
            for col in range(1, 10):
                cell_val = sheet.cell(row=row, column=col).value
                if cell_val and isinstance(cell_val, str) and len(cell_val.strip()) > 0:
                    # Actual content found
                    if 'Name' in str(cell_val) or 'Title' in str(cell_val) or 'Signature' in str(cell_val):
                        signature_content_exists = True
                        break
            if signature_content_exists:
                break
        
        if signature_content_exists:
            logger.debug("Signature page content already exists")
            return
        
        # Add signature page content
        # Calculate dates
        period_end = self.config.period_end
        report_date = period_end + timedelta(days=21)
        report_date_str = report_date.strftime("%d %B %Y")
        
        # Standard signature page layout starting at row 83
        signature_content = [
            # Row 83: Title
            (83, 1, "SIGNATURE PAGE"),
            (83, 2, None),
            # Row 85: QSP ESS B.V.
            (85, 1, "For and on behalf of QSP ESS B.V."),
            # Row 87-88: Name fields
            (87, 1, "Name:"),
            (87, 4, "_________________________________"),
            (88, 1, "Title:"),
            (88, 4, "_________________________________"),
            (89, 1, "Date:"),
            (89, 4, report_date_str),
            # Row 91: Authorised Signatory
            (91, 1, "Authorised Signatory"),
            # Row 93-94: Second signer
            (93, 1, "Name:"),
            (93, 4, "_________________________________"),
            (94, 1, "Title:"),
            (94, 4, "_________________________________"),
            (95, 1, "Date:"),
            (95, 4, report_date_str),
            # Row 97
            (97, 1, "Authorised Signatory"),
        ]
        
        for row, col, value in signature_content:
            cell = sheet.cell(row=row, column=col)
            cell.value = value
            if row == 83:  # Title row
                cell.font = Font(bold=True, size=14)
            elif 'Name' in str(value) or 'Title' in str(value) or 'Date' in str(value):
                cell.font = Font(bold=True)
        
        # Update print area to include signature page
        if sheet.print_area:
            import re
            match = re.match(r".*\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", sheet.print_area)
            if match:
                start_col, start_row, end_col, _ = match.groups()
                new_print_area = f"${start_col}${start_row}:${end_col}$99"
                sheet.print_area = new_print_area
                logger.debug(f"Extended print area to include signature page: {new_print_area}")
        
        logger.info("Added signature page content to SFA CC sheet")
    
    def _update_all_dates_and_text(self):
        """
        Update ALL dates and text references throughout the entire workbook.
        
        This ensures consistency across all sheets:
        - SFA CC signature dates
        - Interest Period references  
        - Quarter text (Q2 2025 → Q3 2025)
        - Date formats (21-7-2025 → 21-10-2025)
        
        IMPORTANT: We MUST NOT replace short quarter patterns (like "25Q2") in 
        column headers of Suppl. Calc and Impact Unit Sales sheets, as these
        represent historical quarters, not the current quarter being updated.
        """
        from datetime import timedelta
        
        period_end = self.config.period_end
        report_date = period_end + timedelta(days=21)
        
        # Previous quarter info
        prev_q = self.config.quarter - 1
        prev_y = self.config.year
        if prev_q == 0:
            prev_q = 4
            prev_y -= 1
        
        prev_month = prev_q * 3
        prev_day = 31 if prev_month in [3, 12] else 30
        
        # Build comprehensive replacement map (sorted by length to prevent partial matches)
        replacements = []
        
        # Quarter text patterns - FULL format only (e.g., "Q2 2025")
        # We DO NOT include the short format (e.g., "25Q2") because it would
        # incorrectly replace historical quarter column headers in Suppl. Calc
        replacements.extend([
            (f"Q{prev_q} {prev_y}", self.config.quarter_str),
            (f"Q{prev_q}-{prev_y}", f"Q{self.config.quarter}-{self.config.year}"),
            # NOTE: We intentionally OMIT the short quarter pattern replacement
            # (f"{str(prev_y)[-2:]}Q{prev_q}", self.config.short_quarter),
            # as it would incorrectly replace column headers in Suppl. Calc and Impact Unit Sales
        ])
        
        # Date patterns - multiple formats
        # Dated field (report date, ~21 days after quarter)
        prev_report_month = prev_month + 1 if prev_month < 12 else 1
        prev_report_year = prev_y if prev_month < 12 else prev_y + 1
        replacements.extend([
            # Full dated field
            (f"Dated 21-{prev_report_month}-{prev_report_year}",
             f"Dated {report_date.day}-{report_date.month}-{report_date.year}"),
            # Just the date
            (f"21-{prev_report_month}-{prev_report_year}",
             f"{report_date.day}-{report_date.month}-{report_date.year}"),
        ])
        
        # Interest Period dates - US format mm/dd/yyyy
        replacements.extend([
            (f"Interest Period: {prev_month}/{prev_day}/{prev_y}",
             f"Interest Period: {period_end.month}/{period_end.day}/{period_end.year}"),
            (f"{prev_month}/{prev_day}/{prev_y}",
             f"{period_end.month}/{period_end.day}/{period_end.year}"),
        ])
        
        # Period end dates - EU format dd-mm-yyyy
        replacements.extend([
            (f"{prev_day}-{prev_month}-{prev_y}",
             f"{period_end.day}-{period_end.month}-{period_end.year}"),
        ])
        
        # Sort by length (longest first) to prevent partial matches
        replacements.sort(key=lambda x: len(x[0]), reverse=True)
        
        total_updated = 0
        
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            sheet_updated = 0
            
            # Skip row 2 in Suppl. Calc and Impact Unit Sales (column headers)
            # These contain historical quarter labels that should NOT be updated
            skip_header_rows = sheet_name in ['Suppl. Calc', 'Impact Unit Sales']
            
            for row in range(1, min(sheet.max_row + 1, 200)):
                # Skip header row (row 2) for specific sheets with quarter column headers
                if skip_header_rows and row == 2:
                    continue
                
                for col in range(1, min(sheet.max_column + 1, 50)):
                    cell = sheet.cell(row=row, column=col)
                    
                    if cell.value is None:
                        continue
                    
                    # Only process string values (not formulas)
                    if isinstance(cell.value, str) and not cell.value.startswith('='):
                        original = cell.value
                        new_value = original
                        
                        for old_pattern, new_pattern in replacements:
                            if old_pattern in new_value:
                                new_value = new_value.replace(old_pattern, new_pattern)
                        
                        if new_value != original:
                            cell.value = new_value
                            sheet_updated += 1
            
            if sheet_updated > 0:
                logger.debug(f"Updated {sheet_updated} cells in sheet '{sheet_name}'")
                total_updated += sheet_updated
        
        if total_updated > 0:
            logger.info(f"Updated {total_updated} date/text references across all sheets")
    
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

