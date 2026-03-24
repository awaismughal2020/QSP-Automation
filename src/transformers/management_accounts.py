"""
Management Accounts Transformer

Builds the Management Accounts Excel file by:
1. Copying ALL data from the BDO file (Cijfers_QSP_) into a new sheet
2. Inserting a new column in Management Cijfers before the LTM column
3. Building formulas for the new column using formula templates (account code mapping)
4. Updating formulas to reference the new BDO sheet with correct row numbers
5. Ensuring Management Cijfers sheet is always at the end

WORKFLOW:
- Copy all data from BDO file → Create new sheet "BDO - Q{quarter}-{YY}"
- Build account code to row mapping from new BDO sheet
- In Management Cijfers sheet, insert new column before LTM column
- Use formula templates to build formulas for new column (looking up account codes)
- Update LTM column to reference new BDO sheet with column H
- Move Management Cijfers sheet to be the last sheet

KEY DESIGN:
- Formula templates define which account codes each Management Cijfers row references
- When building new column, look up account codes in the new BDO sheet to get correct row numbers
- Only the NEW column gets updated formulas; existing columns remain unchanged
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Any
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
    cash_proceeds_sale: float = 0.0  # Row 50: Cash proceeds from sales tracker


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
    4. Build formulas for new quarter column using formula templates
    5. Build formulas for LTM column using same templates but with column H
    6. Move Management Cijfers sheet to be last
    
    KEY DESIGN:
    - Formula templates (config/formula_templates.yaml) define the account codes 
      each Management Cijfers row should reference
    - When adding a new column, we look up those account codes in the NEW BDO sheet
      to get the correct row numbers
    - Only the NEW column gets new formulas; existing columns remain unchanged
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
        
        # Load formula templates
        self.formula_templates_path = Path("config/formula_templates.yaml")
        self.formula_templates = self._load_formula_templates()
        self.balance_sheet_templates = self._load_balance_sheet_templates()
        
        # Track the previous BDO sheet name for formula updates
        self._prev_bdo_sheet_name = None
        
        # Account code to row mapping in the NEW BDO sheet
        self._new_bdo_row_map = {}
        
        # Label to row mapping in the NEW BDO sheet (for non-numeric account codes)
        self._new_bdo_label_map = {}
        
        # Account code to row mapping in the PREVIOUS BDO sheet (for row offset calculation)
        self._prev_bdo_row_map = {}
        
        # Sheets that must never be modified (existing BDO/kwartaal sheets)
        self._protected_sheets: set = set()
    
    def _load_formula_templates(self) -> Dict[int, Dict[str, Any]]:
        """Load formula templates that define the account code mapping for each row."""
        templates = {}
        
        if not self.formula_templates_path.exists():
            logger.warning(f"Formula templates file not found: {self.formula_templates_path}")
            return templates
        
        try:
            with open(self.formula_templates_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            
            # Load P&L section templates
            if 'profit_loss' in config:
                for row_str, template in config['profit_loss'].items():
                    if template:
                        templates[int(row_str)] = template
            
            logger.info(f"Loaded {len(templates)} P&L formula templates from {self.formula_templates_path}")
        except Exception as e:
            logger.warning(f"Error loading formula templates: {e}")
        
        return templates
    
    def _load_balance_sheet_templates(self) -> Dict[int, Dict[str, Any]]:
        """Load Balance Sheet formula templates."""
        templates = {}
        
        if not self.formula_templates_path.exists():
            return templates
        
        try:
            with open(self.formula_templates_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            
            if 'balance_sheet' in config:
                for row_str, template in config['balance_sheet'].items():
                    if template:
                        templates[int(row_str)] = template
            
            logger.info(f"Loaded {len(templates)} Balance Sheet formula templates")
        except Exception as e:
            logger.warning(f"Error loading Balance Sheet templates: {e}")
        
        return templates
    
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
        
        # Mark all existing BDO/kwartaal sheets as protected - they must never be modified
        self._protected_sheets = {
            name for name in self.workbook.sheetnames
            if name.startswith('BDO') or name != self.config.summary_sheet_name
        }
        # The summary sheet (Management Cijfers) is the only pre-existing sheet we modify
        self._protected_sheets.discard(self.config.summary_sheet_name)
        logger.info(f"Protected sheets (will not be modified): {sorted(self._protected_sheets)}")
        
        # Get previous BDO sheet name and build row map
        bdo_sheets = [name for name in self.workbook.sheetnames if name.startswith('BDO')]
        if bdo_sheets:
            self._prev_bdo_sheet_name = bdo_sheets[-1]  # Most recent
            logger.info(f"Previous BDO sheet: {self._prev_bdo_sheet_name}")
            # Build row map from previous BDO sheet
            self._build_row_map(self.workbook[self._prev_bdo_sheet_name], self._prev_bdo_row_map)
        
        # Step 1: Copy ALL data from BDO file into new sheet
        self._copy_bdo_data_to_new_sheet(bdo_result)
        
        # Step 2: Build row map from the NEW BDO sheet (including label map for special rows)
        new_bdo_sheet = self.workbook[self.config.bdo_sheet_name]
        self._build_row_map(new_bdo_sheet, self._new_bdo_row_map, self._new_bdo_label_map)
        logger.info(f"Built row map for new BDO sheet: {len(self._new_bdo_row_map)} accounts, {len(self._new_bdo_label_map)} labels")
        
        # Step 3: Update summary sheet - insert new column, copy style/formulas
        self._update_summary_sheet()
        
        # Step 4: Update all date references throughout the workbook
        self._update_date_references()
        
        # Step 5: Move BDO sheet to second-to-last position
        self._move_bdo_sheet_to_second_last()
        
        # Step 6: Move Management Cijfers sheet to the end (after BDO)
        self._move_summary_sheet_to_end()
        
        # Step 7: Validate calculations
        self.validation_result = self._validate_calculations(bdo_result)
        
        # Save output
        self.workbook.save(self.output_path)
        logger.info(f"Saved to {self.output_path}")
        
        return self.output_path
    
    def _build_row_map(self, sheet, row_map: dict, label_map: dict = None):
        """
        Build mapping of account codes to row numbers in a BDO sheet.
        This is critical for validating that formula row references are correct.
        
        Also builds a label map for non-numeric identifiers (like "SWAP").
        """
        row_map.clear()
        if label_map is not None:
            label_map.clear()
        
        for row_idx in range(1, sheet.max_row + 1):
            code = sheet.cell(row=row_idx, column=1).value
            label = sheet.cell(row=row_idx, column=2).value
            
            if code and isinstance(code, (str, int, float)):
                code_str = str(code).strip()
                if code_str and code_str[0].isdigit():
                    row_map[code_str] = row_idx
                elif code_str and label_map is not None:
                    # Non-numeric code - store in label map
                    label_map[code_str.lower()] = row_idx
            
            # Also map by label column (column B) for special rows like "SWAP"
            if label and isinstance(label, str) and label_map is not None:
                label_map[label.strip().lower()] = row_idx
    
    def _copy_bdo_data_to_new_sheet(self, bdo_result: BDOParseResult):
        """
        Create new BDO sheet as a verbatim copy of the uploaded Cijfers file.

        Only the sheet title is changed to match the automation naming convention.
        All cells (values + formulas), formatting, merged cells, column widths,
        and row heights are preserved exactly as they appear in the source file.
        """
        new_sheet_name = self.config.bdo_sheet_name

        if new_sheet_name in self.workbook.sheetnames:
            logger.warning(f"Sheet {new_sheet_name} already exists, will remove and recreate")
            del self.workbook[new_sheet_name]

        if not self.bdo_source_path or not self.bdo_source_path.exists():
            logger.warning("BDO source file not provided, will clone previous BDO sheet")
            self._clone_and_update_bdo_sheet(bdo_result)
            return

        bdo_wb = openpyxl.load_workbook(self.bdo_source_path, data_only=False)
        source_sheet = bdo_wb.active
        logger.info(f"Loading BDO source (verbatim copy) from: {self.bdo_source_path}")

        insert_index = len(self.workbook.sheetnames)
        for i, name in enumerate(self.workbook.sheetnames):
            if 'Management Cijfers' in name or 'Cijfers' in name:
                insert_index = i
                break

        new_sheet = self.workbook.create_sheet(title=new_sheet_name, index=insert_index)

        max_row = source_sheet.max_row
        max_col = source_sheet.max_column

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                src_cell = source_sheet.cell(row=row, column=col)
                tgt_cell = new_sheet.cell(row=row, column=col)
                tgt_cell.value = src_cell.value
                tgt_cell.font = copy(src_cell.font)
                tgt_cell.alignment = copy(src_cell.alignment)
                tgt_cell.number_format = src_cell.number_format
                tgt_cell.border = copy(src_cell.border)
                tgt_cell.fill = copy(src_cell.fill)

        for col_letter, dim in source_sheet.column_dimensions.items():
            new_sheet.column_dimensions[col_letter].width = dim.width

        for row_idx, dim in source_sheet.row_dimensions.items():
            new_sheet.row_dimensions[row_idx].height = dim.height

        for merged_range in source_sheet.merged_cells.ranges:
            new_sheet.merge_cells(str(merged_range))

        logger.info(f"Created verbatim BDO sheet: {new_sheet_name} ({max_row} rows x {max_col} cols)")

        bdo_wb.close()
    
    def _format_quarter_date(self, quarter: str, offset: int = 0) -> str:
        """Format a quarter string into a date string for column headers."""
        import re
        match = re.match(r'Q(\d)\s*(\d{4})', quarter)
        if not match:
            return quarter
        
        q_num, year = int(match.group(1)), int(match.group(2))
        
        # Apply offset (in quarters)
        total_quarters = (year * 4 + q_num) + offset
        new_year = (total_quarters - 1) // 4
        new_q = ((total_quarters - 1) % 4) + 1
        
        # Quarter end dates
        end_dates = {1: "31-03", 2: "30-06", 3: "30-09", 4: "31-12"}
        return f"{end_dates[new_q]}-{new_year}"
    
    def _set_column_h_formulas(self, sheet, max_row: int):
        """
        Set formulas in column H for the BDO sheet.
        
        Column H contains the LTM (Last Twelve Months) sums.
        
        Structure (based on reference file):
        - Rows 6-124: =SUM(C{row}:G{row}) - Balance sheet items with opening balance
        - Row 125 (Resultaat na belasting): =SUM(C125:G125)
        - Row 126: Empty
        - Row 127 (Verschil balans): =SUM(H5:H123) - Note: Sum of H column, not C-G row
        - Row 128: Empty
        - Row 129 (afschrijving SWAP): No H formula
        - Row 130 (Inkomsten SWAP): No H formula
        - Row 131: Empty
        - Row 132 (Rente): =H117+H114+H116+H115
        - Row 133 (SWAP): =SUM(D133:G133)
        - Row 134 (Afschrijving prepaid): =SUM(D134:G134)
        - Row 135 (IC interest): =SUM(D135:G135)
        - Row 136 (Total): =SUM(D136:G136)
        
        This ensures the formulas remain as formulas, not just values.
        """
        formulas_set = 0
        
        # Find the first special row (where labels like 'afschrijving SWAP' start)
        special_start_row = 129  # Default based on reference
        for row in range(120, min(max_row + 1, 140)):
            label = sheet.cell(row=row, column=1).value
            if label and 'afschrijving swap' in str(label).lower():
                special_start_row = row
                break
        
        # Rows 6 to 124: =SUM(C{row}:G{row})
        # These are balance sheet items that include opening balance (column C)
        for row in range(6, 125):
            cell = sheet.cell(row=row, column=8)  # Column H
            # Only set formula if there's data in the row (check column A)
            if sheet.cell(row=row, column=1).value is not None:
                cell.value = f"=SUM(C{row}:G{row})"
                formulas_set += 1
        
        # Row 125 (Resultaat na belasting): =SUM(C125:G125)
        sheet.cell(row=125, column=8).value = "=SUM(C125:G125)"
        formulas_set += 1
        
        # Row 126: Empty (no formula)
        
        # Row 127 (Verschil balans en winst-en-verlies): =SUM(H5:H123)
        # Note: This sums the H column, not the C-G of the same row
        sheet.cell(row=127, column=8).value = "=SUM(H5:H123)"
        formulas_set += 1
        
        # Rows 128-136 have their H formulas set in _add_special_rows_with_formulas()
        
        logger.info(f"Set {formulas_set} formulas in column H (special rows handled separately)")
    
    def _extract_swap_income_from_sheet(self, sheet) -> float:
        """
        Extract SWAP income value for the new quarter from the BDO sheet.
        Looks up account 4643000 in column A and reads its column G value.
        """
        for row_idx in range(1, sheet.max_row + 1):
            code = sheet.cell(row=row_idx, column=1).value
            if code and str(code).strip() == "4643000":
                val = sheet.cell(row=row_idx, column=7).value
                if val is not None:
                    try:
                        return float(val)
                    except (ValueError, TypeError):
                        pass
                logger.warning("Account 4643000 found but column G has no numeric value")
                return 0.0
        logger.warning("Account 4643000 not found in BDO sheet for SWAP income extraction")
        return 0.0

    def _add_special_rows_with_formulas(self, sheet, start_row: int):
        """
        Add special rows with formulas matching the reference file structure.
        
        Reference structure (from correctFiles.xlsx):
        - Row 126: Empty
        - Row 127: "Verschil balans en winst-en-verlies" with SUM(C5:C123) etc.
        - Row 128: Empty
        - Row 129: "afschrijving SWAP" with =36883.33*3 in D-G
        - Row 130: "Inkomsten SWAP" with shifted values
        - Row 131: Empty
        - Row 132: "Rente" with =D117, =E117, etc.
        - Row 133: "SWAP" with =D130, =E130, etc. and G value
        - Row 134: "Afschrijving prepaid" with =D129, =E129, etc.
        - Row 135: "IC interest" with =D115+D120, etc.
        - Row 136: Total row with SUM formulas
        """
        prev_inkomsten_values = self._get_shifted_inkomsten_swap_values()
        
        swap_income = self._extract_swap_income_from_sheet(sheet)
        logger.info(f"Extracted SWAP income for new quarter: {swap_income:,.2f}")
        
        # Row 129 (start_row + 2): afschrijving SWAP — use formula matching client template
        row_129 = start_row + 2
        sheet.cell(row=row_129, column=1).value = "afschrijving SWAP"
        sheet.cell(row=row_129, column=4).value = "=36883.33*3"  # D
        sheet.cell(row=row_129, column=5).value = "=36883.33*3"  # E
        sheet.cell(row=row_129, column=6).value = "=36883.33*3"  # F
        sheet.cell(row=row_129, column=7).value = "=36883.33*3"  # G
        
        # Row 130 (start_row + 3): Inkomsten SWAP — shifted values + new quarter from BDO
        row_130 = start_row + 3
        sheet.cell(row=row_130, column=1).value = "Inkomsten SWAP"
        sheet.cell(row=row_130, column=4).value = prev_inkomsten_values.get('E', 0)  # D = old E
        sheet.cell(row=row_130, column=5).value = prev_inkomsten_values.get('F', 0)  # E = old F
        sheet.cell(row=row_130, column=6).value = prev_inkomsten_values.get('G', 0)  # F = old G
        sheet.cell(row=row_130, column=7).value = swap_income  # G = new quarter SWAP income
        
        # Row 131 (start_row + 4): Empty row
        
        # Row 132 (start_row + 5): Rente
        row_132 = start_row + 5
        sheet.cell(row=row_132, column=1).value = "Rente"
        sheet.cell(row=row_132, column=4).value = "=D117"
        sheet.cell(row=row_132, column=5).value = "=E117"
        sheet.cell(row=row_132, column=6).value = "=F117"
        sheet.cell(row=row_132, column=7).value = "=G117"
        sheet.cell(row=row_132, column=8).value = "=H117+H114+H116+H115"
        
        # Row 133 (start_row + 6): SWAP — all columns reference Inkomsten SWAP row
        row_133 = start_row + 6
        sheet.cell(row=row_133, column=1).value = "SWAP"
        sheet.cell(row=row_133, column=4).value = f"=D{row_130}"
        sheet.cell(row=row_133, column=5).value = f"=E{row_130}"
        sheet.cell(row=row_133, column=6).value = f"=F{row_130}"
        sheet.cell(row=row_133, column=7).value = f"=G{row_130}"
        sheet.cell(row=row_133, column=8).value = f"=SUM(D{row_133}:G{row_133})"
        
        # Row 134 (start_row + 7): Afschrijving prepaid — references afschrijving SWAP row
        row_134 = start_row + 7
        sheet.cell(row=row_134, column=1).value = "Afschrijving prepaid derivatives transaction"
        sheet.cell(row=row_134, column=4).value = f"=D{row_129}"
        sheet.cell(row=row_134, column=5).value = f"=E{row_129}"
        sheet.cell(row=row_134, column=6).value = f"=F{row_129}"
        sheet.cell(row=row_134, column=7).value = f"=G{row_129}"
        sheet.cell(row=row_134, column=8).value = f"=SUM(D{row_134}:G{row_134})"
        
        # Row 135 (start_row + 8): IC interest
        row_135 = start_row + 8
        sheet.cell(row=row_135, column=1).value = "IC interest"
        sheet.cell(row=row_135, column=4).value = "=D115+D120"
        sheet.cell(row=row_135, column=5).value = "=E115+E120"
        sheet.cell(row=row_135, column=6).value = "=F115+F120"
        sheet.cell(row=row_135, column=7).value = "=G115+G120"
        sheet.cell(row=row_135, column=8).value = f"=SUM(D{row_135}:G{row_135})"
        
        # Row 136 (start_row + 9): Total row
        row_136 = start_row + 9
        sheet.cell(row=row_136, column=4).value = f"=SUM(D{row_132}:D{row_135})"
        sheet.cell(row=row_136, column=5).value = f"=SUM(E{row_132}:E{row_135})"
        sheet.cell(row=row_136, column=6).value = f"=SUM(F{row_132}:F{row_135})"
        sheet.cell(row=row_136, column=7).value = f"=SUM(G{row_132}:G{row_135})"
        sheet.cell(row=row_136, column=8).value = f"=SUM(D{row_136}:G{row_136})"
        
        logger.info(f"Added special rows with formulas (rows {row_129}-{row_136})")
    
    def _get_shifted_inkomsten_swap_values(self) -> dict:
        """Get Inkomsten SWAP values from previous BDO sheet for column shifting."""
        values = {'D': 0, 'E': 0, 'F': 0, 'G': 0}
        
        if not self._prev_bdo_sheet_name:
            return values
        
        prev_wb = None
        try:
            prev_wb = openpyxl.load_workbook(str(self.previous_path), data_only=True)
            if self._prev_bdo_sheet_name in prev_wb.sheetnames:
                prev_sheet = prev_wb[self._prev_bdo_sheet_name]
                for row_idx in range(1, prev_sheet.max_row + 1):
                    label = prev_sheet.cell(row=row_idx, column=1).value
                    if label and 'inkomsten swap' in str(label).lower():
                        values['D'] = prev_sheet.cell(row=row_idx, column=4).value or 0
                        values['E'] = prev_sheet.cell(row=row_idx, column=5).value or 0
                        values['F'] = prev_sheet.cell(row=row_idx, column=6).value or 0
                        values['G'] = prev_sheet.cell(row=row_idx, column=7).value or 0
                        logger.debug(f"Found Inkomsten SWAP at row {row_idx}")
                        break
        except Exception as e:
            logger.warning(f"Error getting Inkomsten SWAP values: {e}")
        finally:
            if prev_wb:
                prev_wb.close()
        
        return values
    
    
    def _copy_bdo_data_direct(self, bdo_result: BDOParseResult):
        """
        Fallback: Copy directly from BDO source file when no previous sheet exists.
        """
        new_sheet_name = self.config.bdo_sheet_name
        
        bdo_wb_values = openpyxl.load_workbook(self.bdo_source_path, data_only=True)
        bdo_wb_formatting = openpyxl.load_workbook(self.bdo_source_path, data_only=False)
        source_sheet_values = bdo_wb_values.active
        source_sheet_formatting = bdo_wb_formatting.active
        
        insert_index = len(self.workbook.sheetnames)
        for i, name in enumerate(self.workbook.sheetnames):
            if 'Management Cijfers' in name or 'Cijfers' in name:
                insert_index = i
                break
        
        new_sheet = self.workbook.create_sheet(title=new_sheet_name, index=insert_index)
        
        max_row = source_sheet_values.max_row
        max_col = source_sheet_values.max_column
        
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                value_cell = source_sheet_values.cell(row=row, column=col)
                format_cell = source_sheet_formatting.cell(row=row, column=col)
                target_cell = new_sheet.cell(row=row, column=col)
                
                target_cell.value = value_cell.value
                target_cell.font = copy(format_cell.font)
                target_cell.alignment = copy(format_cell.alignment)
                target_cell.number_format = format_cell.number_format
                target_cell.border = copy(format_cell.border)
                target_cell.fill = copy(format_cell.fill)
        
        for col_letter, dim in source_sheet_formatting.column_dimensions.items():
            new_sheet.column_dimensions[col_letter].width = dim.width
        
        for row_idx, dim in source_sheet_formatting.row_dimensions.items():
            new_sheet.row_dimensions[row_idx].height = dim.height
        
        for merged_range in source_sheet_formatting.merged_cells.ranges:
            new_sheet.merge_cells(str(merged_range))
        
        logger.info(f"Copied BDO data directly to new sheet: {new_sheet_name}")
        
        bdo_wb_values.close()
        bdo_wb_formatting.close()
    
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
        2. For P&L section (row 22+): Build formulas from templates using account code lookup
        3. For Balance Sheet section (rows 1-20): Copy from LTM column and update BDO refs
        4. Build LTM formulas (same as quarter but with column H instead of G)
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
        
        prev_quarter_col = ltm_col - 1  # Previous quarter column
        insert_position = ltm_col
        
        logger.info(f"LTM column: {get_column_letter(ltm_col)}, Previous quarter: {get_column_letter(prev_quarter_col)}")
        
        # Step 1: Insert new column at LTM position
        # This shifts the old LTM column to the right by 1
        summary_sheet.insert_cols(insert_position)
        
        new_quarter_col = insert_position  # The newly inserted blank column
        new_ltm_col = insert_position + 1  # Where old LTM data now is
        
        new_quarter_letter = get_column_letter(new_quarter_col)
        new_ltm_letter = get_column_letter(new_ltm_col)
        
        logger.info(f"After insert: New quarter col={new_quarter_letter}, LTM col={new_ltm_letter}")
        
        # Step 2: Balance Sheet section (rows 1-20)
        # In the Balance Sheet, only the LTM column has formulas - quarterly columns are empty
        # The new quarter column should remain empty, we only update the LTM column
        # Clear any data that might have been shifted into the new column
        for row_idx in range(1, 21):
            summary_sheet.cell(row=row_idx, column=new_quarter_col).value = None
        
        # Step 3: Build P&L formulas from templates (row 22 onwards)
        self._build_pl_formulas_from_templates(summary_sheet, new_quarter_col, bdo_column='G',
                                               prev_col_idx=prev_quarter_col,
                                               quarter_col_idx=new_quarter_col)
        
        # Step 4: Build LTM formulas (same structure but with column H)
        self._build_pl_formulas_from_templates(summary_sheet, new_ltm_col, bdo_column='H',
                                               prev_col_idx=new_quarter_col,
                                               quarter_col_idx=new_quarter_col)
        
        # Step 4b: Set manual values (Row 50 - Cash proceeds sale)
        if self.config.cash_proceeds_sale:
            cash_cell = summary_sheet.cell(row=50, column=new_quarter_col)
            cash_cell.value = self.config.cash_proceeds_sale
            # Copy number format from previous column
            prev_cash_cell = summary_sheet.cell(row=50, column=prev_quarter_col)
            if prev_cash_cell.number_format:
                cash_cell.number_format = prev_cash_cell.number_format
            logger.info(f"Set Row 50 (Cash proceeds sale) to {self.config.cash_proceeds_sale:,.0f}")
        
        # Step 5: Build Balance Sheet formulas in LTM column
        # Balance Sheet only exists in LTM column - build from templates
        self._build_balance_sheet_formulas(summary_sheet, new_ltm_col, bdo_column='H')
        
        # Step 6: Update headers and copy styling from previous columns
        # Copy header styling from previous quarter column for new quarter (Row 22)
        prev_header_cell = summary_sheet.cell(row=22, column=prev_quarter_col)
        header_cell = summary_sheet.cell(row=22, column=new_quarter_col)
        header_cell.value = self.config.quarter
        # Copy ALL styling including fill (background color) from previous quarter header
        if prev_header_cell.has_style:
            header_cell.font = copy(prev_header_cell.font)
            header_cell.alignment = copy(prev_header_cell.alignment)
            header_cell.border = copy(prev_header_cell.border)
            header_cell.fill = copy(prev_header_cell.fill)  # Copy background color
        
        # Row 2: Balance sheet date header - only in LTM column, quarter column stays empty
        # The date cell in the new quarter column should remain None (cleared earlier)
        # Update the LTM column with the current period end date
        ltm_date_cell = summary_sheet.cell(row=2, column=new_ltm_col)
        ltm_date_cell.value = self.config.period_end
        # Copy styling from previous LTM date cell
        prev_ltm_date_cell = summary_sheet.cell(row=2, column=new_quarter_col + 1) if new_quarter_col + 1 <= summary_sheet.max_column else None
        if prev_ltm_date_cell and prev_ltm_date_cell.has_style:
            ltm_date_cell.number_format = prev_ltm_date_cell.number_format or 'YYYY-MM-DD'
            ltm_date_cell.font = copy(prev_ltm_date_cell.font)
            ltm_date_cell.alignment = copy(prev_ltm_date_cell.alignment)
            ltm_date_cell.border = copy(prev_ltm_date_cell.border)
            ltm_date_cell.fill = copy(prev_ltm_date_cell.fill)
        
        # LTM column header (row 22): copy styling from old LTM header (now
        # at new_ltm_col after the column insert) instead of hardcoding a font.
        ltm_header = summary_sheet.cell(row=22, column=new_ltm_col)
        ltm_header.value = f"LTM {self.config.quarter}"
        
        client_number_format = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
        
        # New quarter column: copy ALL styling from the previous quarter column.
        # The previous column already carries the correct Avenir book/10/bold font
        # and per-row colours; we must not replace it with a hardcoded font.
        for row_idx in range(23, summary_sheet.max_row + 1):
            prev_cell = summary_sheet.cell(row=row_idx, column=prev_quarter_col)
            cell = summary_sheet.cell(row=row_idx, column=new_quarter_col)
            if prev_cell.has_style:
                cell.font = copy(prev_cell.font)
                cell.alignment = copy(prev_cell.alignment)
                cell.border = copy(prev_cell.border)
                cell.fill = copy(prev_cell.fill)
                cell.number_format = prev_cell.number_format if prev_cell.number_format != 'General' else client_number_format
            elif cell.number_format == 'General':
                cell.number_format = client_number_format
        
        # LTM column: explicitly copy full styling from the previous quarter
        # column to ensure fill (blue bar rows), font (green color for LTM
        # values), border, and alignment are consistent across all rows.
        for row_idx in range(23, summary_sheet.max_row + 1):
            prev_cell = summary_sheet.cell(row=row_idx, column=prev_quarter_col)
            cell = summary_sheet.cell(row=row_idx, column=new_ltm_col)
            if prev_cell.has_style:
                cell.font = copy(prev_cell.font)
                cell.alignment = copy(prev_cell.alignment)
                cell.border = copy(prev_cell.border)
                cell.fill = copy(prev_cell.fill)
                cell.number_format = prev_cell.number_format if prev_cell.number_format != 'General' else client_number_format
            elif cell.number_format == 'General':
                cell.number_format = client_number_format
        
        # Copy column width from previous quarter column
        prev_col_letter = get_column_letter(prev_quarter_col)
        new_q_letter = get_column_letter(new_quarter_col)
        new_ltm_letter_dim = get_column_letter(new_ltm_col)
        if prev_col_letter in summary_sheet.column_dimensions:
            prev_width = summary_sheet.column_dimensions[prev_col_letter].width
            if prev_width:
                summary_sheet.column_dimensions[new_q_letter].width = prev_width
                summary_sheet.column_dimensions[new_ltm_letter_dim].width = prev_width
        
        # Step 7: Update LTM column SUM formulas to include new quarter column
        self._update_ltm_sum_ranges(summary_sheet, new_ltm_col, new_quarter_col)
        
        # Step 8: Update Bank Account Overview section (rows 104-120)
        self._update_bank_account_overview_section(summary_sheet, new_ltm_col)
        
        # Step 9: Add new quarter column to column group (outline)
        new_col_letter = get_column_letter(new_quarter_col)
        summary_sheet.column_dimensions[new_col_letter].outlineLevel = 1
        summary_sheet.column_dimensions[new_col_letter].hidden = False
        logger.info(f"Added column {new_col_letter} to group (outline_level=1, hidden=False)")
        
        # Update title
        summary_sheet.cell(row=1, column=1).value = f"Management Accounts QSP ESS B.V. - {self.config.quarter}"
        
        logger.info(f"Updated summary sheet columns: {new_quarter_letter} and {new_ltm_letter}")
    
    def _copy_number_formatting(self, sheet, source_col: int, target_col: int, 
                                 start_row: int, end_row: int):
        """Copy number formatting from source column to target column for specified rows."""
        for row_idx in range(start_row, end_row + 1):
            source_cell = sheet.cell(row=row_idx, column=source_col)
            target_cell = sheet.cell(row=row_idx, column=target_col)
            
            if source_cell.number_format and source_cell.number_format != 'General':
                target_cell.number_format = source_cell.number_format
    
    def _update_ltm_sum_ranges(self, sheet, ltm_col: int, new_quarter_col: int):
        """
        Update SUM formulas in LTM column to include the new quarter column.
        
        After inserting a column, SUM ranges shift but don't extend to include
        the new column. This method fixes that by updating ranges like:
        =SUM(W50:Z50) -> =SUM(X50:AA50)
        """
        ltm_letter = get_column_letter(ltm_col)
        new_q_letter = get_column_letter(new_quarter_col)
        
        updated_count = 0
        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=ltm_col)
            
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                formula = cell.value
                
                # Pattern to match SUM formulas like =SUM(W50:Z50)
                # After insert, old range became shifted but excluded new column
                # Need to update end column to include new quarter
                
                # Match SUM formulas that reference same row
                pattern = rf'=SUM\(([A-Z]+){row_idx}:([A-Z]+){row_idx}\)'
                match = re.match(pattern, formula)
                
                if match:
                    start_col_letter = match.group(1)
                    end_col_letter = match.group(2)
                    
                    start_col_idx = column_index_from_string(start_col_letter)
                    end_col_idx = column_index_from_string(end_col_letter)
                    
                    # LTM should always be sum of 4 quarters
                    # After insert: start should shift +1, end should be new quarter column
                    # Example: =SUM(W50:Z50) should become =SUM(X50:AA50)
                    
                    # Only update if the range doesn't already include new quarter
                    if end_col_idx < new_quarter_col:
                        # Shift start by 1 to drop oldest quarter
                        new_start = get_column_letter(start_col_idx + 1)
                        new_formula = f"=SUM({new_start}{row_idx}:{new_q_letter}{row_idx})"
                        cell.value = new_formula
                        updated_count += 1
        
        if updated_count > 0:
            logger.info(f"Updated {updated_count} LTM SUM formulas to include column {new_q_letter}")
    
    def _update_bank_account_overview_section(self, sheet, ltm_col: int):
        """
        Update the Bank Account Overview section (rows 104-120) in the LTM column.
        
        This section contains:
        - Row 106: Date header ("Per DD-MM-YYYY") - needs date update
        - Rows 107-113: BDO sheet references for each bank account - set correct formulas
        - Row 114: SUM formula - needs column reference update
        
        IMPORTANT: We set the CORRECT formulas directly rather than just updating 
        the BDO sheet name, because the input file may have incorrect formulas.
        
        The correct bank account formulas reference column H (closing balance) in BDO:
        - Row 107 (RENT): H35
        - Row 108 (MAINT): H34
        - Row 109 (EXP): H36
        - Row 110 (GEN): H32
        - Row 111 (DEP): H33
        - Row 112 (CAPEX): H37
        - Row 113 (DISPOSAL): H31
        """
        ltm_letter = get_column_letter(ltm_col)
        new_bdo_name = self.config.bdo_sheet_name
        period_end = self.config.period_end
        updated_count = 0
        
        # Define the correct BDO row references for each bank account
        # These are the closing balance (column H) rows in the BDO sheet
        bank_account_bdo_rows = {
            107: 'H35',  # ABN AMRO RENT account
            108: 'H34',  # ABN AMRO MAINT account
            109: 'H36',  # ABN AMRO EXP account
            110: 'H32',  # ABN AMRO GEN account
            111: 'H33',  # ABN AMRO DEP account
            112: 'H37',  # ABN AMRO CAPEX account
            113: 'H31',  # ABN AMRO DISPOSAL account
        }
        
        # Define the Bank Account Overview section range
        start_row = 104
        end_row = min(sheet.max_row, 125)  # Safety limit
        
        for row_idx in range(start_row, end_row + 1):
            cell = sheet.cell(row=row_idx, column=ltm_col)
            cell_value = cell.value
            
            # Handle date header (row 106 typically: "Per 30-6-2025")
            if isinstance(cell_value, str) and cell_value.startswith('Per '):
                # Update the date to current quarter end
                new_date_str = f"Per {period_end.day}-{period_end.month}-{period_end.year}"
                cell.value = new_date_str
                updated_count += 1
                logger.debug(f"Row {row_idx}: Updated date header to '{new_date_str}'")
                continue
            
            # Set correct bank account formulas (rows 107-113)
            if row_idx in bank_account_bdo_rows:
                bdo_cell_ref = bank_account_bdo_rows[row_idx]
                correct_formula = f"='{new_bdo_name}'!{bdo_cell_ref}"
                old_value = cell.value
                cell.value = correct_formula
                updated_count += 1
                logger.debug(f"Row {row_idx}: Set formula to '{correct_formula}' (was: {old_value})")
                continue
            
            # Handle SUM formula for Total row (row 114)
            if isinstance(cell_value, str) and cell_value.startswith('=SUM('):
                old_formula = cell_value
                
                # Pattern to match SUM formulas like =SUM(AA107:AA113)
                sum_pattern = r'=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)'
                match = re.match(sum_pattern, old_formula, re.IGNORECASE)
                
                if match:
                    start_row_num = match.group(2)
                    end_row_num = match.group(4)
                    
                    # Set correct SUM formula with LTM column
                    new_formula = f"=SUM({ltm_letter}{start_row_num}:{ltm_letter}{end_row_num})"
                    cell.value = new_formula
                    updated_count += 1
                    logger.debug(f"Row {row_idx}: Updated SUM formula: {old_formula} -> {new_formula}")
        
        if updated_count > 0:
            logger.info(f"Updated {updated_count} cells in Bank Account Overview section (column {ltm_letter})")
    
    def _build_pl_formulas_from_templates(self, sheet, col_idx: int, bdo_column: str = 'G',
                                          prev_col_idx: int = None,
                                          quarter_col_idx: int = None):
        """
        Build P&L formulas for a column using formula templates.
        
        For each row defined in the templates:
        - bdo_ref type: Look up account codes in BDO sheet, build formula with correct row numbers
        - calc type: Replace {COL} placeholder with actual column letter
        - manual type: Copy value from previous column (for Cash proceeds sale, etc.)
        - ltm_sum_quarters: SUM of the last N quarterly columns (for LTM)
        - manual_with_ltm: Manual for quarter columns, ltm_sum_quarters for LTM
        
        Args:
            sheet: The summary sheet
            col_idx: Column index to write formulas to
            bdo_column: BDO column to reference (G for quarter, H for LTM)
            prev_col_idx: Previous column index (for copying manual values)
            quarter_col_idx: New quarter column index (needed for ltm_sum_quarters)
        """
        col_letter = get_column_letter(col_idx)
        bdo_sheet_name = self.config.bdo_sheet_name
        formulas_built = 0
        
        for row_idx, template in self.formula_templates.items():
            formula_type = template.get('type', 'calc')
            target_cell = sheet.cell(row=row_idx, column=col_idx)
            
            if formula_type == 'bdo_ref':
                formula = self._build_bdo_ref_formula(template, bdo_sheet_name, bdo_column)
                if formula:
                    target_cell.value = formula
                    formulas_built += 1
                    
            elif formula_type == 'bdo_ref_label':
                formula = self._build_bdo_ref_label_formula(template, bdo_sheet_name, bdo_column)
                if formula:
                    target_cell.value = formula
                    formulas_built += 1
            
            elif formula_type == 'bdo_ref_conditional':
                if bdo_column == 'G':
                    sub_template = template.get('q_config', {})
                else:
                    sub_template = template.get('ltm_config', {})
                
                sub_type = sub_template.get('type', 'bdo_ref')
                if sub_type == 'bdo_ref':
                    formula = self._build_bdo_ref_formula(sub_template, bdo_sheet_name, bdo_column)
                elif sub_type == 'bdo_ref_label':
                    formula = self._build_bdo_ref_label_formula(sub_template, bdo_sheet_name, bdo_column)
                elif sub_type == 'ltm_sum_quarters':
                    formula = self._build_ltm_sum_quarters_formula(
                        row_idx, sub_template.get('num_quarters', 4), quarter_col_idx
                    )
                else:
                    formula = None
                    
                if formula:
                    target_cell.value = formula
                    formulas_built += 1
                    
            elif formula_type == 'calc':
                pattern = template.get('pattern', '')
                formula = pattern.replace('{COL}', col_letter)
                target_cell.value = formula
                formulas_built += 1
                
            elif formula_type == 'manual':
                pass
            
            elif formula_type == 'manual_with_ltm':
                if bdo_column == 'H' and quarter_col_idx:
                    ltm_type = template.get('ltm_type', '')
                    if ltm_type == 'ltm_sum_quarters':
                        formula = self._build_ltm_sum_quarters_formula(
                            row_idx, template.get('num_quarters', 4), quarter_col_idx
                        )
                        if formula:
                            target_cell.value = formula
                            formulas_built += 1
            
            elif formula_type == 'constant':
                target_cell.value = template.get('value', 0)
                formulas_built += 1
            
            if prev_col_idx:
                prev_cell = sheet.cell(row=row_idx, column=prev_col_idx)
                if prev_cell.number_format:
                    target_cell.number_format = prev_cell.number_format
        
        logger.info(f"Built {formulas_built} P&L formulas for column {col_letter} (BDO col {bdo_column})")
    
    def _build_ltm_sum_quarters_formula(self, row_idx: int, num_quarters: int,
                                         quarter_col_idx: int) -> Optional[str]:
        """
        Build =SUM(start:end) formula covering the last N quarterly columns.
        
        The quarter_col_idx is the newest quarter column. The formula sums from
        (quarter_col_idx - num_quarters + 1) to quarter_col_idx.
        """
        if not quarter_col_idx:
            logger.warning(f"Row {row_idx}: Cannot build ltm_sum_quarters without quarter_col_idx")
            return None
        start_col = quarter_col_idx - num_quarters + 1
        if start_col < 1:
            start_col = 1
        start_letter = get_column_letter(start_col)
        end_letter = get_column_letter(quarter_col_idx)
        formula = f"=SUM({start_letter}{row_idx}:{end_letter}{row_idx})"
        logger.debug(f"Row {row_idx}: LTM SUM formula: {formula}")
        return formula
    
    def _build_bdo_ref_formula(self, template: Dict[str, Any], bdo_sheet_name: str, 
                                bdo_column: str) -> Optional[str]:
        """
        Build a formula that references BDO sheet cells by account codes.
        
        Example template:
            account_codes: ["4100410", "4101010", "4101020"]
            sign: "-"
            operator: "-"
        
        Result: =-'BDO - Q3-25'!G91-'BDO - Q3-25'!G92-'BDO - Q3-25'!G93
        """
        account_codes = template.get('account_codes', [])
        sign = template.get('sign', '-')
        operator = template.get('operator', '-')
        
        if not account_codes:
            return None
        
        refs = []
        for code in account_codes:
            row_num = self._new_bdo_row_map.get(code)
            if row_num:
                refs.append(f"'{bdo_sheet_name}'!{bdo_column}{row_num}")
            else:
                # Try to find by partial match or similar code
                row_num = self._find_account_row(code)
                if row_num:
                    refs.append(f"'{bdo_sheet_name}'!{bdo_column}{row_num}")
                else:
                    logger.warning(f"Account code {code} not found in BDO sheet")
        
        if not refs:
            return None
        
        # Build formula with sign and operator
        if len(refs) == 1:
            return f"={sign}{refs[0]}"
        else:
            return f"={sign}{operator.join(refs)}"
    
    def _build_bdo_ref_label_formula(self, template: Dict[str, Any], bdo_sheet_name: str,
                                      bdo_column: str) -> Optional[str]:
        """
        Build a formula that references a BDO sheet cell by label search.
        
        Example template:
            label_search: "SWAP"
            sign: "-"
        
        Result: =-'BDO - Q3-25'!G133
        
        Searches in column A of the BDO sheet for exact or partial matches.
        """
        label_search = template.get('label_search', '')
        sign = template.get('sign', '-')
        
        if not label_search:
            return None
        
        label_search_lower = label_search.lower().strip()
        
        # First try exact match in label map
        row_num = self._new_bdo_label_map.get(label_search_lower)
        
        if not row_num:
            # Try exact match with the full text (case-insensitive)
            for label, row in self._new_bdo_label_map.items():
                if label == label_search_lower:
                    row_num = row
                    break
        
        if not row_num:
            # Try partial match only if exact match not found
            # Be careful with partial matches - "SWAP" shouldn't match "Swap fee"
            # Only match if the label is at the start or is the complete text
            for label, row in self._new_bdo_label_map.items():
                # Match if label is exactly the search term (already checked above)
                # or if search term is at the beginning followed by space or end
                if label.startswith(label_search_lower + ' ') or label.startswith(label_search_lower + '\t'):
                    row_num = row
                    break
        
        if row_num:
            return f"={sign}'{bdo_sheet_name}'!{bdo_column}{row_num}"
        else:
            logger.warning(f"Label '{label_search}' not found in BDO sheet")
            return None
    
    def _find_account_row(self, account_code: str) -> Optional[int]:
        """Find a row for an account code, trying partial matches if exact not found."""
        # First try exact match
        if account_code in self._new_bdo_row_map:
            return self._new_bdo_row_map[account_code]
        
        # Try prefix match (e.g., "4100400" might be stored as "4100400.0")
        for code, row in self._new_bdo_row_map.items():
            if code.startswith(account_code) or account_code.startswith(code.split('.')[0]):
                return row
        
        return None
    
    def _build_balance_sheet_formulas(self, sheet, col_idx: int, bdo_column: str = 'H'):
        """
        Build Balance Sheet formulas for the LTM column using templates.
        
        Balance Sheet section (rows 3-19) only has formulas in the LTM column.
        Uses balance_sheet_templates to build formulas with correct BDO row references.
        """
        col_letter = get_column_letter(col_idx)
        bdo_sheet_name = self.config.bdo_sheet_name
        formulas_built = 0
        
        for row_idx, template in self.balance_sheet_templates.items():
            formula_type = template.get('type', 'bdo_ref')
            target_cell = sheet.cell(row=row_idx, column=col_idx)
            
            if formula_type == 'bdo_ref':
                # Build formula from account codes
                formula = self._build_balance_sheet_ref_formula(template, bdo_sheet_name, bdo_column)
                if formula:
                    target_cell.value = formula
                    formulas_built += 1
                    
            elif formula_type == 'bdo_sum_range':
                # Build SUM formula for a range of accounts
                formula = self._build_balance_sheet_sum_formula(template, bdo_sheet_name, bdo_column)
                if formula:
                    target_cell.value = formula
                    formulas_built += 1
                    
            elif formula_type == 'calc':
                # Replace {COL} with actual column letter
                pattern = template.get('pattern', '')
                formula = pattern.replace('{COL}', col_letter)
                target_cell.value = formula
                formulas_built += 1
        
        logger.info(f"Built {formulas_built} Balance Sheet formulas in column {col_letter}")
    
    def _build_balance_sheet_ref_formula(self, template: Dict[str, Any], bdo_sheet_name: str,
                                          bdo_column: str) -> Optional[str]:
        """Build a Balance Sheet formula that references BDO cells by account codes."""
        account_codes = template.get('account_codes', [])
        sign = template.get('sign', '')
        operator = template.get('operator', '+')
        
        if not account_codes:
            return None
        
        refs = []
        for code in account_codes:
            row_num = self._find_account_row(code)
            if row_num:
                refs.append(f"'{bdo_sheet_name}'!{bdo_column}{row_num}")
            else:
                logger.debug(f"Balance Sheet: Account code {code} not found in BDO sheet")
        
        if not refs:
            return None
        
        if len(refs) == 1:
            return f"={sign}{refs[0]}"
        else:
            return f"={sign}{operator.join(refs)}"
    
    def _build_balance_sheet_sum_formula(self, template: Dict[str, Any], bdo_sheet_name: str,
                                          bdo_column: str) -> Optional[str]:
        """Build a Balance Sheet SUM formula for a range of accounts."""
        start_account = template.get('start_account')
        end_account = template.get('end_account')
        count = template.get('count')  # Optional: number of rows in range
        additional = template.get('additional', [])
        
        start_row = self._find_account_row(start_account) if start_account else None
        
        # Calculate end_row
        if count and start_row:
            # If count is specified, calculate end_row from start_row + count - 1
            end_row = start_row + count - 1
        elif end_account:
            # Find end account row
            end_row = self._find_account_row(end_account)
        else:
            end_row = None
        
        parts = []
        
        if start_row and end_row and end_row > start_row:
            # Range with multiple cells
            parts.append(f"SUM('{bdo_sheet_name}'!{bdo_column}{start_row}:{bdo_column}{end_row})")
        elif start_row and count == 1:
            # Single cell SUM (for consistency with expected format)
            parts.append(f"SUM('{bdo_sheet_name}'!{bdo_column}{start_row})")
        elif start_row:
            parts.append(f"'{bdo_sheet_name}'!{bdo_column}{start_row}")
        
        # Add additional accounts
        for code in additional:
            row_num = self._find_account_row(code)
            if row_num:
                parts.append(f"'{bdo_sheet_name}'!{bdo_column}{row_num}")
        
        if not parts:
            return None
        
        return "=" + "+".join(parts)
    
    def _update_balance_sheet_bdo_refs(self, sheet, col_idx: int, bdo_column: str = 'G'):
        """
        Update Balance Sheet section (rows 1-20) BDO references to point to new BDO sheet.
        This preserves the existing formula structure but updates the sheet name.
        """
        bdo_sheet_name = self.config.bdo_sheet_name
        updated_count = 0
        
        for row_idx in range(1, 21):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                old_formula = cell.value
                if 'BDO' in old_formula:
                    # Update BDO sheet name and column
                    new_formula = re.sub(
                        r"'BDO[^']+'\!([GH])(\d+)",
                        lambda m: f"'{bdo_sheet_name}'!{bdo_column}{m.group(2)}",
                        old_formula
                    )
                    if new_formula != old_formula:
                        cell.value = new_formula
                        updated_count += 1
        
        if updated_count > 0:
            logger.info(f"Updated {updated_count} Balance Sheet BDO refs in column {get_column_letter(col_idx)}")
    
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
                target_cell.fill = copy(source_cell.fill)
        
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
                target_cell.fill = copy(source_cell.fill)
        
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
        """Update date references only in the Management Cijfers summary sheet.
        
        IMPORTANT: Only the summary sheet is modified. All other sheets
        (especially BDO/kwartaal sheets) are protected and never touched.
        """
        old_date_patterns = self._generate_date_patterns(self._get_previous_period_end())
        new_date_patterns = self._generate_date_patterns(self.config.period_end)
        
        target_sheet_name = self.config.summary_sheet_name
        if target_sheet_name not in self.workbook.sheetnames:
            logger.warning(f"Summary sheet '{target_sheet_name}' not found for date reference update")
            return
        
        sheet = self.workbook[target_sheet_name]
        
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
    
    def _compute_bdo_ref_value(self, template: dict, bdo_result: BDOParseResult) -> Tuple[float, List[str]]:
        """
        Compute the numeric value a bdo_ref formula would produce from raw BDO data.
        
        Mirrors the Excel formula built by _build_bdo_ref_formula, but returns
        a computed float instead of a formula string.
        
        Returns (value, list_of_missing_account_codes).
        """
        account_codes = template.get('account_codes', [])
        sign = template.get('sign', '-')
        operator = template.get('operator', '-')
        
        if not account_codes:
            return 0.0, []
        
        vals = []
        missing = []
        for code in account_codes:
            entry = bdo_result.accounts.get(code)
            if entry:
                vals.append(entry.closing_balance - entry.opening_balance)
            else:
                vals.append(0.0)
                missing.append(code)
        
        result = -vals[0] if sign == '-' else vals[0]
        for v in vals[1:]:
            if operator == '-':
                result -= v
            else:
                result += v
        
        return result, missing
    
    def _compute_bdo_sum_range_value(self, template: dict, bdo_result: BDOParseResult,
                                       use_mutation: bool = True) -> Tuple[float, List[str]]:
        """
        Compute the numeric value a bdo_sum_range formula would produce from raw BDO data.

        Uses raw_row from AccountEntry to replicate the Excel SUM range logic.
        When use_mutation=True, computes closing_balance - opening_balance (period change).
        """
        start_account = template.get('start_account')
        end_account = template.get('end_account')
        count = template.get('count')
        additional = template.get('additional', [])

        total = 0.0
        missing = []

        if start_account:
            start_entry = bdo_result.accounts.get(start_account)
            if start_entry:
                row_to_entry = {}
                for entry in bdo_result.accounts.values():
                    row_to_entry[entry.raw_row] = entry

                start_raw = start_entry.raw_row
                if count:
                    for r in range(start_raw, start_raw + count):
                        if r in row_to_entry:
                            e = row_to_entry[r]
                            total += (e.closing_balance - e.opening_balance) if use_mutation else e.closing_balance
                elif end_account:
                    end_entry = bdo_result.accounts.get(end_account)
                    if end_entry:
                        for r in range(start_raw, end_entry.raw_row + 1):
                            if r in row_to_entry:
                                e = row_to_entry[r]
                                total += (e.closing_balance - e.opening_balance) if use_mutation else e.closing_balance
                    else:
                        missing.append(end_account)
                else:
                    total += (start_entry.closing_balance - start_entry.opening_balance) if use_mutation else start_entry.closing_balance
            else:
                missing.append(start_account)

        for code in additional:
            entry = bdo_result.accounts.get(code)
            if entry:
                total += (entry.closing_balance - entry.opening_balance) if use_mutation else entry.closing_balance
            else:
                missing.append(code)

        return total, missing

    def _evaluate_calc_pattern(self, pattern: str, shadow: dict) -> Optional[float]:
        """
        Evaluate a calc-type formula pattern using pre-computed shadow values.
        
        Supports: SUM range, addition chains, negated subtraction, division.
        Returns None if the pattern cannot be parsed.
        """
        p = pattern.strip()
        
        sum_match = re.match(r'^=SUM\(\{COL\}(\d+):\{COL\}(\d+)\)$', p)
        if sum_match:
            start_row = int(sum_match.group(1))
            end_row = int(sum_match.group(2))
            total = 0.0
            for r in range(start_row, end_row + 1):
                if r in shadow:
                    total += shadow[r]['value']
            return total
        
        neg_sub_match = re.match(r'^=-\(\{COL\}(\d+)-\{COL\}(\d+)\)$', p)
        if neg_sub_match:
            r1 = int(neg_sub_match.group(1))
            r2 = int(neg_sub_match.group(2))
            v1 = shadow.get(r1, {}).get('value', 0.0)
            v2 = shadow.get(r2, {}).get('value', 0.0)
            return -(v1 - v2)
        
        div_match = re.match(r'^=\{COL\}(\d+)/\{COL\}(\d+)-1$', p)
        if div_match:
            r1 = int(div_match.group(1))
            r2 = int(div_match.group(2))
            v1 = shadow.get(r1, {}).get('value', 0.0)
            v2 = shadow.get(r2, {}).get('value', 0.0)
            if v2 != 0:
                return v1 / v2 - 1
            return 0.0
        
        add_refs = re.findall(r'\{COL\}(\d+)', p)
        if add_refs and '+' in p and '-' not in p.replace('=-', ''):
            total = 0.0
            for ref in add_refs:
                r = int(ref)
                total += shadow.get(r, {}).get('value', 0.0)
            return total
        
        if add_refs and len(add_refs) >= 2:
            total = 0.0
            for ref in add_refs:
                r = int(ref)
                total += shadow.get(r, {}).get('value', 0.0)
            return total
        
        logger.debug(f"Could not parse calc pattern: {p}")
        return None
    
    def _compute_shadow_pl(self, bdo_result: BDOParseResult) -> dict:
        """
        Compute a 'shadow P&L' from BDO raw data, mirroring the formula chain.
        
        For each row in the formula templates, computes the expected numeric value
        that the Excel formula would produce if evaluated. This allows comparing
        the P&L chain output against equity movement to find divergence points.
        
        Returns: {row_num: {'label': str, 'value': float, 'type': str,
                            'account_codes': list, 'missing_codes': list}, ...}
        """
        shadow = {}
        
        for row_idx in sorted(self.formula_templates.keys()):
            template = self.formula_templates[row_idx]
            formula_type = template.get('type', 'calc')
            label = template.get('label', f'Row {row_idx}')
            entry = {'label': label, 'value': 0.0, 'type': formula_type,
                     'account_codes': [], 'missing_codes': []}
            
            if formula_type == 'bdo_ref':
                val, missing = self._compute_bdo_ref_value(template, bdo_result)
                entry['value'] = val
                entry['account_codes'] = template.get('account_codes', [])
                entry['missing_codes'] = missing
            
            elif formula_type == 'bdo_ref_conditional':
                sub = template.get('q_config', {})
                sub_type = sub.get('type', 'bdo_ref')
                if sub_type == 'bdo_ref':
                    val, missing = self._compute_bdo_ref_value(sub, bdo_result)
                    entry['value'] = val
                    entry['account_codes'] = sub.get('account_codes', [])
                    entry['missing_codes'] = missing
                else:
                    entry['value'] = 0.0
            
            elif formula_type == 'calc':
                pattern = template.get('pattern', '')
                val = self._evaluate_calc_pattern(pattern, shadow)
                entry['value'] = val if val is not None else 0.0
            
            elif formula_type == 'constant':
                entry['value'] = float(template.get('value', 0))
            
            elif formula_type in ('manual', 'manual_with_ltm'):
                if row_idx == 50:
                    entry['value'] = self.config.cash_proceeds_sale or 0.0
                else:
                    entry['value'] = 0.0
            
            shadow[row_idx] = entry
        
        logger.info(f"Shadow P&L computed for {len(shadow)} rows")
        return shadow
    
    def _compute_shadow_bs(self, bdo_result: BDOParseResult) -> dict:
        """
        Compute a 'shadow Balance Sheet' from BDO raw data using balance_sheet_templates.

        For each row in the balance sheet templates (rows 3-18), computes the expected
        numeric value using the same account codes as the Excel formulas. Row 19 is
        the SUM of rows 3-18, representing the Total Equity Movement.

        Uses mutations (closing_balance - opening_balance) so the result is comparable
        to the P&L-based Direct Result (row 68).

        Returns: {row_num: {'label': str, 'value': float, 'type': str,
                            'account_codes': list, 'missing_codes': list}, ...}
        """
        shadow = {}

        for row_idx in sorted(self.balance_sheet_templates.keys()):
            template = self.balance_sheet_templates[row_idx]
            formula_type = template.get('type', 'bdo_ref')
            label = template.get('label', f'Row {row_idx}')
            entry = {'label': label, 'value': 0.0, 'type': formula_type,
                     'account_codes': [], 'missing_codes': []}

            if formula_type == 'bdo_ref':
                val, missing = self._compute_bdo_ref_value(template, bdo_result)
                entry['value'] = val
                entry['account_codes'] = template.get('account_codes', [])
                entry['missing_codes'] = missing

            elif formula_type == 'bdo_sum_range':
                val, missing = self._compute_bdo_sum_range_value(template, bdo_result)
                entry['value'] = val
                codes = []
                if template.get('start_account'):
                    codes.append(template['start_account'])
                if template.get('end_account'):
                    codes.append(f"..{template['end_account']}")
                codes.extend(template.get('additional', []))
                entry['account_codes'] = codes
                entry['missing_codes'] = missing

            elif formula_type == 'calc':
                pattern = template.get('pattern', '')
                val = self._evaluate_calc_pattern(pattern, shadow)
                entry['value'] = val if val is not None else 0.0

            shadow[row_idx] = entry

        logger.info(f"Shadow Balance Sheet computed for {len(shadow)} rows, "
                     f"row 19 = {shadow.get(19, {}).get('value', 'N/A')}")
        return shadow
    
    def _reconcile_pl_chain(self, shadow: dict, equity_movement: float) -> dict:
        """
        Compare shadow P&L chain against equity movement to find divergence.
        
        Walks backwards from row 68 (Direct Result) through the formula chain
        to identify which specific row(s) contribute to any mismatch.
        """
        shadow_direct = shadow.get(68, {}).get('value', 0.0)
        tolerance = 1.0
        difference = abs(shadow_direct - equity_movement)
        is_reconciled = difference <= tolerance
        
        key_rows = [68, 66, 57, 55, 48, 45, 25, 67, 60, 61, 64, 65, 56, 53, 50,
                    44, 30, 23, 24, 27, 28, 29, 46, 47]
        
        breakdown = []
        for r in key_rows:
            if r in shadow:
                s = shadow[r]
                breakdown.append({
                    'row': r,
                    'label': s['label'],
                    'value': s['value'],
                    'type': s['type'],
                    'missing_codes': s.get('missing_codes', []),
                })
        
        divergence_point = None
        if not is_reconciled:
            chain_pairs = [
                (68, [67, 66], 'Direct result = Delta DTA + EBT'),
                (66, list(range(57, 66)), 'EBT = SUM(57:65)'),
                (57, [56, 55], 'EBIT = Depreciation + Total EBITDA'),
                (55, [53, 48], 'Total EBITDA = Sales + Rental'),
                (48, [47, 46, 45], 'EBITDA = Mgmt fees + Net rental'),
                (45, [25, 30, 44], 'Net rental = Gross + Service + Property'),
            ]
            
            for parent_row, child_rows, desc in chain_pairs:
                parent_val = shadow.get(parent_row, {}).get('value', 0.0)
                children_sum = sum(
                    shadow.get(r, {}).get('value', 0.0) for r in child_rows
                    if r in shadow
                )
                if abs(parent_val - children_sum) > tolerance:
                    divergence_point = {
                        'row': parent_row,
                        'label': shadow.get(parent_row, {}).get('label', ''),
                        'expected': children_sum,
                        'actual': parent_val,
                        'detail': f"{desc}: children sum={children_sum:,.2f}, parent={parent_val:,.2f}"
                    }
                    break
            
            if not divergence_point:
                bdo_ref_rows = [r for r in shadow if shadow[r]['type'] in ('bdo_ref', 'bdo_ref_conditional')]
                for r in bdo_ref_rows:
                    if shadow[r].get('missing_codes'):
                        codes_str = ', '.join(shadow[r]['missing_codes'])
                        divergence_point = {
                            'row': r,
                            'label': shadow[r]['label'],
                            'expected': None,
                            'actual': shadow[r]['value'],
                            'detail': f"Missing account codes in BDO data: {codes_str}"
                        }
                        break
        
        result = {
            'shadow_direct_result': shadow_direct,
            'equity_movement': equity_movement,
            'is_reconciled': is_reconciled,
            'difference': difference,
            'row_breakdown': breakdown,
            'divergence_point': divergence_point,
        }
        
        if is_reconciled:
            logger.info(f"Shadow P&L reconciliation passed: shadow={shadow_direct:,.2f}, equity={equity_movement:,.2f}")
        else:
            logger.warning(
                f"Shadow P&L reconciliation FAILED: shadow={shadow_direct:,.2f}, "
                f"equity={equity_movement:,.2f}, diff={difference:,.2f}"
            )
            if divergence_point:
                logger.warning(f"  Divergence at row {divergence_point['row']}: {divergence_point['detail']}")
        
        return result
    
    def _validate_calculations(self, bdo_result: BDOParseResult) -> dict:
        """
        Validate that calculations are correct and return structured results.
        
        Three-pass validation:
        1. Template-based check: shadow BS row 19 vs shadow P&L row 68
           (uses the exact same account codes as the Excel formulas)
        2. Structural formula check: verify key Excel formulas are present and correct
        3. Shadow reconciliation: row-by-row check to pinpoint divergence
        
        If any check fails, writes warning rows into the Management Cijfers sheet.
        """
        # --- Compute shadow models from BDO data using template account codes ---
        shadow_pl = self._compute_shadow_pl(bdo_result)
        shadow_bs = self._compute_shadow_bs(bdo_result)
        
        direct_result = shadow_pl.get(68, {}).get('value', 0.0)
        equity_movement = shadow_bs.get(19, {}).get('value', 0.0)
        
        tolerance = 1.0
        difference = abs(equity_movement - direct_result)
        is_aligned = difference <= tolerance
        
        validation = {
            'equity_movement': equity_movement,
            'direct_result': direct_result,
            'is_aligned': is_aligned,
            'difference': difference,
            'formula_checks': {},
            'shadow_bs': shadow_bs,
            'shadow_pl': shadow_pl,
            'messages': []
        }
        
        if not is_aligned:
            msg = (
                f"Template check: Total Equity Movement (BS row 19 = {equity_movement:,.2f}) != "
                f"Direct Result (P&L row 68 = {direct_result:,.2f}), difference: {difference:,.2f}"
            )
            validation['messages'].append(msg)
            logger.warning(f"VALIDATION WARNING: {msg}")
            
            bs_missing = []
            for r in sorted(shadow_bs.keys()):
                if r == 19:
                    continue
                codes = shadow_bs[r].get('missing_codes', [])
                if codes:
                    bs_missing.append(f"BS row {r} ({shadow_bs[r]['label']}): {', '.join(codes)}")
            if bs_missing:
                validation['messages'].append(
                    f"Missing BDO accounts in Balance Sheet: {'; '.join(bs_missing)}"
                )
        else:
            logger.info(
                f"Template check passed: BS row 19 = {equity_movement:,.2f} "
                f"≈ P&L row 68 = {direct_result:,.2f}"
            )
        
        # --- Pass 1b: BDO ground truth check ---
        # Profits are stored as NEGATIVE in BDO (Dutch trial balance convention).
        # Management Cijfers uses positive-for-profit (formula templates negate via sign: "-").
        # The ground truth comparison must account for this sign difference.
        bdo_resultaat_raw = self._read_bdo_resultaat_na_belasting()
        validation['bdo_resultaat_raw'] = bdo_resultaat_raw
        
        if bdo_resultaat_raw is not None:
            # Negate BDO value: BDO stores profit as negative, MC stores as positive
            bdo_resultaat = -bdo_resultaat_raw
            validation['bdo_resultaat'] = bdo_resultaat
            
            diff_dr = abs(bdo_resultaat - direct_result)
            diff_eq = abs(bdo_resultaat - equity_movement)
            dr_ok = diff_dr <= tolerance
            eq_ok = diff_eq <= tolerance
            
            validation['bdo_vs_direct_result'] = diff_dr
            validation['bdo_vs_equity_movement'] = diff_eq
            
            if not dr_ok or not eq_ok:
                validation['is_aligned'] = False
                validation['messages'].append(
                    f"BDO ground truth: Resultaat na belasting (kwartaal D-G) = {bdo_resultaat_raw:,.2f} "
                    f"(sign-adjusted: {bdo_resultaat:,.2f})"
                )
                if not dr_ok:
                    validation['messages'].append(
                        f"  Expected (BDO sign-adjusted): {bdo_resultaat:,.2f}"
                    )
                    validation['messages'].append(
                        f"  Current  (Direct Result P&L row 68): {direct_result:,.2f}"
                    )
                    validation['messages'].append(
                        f"  Difference: {diff_dr:,.2f}"
                    )
                if not eq_ok:
                    validation['messages'].append(
                        f"  Expected (BDO sign-adjusted): {bdo_resultaat:,.2f}"
                    )
                    validation['messages'].append(
                        f"  Current  (Equity Movement BS row 19): {equity_movement:,.2f}"
                    )
                    validation['messages'].append(
                        f"  Difference: {diff_eq:,.2f}"
                    )
                logger.warning(
                    f"VALIDATION WARNING: BDO Resultaat na belasting "
                    f"(raw={bdo_resultaat_raw:,.2f}, adjusted={bdo_resultaat:,.2f}) "
                    f"differs from Direct Result ({direct_result:,.2f}, Δ={diff_dr:,.2f}) "
                    f"and/or Equity Movement ({equity_movement:,.2f}, Δ={diff_eq:,.2f})"
                )
            else:
                logger.info(
                    f"BDO ground truth check passed: Resultaat(raw={bdo_resultaat_raw:,.2f}, "
                    f"adjusted={bdo_resultaat:,.2f}) ≈ DR={direct_result:,.2f} ≈ EM={equity_movement:,.2f}"
                )
        else:
            validation['bdo_resultaat'] = None
            validation['messages'].append(
                "BDO ground truth: Could not locate 'Resultaat na belasting' row in BDO sheet"
            )
        
        # --- Pass 2: Structural formula check ---
        formula_ok = self._validate_structural_formulas(validation)
        
        # --- Pass 3: Shadow reconciliation (detailed row-by-row) ---
        reconciliation = self._reconcile_pl_chain(shadow_pl, equity_movement)
        validation['reconciliation'] = reconciliation
        
        if not reconciliation['is_reconciled']:
            validation['messages'].append(
                f"Row-by-row: P&L chain result = {reconciliation['shadow_direct_result']:,.2f}, "
                f"BS equity movement = {reconciliation['equity_movement']:,.2f}, "
                f"difference = {reconciliation['difference']:,.2f}"
            )
            
            if reconciliation.get('divergence_point'):
                dp = reconciliation['divergence_point']
                validation['messages'].append(
                    f"Divergence point: Row {dp['row']} ({dp['label']}) - {dp['detail']}"
                )
            
            for row_info in reconciliation['row_breakdown']:
                if row_info.get('missing_codes'):
                    codes = ', '.join(row_info['missing_codes'])
                    validation['messages'].append(
                        f"Row {row_info['row']} ({row_info['label']}): "
                        f"missing BDO accounts: {codes}"
                    )
            
            validation['is_aligned'] = False
        
        if not validation['is_aligned'] or not formula_ok or not reconciliation['is_reconciled']:
            self._add_validation_warning_to_sheet(validation['messages'])
        
        return validation
    
    def _read_bdo_resultaat_na_belasting(self) -> Optional[float]:
        """
        Read the 'Resultaat na belasting' ground truth value from the BDO sheet.

        Locates the row by label, then for each of columns D-G either reads the
        numeric value directly or, if the cell contains a SUM formula (e.g.
        =SUM(D76:D123)), evaluates it by summing the referenced cell range.
        Returns None if the row or sheet cannot be found.
        """
        bdo_sheet_name = self.config.bdo_sheet_name
        if bdo_sheet_name not in self.workbook.sheetnames:
            logger.warning("BDO ground truth check: sheet not found")
            return None

        sheet = self.workbook[bdo_sheet_name]

        target_row = None

        label_key = 'resultaat na belasting'
        if label_key in self._new_bdo_label_map:
            target_row = self._new_bdo_label_map[label_key]
        else:
            for row_idx in range(1, min(sheet.max_row + 1, 140)):
                for col in (1, 2):
                    cell_val = sheet.cell(row=row_idx, column=col).value
                    if cell_val and isinstance(cell_val, str):
                        if 'resultaat na belasting' in cell_val.lower():
                            target_row = row_idx
                            break
                if target_row:
                    break

        if target_row is None:
            logger.warning("BDO ground truth check: 'Resultaat na belasting' row not found")
            return None

        total = 0.0
        for col_idx in range(4, 8):  # columns D(4), E(5), F(6), G(7)
            val = sheet.cell(row=target_row, column=col_idx).value
            if isinstance(val, (int, float)):
                total += val
            elif isinstance(val, str) and val.startswith('='):
                total += self._eval_simple_sum_formula(sheet, val, col_idx)
            elif isinstance(val, str):
                try:
                    total += float(val.replace(',', '.').replace(' ', ''))
                except (ValueError, AttributeError):
                    pass

        logger.info(f"BDO ground truth: Resultaat na belasting (row {target_row}, D-G) = {total:,.2f}")
        return total

    def _eval_simple_sum_formula(self, sheet, formula: str, default_col: int) -> float:
        """
        Evaluate a simple SUM formula like =SUM(D76:D123) by reading cell values.

        Only supports =SUM(COLn:COLm) patterns. Returns 0.0 for unparseable formulas.
        """
        m = re.match(r'^=SUM\(([A-Z])(\d+):([A-Z])(\d+)\)$', formula, re.IGNORECASE)
        if not m:
            logger.debug(f"Cannot evaluate formula: {formula}")
            return 0.0

        col_letter_start = m.group(1).upper()
        row_start = int(m.group(2))
        row_end = int(m.group(4))
        col_num = ord(col_letter_start) - ord('A') + 1

        total = 0.0
        for r in range(row_start, row_end + 1):
            cell_val = sheet.cell(row=r, column=col_num).value
            if isinstance(cell_val, (int, float)):
                total += cell_val
        return total

    def _validate_structural_formulas(self, validation: dict) -> bool:
        """
        Verify that key Excel formulas in Management Cijfers are structurally correct.
        
        Checks row 19 (Total Equity Movement), row 66 (EBT), row 68 (Direct Result),
        and intermediate P&L rows (60, 61, 64, 65, 67) for the new quarter and LTM columns.
        
        Returns True if all formula checks pass.
        """
        summary_sheets = [name for name in self.workbook.sheetnames
                          if 'Management Cijfers' in name]
        if not summary_sheets:
            logger.warning("Structural validation skipped: Management Cijfers sheet not found")
            return True
        
        sheet = self.workbook[summary_sheets[-1]]
        
        ltm_col = self._find_ltm_column(sheet)
        if not ltm_col:
            logger.warning("Structural validation skipped: LTM column not found")
            return True
        
        quarter_col = ltm_col - 1
        q_letter = get_column_letter(quarter_col)
        ltm_letter = get_column_letter(ltm_col)
        
        all_ok = True
        checks = {}
        
        expected_formulas = {
            19: {'col': ltm_col, 'letter': ltm_letter,
                 'expected': f'=SUM({ltm_letter}3:{ltm_letter}18)',
                 'label': 'Total Equity Movement (LTM)'},
            66: {'col': quarter_col, 'letter': q_letter,
                 'expected': f'=SUM({q_letter}57:{q_letter}65)',
                 'label': 'EBT (quarter)'},
            68: {'col': quarter_col, 'letter': q_letter,
                 'expected': f'={q_letter}67+{q_letter}66',
                 'label': 'Direct Result (quarter)'},
        }
        
        for row_num, spec in expected_formulas.items():
            cell = sheet.cell(row=row_num, column=spec['col'])
            actual = str(cell.value) if cell.value else '(empty)'
            ok = actual.upper() == spec['expected'].upper()
            checks[f'row_{row_num}'] = {
                'label': spec['label'],
                'expected': spec['expected'],
                'actual': actual,
                'ok': ok
            }
            if not ok:
                all_ok = False
                msg = (
                    f"Formula check failed: {spec['label']} row {row_num} col {spec['letter']} - "
                    f"expected '{spec['expected']}', got '{actual}'"
                )
                validation['messages'].append(msg)
                logger.warning(f"VALIDATION WARNING: {msg}")
        
        intermediate_rows = {
            60: 'Net interest expenses - SFA',
            61: 'Prepaid derivatives - SFA',
            64: 'Net interest expenses - Hedge',
            65: 'Net interest income - DMRRP',
            67: 'Delta DTA & CIT',
        }
        for row_num, label in intermediate_rows.items():
            q_cell = sheet.cell(row=row_num, column=quarter_col)
            ltm_cell = sheet.cell(row=row_num, column=ltm_col)
            q_val = q_cell.value
            ltm_val = ltm_cell.value
            
            q_ok = q_val is not None
            ltm_ok = ltm_val is not None
            checks[f'row_{row_num}_q'] = {
                'label': f'{label} (quarter)',
                'has_formula': q_ok,
                'actual': str(q_val) if q_val else '(empty)',
                'ok': q_ok
            }
            checks[f'row_{row_num}_ltm'] = {
                'label': f'{label} (LTM)',
                'has_formula': ltm_ok,
                'actual': str(ltm_val) if ltm_val else '(empty)',
                'ok': ltm_ok
            }
            if not q_ok:
                all_ok = False
                msg = f"Formula missing: {label} (quarter) row {row_num} col {q_letter} is empty"
                validation['messages'].append(msg)
                logger.warning(f"VALIDATION WARNING: {msg}")
            if not ltm_ok:
                all_ok = False
                msg = f"Formula missing: {label} (LTM) row {row_num} col {ltm_letter} is empty"
                validation['messages'].append(msg)
                logger.warning(f"VALIDATION WARNING: {msg}")
        
        validation['formula_checks'] = checks
        
        if all_ok:
            logger.info("Structural formula validation passed: all key formulas present and correct")
        else:
            validation['is_aligned'] = False
        
        return all_ok
    
    def _add_validation_warning_to_sheet(self, messages: list):
        """Write visible red warning rows in Management Cijfers if validation fails."""
        summary_sheets = [name for name in self.workbook.sheetnames
                          if 'Management Cijfers' in name]
        if not summary_sheets:
            return
        
        sheet = self.workbook[summary_sheets[-1]]
        warning_row = 70
        
        red_font = Font(name='Calibri', size=11, bold=True, color='FF0000')
        red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
        
        header_cell = sheet.cell(row=warning_row, column=2)
        header_cell.value = "VALIDATION WARNING: Reconciliation check failed — see details below"
        header_cell.font = red_font
        header_cell.fill = red_fill
        
        for i, msg in enumerate(messages):
            detail_cell = sheet.cell(row=warning_row + 1 + i, column=2)
            detail_cell.value = f"  - {msg}"
            detail_cell.font = Font(name='Calibri', size=10, color='FF0000')
        
        logger.info(f"Added validation warnings in Management Cijfers rows {warning_row}-{warning_row + len(messages)}")

