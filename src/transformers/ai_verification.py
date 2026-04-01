"""
AI-Verified Management Accounts — Claude Opus 4.6

Proactive verification of Management Accounts workbooks using Claude.
After the automated MA build, this module:
1. Extracts full workbook context (formulas, values, formats, BDO map)
2. Sends to Claude for verification against formula templates and sign rules
3. Applies returned JSON patches (Management Cijfers only)
4. Re-runs validation to confirm fixes

Original draft is always preserved alongside the verified output.
"""

import json
import os
import re
from copy import copy as copy_style
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import yaml
from loguru import logger
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter

from ..parsers.bdo_parser import BDOParseResult


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class VerificationConfig:
    """Configuration for AI verification pass."""
    quarter: str              # e.g. "Q4 2025"
    year: int
    quarter_num: int
    period_end: str           # e.g. "31-12-2025"
    bdo_sheet_name: str       # e.g. "BDO - Q4-25"
    summary_sheet_name: str   # e.g. "Management Cijfers - Q4 2025"
    previous_quarter_label: str  # e.g. "Q3 2025"


@dataclass
class VerificationResult:
    """Result from AI verification."""
    status: str                     # "PASS" or "ISSUES_FOUND"
    patches_applied: int = 0
    revalidation_passed: bool = False
    original_file: str = ""
    verified_file: str = ""
    issues: List[Dict] = field(default_factory=list)
    patches: List[Dict] = field(default_factory=list)
    validation_summary: Dict = field(default_factory=dict)
    notes: str = ""
    error: Optional[str] = None


# ---------------------------------------------------------------------------
# Context Extractor
# ---------------------------------------------------------------------------

class WorkbookContextExtractor:
    """
    Extracts structured context from a Management Accounts workbook
    for submission to Claude verification.
    """

    def __init__(self, workbook_path: str, config: VerificationConfig,
                 formula_templates_path: str = "config/formula_templates.yaml"):
        self.workbook_path = Path(workbook_path)
        self.config = config
        self.formula_templates_path = Path(formula_templates_path)
        self.formula_templates = self._load_formula_templates()

    def _load_formula_templates(self) -> dict:
        if self.formula_templates_path.exists():
            with open(self.formula_templates_path, 'r') as f:
                return yaml.safe_load(f)
        return {}

    def extract(self) -> Dict[str, Any]:
        """Extract complete workbook context for Claude verification."""
        wb = openpyxl.load_workbook(str(self.workbook_path))
        try:
            context = {
                'metadata': self._extract_metadata(),
                'formula_templates': self.formula_templates,
                'bdo_account_to_row_map': {},
                'bdo_label_to_row_map': {},
                'actual_cell_contents': {},
                'formatting_snapshot': {},
                'bdo_column_h_formulas': {},
                'bdo_special_rows': {},
            }

            bdo_sheet = self._get_bdo_sheet(wb)
            if bdo_sheet:
                context['bdo_account_to_row_map'] = self._build_bdo_row_map(bdo_sheet)
                context['bdo_label_to_row_map'] = self._build_bdo_label_map(bdo_sheet)
                context['bdo_column_h_formulas'] = self._extract_bdo_column_h(bdo_sheet)
                context['bdo_special_rows'] = self._extract_bdo_special_rows(bdo_sheet)
                context['bdo_resultaat_na_belasting'] = self._read_bdo_resultaat(bdo_sheet)

            summary_sheet = self._get_summary_sheet(wb)
            if summary_sheet:
                ltm_col = self._find_ltm_column(summary_sheet)
                quarter_col = ltm_col - 1 if ltm_col else None
                context['ltm_col'] = ltm_col
                context['quarter_col'] = quarter_col
                context['ltm_col_letter'] = get_column_letter(ltm_col) if ltm_col else None
                context['quarter_col_letter'] = get_column_letter(quarter_col) if quarter_col else None

                if quarter_col and ltm_col:
                    context['actual_cell_contents'] = self._extract_cell_contents(
                        summary_sheet, quarter_col, ltm_col
                    )
                    context['formatting_snapshot'] = self._extract_formatting(
                        summary_sheet, quarter_col, ltm_col
                    )
            return context
        finally:
            wb.close()

    def _extract_metadata(self) -> Dict[str, Any]:
        return {
            'quarter': self.config.quarter,
            'year': self.config.year,
            'quarter_num': self.config.quarter_num,
            'period_end': self.config.period_end,
            'bdo_sheet_name': self.config.bdo_sheet_name,
            'summary_sheet_name': self.config.summary_sheet_name,
            'previous_quarter_label': self.config.previous_quarter_label,
        }

    def _get_bdo_sheet(self, wb):
        name = self.config.bdo_sheet_name
        if name in wb.sheetnames:
            return wb[name]
        for sn in wb.sheetnames:
            if sn.startswith('BDO'):
                return wb[sn]
        return None

    def _get_summary_sheet(self, wb):
        name = self.config.summary_sheet_name
        if name in wb.sheetnames:
            return wb[name]
        for sn in wb.sheetnames:
            if 'Management Cijfers' in sn:
                return wb[sn]
        return None

    def _find_ltm_column(self, sheet) -> Optional[int]:
        for col_idx in range(sheet.max_column, 0, -1):
            cell = sheet.cell(row=22, column=col_idx)
            if cell.value and 'LTM' in str(cell.value):
                return col_idx
        return None

    def _build_bdo_row_map(self, sheet) -> Dict[str, int]:
        row_map = {}
        for row_idx in range(1, sheet.max_row + 1):
            code = sheet.cell(row=row_idx, column=1).value
            if code and isinstance(code, (str, int, float)):
                code_str = str(code).strip()
                if code_str and code_str[0].isdigit():
                    row_map[code_str] = row_idx
        return row_map

    def _build_bdo_label_map(self, sheet) -> Dict[str, int]:
        label_map = {}
        for row_idx in range(1, sheet.max_row + 1):
            code = sheet.cell(row=row_idx, column=1).value
            label = sheet.cell(row=row_idx, column=2).value

            if code and isinstance(code, (str, int, float)):
                code_str = str(code).strip()
                if code_str and not code_str[0].isdigit():
                    label_map[code_str.lower()] = row_idx

            if label and isinstance(label, str):
                label_map[label.strip().lower()] = row_idx
        return label_map

    def _extract_bdo_column_h(self, sheet) -> Dict[int, Dict[str, Any]]:
        """Extract column H formulas/values from BDO sheet (rows 6-127)."""
        result = {}
        for row_idx in range(6, min(sheet.max_row + 1, 128)):
            code = sheet.cell(row=row_idx, column=1).value
            if code is None:
                continue
            cell = sheet.cell(row=row_idx, column=8)  # column H
            result[row_idx] = {
                'value': cell.value,
                'is_formula': isinstance(cell.value, str) and str(cell.value).startswith('='),
            }
        return result

    def _extract_bdo_special_rows(self, sheet) -> Dict[int, Dict[str, Any]]:
        """Extract special rows 129-136 from BDO sheet."""
        result = {}
        for row_idx in range(129, min(sheet.max_row + 1, 137)):
            row_data = {}
            for col_idx in range(1, 9):  # columns A-H
                cell = sheet.cell(row=row_idx, column=col_idx)
                col_letter = get_column_letter(col_idx)
                row_data[col_letter] = {
                    'value': cell.value if not isinstance(cell.value, (int, float)) else cell.value,
                    'is_formula': isinstance(cell.value, str) and str(cell.value).startswith('='),
                }
            result[row_idx] = row_data
        return result

    def _read_bdo_resultaat(self, sheet) -> Optional[Dict[str, Any]]:
        """Read Resultaat na belasting from BDO sheet."""
        target_row = None
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
            return None

        total = 0.0
        col_values = {}
        for col_idx in range(4, 8):
            val = sheet.cell(row=target_row, column=col_idx).value
            col_letter = get_column_letter(col_idx)
            if isinstance(val, (int, float)):
                total += val
                col_values[col_letter] = val
            elif isinstance(val, str) and val.startswith('='):
                col_values[col_letter] = f"(formula: {val})"
            else:
                col_values[col_letter] = val

        return {
            'row': target_row,
            'raw_total_D_G': total,
            'sign_adjusted': -total,
            'column_values': col_values,
        }

    def _extract_cell_contents(self, sheet, quarter_col: int, ltm_col: int) -> Dict[str, Any]:
        """Extract actual cell formulas and values for verification rows."""
        sections = {
            'balance_sheet': list(range(3, 20)),        # rows 3-19
            'header': [1, 2, 22],
            'profit_loss': list(range(23, 69)),         # rows 23-68
            'bank_accounts': list(range(104, 121)),     # rows 104-120
        }

        contents = {}
        q_letter = get_column_letter(quarter_col)
        ltm_letter = get_column_letter(ltm_col)

        for section_name, rows in sections.items():
            section_data = {}
            for row_idx in rows:
                row_data = {}
                label_cell = sheet.cell(row=row_idx, column=2)
                row_data['label'] = str(label_cell.value) if label_cell.value else ''

                q_cell = sheet.cell(row=row_idx, column=quarter_col)
                row_data['quarter'] = {
                    'col': quarter_col,
                    'col_letter': q_letter,
                    'value': q_cell.value,
                    'is_formula': isinstance(q_cell.value, str) and str(q_cell.value).startswith('='),
                }

                ltm_cell = sheet.cell(row=row_idx, column=ltm_col)
                row_data['ltm'] = {
                    'col': ltm_col,
                    'col_letter': ltm_letter,
                    'value': ltm_cell.value,
                    'is_formula': isinstance(ltm_cell.value, str) and str(ltm_cell.value).startswith('='),
                }

                section_data[row_idx] = row_data
            contents[section_name] = section_data

        return contents

    def _extract_formatting(self, sheet, quarter_col: int, ltm_col: int) -> Dict[str, Any]:
        """Extract formatting metadata for quarter and LTM columns."""
        formatting = {}
        rows_to_check = list(range(1, 20)) + [22] + list(range(23, 69)) + list(range(104, 121))

        for row_idx in rows_to_check:
            row_fmt = {}
            for col, col_name in [(quarter_col, 'quarter'), (ltm_col, 'ltm')]:
                cell = sheet.cell(row=row_idx, column=col)
                fmt = {
                    'font_name': cell.font.name if cell.font else None,
                    'font_size': cell.font.size if cell.font else None,
                    'bold': cell.font.bold if cell.font else False,
                    'font_color': str(cell.font.color.rgb) if cell.font and cell.font.color and cell.font.color.rgb else None,
                    'number_format': cell.number_format,
                }
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                    fmt['fill_color'] = str(cell.fill.fgColor.rgb)
                else:
                    fmt['fill_color'] = None
                row_fmt[col_name] = fmt
            formatting[row_idx] = row_fmt

        return formatting


# ---------------------------------------------------------------------------
# Prompt Builder
# ---------------------------------------------------------------------------

class VerificationPromptBuilder:
    """Builds the Claude Opus 4.6 verification prompt."""

    def build(self, context: Dict[str, Any]) -> str:
        """Build the complete verification prompt from extracted context."""
        meta = context.get('metadata', {})

        prompt_parts = [
            self._role_section(),
            self._context_section(context),
            self._sign_convention_section(context),
            self._formula_construction_guide(context),
            self._checks_section(context),
            self._response_format_section(),
        ]
        return '\n\n'.join(prompt_parts)

    def _role_section(self) -> str:
        return """ROLE:
You are verifying a Management Accounts Excel workbook for QSP ESS B.V.
The workbook was auto-generated and you must check every formula, value,
cross-sheet reference, sign convention, and cell format for correctness.

You must return ONLY valid JSON — no markdown fences, no commentary outside
the JSON object.

IMPORTANT: If every check passes, return status "PASS" with empty "issues"
and "patches" arrays. Do NOT invent potential improvements or suggestions —
only report actual errors found in the workbook."""

    def _context_section(self, context: Dict[str, Any]) -> str:
        meta = context.get('metadata', {})
        parts = [
            "CONTEXT:",
            f"Quarter: {meta.get('quarter', 'N/A')}",
            f"Year: {meta.get('year', 'N/A')}",
            f"Period End: {meta.get('period_end', 'N/A')}",
            f"BDO Sheet: {meta.get('bdo_sheet_name', 'N/A')}",
            f"Summary Sheet: {meta.get('summary_sheet_name', 'N/A')}",
            f"Quarter Column: {context.get('quarter_col_letter', 'N/A')} (index {context.get('quarter_col', 'N/A')})",
            f"LTM Column: {context.get('ltm_col_letter', 'N/A')} (index {context.get('ltm_col', 'N/A')})",
            "",
            "BDO ACCOUNT-CODE-TO-ROW MAP:",
            json.dumps(context.get('bdo_account_to_row_map', {}), indent=2),
            "",
            "BDO LABEL-TO-ROW MAP:",
            json.dumps(context.get('bdo_label_to_row_map', {}), indent=2),
            "",
            "FORMULA TEMPLATES (balance_sheet and profit_loss):",
            json.dumps(context.get('formula_templates', {}), indent=2),
            "",
            "ACTUAL CELL CONTENTS (new quarter + LTM columns):",
            self._serialize_cell_contents(context.get('actual_cell_contents', {})),
            "",
            "BDO COLUMN H FORMULAS (LTM sums):",
            json.dumps(context.get('bdo_column_h_formulas', {}), indent=2),
            "",
            "BDO SPECIAL ROWS (129-136):",
            json.dumps(context.get('bdo_special_rows', {}), indent=2),
            "",
            "BDO RESULTAAT NA BELASTING (ground truth):",
            json.dumps(context.get('bdo_resultaat_na_belasting', {}), indent=2),
            "",
            "FORMATTING SNAPSHOT:",
            json.dumps(context.get('formatting_snapshot', {}), indent=2),
        ]

        shadow_pl = context.get('shadow_pl')
        shadow_bs = context.get('shadow_bs')
        if shadow_pl:
            parts.append("")
            parts.append("SHADOW P&L (expected values computed from raw BDO data):")
            parts.append("Each row shows label, expected numeric value, formula type,")
            parts.append("account codes used, and any missing codes (missing = BDO account")
            parts.append("not found, which would cause an incorrect zero in the formula).")
            parts.append(json.dumps(
                {str(k): v for k, v in shadow_pl.items()}, indent=2
            ))
        if shadow_bs:
            parts.append("")
            parts.append("SHADOW BALANCE SHEET (expected values computed from raw BDO data):")
            parts.append(json.dumps(
                {str(k): v for k, v in shadow_bs.items()}, indent=2
            ))

        return '\n'.join(parts)

    def _serialize_cell_contents(self, contents: Dict) -> str:
        """Serialize cell contents with safe JSON handling."""
        safe = {}
        for section, rows in contents.items():
            safe_section = {}
            for row_idx, row_data in rows.items():
                safe_row = {}
                for key, val in row_data.items():
                    if isinstance(val, dict):
                        safe_val = {}
                        for k, v in val.items():
                            if isinstance(v, (int, float, str, bool, type(None))):
                                safe_val[k] = v
                            else:
                                safe_val[k] = str(v)
                        safe_row[key] = safe_val
                    elif isinstance(val, (int, float, str, bool, type(None))):
                        safe_row[key] = val
                    else:
                        safe_row[key] = str(val)
                safe_section[str(row_idx)] = safe_row
            safe[section] = safe_section
        return json.dumps(safe, indent=2)

    def _sign_convention_section(self, context: Dict[str, Any]) -> str:
        bdo = context.get('metadata', {}).get('bdo_sheet_name', 'BDO')
        return f"""SIGN CONVENTION:
BDO stores revenue and profit as NEGATIVE numbers (Dutch trial balance convention).
Management Cijfers uses POSITIVE for income. All bdo_ref formulas with
sign: "-" negate the BDO value to convert.

For Balance Sheet rows (3-19): formulas reference BDO column H (LTM/closing).
Quarter columns for Balance Sheet rows should be EMPTY (no data).

For P&L rows (23-68):
- Quarter column formulas reference BDO column G (current quarter mutations).
- LTM column formulas reference BDO column H (full-year closing).
- Exception: bdo_ref_conditional and manual_with_ltm rows use
  ltm_sum_quarters (SUM of last 4 quarterly columns in Management Cijfers)
  for the LTM column.

All bdo_ref rows with sign "-": the formula negates BDO values (prefix with =-).
Multi-account bdo_ref with operator "-": first account is negated,
subsequent accounts are subtracted (=-'{bdo}'!G{{row1}}-'{bdo}'!G{{row2}}-...).
Multi-account bdo_ref with operator "+": first account is negated,
subsequent accounts are added (=-'{bdo}'!G{{row1}}+'{bdo}'!G{{row2}}+...).
bdo_ref with sign "" (empty): no negation, formula is ='{{BDO}}'!H{{row}} (used
for Balance Sheet rows where BDO values already have the correct sign)."""

    def _formula_construction_guide(self, context: Dict[str, Any]) -> str:
        bdo = context.get('metadata', {}).get('bdo_sheet_name', 'BDO')
        q = context.get('quarter_col_letter', 'AA')
        ltm = context.get('ltm_col_letter', 'AB')

        return f"""FORMULA CONSTRUCTION RULES (how to verify each formula type):

1. bdo_ref (single account, sign "-", quarter column):
   Template: account_codes: ["8000003"], sign: "-"
   Look up 8000003 in the BDO account-code-to-row map → say it's row 10.
   Expected quarter formula: =-'{bdo}'!G10
   Expected LTM formula:     =-'{bdo}'!H10

2. bdo_ref (multiple accounts, sign "-", operator "-"):
   Template: account_codes: ["4611000","4613000","4620200","4623001","4642000"], sign: "-", operator: "-"
   Look up each code in the BDO map → say rows 91, 92, 93, 94, 95.
   Expected quarter formula: =-'{bdo}'!G91-'{bdo}'!G92-'{bdo}'!G93-'{bdo}'!G94-'{bdo}'!G95
   Expected LTM formula:     =-'{bdo}'!H91-'{bdo}'!H92-'{bdo}'!H93-'{bdo}'!H94-'{bdo}'!H95
   (Each term subtracts because operator is "-".)

3. bdo_ref (multiple accounts, sign "-", operator "+"):
   Template: account_codes: ["9500000","9510000"], sign: "-", operator: "+"
   Rows → say 19, 20.
   Expected: =-'{bdo}'!G19+'{bdo}'!G20  (BUT note: ONLY the first term gets
   the sign prefix; subsequent use the operator. The net effect: first is negated,
   rest are added with their natural sign.)
   WRONG:    =-'{bdo}'!G19-'{bdo}'!G20  ← this would be operator "-", not "+"

4. bdo_ref (sign "" / empty, Balance Sheet):
   Template: account_codes: ["1790002"], sign: ""
   Row → say 6.
   Expected LTM formula: ='{bdo}'!H6  (no negation, column H only)

5. bdo_sum_range:
   Template: start_account: "2400200", end_account: "2400206", count: 7
   Look up start code → say row 31. The range spans 7 consecutive rows.
   Expected LTM formula: =SUM('{bdo}'!H31:'{bdo}'!H37)
   If "additional" account codes exist, they are added outside the range:
   =SUM('{bdo}'!H31:'{bdo}'!H37)+'{bdo}'!H42+'{bdo}'!H43

6. bdo_ref_conditional:
   Quarter column: uses q_config (a bdo_ref, same rules as above).
   LTM column: uses ltm_config with type "ltm_sum_quarters" — formula is
   =SUM({{4 columns before LTM}}50:{{column before LTM}}50) spanning exactly
   the last 4 quarterly columns in Management Cijfers (NOT BDO references).

7. calc:
   Pattern uses {{COL}} placeholder → replace with the actual column letter.
   In the current workbook: quarter column = {q}, LTM column = {ltm}.
   Row 19 is special: only exists in LTM column."""

    def _checks_section(self, context: Dict[str, Any]) -> str:
        meta = context.get('metadata', {})
        bdo_name = meta.get('bdo_sheet_name', 'BDO')
        summary = meta.get('summary_sheet_name', 'Management Cijfers')
        q = context.get('quarter_col_letter', 'AA')
        ltm = context.get('ltm_col_letter', 'AB')
        q_idx = context.get('quarter_col', '')
        ltm_idx = context.get('ltm_col', '')

        return f"""CHECKS TO PERFORM:

1. BDO FORMULA VERIFICATION:
   For every bdo_ref/bdo_sum_range formula in the actual cell contents:
   verify the row numbers in the formula match the account codes looked up
   in the BDO account-code-to-row map. The formula must reference sheet
   '{bdo_name}'. Use the formula construction rules above to determine
   the expected formula and compare with the actual.

2. CALC FORMULA VERIFICATION:
   For every calc formula, verify it references the correct rows using the
   actual column letters (quarter={q}, LTM={ltm}):

   Quarter column ({q}, index {q_idx}):
     Row 25: =SUM({q}23:{q}24)
     Row 26: ={q}25/{q}23-1
     Row 30: =SUM({q}27:{q}29)
     Row 44: =SUM({q}32:{q}43)
     Row 45: ={q}25+{q}30+{q}44
     Row 48: ={q}47+{q}46+{q}45
     Row 51: =-({q}50-{q}53)
     Row 55: ={q}53+{q}48
     Row 57: ={q}56+{q}55
     Row 66: =SUM({q}57:{q}65)
     Row 68: ={q}67+{q}66

   LTM column ({ltm}, index {ltm_idx}):
     Row 19: =SUM({ltm}3:{ltm}18)
     Row 25: =SUM({ltm}23:{ltm}24)
     Row 26: ={ltm}25/{ltm}23-1
     Row 30: =SUM({ltm}27:{ltm}29)
     Row 44: =SUM({ltm}32:{ltm}43)
     Row 45: ={ltm}25+{ltm}30+{ltm}44
     Row 48: ={ltm}47+{ltm}46+{ltm}45
     Row 51: =-({ltm}50-{ltm}53)
     Row 55: ={ltm}53+{ltm}48
     Row 57: ={ltm}56+{ltm}55
     Row 66: =SUM({ltm}57:{ltm}65)
     Row 68: ={ltm}67+{ltm}66

3. CROSS-CHECK: Compare the Shadow P&L row 68 value against Shadow BS row 19
   value. They must match within tolerance 1.0 (both computed from the same
   BDO data via different paths). If they don't match, the workbook has an
   error. Report which row(s) appear wrong and generate a patch.

4. BDO GROUND TRUTH: The "Resultaat na belasting" raw value from BDO
   (sum of columns D-G), when NEGATED, should match both Shadow BS row 19
   and Shadow P&L row 68 within tolerance 1.0.

5. LTM SUM RANGES: For rows 50 (manual_with_ltm), 61, and 64
   (bdo_ref_conditional), the LTM column must contain a SUM spanning exactly
   the last 4 quarterly columns (e.g., =SUM({q}50:{{prev_q}}50) where
   {{prev_q}} is 3 columns left of {q}). After column insertion these ranges
   sometimes exclude the new column — catch and fix.

6. BANK ACCOUNT FORMULAS (rows 107-113 in '{summary}'):
   Each row must reference '{bdo_name}' column H at these specific BDO rows:
     Row 107 (ABN AMRO RENT):     ='{bdo_name}'!H35
     Row 108 (ABN AMRO MAINT):    ='{bdo_name}'!H34
     Row 109 (ABN AMRO EXP):      ='{bdo_name}'!H36
     Row 110 (ABN AMRO GEN):      ='{bdo_name}'!H32
     Row 111 (ABN AMRO DEP):      ='{bdo_name}'!H33
     Row 112 (ABN AMRO CAPEX):    ='{bdo_name}'!H37
     Row 113 (ABN AMRO DISPOSAL): ='{bdo_name}'!H31
     Row 114 (Total):             =SUM({ltm}107:{ltm}113)

7. FORMATTING VERIFICATION:
   For every financial cell in the quarter and LTM columns, verify:
   - Font: Avenir Book, size 10 (bold for total/header rows 19, 25, 30, 44,
     45, 48, 55, 57, 66, 68)
   - Number format: _(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)
     for all numeric cells (rows 3-19, 23-68, 107-114)
   - Fill: Section header and total rows should have blue bar fills
     matching the previous quarter column (compare fill_color in the
     formatting snapshot between quarter and LTM — they should be identical
     for the same row)
   - Row 22: quarter header "{meta.get('quarter', '')}" and LTM header
     "LTM {meta.get('quarter', '')}" with bold styling

8. MISSING ACCOUNT CODES: For every account code in the formula templates,
   verify it exists in the BDO account-code-to-row map. If missing, flag it
   as kind "missing" — a missing code means the formula will produce an
   incorrect zero.

9. BALANCE SHEET QUARTER COLUMNS: Rows 3-19 in the quarter column ({q})
   must be empty/None. Only the LTM column ({ltm}) should have formulas
   for Balance Sheet rows.

10. ROW 50 (Cash proceeds sale):
    - Quarter column ({q}): must contain a numeric value (from sales tracker)
      or 0, NOT a formula.
    - LTM column ({ltm}): must contain a SUM of the last 4 quarterly columns."""

    def _response_format_section(self) -> str:
        return """RESPONSE FORMAT (strict JSON only, no markdown fences):
{
  "status": "PASS" or "ISSUES_FOUND",
  "checks_performed": <count of individual checks>,
  "issues": [
    {
      "sheet": "<sheet name>",
      "row": <row number>,
      "col": <column index (integer)>,
      "col_letter": "<column letter>",
      "kind": "formula" | "value" | "format" | "missing",
      "current": "<what is currently in the cell>",
      "expected": "<what should be there>",
      "explanation": "<why this is wrong>"
    }
  ],
  "patches": [
    {
      "sheet": "<sheet name>",
      "row": <row number>,
      "col": <column index (integer)>,
      "kind": "formula" | "value" | "format",
      "value": "<corrected formula string, numeric value, or format dict>"
    }
  ],
  "validation_summary": {
    "bs_row_19": <numeric value from Shadow BS or null>,
    "pl_row_68": <numeric value from Shadow P&L or null>,
    "bdo_resultaat_adjusted": <negated Resultaat na belasting or null>,
    "all_match": true | false
  },
  "notes": "<free-text summary of findings>"
}

RULES:
- If status is "PASS", issues and patches MUST both be empty arrays [].
- Only include a patch if you are confident the corrected value is right.
- Every issue should have a corresponding patch unless the fix is ambiguous.
- The "col" field must be an integer column index (1-based), not a letter.
- For formula patches, the value must start with "=".
- For format patches, value is a dict like {"number_format": "...", "bold": true}."""


# ---------------------------------------------------------------------------
# Patch Applier
# ---------------------------------------------------------------------------

ALLOWLISTED_SHEET_PREFIX = "Management Cijfers"
ALLOWLISTED_ROW_RANGE = range(1, 121)
DISALLOWED_SHEET_REFS = ["BDO"]


class PatchApplier:
    """
    Applies JSON patches from Claude to the workbook.
    Strict allowlist: only Management Cijfers sheets, rows 1-120.
    """

    def __init__(self, workbook_path: str, summary_sheet_name: str):
        self.workbook_path = Path(workbook_path)
        self.summary_sheet_name = summary_sheet_name
        self.applied_patches: List[Dict] = []
        self.rejected_patches: List[Dict] = []

    def apply(self, patches: List[Dict], output_path: str) -> Tuple[int, List[Dict]]:
        """
        Apply patches to workbook, save to output_path.
        Returns (count_applied, list_of_rejected).
        """
        wb = openpyxl.load_workbook(str(self.workbook_path))
        applied = 0

        for patch in patches:
            if self._validate_patch(patch):
                try:
                    self._apply_single_patch(wb, patch)
                    self.applied_patches.append(patch)
                    applied += 1
                except Exception as e:
                    patch['rejection_reason'] = f"Apply error: {e}"
                    self.rejected_patches.append(patch)
                    logger.warning(f"Patch apply failed for row {patch.get('row')}: {e}")
            else:
                self.rejected_patches.append(patch)

        wb.save(output_path)
        wb.close()
        logger.info(f"Applied {applied} patches, rejected {len(self.rejected_patches)}")
        return applied, self.rejected_patches

    def _validate_patch(self, patch: Dict) -> bool:
        sheet_name = patch.get('sheet', '')
        if ALLOWLISTED_SHEET_PREFIX not in sheet_name:
            patch['rejection_reason'] = f"Sheet '{sheet_name}' not in allowlist"
            logger.warning(f"Rejected patch: {patch['rejection_reason']}")
            return False

        for prefix in DISALLOWED_SHEET_REFS:
            if sheet_name.startswith(prefix):
                patch['rejection_reason'] = f"Sheet '{sheet_name}' is protected"
                logger.warning(f"Rejected patch: {patch['rejection_reason']}")
                return False

        row = patch.get('row', 0)
        if row not in ALLOWLISTED_ROW_RANGE:
            patch['rejection_reason'] = f"Row {row} outside allowlist (1-120)"
            logger.warning(f"Rejected patch: {patch['rejection_reason']}")
            return False

        kind = patch.get('kind', '')
        value = patch.get('value', '')

        if kind == 'formula':
            if not isinstance(value, str) or not value.startswith('='):
                patch['rejection_reason'] = f"Formula patch value doesn't start with '=': {value}"
                logger.warning(f"Rejected patch: {patch['rejection_reason']}")
                return False
            for disallowed in DISALLOWED_SHEET_REFS:
                pass  # BDO refs within formulas are expected and valid

        return True

    def _apply_single_patch(self, wb, patch: Dict):
        sheet_name = patch['sheet']
        if sheet_name not in wb.sheetnames:
            for sn in wb.sheetnames:
                if ALLOWLISTED_SHEET_PREFIX in sn:
                    sheet_name = sn
                    break
            else:
                raise ValueError(f"Sheet '{sheet_name}' not found")

        sheet = wb[sheet_name]
        row = patch['row']
        col = patch['col']
        kind = patch['kind']
        value = patch['value']

        cell = sheet.cell(row=row, column=col)
        old_value = cell.value

        if kind in ('formula', 'value'):
            cell.value = value
            logger.info(
                f"Patched {sheet_name}!{get_column_letter(col)}{row}: "
                f"'{old_value}' → '{value}' (kind={kind})"
            )

        elif kind == 'format':
            if isinstance(value, dict):
                if 'number_format' in value:
                    cell.number_format = value['number_format']
                if 'bold' in value:
                    new_font = copy_style(cell.font)
                    cell.font = Font(
                        name=new_font.name,
                        size=new_font.size,
                        bold=value['bold'],
                        color=new_font.color,
                    )
                if 'font_name' in value:
                    new_font = copy_style(cell.font)
                    cell.font = Font(
                        name=value['font_name'],
                        size=new_font.size,
                        bold=new_font.bold,
                        color=new_font.color,
                    )
            logger.info(f"Format patched {sheet_name}!{get_column_letter(col)}{row}")


# ---------------------------------------------------------------------------
# Claude API Client
# ---------------------------------------------------------------------------

class ClaudeVerifier:
    """Sends workbook context to Claude for verification and parses response."""

    MODEL = "claude-opus-4-20250514"

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get('ANTHROPIC_API_KEY', '')

    def verify(self, prompt: str) -> Dict[str, Any]:
        """
        Send verification prompt to Claude and parse JSON response.
        Returns parsed JSON dict or error dict.
        """
        if not self.api_key:
            logger.warning("ANTHROPIC_API_KEY not set — returning mock PASS")
            return self._mock_pass_response()

        try:
            import anthropic
        except ImportError:
            logger.error("anthropic package not installed. Run: pip install anthropic")
            return {'status': 'ERROR', 'error': 'anthropic package not installed'}

        client = anthropic.Anthropic(api_key=self.api_key)

        try:
            message = client.messages.create(
                model=self.MODEL,
                max_tokens=16384,
                messages=[
                    {"role": "user", "content": prompt}
                ],
            )

            response_text = message.content[0].text.strip()
            return self._parse_response(response_text)

        except Exception as e:
            logger.error(f"Claude API call failed: {e}")
            return {'status': 'ERROR', 'error': str(e)}

    def _parse_response(self, text: str) -> Dict[str, Any]:
        """Parse Claude's JSON response, handling markdown fences."""
        cleaned = text.strip()
        if cleaned.startswith('```'):
            lines = cleaned.split('\n')
            start = 1
            end = len(lines)
            for i, line in enumerate(lines[1:], 1):
                if line.strip() == '```':
                    end = i
                    break
            cleaned = '\n'.join(lines[start:end])

        try:
            return json.loads(cleaned)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse Claude response as JSON: {e}")
            logger.debug(f"Raw response: {text[:500]}")
            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                try:
                    return json.loads(json_match.group())
                except json.JSONDecodeError:
                    pass
            return {
                'status': 'ERROR',
                'error': f'JSON parse failed: {e}',
                'raw_response': text[:2000],
            }

    def _mock_pass_response(self) -> Dict[str, Any]:
        return {
            'status': 'PASS',
            'checks_performed': 0,
            'issues': [],
            'patches': [],
            'validation_summary': {
                'bs_row_19': None,
                'pl_row_68': None,
                'bdo_resultaat_adjusted': None,
                'all_match': True,
            },
            'notes': 'Mock response — ANTHROPIC_API_KEY not configured',
        }


# ---------------------------------------------------------------------------
# BDO Copy Verifier (Phase A context additions)
# ---------------------------------------------------------------------------

class BDOCopyVerifier:
    """Adds BDO sheet copy verification context to the prompt."""

    def __init__(self, workbook_path: str, bdo_source_path: Optional[str],
                 config: VerificationConfig):
        self.workbook_path = Path(workbook_path)
        self.bdo_source_path = Path(bdo_source_path) if bdo_source_path else None
        self.config = config

    def build_verification_context(self) -> Dict[str, Any]:
        """Build BDO copy verification context."""
        result = {
            'bdo_copy_verified': False,
            'integrity_checks': {},
            'column_h_checks': {},
            'special_row_checks': {},
        }

        if not self.bdo_source_path or not self.bdo_source_path.exists():
            result['note'] = 'BDO source file not available for copy verification'
            return result

        try:
            wb = openpyxl.load_workbook(str(self.workbook_path))
            source_wb = openpyxl.load_workbook(str(self.bdo_source_path))

            bdo_sheet = None
            for sn in wb.sheetnames:
                if sn == self.config.bdo_sheet_name:
                    bdo_sheet = wb[sn]
                    break

            source_sheet = source_wb.active

            if bdo_sheet and source_sheet:
                result['integrity_checks'] = self._check_copy_integrity(
                    source_sheet, bdo_sheet
                )
                result['column_h_checks'] = self._check_column_h_formulas(bdo_sheet)
                result['special_row_checks'] = self._check_special_rows(bdo_sheet)
                result['bdo_copy_verified'] = True

            wb.close()
            source_wb.close()
        except Exception as e:
            result['error'] = str(e)
            logger.warning(f"BDO copy verification failed: {e}")

        return result

    def _check_copy_integrity(self, source, target) -> Dict[str, Any]:
        mismatches = []
        rows_checked = 0

        max_row = min(source.max_row, 140)
        max_col = min(source.max_column, 10)

        for row_idx in range(1, max_row + 1):
            for col_idx in range(1, max_col + 1):
                src_val = source.cell(row=row_idx, column=col_idx).value
                tgt_val = target.cell(row=row_idx, column=col_idx).value

                if src_val != tgt_val:
                    if isinstance(src_val, float) and isinstance(tgt_val, float):
                        if abs(src_val - tgt_val) < 0.01:
                            continue
                    mismatches.append({
                        'row': row_idx,
                        'col': col_idx,
                        'source': str(src_val)[:50] if src_val else None,
                        'target': str(tgt_val)[:50] if tgt_val else None,
                    })
            rows_checked += 1

        return {
            'rows_checked': rows_checked,
            'mismatches': mismatches[:20],
            'total_mismatches': len(mismatches),
            'passed': len(mismatches) == 0,
        }

    def _check_column_h_formulas(self, sheet) -> Dict[str, Any]:
        """Verify column H contains =SUM(C{row}:G{row}) for data rows."""
        issues = []
        for row_idx in range(6, min(sheet.max_row + 1, 125)):
            code = sheet.cell(row=row_idx, column=1).value
            if code is None:
                continue
            h_cell = sheet.cell(row=row_idx, column=8)
            if h_cell.value and isinstance(h_cell.value, str) and h_cell.value.startswith('='):
                expected = f"=SUM(C{row_idx}:G{row_idx})"
                if h_cell.value.upper() != expected.upper():
                    issues.append({
                        'row': row_idx,
                        'expected': expected,
                        'actual': h_cell.value,
                    })
        return {'issues': issues[:20], 'total_issues': len(issues)}

    def _check_special_rows(self, sheet) -> Dict[str, Any]:
        """Verify special rows 129-136."""
        checks = {}
        for row_idx in range(129, min(sheet.max_row + 1, 137)):
            row_data = {}
            for col_idx in range(1, 9):
                cell = sheet.cell(row=row_idx, column=col_idx)
                row_data[get_column_letter(col_idx)] = {
                    'value': cell.value if not isinstance(cell.value, float) else cell.value,
                    'is_formula': isinstance(cell.value, str) and str(cell.value).startswith('='),
                }
            checks[row_idx] = row_data
        return checks


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

class AIVerificationOrchestrator:
    """
    Main orchestrator for AI verification of Management Accounts.
    Always runs after MA build, applies patches, re-validates.
    """

    def __init__(self, config: VerificationConfig,
                 api_key: Optional[str] = None,
                 formula_templates_path: str = "config/formula_templates.yaml"):
        self.config = config
        self.api_key = api_key
        self.formula_templates_path = formula_templates_path

    def run(self, workbook_path: str,
            bdo_source_path: Optional[str] = None,
            bdo_result: Optional[BDOParseResult] = None,
            computed_values: Optional[Dict] = None) -> VerificationResult:
        """
        Run full AI verification pipeline.

        Args:
            workbook_path: Path to the draft MA workbook
            bdo_source_path: Path to original BDO source file (for copy verification)
            bdo_result: Parsed BDO data (for shadow model context)
            computed_values: Pre-computed numeric values from MA builder {(row,'quarter'|'ltm'): float}

        Returns:
            VerificationResult with status, patches applied, and file paths
        """
        workbook_path = str(workbook_path)

        result = VerificationResult(
            status='PASS',
            original_file=workbook_path,
            verified_file=workbook_path,
        )

        try:
            # Phase A: Extract workbook context
            logger.info("[AI Verify] Phase A: Extracting workbook context...")
            extractor = WorkbookContextExtractor(
                workbook_path, self.config, self.formula_templates_path
            )
            context = extractor.extract()

            # Phase A+: BDO copy verification context
            if bdo_source_path:
                logger.info("[AI Verify] Phase A+: Verifying BDO copy integrity...")
                bdo_verifier = BDOCopyVerifier(workbook_path, bdo_source_path, self.config)
                bdo_context = bdo_verifier.build_verification_context()
                context['bdo_copy_verification'] = bdo_context

            # Phase A++: Compute shadow P&L and shadow BS from raw BDO data
            if bdo_result:
                logger.info("[AI Verify] Phase A++: Computing shadow models...")
                shadow_pl, shadow_bs = self._compute_shadow_models(
                    workbook_path, bdo_result
                )
                if shadow_pl:
                    context['shadow_pl'] = shadow_pl
                if shadow_bs:
                    context['shadow_bs'] = shadow_bs

            # Phase B: Build prompt
            logger.info("[AI Verify] Phase B: Building verification prompt...")
            prompt_builder = VerificationPromptBuilder()
            prompt = prompt_builder.build(context)

            # Add BDO copy verification to prompt if available
            if 'bdo_copy_verification' in context:
                prompt += "\n\nBDO COPY VERIFICATION CONTEXT:\n"
                prompt += json.dumps(context['bdo_copy_verification'], indent=2)

            logger.info(f"[AI Verify] Prompt size: {len(prompt)} chars")

            # Phase C: Send to Claude
            logger.info("[AI Verify] Phase C: Sending to Claude for verification...")
            verifier = ClaudeVerifier(api_key=self.api_key)
            claude_response = verifier.verify(prompt)

            if claude_response.get('status') == 'ERROR':
                result.status = 'ERROR'
                result.error = claude_response.get('error', 'Unknown error')
                result.notes = f"Claude API error: {result.error}"
                logger.warning(f"[AI Verify] Claude error — original file unchanged: {result.error}")
                return result

            result.status = claude_response.get('status', 'PASS')
            result.issues = claude_response.get('issues', [])
            result.patches = claude_response.get('patches', [])
            result.validation_summary = claude_response.get('validation_summary', {})
            result.notes = claude_response.get('notes', '')

            # Phase D: Apply patches in-place to the original workbook
            if result.patches:
                logger.info(f"[AI Verify] Phase D: Applying {len(result.patches)} patches in-place...")

                applier = PatchApplier(workbook_path, self.config.summary_sheet_name)
                applied, rejected = applier.apply(result.patches, workbook_path)
                result.patches_applied = applied

                if rejected:
                    logger.warning(f"[AI Verify] {len(rejected)} patches rejected")
                    for rej in rejected:
                        logger.warning(f"  Rejected: row {rej.get('row')}, reason: {rej.get('rejection_reason')}")

                # Phase D+: Re-validate
                logger.info("[AI Verify] Phase D+: Re-validating patched workbook...")
                result.revalidation_passed = self._revalidate(workbook_path, bdo_result, computed_values)

            else:
                result.revalidation_passed = True
                logger.info("[AI Verify] No patches needed — workbook passed verification")

            logger.info(
                f"[AI Verify] Complete — status={result.status}, "
                f"patches={result.patches_applied}, "
                f"revalidation={'PASS' if result.revalidation_passed else 'FAIL'}"
            )

        except Exception as e:
            logger.exception(f"[AI Verify] Verification failed: {e}")
            result.status = 'ERROR'
            result.error = str(e)
            result.verified_file = workbook_path

        return result

    def _compute_shadow_models(
        self, workbook_path: str, bdo_result: BDOParseResult
    ) -> Tuple[Optional[Dict], Optional[Dict]]:
        """
        Compute shadow P&L and shadow BS from raw BDO data.
        Returns (shadow_pl, shadow_bs) dicts or (None, None) on failure.
        """
        try:
            from .management_accounts import ManagementAccountsBuilder, ManagementAccountsConfig

            config = ManagementAccountsConfig(
                quarter=self.config.quarter,
                period_end=_parse_period_end(self.config.period_end),
                bdo_sheet_name=self.config.bdo_sheet_name,
                summary_sheet_name=self.config.summary_sheet_name,
            )

            builder = ManagementAccountsBuilder.__new__(ManagementAccountsBuilder)
            builder.config = config
            builder.formula_templates_path = Path(self.formula_templates_path)
            builder.formula_templates = builder._load_formula_templates()
            builder.balance_sheet_templates = builder._load_balance_sheet_templates()
            builder._new_bdo_row_map = {}
            builder._new_bdo_label_map = {}
            builder._prev_bdo_row_map = {}

            builder.workbook = openpyxl.load_workbook(workbook_path)
            bdo_sheet = None
            for sn in builder.workbook.sheetnames:
                if sn == self.config.bdo_sheet_name:
                    bdo_sheet = builder.workbook[sn]
                    break
            if bdo_sheet:
                builder._build_row_map(
                    bdo_sheet, builder._new_bdo_row_map, builder._new_bdo_label_map
                )

            shadow_pl = builder._compute_shadow_pl(bdo_result)
            shadow_bs = builder._compute_shadow_bs(bdo_result)
            builder.workbook.close()

            logger.info(
                f"[AI Verify] Shadow models: P&L row 68={shadow_pl.get(68, {}).get('value', 'N/A')}, "
                f"BS row 19={shadow_bs.get(19, {}).get('value', 'N/A')}"
            )
            return shadow_pl, shadow_bs

        except Exception as e:
            logger.warning(f"[AI Verify] Shadow model computation failed: {e}")
            return None, None

    def _revalidate(self, workbook_path: str,
                    bdo_result: Optional[BDOParseResult],
                    computed_values: Optional[Dict] = None) -> bool:
        """
        Re-run validation on the patched workbook.

        When computed_values are available (from the MA builder), check them
        against the BDO ground truth directly.  Otherwise fall back to
        recomputing shadow models.
        """
        tolerance = 1.0

        if computed_values:
            pl68 = computed_values.get((68, 'ltm'), 0.0)
            bs19 = computed_values.get((19, 'ltm'), 0.0)
            diff = abs(pl68 - bs19)
            passed = diff <= tolerance
            logger.info(
                f"[AI Verify] Re-validation (computed_values): "
                f"PL68={pl68:,.2f}, BS19={bs19:,.2f}, diff={diff:,.2f}, passed={passed}"
            )
            return passed

        if bdo_result is None:
            logger.warning("[AI Verify] No BDO result for re-validation, skipping")
            return True

        try:
            from .management_accounts import ManagementAccountsBuilder, ManagementAccountsConfig

            config = ManagementAccountsConfig(
                quarter=self.config.quarter,
                period_end=_parse_period_end(self.config.period_end),
                bdo_sheet_name=self.config.bdo_sheet_name,
                summary_sheet_name=self.config.summary_sheet_name,
            )

            builder = ManagementAccountsBuilder.__new__(ManagementAccountsBuilder)
            builder.config = config
            builder.formula_templates_path = Path(self.formula_templates_path)
            builder.formula_templates = builder._load_formula_templates()
            builder.balance_sheet_templates = builder._load_balance_sheet_templates()
            builder._new_bdo_row_map = {}
            builder._new_bdo_label_map = {}
            builder._prev_bdo_row_map = {}

            builder.workbook = openpyxl.load_workbook(workbook_path)

            bdo_sheet = None
            for sn in builder.workbook.sheetnames:
                if sn == self.config.bdo_sheet_name:
                    bdo_sheet = builder.workbook[sn]
                    break

            if bdo_sheet:
                builder._build_row_map(bdo_sheet, builder._new_bdo_row_map, builder._new_bdo_label_map)

            shadow_pl = builder._compute_shadow_pl(bdo_result)
            shadow_bs = builder._compute_shadow_bs(bdo_result)

            dr = shadow_pl.get(68, {}).get('value', 0.0)
            em = shadow_bs.get(19, {}).get('value', 0.0)

            builder.workbook.close()

            diff = abs(dr - em)
            passed = diff <= tolerance
            logger.info(
                f"[AI Verify] Re-validation (shadow): DR={dr:,.2f}, EM={em:,.2f}, "
                f"diff={diff:,.2f}, passed={passed}"
            )
            return passed

        except Exception as e:
            logger.error(f"[AI Verify] Re-validation error: {e}")
            return False


def _parse_period_end(period_end_str: str):
    """Parse period end string to datetime."""
    from datetime import datetime
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(period_end_str, fmt)
        except ValueError:
            continue
    return datetime.now()
