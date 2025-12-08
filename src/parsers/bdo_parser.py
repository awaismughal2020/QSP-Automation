"""
BDO Excel Parser - Schema Drift Resistant

This parser handles quarterly BDO financial exports which have varying schemas:
- Column count varies between 8-14 columns across quarters
- Header row position can shift
- Account codes are occasionally renamed
- New accounts are introduced over time

Strategy: Parse by account code recognition, not column position.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import pandas as pd
import numpy as np
import yaml
import re
from loguru import logger


@dataclass
class AccountEntry:
    """Represents a single account line from BDO data."""
    code: str
    name: str
    opening_balance: float = 0.0
    mutations: Dict[str, float] = field(default_factory=dict)  # period -> amount
    closing_balance: float = 0.0
    raw_row: int = 0
    

@dataclass
class BDOParseResult:
    """Result of parsing a BDO Excel file."""
    accounts: Dict[str, AccountEntry]  # code -> entry
    period_end: str  # e.g., "2025-09-30"
    quarter: str  # e.g., "Q3 2025"
    column_mapping: Dict[str, int]  # detected column positions
    schema_version: str  # detected schema pattern
    warnings: List[str] = field(default_factory=list)
    raw_dataframe: pd.DataFrame = None


class BDOParser:
    """
    Schema-drift resistant parser for BDO quarterly Excel exports.
    
    Usage:
        parser = BDOParser(account_mappings_path="config/account_mappings.yaml")
        result = parser.parse("Cijfers QSP 30-09-2025.xlsx")
        
        # Access parsed data
        real_estate = result.accounts.get("1600000")
        print(f"Property value: {real_estate.closing_balance:,.2f}")
    """
    
    def __init__(self, account_mappings_path: str = "config/account_mappings.yaml"):
        self.mappings = self._load_mappings(account_mappings_path)
        self.known_account_codes = self._extract_known_codes()
        
    def _load_mappings(self, path: str) -> dict:
        """Load account code mappings from YAML."""
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    
    def _extract_known_codes(self) -> set:
        """Extract all known account codes for validation."""
        codes = set()
        for account in self.mappings.get('account_codes', {}).values():
            for code in account.get('codes', []):
                # Handle wildcard patterns
                codes.add(code.replace('*', ''))
        return codes
    
    def parse(self, file_path: str, sheet_name: str = None) -> BDOParseResult:
        """
        Parse BDO Excel file with automatic schema detection.
        
        Args:
            file_path: Path to BDO Excel file
            sheet_name: Specific sheet to parse (auto-detects if None)
            
        Returns:
            BDOParseResult with parsed accounts and metadata
        """
        logger.info(f"Parsing BDO file: {file_path}")
        
        # Load Excel file
        xlsx = pd.ExcelFile(file_path)
        
        # Auto-detect the main data sheet
        if sheet_name is None:
            sheet_name = self._detect_main_sheet(xlsx.sheet_names)
        
        logger.info(f"Using sheet: {sheet_name}")
        
        # Read raw data without headers (we'll detect them)
        df_raw = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
        
        # Detect schema structure
        header_row, data_start_row = self._detect_data_boundaries(df_raw)
        column_mapping = self._detect_columns(df_raw, header_row)
        
        logger.info(f"Detected header row: {header_row}, data starts: {data_start_row}")
        logger.info(f"Column mapping: {column_mapping}")
        
        # Parse accounts
        accounts = self._parse_accounts(df_raw, data_start_row, column_mapping)
        
        # Extract period information
        period_end, quarter = self._extract_period_info(df_raw, file_path)
        
        # Determine schema version
        schema_version = self._classify_schema(column_mapping, len(df_raw.columns))
        
        result = BDOParseResult(
            accounts=accounts,
            period_end=period_end,
            quarter=quarter,
            column_mapping=column_mapping,
            schema_version=schema_version,
            raw_dataframe=df_raw
        )
        
        # Validate and add warnings
        self._validate_result(result)
        
        return result
    
    def _detect_main_sheet(self, sheet_names: List[str]) -> str:
        """Detect the main data sheet from available sheets."""
        # Priority patterns for main data sheet
        priority_patterns = [
            r'BalansenWinstverlies',
            r'Balansen',
            r'P&L',
            r'Winst.*verlies',
        ]
        
        for pattern in priority_patterns:
            for name in sheet_names:
                if re.search(pattern, name, re.IGNORECASE):
                    return name
        
        # Fallback to first sheet
        return sheet_names[0]
    
    def _detect_data_boundaries(self, df: pd.DataFrame) -> Tuple[int, int]:
        """
        Detect where headers and data actually start.
        
        BDO files structure:
        - Row 0: Often contains column headers (Eindbalans, Mutaties, Saldi)
        - Some rows may be category labels (Balans, Materiële vaste activa, etc.)
        - Data rows have 7-digit account codes in column 0
        
        Returns:
            (header_row_index, data_start_row_index)
        """
        header_row = None
        data_start = None
        
        for idx, row in df.iterrows():
            row_values = [str(v).strip() for v in row.values if pd.notna(v)]
            row_text = ' '.join(row_values).lower()
            
            # Look for header indicators in early rows
            # BDO files often have balance/mutation headers in row 0
            if header_row is None:
                header_patterns = [
                    'eindbalans', 'beginsaldo', 'saldi', 'mutaties',
                    'grootboek', 'account', 'omschrijving'
                ]
                if any(pattern in row_text for pattern in header_patterns):
                    header_row = idx
            
            # Look for first numeric account code (7 digits starting with 0-9)
            first_cell = str(row.iloc[0]).strip()
            if re.match(r'^\d{7}$', first_cell):
                if data_start is None:
                    data_start = idx
                    # If we found data but no header, header is likely row 0
                    if header_row is None:
                        header_row = 0
                    break
        
        # Fallbacks based on typical BDO structure
        if header_row is None:
            header_row = 0  # Assume row 0 has headers
        if data_start is None:
            # Scan for first data row
            for idx in range(len(df)):
                first_cell = str(df.iloc[idx, 0]).strip()
                if re.match(r'^\d{7}$', first_cell):
                    data_start = idx
                    break
            if data_start is None:
                data_start = 5  # Common position
            
        return header_row, data_start
    
    def _detect_columns(self, df: pd.DataFrame, header_row: int) -> Dict[str, int]:
        """
        Detect column positions by header text patterns.
        
        BDO file structure (typical):
        - Col 0: Account code (7-digit number)
        - Col 1: Account name
        - Col 2: Opening balance (Eindbalans/Beginsaldo)
        - Col 3-N: Mutations (Mutaties Q1, Q2, etc.)
        - Last col: Closing balance (Saldi/Eindsaldo)
        
        Returns dict mapping column type to column index.
        """
        if header_row >= len(df):
            return self._fallback_column_mapping(df)
        
        # Account code and name are always in first two columns
        mapping = {
            'account_code': 0,
            'account_name': 1
        }
        
        # Search header row for balance column indicators
        headers = df.iloc[header_row].fillna('').astype(str).str.lower()
        
        for idx, header in enumerate(headers):
            # Opening balance patterns (Dutch: Eindbalans [previous period], Beginsaldo)
            if any(pattern in header for pattern in ['eindbalans', 'beginsaldo', 'opening', 'start']):
                if 'opening_balance' not in mapping:
                    mapping['opening_balance'] = idx
            
            # Closing balance patterns (Dutch: Saldi per, Eindsaldo)
            if any(pattern in header for pattern in ['saldi per', 'eindsaldo', 'closing', 'end balance']):
                mapping['closing_balance'] = idx
        
        # If headers weren't detected, use column position heuristics
        if 'opening_balance' not in mapping or 'closing_balance' not in mapping:
            # Find data row to analyze numeric columns
            data_start = header_row + 1
            for idx in range(header_row, min(header_row + 20, len(df))):
                first_cell = str(df.iloc[idx, 0]).strip()
                if re.match(r'^\d{7}$', first_cell):
                    data_start = idx
                    break
            
            # Analyze which columns have numeric data
            numeric_cols = []
            for col_idx in range(2, df.shape[1]):  # Start from col 2 (after code/name)
                has_numeric = False
                for row_idx in range(data_start, min(data_start + 10, len(df))):
                    val = df.iloc[row_idx, col_idx]
                    if pd.notna(val):
                        if isinstance(val, (int, float)):
                            has_numeric = True
                            break
                        elif isinstance(val, str):
                            try:
                                float(val.replace(',', '.').replace(' ', ''))
                                has_numeric = True
                                break
                            except ValueError:
                                pass
                if has_numeric:
                    numeric_cols.append(col_idx)
            
            # Opening balance is first numeric column, closing is last
            if numeric_cols:
                if 'opening_balance' not in mapping:
                    mapping['opening_balance'] = numeric_cols[0]
                if 'closing_balance' not in mapping:
                    mapping['closing_balance'] = numeric_cols[-1]
        
        # Final fallback: use standard positions
        if 'opening_balance' not in mapping:
            mapping['opening_balance'] = 2
        if 'closing_balance' not in mapping:
            mapping['closing_balance'] = df.shape[1] - 2  # Second to last (last is often empty)
            
        return mapping
    
    def _fallback_column_mapping(self, df: pd.DataFrame) -> Dict[str, int]:
        """Fallback column mapping when header detection fails."""
        return {
            'account_code': 0,
            'account_name': 1,
            'opening_balance': 2,
            'closing_balance': df.shape[1] - 1
        }
    
    def _parse_accounts(self, df: pd.DataFrame, data_start: int, 
                        column_mapping: Dict[str, int]) -> Dict[str, AccountEntry]:
        """Parse account entries from data rows."""
        accounts = {}
        
        code_col = column_mapping.get('account_code', 0)
        name_col = column_mapping.get('account_name', 1)
        opening_col = column_mapping.get('opening_balance')
        closing_col = column_mapping.get('closing_balance')
        
        for idx in range(data_start, len(df)):
            row = df.iloc[idx]
            
            # Extract account code
            code = str(row.iloc[code_col]).strip()
            
            # Skip non-account rows
            if not re.match(r'^\d{7}$', code):
                continue
            
            # Extract values
            name = str(row.iloc[name_col]).strip() if name_col < len(row) else ""
            
            opening = self._parse_number(row.iloc[opening_col]) if opening_col else 0.0
            closing = self._parse_number(row.iloc[closing_col]) if closing_col else 0.0
            
            # Parse mutation columns (between opening and closing)
            mutations = {}
            if opening_col and closing_col:
                for mut_idx in range(opening_col + 1, closing_col):
                    mut_value = self._parse_number(row.iloc[mut_idx])
                    if mut_value != 0:
                        mutations[f"mutation_{mut_idx}"] = mut_value
            
            accounts[code] = AccountEntry(
                code=code,
                name=name,
                opening_balance=opening,
                mutations=mutations,
                closing_balance=closing,
                raw_row=idx
            )
        
        logger.info(f"Parsed {len(accounts)} accounts")
        return accounts
    
    def _parse_number(self, value) -> float:
        """Parse a number from various formats."""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        
        # Handle string formats
        s = str(value).strip()
        if not s or s == '-':
            return 0.0
            
        # Remove thousand separators and handle comma decimal
        s = s.replace(' ', '').replace('.', '').replace(',', '.')
        
        try:
            return float(s)
        except ValueError:
            return 0.0
    
    def _extract_period_info(self, df: pd.DataFrame, file_path: str) -> Tuple[str, str]:
        """Extract period end date and quarter from file."""
        # Try to extract from filename first
        filename = Path(file_path).stem
        date_match = re.search(r'(\d{2})-(\d{2})-(\d{4})', filename)
        
        if date_match:
            day, month, year = date_match.groups()
            period_end = f"{year}-{month}-{day}"
            
            # Determine quarter
            month_int = int(month)
            quarter_num = (month_int - 1) // 3 + 1
            quarter = f"Q{quarter_num} {year}"
            
            return period_end, quarter
        
        # Fallback: look in first few rows
        for idx in range(min(5, len(df))):
            row_text = ' '.join(str(v) for v in df.iloc[idx].values if pd.notna(v))
            date_match = re.search(r'(\d{2})-(\d{2})-(\d{4})', row_text)
            if date_match:
                day, month, year = date_match.groups()
                period_end = f"{year}-{month}-{day}"
                quarter_num = (int(month) - 1) // 3 + 1
                return period_end, f"Q{quarter_num} {year}"
        
        return "unknown", "unknown"
    
    def _classify_schema(self, column_mapping: Dict[str, int], total_cols: int) -> str:
        """Classify the schema version for debugging."""
        if total_cols <= 8:
            return "compact"
        elif total_cols <= 10:
            return "standard"
        elif total_cols <= 12:
            return "extended"
        else:
            return "full"
    
    def _validate_result(self, result: BDOParseResult):
        """Add validation warnings to result."""
        # Check for missing critical accounts
        critical_codes = ['1600000', '1600200']  # Property, depreciation
        for code in critical_codes:
            if code not in result.accounts:
                # Check if there's a similar account (might be a variant)
                found_similar = False
                for existing_code in result.accounts.keys():
                    if existing_code.startswith(code[:4]):
                        found_similar = True
                        break
                if not found_similar:
                    result.warnings.append(f"Missing critical account: {code}")
        
        # Check for unexpected new accounts (only warn, don't fail)
        # Use a more lenient check - check if code prefix matches known patterns
        for code in result.accounts.keys():
            if len(code) == 7 and code.isdigit():
                # Check if this code prefix is known
                code_prefix_4 = code[:4]
                code_prefix_3 = code[:3]
                known = False
                
                # Check against known code prefixes
                for known_code in self.known_account_codes:
                    if len(known_code) >= 4:
                        if code.startswith(known_code[:4]):
                            known = True
                            break
                    elif len(known_code) >= 3:
                        if code.startswith(known_code[:3]):
                            known = True
                            break
                
                # Also check wildcard patterns in mappings
                if not known:
                    for account_def in self.mappings.get('account_codes', {}).values():
                        for pattern in account_def.get('codes', []):
                            if pattern.endswith('*'):
                                prefix = pattern[:-1]
                                if code.startswith(prefix):
                                    known = True
                                    break
                        if known:
                            break
                
                if not known:
                    result.warnings.append(f"Unknown account code: {code}")
    
    def get_account_by_category(self, result: BDOParseResult, category: str) -> List[AccountEntry]:
        """Get all accounts matching a category from mappings."""
        matching = []
        category_codes = []
        
        for account_def in self.mappings.get('account_codes', {}).values():
            if account_def.get('category') == category:
                category_codes.extend(account_def.get('codes', []))
        
        for code, entry in result.accounts.items():
            for pattern in category_codes:
                if pattern.endswith('*'):
                    if code.startswith(pattern[:-1]):
                        matching.append(entry)
                        break
                elif code == pattern:
                    matching.append(entry)
                    break
        
        return matching
    
    def calculate_totals(self, result: BDOParseResult) -> Dict[str, float]:
        """
        Calculate summary totals from parsed accounts.
        
        Dutch chart of accounts structure (QSP):
        - 10xxxxx, 11xxxxx = Equity (EV - Eigen Vermogen)
        - 16xxxxx = Real estate assets (VG - Vastgoed)
        - 17xxxxx = Financial fixed assets (FVA)
        - 19xxxxx = Long-term liabilities (LLS - Loans)
        - 20xxxxx = Receivables (VOR - Vorderingen)
        - 23-27xxxxx = Current liabilities (LM, KLS)
        - 40-47xxxxx = Expenses (KN, FINBL, AFKN)
        - 80-84xxxxx = Income (OPB - Opbrengsten)
        - 95xxxxx = Tax/Result (VPB)
        """
        totals = {
            'total_assets': 0.0,
            'real_estate_net': 0.0,
            'financial_fixed_assets': 0.0,
            'current_assets': 0.0,
            'total_equity': 0.0,
            'total_liabilities': 0.0,
            'long_term_liabilities': 0.0,
            'current_liabilities': 0.0,
            'total_income': 0.0,
            'total_expenses': 0.0,
            'direct_result': 0.0,
        }
        
        for code, entry in result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            first_digit = code[0] if code else ''
            
            # EQUITY (10xxxxx, 11xxxxx - EV accounts)
            if first_two in ('10', '11'):
                totals['total_equity'] += entry.closing_balance
            
            # REAL ESTATE ASSETS (16xxxxx - VG accounts)
            elif first_two == '16':
                totals['real_estate_net'] += entry.closing_balance
            
            # FINANCIAL FIXED ASSETS (17xxxxx, 18xxxxx - FVA accounts)
            elif first_two in ('17', '18'):
                totals['financial_fixed_assets'] += entry.closing_balance
            
            # LONG-TERM LIABILITIES (19xxxxx - LLS loan accounts)
            elif first_two == '19':
                totals['long_term_liabilities'] += entry.closing_balance
            
            # CURRENT ASSETS - Receivables (20xxxxx, 21xxxxx, 22xxxxx - VOR accounts)
            elif first_two in ('20', '21', '22'):
                totals['current_assets'] += entry.closing_balance
            
            # CURRENT LIABILITIES (23-29xxxxx - LM, KLS, BEL, RC accounts)
            elif first_two in ('23', '24', '25', '26', '27', '28', '29'):
                totals['current_liabilities'] += entry.closing_balance
            
            # EXPENSES (40-49xxxxx - KN, FINBL, AFKN accounts)
            elif first_digit == '4':
                totals['total_expenses'] += entry.closing_balance
            
            # INCOME (80-89xxxxx - OPB accounts)
            elif first_digit == '8':
                totals['total_income'] += entry.closing_balance
            
            # TAX/RESULT (95xxxxx - VPB accounts)
            elif first_two == '95':
                totals['direct_result'] += entry.closing_balance
        
        # Calculate totals
        totals['total_assets'] = (totals['real_estate_net'] + 
                                  totals['financial_fixed_assets'] + 
                                  totals['current_assets'])
        
        totals['total_liabilities'] = (totals['long_term_liabilities'] + 
                                       totals['current_liabilities'])
        
        # Direct result from P&L if not calculated from result accounts
        if totals['direct_result'] == 0:
            totals['direct_result'] = totals['total_income'] + totals['total_expenses']
        
        return totals


# Convenience function for quick parsing
def parse_bdo_file(file_path: str, mappings_path: str = "config/account_mappings.yaml") -> BDOParseResult:
    """Quick parse function for BDO files."""
    parser = BDOParser(mappings_path)
    return parser.parse(file_path)

