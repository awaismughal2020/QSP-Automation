"""
Rent Roll Parser

Parses the QSP rent roll Excel file containing unit-level rental data.
The file contains rental unit information.

Key outputs:
- Total units count
- Total annual rent
- Vacancy rate calculation
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional
from pathlib import Path
import pandas as pd
import re
from loguru import logger


@dataclass
class RentRollEntry:
    """Represents a single unit in the rent roll."""
    unit_id: str
    owner: Optional[str] = None
    complex: Optional[str] = None
    subcomplex: Optional[str] = None
    address: Optional[str] = None
    monthly_rent: float = 0.0
    annual_rent: float = 0.0
    is_vacant: bool = False
    raw_row: int = 0


@dataclass
class RentRollResult:
    """Result of parsing a rent roll Excel file."""
    entries: List[RentRollEntry]
    total_units: int
    total_annual_rent: float
    vacancy_rate: float  # Percentage as decimal (e.g., 0.05 for 5%)
    vacant_units: int
    occupied_units: int
    warnings: List[str] = field(default_factory=list)
    raw_dataframe: pd.DataFrame = None


class RentRollParser:
    """
    Parser for QSP rent roll Excel files.
    
    Usage:
        parser = RentRollParser()
        result = parser.parse("QSP huurlijst 1-10-2025.xlsx")
        
        print(f"Total units: {result.total_units}")
        print(f"Annual rent: €{result.total_annual_rent:,.2f}")
        print(f"Vacancy rate: {result.vacancy_rate:.2%}")
    """
    
    def __init__(self):
        """Initialize the rent roll parser."""
        pass
    
    def parse(self, file_path: str, sheet_name: str = None) -> RentRollResult:
        """
        Parse rent roll Excel file.
        
        Args:
            file_path: Path to rent roll Excel file
            sheet_name: Specific sheet to parse (auto-detects if None)
            
        Returns:
            RentRollResult with parsed units and aggregated totals
        """
        logger.info(f"Parsing rent roll file: {file_path}")
        
        # Load Excel file
        xlsx = pd.ExcelFile(file_path)
        
        # Auto-detect the main data sheet
        if sheet_name is None:
            sheet_name = self._detect_main_sheet(xlsx.sheet_names)
        
        logger.info(f"Using sheet: {sheet_name}")
        
        # Read raw data
        df_raw = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
        
        # Detect header row and data start
        header_row, data_start_row = self._detect_data_boundaries(df_raw)
        column_mapping = self._detect_columns(df_raw, header_row)
        
        logger.info(f"Detected header row: {header_row}, data starts: {data_start_row}")
        logger.info(f"Column mapping: {column_mapping}")
        
        # Parse unit entries
        entries = self._parse_entries(df_raw, data_start_row, column_mapping)
        
        # Calculate aggregates
        total_units = len(entries)
        total_annual_rent = sum(entry.annual_rent for entry in entries)
        vacant_units = sum(1 for entry in entries if entry.is_vacant)
        occupied_units = total_units - vacant_units
        vacancy_rate = vacant_units / total_units if total_units > 0 else 0.0
        
        result = RentRollResult(
            entries=entries,
            total_units=total_units,
            total_annual_rent=total_annual_rent,
            vacancy_rate=vacancy_rate,
            vacant_units=vacant_units,
            occupied_units=occupied_units,
            raw_dataframe=df_raw
        )
        
        # Validate and add warnings
        self._validate_result(result)
        
        logger.info(f"Parsed {total_units} units, annual rent: €{total_annual_rent:,.2f}")
        logger.info(f"Vacancy: {vacant_units} units ({vacancy_rate:.2%})")
        
        return result
    
    def _detect_main_sheet(self, sheet_names: List[str]) -> str:
        """Detect the main data sheet from available sheets."""
        # Priority patterns for main data sheet
        priority_patterns = [
            r'huurlijst',
            r'rent.*roll',
            r'units',
            r'data',
        ]
        
        for pattern in priority_patterns:
            for name in sheet_names:
                if re.search(pattern, name, re.IGNORECASE):
                    return name
        
        # Fallback to first sheet
        return sheet_names[0]
    
    def _detect_data_boundaries(self, df: pd.DataFrame) -> tuple:
        """
        Detect where headers and data actually start.
        
        Returns:
            (header_row_index, data_start_row_index)
        """
        header_row = None
        data_start = None
        
        for idx, row in df.iterrows():
            row_values = [str(v).strip() for v in row.values if pd.notna(v)]
            row_text = ' '.join(row_values).lower()
            
            # Look for header indicators
            header_keywords = ['owner', 'complex', 'subcomplex', 'address', 'rent', 
                             'eigenaar', 'adres', 'huur', 'unit']
            if any(keyword in row_text for keyword in header_keywords):
                header_row = idx
                continue
            
            # Look for first data row (has address or unit identifier)
            if header_row is not None:
                # Check if row has meaningful data (not all empty/NaN)
                non_empty = sum(1 for v in row.values if pd.notna(v) and str(v).strip())
                if non_empty >= 3:  # At least 3 non-empty cells
                    data_start = idx
                    break
        
        # Fallbacks
        if header_row is None:
            header_row = 0  # Assume first row is header
        if data_start is None:
            data_start = header_row + 1
            
        return header_row, data_start
    
    def _detect_columns(self, df: pd.DataFrame, header_row: int) -> Dict[str, int]:
        """
        Detect column positions by header text patterns.
        
        Returns dict mapping column type to column index.
        """
        if header_row >= len(df):
            return self._fallback_column_mapping(df)
        
        headers = df.iloc[header_row].fillna('').astype(str).str.lower()
        
        mapping = {}
        
        # Column detection patterns
        patterns = {
            'owner': ['owner', 'eigenaar', 'owner name'],
            'complex': ['complex', 'complex name'],
            'subcomplex': ['subcomplex', 'sub complex', 'sub-complex'],
            'unit_id': ['unit', 'unit id', 'unit number', 'unitnr', 'appartement'],
            'address': ['address', 'adres', 'street', 'straat'],
            'monthly_rent': ['monthly', 'maandelijks', 'rent per month', 'huur per maand'],
            'annual_rent': ['annual', 'jaarlijks', 'yearly', 'rent per year', 'huur per jaar', 'total rent'],
            'status': ['status', 'vacant', 'occupied', 'leeg', 'bezet', 'vacancy'],
        }
        
        for col_type, search_patterns in patterns.items():
            for idx, header in enumerate(headers):
                for pattern in search_patterns:
                    if re.search(pattern, header, re.IGNORECASE):
                        mapping[col_type] = idx
                        break
                if col_type in mapping:
                    break
        
        # Try to detect annual rent column if not found (look for numeric column with large values)
        if 'annual_rent' not in mapping:
            for idx in range(len(headers)):
                # Sample data from this column
                sample_data = df.iloc[header_row + 1:header_row + 20, idx]
                numeric_values = []
                for val in sample_data:
                    if pd.notna(val):
                        num_val = self._parse_number(val)
                        if num_val > 1000:  # Annual rent should be > 1000 EUR
                            numeric_values.append(num_val)
                
                if len(numeric_values) > 5:  # Found multiple large values
                    mapping['annual_rent'] = idx
                    break
        
        # Fallback: if annual_rent not found, try monthly_rent and multiply
        if 'annual_rent' not in mapping and 'monthly_rent' in mapping:
            # Will calculate annual from monthly in parsing
            pass
        
        return mapping
    
    def _fallback_column_mapping(self, df: pd.DataFrame) -> Dict[str, int]:
        """Fallback column mapping when header detection fails."""
        # Common structure: Owner, Complex, Subcomplex, Unit, Address, Rent...
        return {
            'owner': 0,
            'complex': 1,
            'subcomplex': 2,
            'unit_id': 3,
            'address': 4,
            'annual_rent': 5,
        }
    
    def _parse_entries(self, df: pd.DataFrame, data_start: int, 
                       column_mapping: Dict[str, int]) -> List[RentRollEntry]:
        """Parse unit entries from data rows."""
        entries = []
        
        owner_col = column_mapping.get('owner')
        complex_col = column_mapping.get('complex')
        subcomplex_col = column_mapping.get('subcomplex')
        unit_id_col = column_mapping.get('unit_id')
        address_col = column_mapping.get('address')
        monthly_rent_col = column_mapping.get('monthly_rent')
        annual_rent_col = column_mapping.get('annual_rent')
        status_col = column_mapping.get('status')
        
        for idx in range(data_start, len(df)):
            row = df.iloc[idx]
            
            # Skip empty rows
            if row.isna().all():
                continue
            
            # Extract unit identifier (try multiple columns)
            unit_id = None
            if unit_id_col is not None and unit_id_col < len(row):
                unit_id = str(row.iloc[unit_id_col]).strip()
            elif address_col is not None and address_col < len(row):
                # Use address as fallback identifier
                unit_id = str(row.iloc[address_col]).strip()
            
            # Skip if no unit identifier found
            if not unit_id or unit_id == 'nan' or unit_id == '':
                continue
            
            # Extract other fields
            owner = self._safe_get(row, owner_col)
            complex_name = self._safe_get(row, complex_col)
            subcomplex_name = self._safe_get(row, subcomplex_col)
            address = self._safe_get(row, address_col)
            
            # Extract rent values
            monthly_rent = 0.0
            annual_rent = 0.0
            
            if annual_rent_col is not None and annual_rent_col < len(row):
                annual_rent = self._parse_number(row.iloc[annual_rent_col])
            
            if monthly_rent_col is not None and monthly_rent_col < len(row):
                monthly_rent = self._parse_number(row.iloc[monthly_rent_col])
                # Calculate annual if we have monthly but not annual
                if annual_rent == 0.0 and monthly_rent > 0:
                    annual_rent = monthly_rent * 12
            
            # Determine vacancy status
            is_vacant = False
            if status_col is not None and status_col < len(row):
                status = str(row.iloc[status_col]).strip().lower()
                is_vacant = any(keyword in status for keyword in ['vacant', 'leeg', 'empty', '0'])
            else:
                # Infer from rent: if annual rent is 0 or very low, likely vacant
                if annual_rent < 100:
                    is_vacant = True
            
            entry = RentRollEntry(
                unit_id=unit_id,
                owner=owner,
                complex=complex_name,
                subcomplex=subcomplex_name,
                address=address,
                monthly_rent=monthly_rent,
                annual_rent=annual_rent,
                is_vacant=is_vacant,
                raw_row=idx
            )
            
            entries.append(entry)
        
        return entries
    
    def _safe_get(self, row: pd.Series, col_idx: Optional[int]) -> Optional[str]:
        """Safely get string value from row at column index."""
        if col_idx is None or col_idx >= len(row):
            return None
        value = row.iloc[col_idx]
        if pd.isna(value):
            return None
        return str(value).strip()
    
    def _parse_number(self, value) -> float:
        """Parse a number from various formats."""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        
        # Handle string formats
        s = str(value).strip()
        if not s or s == '-' or s.lower() == 'nan':
            return 0.0
        
        # Remove currency symbols and thousand separators
        s = s.replace('€', '').replace('EUR', '').replace('$', '')
        s = s.replace(' ', '').replace('.', '').replace(',', '.')
        
        try:
            return float(s)
        except ValueError:
            return 0.0
    
    def _validate_result(self, result: RentRollResult):
        """Add validation warnings to result."""
        # Data quality checks (no hardcoded expected values)
        
        # Check for suspiciously low unit count
        if result.total_units < 10:
            result.warnings.append(f"Very low unit count: {result.total_units} - verify file is complete")
        
        # Check for zero or negative annual rent
        if result.total_annual_rent <= 0:
            result.warnings.append("Total annual rent is zero or negative - verify data")
        
        # Check for units with zero rent that aren't marked vacant
        zero_rent_occupied = sum(
            1 for entry in result.entries 
            if entry.annual_rent == 0 and not entry.is_vacant
        )
        if zero_rent_occupied > 0:
            result.warnings.append(f"{zero_rent_occupied} units have zero rent but are not marked vacant")
        
        # Check for missing required fields
        missing_address = sum(1 for entry in result.entries if not entry.address)
        if missing_address > 0:
            result.warnings.append(f"{missing_address} units missing address information")


# Convenience function for quick parsing
def parse_rent_roll(file_path: str) -> RentRollResult:
    """Quick parse function for rent roll files."""
    parser = RentRollParser()
    return parser.parse(file_path)

