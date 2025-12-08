"""
Sales Tracker Parser

Parses the QSP unit sales tracker Excel file containing disposal data.
The file tracks approximately 170 unit sales from 2018-2025.

Key outputs:
- Total units sold (all time)
- Current quarter units sold
- Current quarter proceeds
- Last 12 months sales data
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import re
from loguru import logger


@dataclass
class SaleEntry:
    """Represents a single unit sale/disposal."""
    unit_id: str
    address: Optional[str] = None
    valuation: float = 0.0
    sale_price: float = 0.0
    delta: float = 0.0  # Difference between sale price and valuation
    closing_date: Optional[datetime] = None
    status: Optional[str] = None
    raw_row: int = 0


@dataclass
class SalesTrackerResult:
    """Result of parsing a sales tracker Excel file."""
    entries: List[SaleEntry]
    total_units_sold: int  # All time
    quarter_units_sold: int  # Current quarter
    quarter_proceeds: float  # Total proceeds for current quarter
    last_12_months_sales: int  # Units sold in last 12 months
    last_12_months_proceeds: float
    period_end: Optional[datetime] = None  # Current quarter end date
    warnings: List[str] = field(default_factory=list)
    raw_dataframe: pd.DataFrame = None


class SalesTrackerParser:
    """
    Parser for QSP unit sales tracker Excel files.
    
    Usage:
        parser = SalesTrackerParser()
        result = parser.parse("Unit_Sales_tracker_Q3_updated.xlsx", period_end=datetime(2025, 9, 30))
        
        print(f"Total units sold: {result.total_units_sold}")
        print(f"Q3 2025 units: {result.quarter_units_sold}")
        print(f"Q3 2025 proceeds: €{result.quarter_proceeds:,.2f}")
    """
    
    def __init__(self):
        """Initialize the sales tracker parser."""
        pass
    
    def parse(self, file_path: str, period_end: Optional[datetime] = None, 
              sheet_name: str = None) -> SalesTrackerResult:
        """
        Parse sales tracker Excel file.
        
        Args:
            file_path: Path to sales tracker Excel file
            period_end: Period end date for current quarter (auto-detected if None)
            sheet_name: Specific sheet to parse (auto-detects if None)
            
        Returns:
            SalesTrackerResult with parsed sales and aggregated totals
        """
        logger.info(f"Parsing sales tracker file: {file_path}")
        
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
        
        # Extract period end from filename if not provided
        if period_end is None:
            period_end = self._extract_period_from_filename(file_path)
        
        # Parse sale entries
        entries = self._parse_entries(df_raw, data_start_row, column_mapping)
        
        # Calculate aggregates
        total_units_sold = len(entries)
        quarter_units_sold, quarter_proceeds = self._calculate_quarter_stats(
            entries, period_end
        )
        last_12_months_sales, last_12_months_proceeds = self._calculate_12_month_stats(
            entries, period_end
        )
        
        result = SalesTrackerResult(
            entries=entries,
            total_units_sold=total_units_sold,
            quarter_units_sold=quarter_units_sold,
            quarter_proceeds=quarter_proceeds,
            last_12_months_sales=last_12_months_sales,
            last_12_months_proceeds=last_12_months_proceeds,
            period_end=period_end,
            raw_dataframe=df_raw
        )
        
        # Validate and add warnings
        self._validate_result(result)
        
        logger.info(f"Parsed {total_units_sold} total sales")
        logger.info(f"Current quarter: {quarter_units_sold} units, €{quarter_proceeds:,.2f}")
        logger.info(f"Last 12 months: {last_12_months_sales} units, €{last_12_months_proceeds:,.2f}")
        
        return result
    
    def _detect_main_sheet(self, sheet_names: List[str]) -> str:
        """Detect the main data sheet from available sheets."""
        # Priority patterns for main data sheet
        priority_patterns = [
            r'sales',
            r'tracker',
            r'disposals',
            r'data',
            r'units',
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
            header_keywords = ['unit', 'address', 'valuation', 'sale', 'price', 
                             'delta', 'closing', 'date', 'status', 'disposal',
                             'adres', 'verkoop', 'waarde']
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
            'unit_id': ['unit', 'unit id', 'unit number', 'unitnr', 'appartement', 'unit code'],
            'address': ['address', 'adres', 'street', 'straat', 'location'],
            'valuation': ['valuation', 'waarde', 'book value', 'boekwaarde', 'carrying value'],
            'sale_price': ['sale', 'price', 'verkoop', 'prijs', 'proceeds', 'sale price', 'verkoopprijs'],
            'delta': ['delta', 'difference', 'verschil', 'diff', 'gain', 'loss'],
            'closing_date': ['closing', 'date', 'datum', 'sale date', 'verkoopdatum', 'closing date'],
            'status': ['status', 'state', 'condition'],
        }
        
        for col_type, search_patterns in patterns.items():
            for idx, header in enumerate(headers):
                for pattern in search_patterns:
                    if re.search(pattern, header, re.IGNORECASE):
                        mapping[col_type] = idx
                        break
                if col_type in mapping:
                    break
        
        return mapping
    
    def _fallback_column_mapping(self, df: pd.DataFrame) -> Dict[str, int]:
        """Fallback column mapping when header detection fails."""
        # Common structure: Unit, Address, Valuation, Sale Price, Delta, Closing Date, Status
        return {
            'unit_id': 0,
            'address': 1,
            'valuation': 2,
            'sale_price': 3,
            'delta': 4,
            'closing_date': 5,
            'status': 6,
        }
    
    def _parse_entries(self, df: pd.DataFrame, data_start: int, 
                       column_mapping: Dict[str, int]) -> List[SaleEntry]:
        """Parse sale entries from data rows."""
        entries = []
        
        unit_id_col = column_mapping.get('unit_id')
        address_col = column_mapping.get('address')
        valuation_col = column_mapping.get('valuation')
        sale_price_col = column_mapping.get('sale_price')
        delta_col = column_mapping.get('delta')
        closing_date_col = column_mapping.get('closing_date')
        status_col = column_mapping.get('status')
        
        for idx in range(data_start, len(df)):
            row = df.iloc[idx]
            
            # Skip empty rows
            if row.isna().all():
                continue
            
            # Extract unit identifier
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
            address = self._safe_get(row, address_col)
            valuation = self._parse_number(row.iloc[valuation_col]) if valuation_col is not None and valuation_col < len(row) else 0.0
            sale_price = self._parse_number(row.iloc[sale_price_col]) if sale_price_col is not None and sale_price_col < len(row) else 0.0
            
            # Calculate delta if not provided
            delta = 0.0
            if delta_col is not None and delta_col < len(row):
                delta = self._parse_number(row.iloc[delta_col])
            elif sale_price > 0 and valuation > 0:
                delta = sale_price - valuation
            
            # Parse closing date
            closing_date = None
            if closing_date_col is not None and closing_date_col < len(row):
                closing_date = self._parse_date(row.iloc[closing_date_col])
            
            status = self._safe_get(row, status_col)
            
            entry = SaleEntry(
                unit_id=unit_id,
                address=address,
                valuation=valuation,
                sale_price=sale_price,
                delta=delta,
                closing_date=closing_date,
                status=status,
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
        s = s.replace('€', '').replace('EUR', '').replace('$', '').replace('£', '')
        s = s.replace(' ', '').replace('.', '').replace(',', '.')
        
        try:
            return float(s)
        except ValueError:
            return 0.0
    
    def _parse_date(self, value) -> Optional[datetime]:
        """Parse a date from various formats."""
        if pd.isna(value):
            return None
        
        # If already a datetime
        if isinstance(value, datetime):
            return value
        
        # If pandas Timestamp
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        
        # Try parsing string formats
        s = str(value).strip()
        if not s or s.lower() == 'nan':
            return None
        
        # Common date formats
        date_formats = [
            '%Y-%m-%d',
            '%d-%m-%Y',
            '%d/%m/%Y',
            '%m/%d/%Y',
            '%Y/%m/%d',
            '%d.%m.%Y',
            '%d %B %Y',
            '%d %b %Y',
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        
        # Try pandas to_datetime as fallback
        try:
            return pd.to_datetime(s).to_pydatetime()
        except (ValueError, TypeError):
            return None
    
    def _extract_period_from_filename(self, file_path: str) -> Optional[datetime]:
        """Extract period end date from filename."""
        filename = Path(file_path).stem
        
        # Look for quarter pattern: Q1, Q2, Q3, Q4
        quarter_match = re.search(r'Q([1-4])\s*(\d{4})', filename, re.IGNORECASE)
        if quarter_match:
            quarter = int(quarter_match.group(1))
            year = int(quarter_match.group(2))
            
            # Calculate period end date
            if quarter == 1:
                return datetime(year, 3, 31)
            elif quarter == 2:
                return datetime(year, 6, 30)
            elif quarter == 3:
                return datetime(year, 9, 30)
            else:  # Q4
                return datetime(year, 12, 31)
        
        # Look for date pattern: DD-MM-YYYY or YYYY-MM-DD
        date_match = re.search(r'(\d{2})-(\d{2})-(\d{4})', filename)
        if date_match:
            day, month, year = date_match.groups()
            try:
                return datetime(int(year), int(month), int(day))
            except ValueError:
                pass
        
        return None
    
    def _calculate_quarter_stats(self, entries: List[SaleEntry], 
                                  period_end: Optional[datetime]) -> tuple:
        """
        Calculate current quarter sales statistics.
        
        Returns:
            (quarter_units_sold, quarter_proceeds)
        """
        if period_end is None:
            return 0, 0.0
        
        # Determine quarter start date
        if period_end.month == 3:  # Q1
            quarter_start = datetime(period_end.year, 1, 1)
        elif period_end.month == 6:  # Q2
            quarter_start = datetime(period_end.year, 4, 1)
        elif period_end.month == 9:  # Q3
            quarter_start = datetime(period_end.year, 7, 1)
        else:  # Q4
            quarter_start = datetime(period_end.year, 10, 1)
        
        quarter_units = 0
        quarter_proceeds = 0.0
        
        for entry in entries:
            if entry.closing_date is None:
                continue
            
            # Check if sale is in current quarter
            if quarter_start <= entry.closing_date <= period_end:
                quarter_units += 1
                quarter_proceeds += entry.sale_price
        
        return quarter_units, quarter_proceeds
    
    def _calculate_12_month_stats(self, entries: List[SaleEntry], 
                                   period_end: Optional[datetime]) -> tuple:
        """
        Calculate last 12 months sales statistics.
        
        Returns:
            (last_12_months_sales, last_12_months_proceeds)
        """
        if period_end is None:
            return 0, 0.0
        
        twelve_months_ago = period_end - timedelta(days=365)
        
        last_12_months_units = 0
        last_12_months_proceeds = 0.0
        
        for entry in entries:
            if entry.closing_date is None:
                continue
            
            # Check if sale is in last 12 months
            if twelve_months_ago <= entry.closing_date <= period_end:
                last_12_months_units += 1
                last_12_months_proceeds += entry.sale_price
        
        return last_12_months_units, last_12_months_proceeds
    
    def _validate_result(self, result: SalesTrackerResult):
        """Add validation warnings to result."""
        # Check for expected total sales (should be around 170)
        if result.total_units_sold < 150:
            result.warnings.append(
                f"Low total sales count: {result.total_units_sold} (expected ~170)"
            )
        elif result.total_units_sold > 200:
            result.warnings.append(
                f"High total sales count: {result.total_units_sold} (expected ~170)"
            )
        
        # Data quality checks (generic, no quarter-specific hardcoding)
        
        # Check for sales with zero sale price (may indicate incomplete data)
        zero_price_sales = sum(1 for entry in result.entries if entry.sale_price == 0)
        if zero_price_sales > 0:
            result.warnings.append(f"{zero_price_sales} sales have zero sale price")
        
        # Check for sales missing closing dates
        missing_dates = sum(1 for entry in result.entries if entry.closing_date is None)
        if missing_dates > 0:
            result.warnings.append(f"{missing_dates} sales missing closing dates")


# Convenience function for quick parsing
def parse_sales_tracker(file_path: str, period_end: Optional[datetime] = None) -> SalesTrackerResult:
    """Quick parse function for sales tracker files."""
    parser = SalesTrackerParser()
    return parser.parse(file_path, period_end)

