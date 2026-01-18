"""
Tests for Management Accounts Builder
"""

import pytest
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook, load_workbook

from src.transformers.management_accounts import (
    ManagementAccountsBuilder,
    ManagementAccountsConfig
)
from src.parsers.bdo_parser import BDOParseResult, AccountEntry


class TestManagementAccountsBuilder:
    """Test Management Accounts builder."""
    
    @pytest.fixture
    def previous_ma_file(self, tmp_path):
        """Create a previous quarter Management Accounts file."""
        file_path = tmp_path / "prev_ma.xlsx"
        
        wb = Workbook()
        
        # Create BDO sheet
        bdo_sheet = wb.active
        bdo_sheet.title = "BDO - Q2-25"
        bdo_sheet.cell(row=1, column=1, value="Account Code")
        bdo_sheet.cell(row=1, column=2, value="Account Name")
        bdo_sheet.cell(row=1, column=3, value="Opening Balance")
        bdo_sheet.cell(row=1, column=4, value="Mutations")
        bdo_sheet.cell(row=1, column=5, value="Closing Balance")
        
        # Create summary sheet
        summary_sheet = wb.create_sheet("Management Cijfers - Q2 2025")
        summary_sheet.cell(row=1, column=1, value="Management Accounts QSP ESS B.V. - Q2 2025")
        summary_sheet.cell(row=2, column=1, value="Item")
        summary_sheet.cell(row=2, column=2, value="2025-06-30")
        
        # Add line items
        line_items = [
            "Deferred Tax Asset",
            "Real estate",
            "Financial fixed assets",
            "Accounts receivable",
            "Service costs to be settled",
            "Prepaid expenses",
            "Cash",
            "Equity",
            "Senior Loan",
            "Accrued Interest",
        ]
        
        for row_idx, item in enumerate(line_items, 3):
            summary_sheet.cell(row=row_idx, column=1, value=item)
            summary_sheet.cell(row=row_idx, column=2, value=1000000.00)
        
        wb.save(file_path)
        return file_path
    
    @pytest.fixture
    def bdo_result(self):
        """Create sample BDO result."""
        accounts = {
            "1600000": AccountEntry(
                code="1600000",
                name="Verkrijgingsprijs",
                opening_balance=12000000.00,
                closing_balance=12000000.00,
                mutations={},
                raw_row=5
            ),
            "1600200": AccountEntry(
                code="1600200",
                name="Cumulatieve afschrijving",
                opening_balance=-600000.00,
                closing_balance=-630000.00,
                mutations={"mutation_4": -30000.00},
                raw_row=6
            ),
            "1760000": AccountEntry(
                code="1760000",
                name="Te vorderen DMRRP",
                opening_balance=600000.00,
                closing_balance=600000.00,
                mutations={},
                raw_row=7
            ),
            "1790002": AccountEntry(
                code="1790002",
                name="Latente belastingvordering",
                opening_balance=100000.00,
                closing_balance=100000.00,
                mutations={},
                raw_row=8
            ),
            "1100000": AccountEntry(
                code="1100000",
                name="Kas",
                opening_balance=150000.00,
                closing_balance=200000.00,
                mutations={"mutation_4": 50000.00},
                raw_row=9
            ),
            "1110000": AccountEntry(
                code="1110000",
                name="Bank",
                opening_balance=500000.00,
                closing_balance=600000.00,
                mutations={"mutation_4": 100000.00},
                raw_row=10
            ),
            "1300000": AccountEntry(
                code="1300000",
                name="Debiteuren",
                opening_balance=250000.00,
                closing_balance=250000.00,
                mutations={},
                raw_row=11
            ),
            "0500000": AccountEntry(
                code="0500000",
                name="Aandelenkapitaal",
                opening_balance=5000000.00,
                closing_balance=5000000.00,
                mutations={},
                raw_row=12
            ),
            "0600000": AccountEntry(
                code="0600000",
                name="Ingehouden winst",
                opening_balance=5000000.00,
                closing_balance=5120000.00,
                mutations={"mutation_4": 120000.00},
                raw_row=13
            ),
            "0900000": AccountEntry(
                code="0900000",
                name="Senior lening",
                opening_balance=3500000.00,
                closing_balance=3500000.00,
                mutations={},
                raw_row=14
            ),
            "0920000": AccountEntry(
                code="0920000",
                name="Opgelopen rente",
                opening_balance=50000.00,
                closing_balance=75000.00,
                mutations={"mutation_4": 25000.00},
                raw_row=15
            ),
        }
        
        return BDOParseResult(
            accounts=accounts,
            period_end="2025-09-30",
            quarter="Q3 2025",
            column_mapping={},
            schema_version="standard",
            warnings=[]
        )
    
    def test_add_bdo_sheet(self, previous_ma_file, bdo_result, tmp_path):
        """Test adding new BDO sheet."""
        output_path = tmp_path / "new_ma.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q3 2025",
            period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025"
        )
        
        builder = ManagementAccountsBuilder(
            str(previous_ma_file),
            str(output_path),
            config
        )
        
        builder._add_bdo_sheet(bdo_result)
        
        # Verify sheet was added
        wb = load_workbook(output_path)
        assert "BDO - Q3-25" in wb.sheetnames
        
        # Verify data in sheet
        bdo_sheet = wb["BDO - Q3-25"]
        assert bdo_sheet.cell(row=1, column=1).value == "Account Code"
        assert bdo_sheet.cell(row=2, column=1).value == "1600000"
    
    def test_update_summary_sheet(self, previous_ma_file, bdo_result, tmp_path):
        """Test updating summary sheet with new column."""
        output_path = tmp_path / "new_ma.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q3 2025",
            period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025"
        )
        
        builder = ManagementAccountsBuilder(
            str(previous_ma_file),
            str(output_path),
            config
        )
        
        builder._update_summary_sheet(bdo_result)
        
        # Verify new column was added
        wb = load_workbook(output_path)
        summary_sheet = wb["Management Cijfers - Q3 2025"]
        
        # Should have 3 columns now (Item, Q2, Q3)
        assert summary_sheet.cell(row=2, column=3).value == datetime(2025, 9, 30)
        
        # Verify values were populated
        # Real estate should be sum of 1600000, 1600200, etc.
        real_estate_row = None
        for row_idx in range(3, summary_sheet.max_row + 1):
            if summary_sheet.cell(row=row_idx, column=1).value == "Real estate":
                real_estate_row = row_idx
                break
        
        if real_estate_row:
            value = summary_sheet.cell(row=real_estate_row, column=3).value
            assert value is not None
            assert isinstance(value, (int, float))
    
    def test_build_complete(self, previous_ma_file, bdo_result, tmp_path):
        """Test complete build process."""
        output_path = tmp_path / "new_ma.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q3 2025",
            period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025"
        )
        
        builder = ManagementAccountsBuilder(
            str(previous_ma_file),
            str(output_path),
            config
        )
        
        result_path = builder.build(bdo_result)
        
        assert result_path.exists()
        assert result_path == output_path
        
        # Verify file structure
        wb = load_workbook(output_path)
        assert "BDO - Q3-25" in wb.sheetnames
        assert "Management Cijfers - Q3 2025" in wb.sheetnames
    
    def test_validate_calculations(self, previous_ma_file, bdo_result, tmp_path):
        """Test equity movement validation."""
        output_path = tmp_path / "new_ma.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q3 2025",
            period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025"
        )
        
        builder = ManagementAccountsBuilder(
            str(previous_ma_file),
            str(output_path),
            config
        )
        
        # This should not raise an exception
        builder._validate_calculations(bdo_result)

