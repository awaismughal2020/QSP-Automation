"""
Tests for Management Accounts Builder
"""

import pytest
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook, load_workbook

from src.transformers.management_accounts import (
    ManagementAccountsBuilder,
    ManagementAccountsConfig,
    _shift_formula_column_letter,
    find_column_by_header_label,
    find_ltm_column_near_quarter,
    ltm_column_after_insert,
    resolve_quarter_column,
    scan_summary_column_layout,
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
        
        builder._update_summary_sheet()
        
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

    def test_computed_values_populated(self, previous_ma_file, bdo_result, tmp_path):
        """After build(), computed_values dict must be populated for all key rows."""
        output_path = tmp_path / "new_ma.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q3 2025",
            period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025"
        )
        builder = ManagementAccountsBuilder(
            str(previous_ma_file), str(output_path), config
        )
        builder.build(bdo_result)

        cv = builder.computed_values
        assert cv, "computed_values should not be empty"

        assert (19, 'ltm') in cv, "BS row 19 LTM must be in computed_values"
        assert (68, 'quarter') in cv, "P&L row 68 quarter must be in computed_values"
        assert (68, 'ltm') in cv, "P&L row 68 LTM must be in computed_values"

    def test_computed_values_has_key_rows(self, previous_ma_file, bdo_result, tmp_path):
        """computed_values must contain BS and P&L totals after build()."""
        output_path = tmp_path / "new_ma.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q3 2025",
            period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025"
        )
        builder = ManagementAccountsBuilder(
            str(previous_ma_file), str(output_path), config
        )
        builder.build(bdo_result)

        cv = builder.computed_values
        # These keys must always be present even with minimal test data
        for key in [(19, 'ltm'), (68, 'quarter'), (68, 'ltm'), (25, 'quarter'), (66, 'ltm')]:
            assert key in cv, f"computed_values missing key {key}"


def _make_bdo_result(accounts=None):
    """Helper to create a minimal BDOParseResult for tests."""
    return BDOParseResult(
        accounts=accounts or {},
        period_end="2025-09-30",
        quarter="Q3 2025",
        column_mapping={},
        schema_version="v1",
    )


class TestDetectFirstQuarterColumn:
    def test_finds_leftmost_q_header(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=22, column=5, value="Q1 2025")
        ws.cell(row=22, column=10, value="Q2 2025")
        ws.cell(row=22, column=28, value="LTM Q3 2025")
        shell = ManagementAccountsBuilder.__new__(ManagementAccountsBuilder)
        shell._rules = ManagementAccountsBuilder._load_accounting_rules()
        col = shell._detect_first_quarter_column_from_header_row(ws, 22, 27)
        assert col == 5


class TestSummaryColumnLayout:
    """Column insert must follow Q headers, not FY or misplaced LTM anchors."""

    def test_standard_layout_inserts_before_ltm(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=22, column=27, value="Q3 2025")
        ws.cell(row=22, column=28, value="Q4 2025")
        ws.cell(row=22, column=29, value="LTM Q4 2025")
        layout = scan_summary_column_layout(ws, 22)
        assert layout.last_quarter_column == 28
        assert layout.insert_column == 29
        assert layout.previous_quarter_column == 28
        assert layout.ltm_column == 29
        assert ltm_column_after_insert(layout.ltm_column, 29, layout.fy_columns) == 30

    def test_fy_between_quarter_and_ltm_inserts_after_quarter_not_ltm(self):
        """Client template: Q4 | FY | LTM — new Q column belongs after Q4, before FY."""
        wb = Workbook()
        ws = wb.active
        ws.cell(row=22, column=28, value="Q3 2025")
        ws.cell(row=22, column=29, value="Q4 2025")
        ws.cell(row=22, column=30, value="FY 2025")
        ws.cell(row=22, column=31, value="LTM Q4 2025")
        layout = scan_summary_column_layout(ws, 22)
        assert layout.last_quarter_column == 29
        assert layout.insert_column == 30
        assert layout.previous_quarter_column == 29
        assert layout.ltm_column == 31
        assert ltm_column_after_insert(31, 30, layout.fy_columns) == 32

    def test_find_column_by_header_label(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=22, column=30, value="Q1 2026")
        assert find_column_by_header_label(ws, "Q1 2026", 22) == 30
        assert find_column_by_header_label(ws, "q1 2026", 22) == 30

    def test_update_summary_inserts_q1_before_fy(self, tmp_path):
        """Regression: Q1 2026 must appear before FY, not after hidden LTM."""
        from openpyxl.utils import get_column_letter

        prev_path = tmp_path / "prev_q4_fy.xlsx"
        wb = Workbook()
        bdo = wb.active
        bdo.title = "BDO - Q4-25"
        bdo.cell(row=6, column=1, value="8000003")
        bdo.cell(row=6, column=7, value=100)

        mc = wb.create_sheet("Management Cijfers - Q4 2025")
        mc.cell(row=1, column=1, value="Management Accounts QSP ESS B.V. - Q4 2025")
        mc.cell(row=22, column=28, value="Q3 2025")
        mc.cell(row=22, column=29, value="Q4 2025")
        mc.cell(row=22, column=30, value="FY 2025")
        mc.cell(row=22, column=31, value="LTM Q4 2025")
        mc.cell(row=2, column=31, value=datetime(2025, 12, 31))
        mc.cell(row=23, column=29, value=1000)
        mc.cell(row=23, column=31, value="='BDO - Q4-25'!H6")
        wb.save(prev_path)

        out_path = tmp_path / "ma_q1_2026.xlsx"
        config = ManagementAccountsConfig(
            quarter="Q1 2026",
            period_end=datetime(2026, 3, 31),
            bdo_sheet_name="BDO - Q1-26",
            summary_sheet_name="Management Cijfers - Q1 2026",
        )
        builder = ManagementAccountsBuilder(str(prev_path), str(out_path), config)
        builder.workbook = load_workbook(prev_path)
        builder.config = config
        builder._copy_bdo_data_to_new_sheet(_make_bdo_result())
        new_bdo = builder.workbook[config.bdo_sheet_name]
        builder._build_row_map(new_bdo, builder._new_bdo_row_map, builder._new_bdo_label_map)
        builder._update_summary_sheet()
        builder.workbook.save(out_path)

        result = load_workbook(out_path)
        sheet = result["Management Cijfers - Q1 2026"]
        assert sheet.cell(row=22, column=30).value == "Q1 2026"
        assert sheet.cell(row=22, column=31).value == "FY 2025"
        assert "LTM" in str(sheet.cell(row=22, column=32).value)
        assert get_column_letter(builder._new_quarter_col) == "AD"
        result.close()

    def test_standard_layout_matches_legacy_ltm_insert(self):
        """Without FY, insert after last Q is the same as insert before LTM (Q3/Q4 2025 path)."""
        wb = Workbook()
        ws = wb.active
        ws.cell(row=22, column=27, value="Q3 2025")
        ws.cell(row=22, column=28, value="Q4 2025")
        ws.cell(row=22, column=29, value="LTM Q4 2025")
        layout = scan_summary_column_layout(ws, 22)
        assert layout.insert_column == layout.ltm_column
        assert layout.previous_quarter_column == 28

    @pytest.mark.parametrize(
        "headers,expected_insert,expected_prev,expected_ltm_after",
        [
            (
                [(28, "Q3 2025"), (29, "Q4 2025"), (30, "LTM Q4 2025")],
                30,
                29,
                31,
            ),
            (
                [(28, "Q3 2025"), (29, "Q4 2025"), (30, "FY 2025"), (31, "LTM Q4 2025")],
                30,
                29,
                32,
            ),
            (
                [(29, "Q4 2025"), (30, "Q1 2026"), (31, "FY 2025"), (32, "LTM Q1 2026")],
                31,
                30,
                33,
            ),
            (
                [(30, "Q1 2026"), (31, "Q2 2026"), (32, "FY 2026"), (33, "LTM Q2 2026")],
                32,
                31,
                34,
            ),
        ],
    )
    def test_future_quarter_insert_positions(
        self, headers, expected_insert, expected_prev, expected_ltm_after
    ):
        wb = Workbook()
        ws = wb.active
        for col, label in headers:
            ws.cell(row=22, column=col, value=label)
        layout = scan_summary_column_layout(ws, 22)
        assert layout.insert_column == expected_insert
        assert layout.previous_quarter_column == expected_prev
        assert ltm_column_after_insert(
            layout.ltm_column, layout.insert_column, layout.fy_columns
        ) == expected_ltm_after

    def test_resolve_quarter_column_skips_fy(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=22, column=29, value="Q4 2025")
        ws.cell(row=22, column=30, value="Q1 2026")
        ws.cell(row=22, column=31, value="FY 2025")
        ws.cell(row=22, column=32, value="LTM Q1 2026")
        assert resolve_quarter_column(ws, "Q1 2026", 22) == 30
        assert find_ltm_column_near_quarter(ws, 30, 22) == 32


class TestShiftFormulaColumnLetter:
    """Regression tests for interest row 60 formula extension."""

    def test_shifts_prev_column_refs_ab_to_ac(self):
        f = "=AB57+AB60-$AB$61+SUM(Y60:AB60)"
        out = _shift_formula_column_letter(f, "AB", "AC")
        assert "AB57" not in out and "AC57" in out
        assert "AC60" in out
        assert "$AC$61" in out
        assert "SUM(Y60:AC60)" in out

    def test_does_not_replace_a_inside_aa(self):
        f = "=AA10+A11"
        out = _shift_formula_column_letter(f, "A", "B")
        assert out == "=AA10+B11"


class TestFuzzyAccountLookup:
    """Test that shadow model uses fuzzy matching like _find_account_row."""

    def test_exact_match(self):
        entry = AccountEntry(
            code="1600000", name="Test", opening_balance=0, closing_balance=100,
            mutations={}, raw_row=5,
        )
        result = ManagementAccountsBuilder._fuzzy_account_lookup(
            "1600000", _make_bdo_result({"1600000": entry}),
        )
        assert result is entry

    def test_prefix_match(self):
        entry = AccountEntry(
            code="1600000.0", name="Test", opening_balance=0, closing_balance=100,
            mutations={}, raw_row=5,
        )
        result = ManagementAccountsBuilder._fuzzy_account_lookup(
            "1600000", _make_bdo_result({"1600000.0": entry}),
        )
        assert result is entry

    def test_no_match_returns_none(self):
        result = ManagementAccountsBuilder._fuzzy_account_lookup(
            "9999999", _make_bdo_result(),
        )
        assert result is None


class TestEvaluateCalcPattern:
    """Ensure _evaluate_calc_pattern handles every YAML pattern."""

    @pytest.fixture
    def builder(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "BDO - Q3-25"
        wb.create_sheet("Management Cijfers - Q3 2025")
        path = tmp_path / "dummy.xlsx"
        wb.save(path)
        config = ManagementAccountsConfig(
            quarter="Q3 2025", period_end=datetime(2025, 9, 30),
            bdo_sheet_name="BDO - Q3-25",
            summary_sheet_name="Management Cijfers - Q3 2025",
        )
        b = ManagementAccountsBuilder.__new__(ManagementAccountsBuilder)
        b.config = config
        b.formula_templates_path = Path("config/formula_templates.yaml")
        b.formula_templates = {}
        b.balance_sheet_templates = {}
        return b

    def test_sum_range(self, builder):
        shadow = {3: {'value': 10.0}, 4: {'value': 20.0}, 5: {'value': 30.0}}
        assert builder._evaluate_calc_pattern("=SUM({COL}3:{COL}5)", shadow) == 60.0

    def test_negated_subtraction(self, builder):
        shadow = {50: {'value': 100.0}, 53: {'value': 40.0}}
        assert builder._evaluate_calc_pattern("=-({COL}50-{COL}53)", shadow) == -60.0

    def test_division_minus_one(self, builder):
        shadow = {25: {'value': 200.0}, 23: {'value': 100.0}}
        assert builder._evaluate_calc_pattern("={COL}25/{COL}23-1", shadow) == pytest.approx(1.0)

    def test_addition_chain(self, builder):
        shadow = {25: {'value': 10.0}, 30: {'value': 20.0}, 44: {'value': 30.0}}
        assert builder._evaluate_calc_pattern("={COL}25+{COL}30+{COL}44", shadow) == 60.0

    def test_two_term_addition(self, builder):
        shadow = {67: {'value': 5.0}, 66: {'value': 3.0}}
        assert builder._evaluate_calc_pattern("={COL}67+{COL}66", shadow) == 8.0

    def test_two_term_with_plus(self, builder):
        shadow = {53: {'value': 10.0}, 48: {'value': 20.0}}
        assert builder._evaluate_calc_pattern("={COL}53+{COL}48", shadow) == 30.0


class TestCCDirectCopy:
    """Verify the CC builder uses computed_values for direct copy."""

    def test_copy_ma_data_from_cache(self, tmp_path):
        from src.transformers.compliance_builder import ComplianceBuilder, ComplianceConfig

        # Create a minimal previous CC file with a "Q3 Management Accounts" sheet
        prev_path = tmp_path / "prev_cc.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Q2 Management Accounts"
        for r in range(1, 80):
            ws.cell(row=r, column=1, value=f"Row {r}")
        wb.save(prev_path)

        output_path = tmp_path / "new_cc.xlsx"
        config = ComplianceConfig(year=2025, quarter=3)
        builder = ComplianceBuilder(str(prev_path), str(output_path), config)

        computed_values = {
            (3, 'ltm'): 100.0,
            (3, 'quarter'): 50.0,
            (19, 'ltm'): 555.0,
            (23, 'ltm'): 200.0,
            (23, 'quarter'): 150.0,
            (68, 'ltm'): 999.0,
            (68, 'quarter'): 888.0,
            (107, 'ltm'): 303.0,
            (114, 'ltm'): 303.0,
        }

        bdo_result = _make_bdo_result()
        builder.build(bdo_result, management_accounts_path=None, computed_values=computed_values)

        result_wb = load_workbook(str(output_path))
        ma_sheet = result_wb["Q3 Management Accounts"]

        # CC row 2 -> MA row 3 (BS): LTM in col C only, no quarter for BS
        assert ma_sheet.cell(row=2, column=3).value == 100.0
        assert ma_sheet.cell(row=2, column=2).value is None

        # CC row 18 -> MA row 19 (Total Equity Movement): LTM only
        assert ma_sheet.cell(row=18, column=3).value == 555.0

        # CC row 22 -> MA row 23 (P&L): both quarter and LTM
        assert ma_sheet.cell(row=22, column=3).value == 200.0
        assert ma_sheet.cell(row=22, column=2).value == 150.0

        result_wb.close()

