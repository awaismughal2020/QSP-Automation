# QSP Quarterly Reporting - Test Suite

## Test Structure

### Test Files

- **`conftest.py`** - Shared pytest fixtures for all tests
- **`test_bdo_parser.py`** - Tests for BDO parser schema drift handling
- **`test_management_accounts.py`** - Tests for Management Accounts builder
- **`test_compliance_calc.py`** - Tests for compliance calculator
- **`test_orchestrator.py`** - Integration tests for orchestrator

### Fixtures

Test fixtures are created dynamically using pytest's `tmp_path` fixture. The fixtures include:

- **`sample_bdo_file_8col`** - BDO file with 8 columns (old format, Q1 2019)
- **`sample_bdo_file_9col`** - BDO file with 9 columns (new format, Q3 2025)
- **`sample_rent_roll_file`** - Sample rent roll with 4 units
- **`sample_sales_tracker_file`** - Sample sales tracker with 2 sales
- **`sample_bdo_result`** - Pre-parsed BDOParseResult for testing

## Running Tests

```bash
# Run all tests
pytest

# Run specific test file
pytest tests/test_bdo_parser.py

# Run with verbose output
pytest -v

# Run with coverage
pytest --cov=src --cov-report=html
```

## Test Coverage

### BDO Parser Tests
- Schema drift handling (8 vs 9 columns)
- Account code matching with wildcards
- Total calculations
- Missing critical accounts validation
- Period extraction from filename
- Number format parsing

### Management Accounts Tests
- Adding new BDO sheet
- Updating summary sheet with new column
- Complete build process
- Equity movement validation

### Compliance Calculator Tests
- DSCR, NDY, LTV calculations
- Threshold pass/fail logic
- Warning margin detection
- Projected vs historic calculations
- Result formatting

### Orchestrator Tests
- Initialization
- Dry run mode
- Parse steps
- Results structure
- Error handling
- Config properties

## Fixture Files

The `tests/fixtures/` directory contains:
- `create_fixtures.py` - Script to generate fixture Excel files (optional, fixtures are created dynamically)

Fixtures are created on-the-fly during test execution, so no pre-created files are required.

