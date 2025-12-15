# QSP ESS B.V. Quarterly Report Automation

Automated quarterly reporting pipeline for QSP ESS B.V. that processes financial data, builds management accounts, calculates compliance metrics, and generates the final quarterly report.

## Overview

This automation replicates the manual quarterly reporting workflow:

1. **BDO Quarterly Financials** → Parse account data with schema-drift resistance
2. **Management Accounts** → Add new BDO sheet + quarterly column with formulas
3. **Compliance Certificate** → Update calculations and formulas for new quarter
4. **Rent Roll** → Extract unit data and calculate totals
5. **Unit Sales Tracker** → Track disposals and proceeds
6. **Word Report** → Update executive summary with calculated values
7. **PDF Assembly** → Combine all documents into final report (requires LibreOffice)

## Quick Start

### Installation

```bash
# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Optional: Install LibreOffice for PDF generation
# macOS: brew install --cask libreoffice
# Linux: sudo apt install libreoffice
```

### Option 1: Command Line Interface

```bash
# Full generation
python -m src.main generate \
    --year 2025 \
    --quarter 3 \
    --bdo-file "inputs/Cijfers_QSP_30-09-2025_d_d__14-10-2025.xlsx" \
    --prev-ma "inputs/Management Accounts Q2 2025 - Draft 1.xlsx" \
    --rent-roll "inputs/QSP_huurlijst_1-10-2025.xlsx" \
    --sales-tracker "inputs/Unit_Sales_tracker_Q3_updated.xlsx" \
    --prev-compliance "inputs/Compliance Certificate Berekening QSP - Q2 2025_updated.xlsx" \
    --word-template "inputs/Quarterly_QSP_-_Q3_2025_-_Draft.docx" \
    --output-dir "outputs"

# Dry run (validate without generating)
python -m src.main generate --dry-run \
    --year 2025 --quarter 3 \
    --bdo-file "inputs/Cijfers_QSP_30-09-2025_d_d__14-10-2025.xlsx" \
    --prev-ma "inputs/Management Accounts Q2 2025 - Draft 1.xlsx" \
    --rent-roll "inputs/QSP_huurlijst_1-10-2025.xlsx" \
    --sales-tracker "inputs/Unit_Sales_tracker_Q3_updated.xlsx" \
    --prev-compliance "inputs/Compliance Certificate Berekening QSP - Q2 2025_updated.xlsx" \
    --word-template "inputs/Quarterly_QSP_-_Q3_2025_-_Draft.docx" \
    --output-dir "outputs"
```

### Option 2: Docker (Recommended for Production)

```bash
# Copy environment file
cp env.example .env

# Build and start all services (API + n8n)
docker-compose up -d --build

# View logs
docker-compose logs -f

# Stop services
docker-compose down
```

Services available:
- **QSP API**: http://localhost:8000
- **n8n Workflow**: http://localhost:5678 (login: admin/qsp2025)
- **API Docs**: http://localhost:8000/docs

### Option 3: REST API (Development)

Start the API server locally:

```bash
# Start server
uvicorn src.api:app --reload --port 8000

# Or run directly
python -m src.api
```

API documentation available at:
- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc

#### Generate Report via API

```bash
# POST request to generate report
curl -X POST "http://localhost:8000/api/v1/generate" \
  -H "Content-Type: application/json" \
  -d '{
    "year": 2025,
    "quarter": 3,
    "bdo_file": "inputs/Cijfers_QSP_30-09-2025_d_d__14-10-2025.xlsx",
    "prev_ma_file": "inputs/Management Accounts Q2 2025 - Draft 1.xlsx",
    "rent_roll_file": "inputs/QSP_huurlijst_1-10-2025.xlsx",
    "sales_tracker_file": "inputs/Unit_Sales_tracker_Q3_updated.xlsx",
    "prev_compliance_file": "inputs/Compliance Certificate Berekening QSP - Q2 2025_updated.xlsx",
    "word_template_file": "inputs/Quarterly_QSP_-_Q3_2025_-_Draft.docx",
    "dry_run": false
  }'
```

#### API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/` | Health check |
| GET | `/health` | Health check |
| POST | `/api/v1/generate` | Generate quarterly report (synchronous) |
| POST | `/api/v1/generate/async` | Generate report asynchronously |
| GET | `/api/v1/jobs/{job_id}` | Check async job status |
| GET | `/api/v1/inputs` | List input files |
| POST | `/api/v1/upload` | Upload input file |
| GET | `/api/v1/outputs` | List generated output files |
| GET | `/api/v1/outputs/{filename}` | Download output file |
| DELETE | `/api/v1/outputs/{filename}` | Delete output file |

#### Example API Response

```json
{
  "status": "success",
  "job_id": "a1b2c3d4",
  "message": "Quarterly report Q3 2025 generated successfully",
  "output_files": [
    "outputs/Management Accounts Q3 2025 - Draft 1.xlsx",
    "outputs/Compliance Certificate Berekening QSP - Q3 2025_updated.xlsx",
    "outputs/Quarterly QSP - Q3 2025 - Draft.docx",
    "outputs/Quarterly QSP - Q3 2025 - ASX.pdf"
  ],
  "warnings": [],
  "execution_time_seconds": 15.3
}
```

## Input Files

Place these files in the `inputs/` folder:

| File | Description | Source |
|------|-------------|--------|
| `Cijfers_QSP_*.xlsx` | BDO quarterly financials | BDO |
| `Management Accounts Q{N} {YYYY} - Draft 1.xlsx` | Previous quarter Management Accounts | Internal |
| `QSP_huurlijst_*.xlsx` | Rent roll | Hoen (Property Manager) |
| `Unit_Sales_tracker_*.xlsx` | Unit sales tracker | Hoen (Property Manager) |
| `Compliance Certificate Berekening QSP - Q{N} {YYYY}_updated.xlsx` | Previous compliance calculator | Internal |
| `Quarterly_QSP_-_Q{N}_{YYYY}_-_Draft.docx` | Word report template | Previous quarter |

## Output Files

Generated files are saved to the `outputs/` folder:

| File | Description |
|------|-------------|
| `Management Accounts Q{N} {YYYY} - Draft 1.xlsx` | Updated Management Accounts with BDO sheet and formulas |
| `Compliance Certificate Berekening QSP - Q{N} {YYYY}_updated.xlsx` | Updated Compliance Certificate |
| `Quarterly QSP - Q{N} {YYYY} - Draft.docx` | Updated Word report |
| `Quarterly QSP - Q{N} {YYYY} - ASX.pdf` | Final combined PDF (requires LibreOffice) |

## Workflow Steps

### Step 1: Parse BDO Quarterly Financials
- Reads the BDO Excel file with trial balance data
- **Schema-drift resistant**: Dynamically detects column positions
- Extracts 93+ account codes with opening/closing balances
- Organizes accounts by Dutch Chart of Accounts categories

### Step 2: Parse Rent Roll
- Extracts unit details from property manager data
- Calculates total units, annual rent, and vacancy rate

### Step 3: Parse Unit Sales Tracker
- Tracks property disposals
- Calculates quarterly and trailing 12-month sales metrics

### Step 4: Build Management Accounts
- Copies previous quarter file structure
- Adds new `BDO - Q{N}-{YY}` sheet with 8-column structure:
  - Column A: Account Code
  - Column B: Account Name
  - Column C: Opening Balance
  - Column G: Current Quarter Mutations
  - Column H: Closing Balance
- Updates `Management Cijfers` summary sheet:
  - Converts LTM column to new quarter data column
  - Adds new LTM column with closing balance formulas
  - Creates formulas like `=-'BDO - Q3-25'!G88` referencing BDO sheet

### Step 4b: Build Compliance Certificate
- Updates Management Accounts sheet with Q3 data
- Adds next forecast quarter (e.g., 26Q3) to `Suppl. Calc`
- Shifts NTM column and updates formula references
- Updates `Impact Unit Sales` calculations

### Step 5: Calculate Compliance Metrics
- Historic Debt Service Cover Ratio (HDSCR) >= 120%
- Projected Debt Service Cover Ratio (PDSCR) >= 120%
- Historic Net Debt Yield (HNDY) >= 6.20%
- Projected Net Debt Yield (PNDY) >= 6.20%
- Loan-to-Value Ratio (LTV) <= 62%

### Step 6: Validate Data
- Balance sheet equation: Assets = Liabilities + Equity
- Cross-file reconciliation checks

### Step 7: Update Word Template
- Replaces quarter references (Q2 → Q3)
- Updates placeholders with calculated values

### Step 8: Assemble PDF (Optional)
- Converts Word and Excel to PDF using LibreOffice
- Merges all documents in order

## Configuration

### Line Item Mappings

The mapping between Management Accounts rows and BDO account codes is in `config/line_item_mappings.yaml`:

```yaml
balance_sheet:
  "Real estate":
    accounts: ["1600000", "1600200", "1601000", "1610803", "1611003"]
    calc_type: "sum"
    
  "Cash":
    accounts: ["2400*"]  # Wildcard matches all 2400xxx accounts
    calc_type: "sum"

profit_loss:
  "Gross Theoretical rental income":
    accounts: ["8000003"]
    calc_type: "direct"
```

### Dutch Chart of Accounts Reference

| Prefix | Category |
|--------|----------|
| 10xxxxx, 11xxxxx | Equity (Eigen Vermogen) |
| 16xxxxx | Real Estate (Vastgoed) |
| 17xxxxx | Financial Fixed Assets |
| 19xxxxx | Long-term Liabilities |
| 20-22xxxxx | Current Assets (Receivables) |
| 24xxxxx | Bank Accounts |
| 23, 25-29xxxxx | Current Liabilities |
| 40-47xxxxx | Expenses |
| 80-84xxxxx | Income |
| 95xxxxx | Tax/Result |

### Covenant Thresholds

Defined in `config/covenant_thresholds.yaml`:

```yaml
covenants:
  historic_debt_service_cover:
    threshold: 1.20
    operator: ">="
  historic_net_debt_yield:
    threshold: 0.062
    operator: ">="
  loan_to_value:
    threshold: 0.62
    operator: "<="
```

## Project Structure

```
QSP-Automation/
├── config/
│   ├── config.yaml              # Main configuration
│   ├── account_mappings.yaml    # BDO account mappings
│   ├── line_item_mappings.yaml  # Management Accounts mappings
│   ├── covenant_thresholds.yaml # Compliance thresholds
│   └── validation_rules.yaml    # Validation rules
├── src/
│   ├── parsers/
│   │   ├── bdo_parser.py        # BDO Excel parser
│   │   ├── rent_roll_parser.py  # Rent roll parser
│   │   └── sales_tracker_parser.py
│   ├── transformers/
│   │   ├── management_accounts.py   # Management Accounts builder
│   │   ├── compliance_builder.py    # Compliance Certificate builder
│   │   └── compliance_calc.py       # Covenant calculations
│   ├── generators/
│   │   ├── word_updater.py      # Word template updater
│   │   └── pdf_assembler.py     # PDF assembly
│   ├── validators/
│   │   ├── balance_validator.py
│   │   ├── covenant_validator.py
│   │   └── reconciliation.py
│   ├── api.py                   # FastAPI server
│   ├── orchestrator.py          # Workflow coordinator
│   └── main.py                  # CLI entry point
├── inputs/                      # Input files
├── outputs/                     # Generated files
├── tests/                       # Test suite
├── logs/                        # Log files
└── requirements.txt             # Dependencies
```

## Troubleshooting

### LibreOffice Not Found
PDF generation requires LibreOffice:
```bash
# macOS
brew install --cask libreoffice

# Linux
sudo apt install libreoffice

# Windows
# Download from https://www.libreoffice.org/download/
```

### Schema Drift in BDO Files
The BDO parser dynamically detects columns. If issues occur:
1. Check logs for column mapping info
2. Verify headers match expected patterns
3. Adjust `config/account_mappings.yaml` if needed

### Missing Account Codes
If accounts show as "unknown":
1. Check `config/account_mappings.yaml`
2. Add new account codes under appropriate category
3. Update `config/line_item_mappings.yaml` for Management Accounts

### Formula Errors in Excel
If formulas show `#REF!` errors:
1. Verify BDO sheet exists with correct name (e.g., `BDO - Q3-25`)
2. Check that account rows match formula references
3. The formulas reference specific row numbers in the BDO sheet

## Development

### Running Tests
```bash
# Run all tests
pytest tests/ -v

# Run specific test file
pytest tests/test_bdo_parser.py -v

# Run with coverage
pytest tests/ --cov=src --cov-report=html
```

### Running the API Server
```bash
# Development mode with auto-reload
uvicorn src.api:app --reload --port 8000

# Production mode
uvicorn src.api:app --host 0.0.0.0 --port 8000 --workers 4
```

### Adding New Account Mappings
1. Identify the BDO account code (7 digits)
2. Add to `config/account_mappings.yaml` under appropriate category
3. If it maps to a Management Accounts row, add to `config/line_item_mappings.yaml`

### Debugging
Enable verbose logging:
```bash
python -m src.main generate -v --log-file logs/debug.log ...
```

## API Python Client Example

```python
import requests

# Generate report
response = requests.post(
    "http://localhost:8000/api/v1/generate",
    json={
        "year": 2025,
        "quarter": 3,
        "bdo_file": "inputs/Cijfers_QSP_30-09-2025_d_d__14-10-2025.xlsx",
        "prev_ma_file": "inputs/Management Accounts Q2 2025 - Draft 1.xlsx",
        "rent_roll_file": "inputs/QSP_huurlijst_1-10-2025.xlsx",
        "sales_tracker_file": "inputs/Unit_Sales_tracker_Q3_updated.xlsx",
        "prev_compliance_file": "inputs/Compliance Certificate Berekening QSP - Q2 2025_updated.xlsx",
        "word_template_file": "inputs/Quarterly_QSP_-_Q3_2025_-_Draft.docx"
    }
)

result = response.json()
print(f"Status: {result['status']}")
print(f"Output files: {result['output_files']}")

# Download generated file
if result['status'] == 'success':
    for filename in result['output_files']:
        file_response = requests.get(
            f"http://localhost:8000/api/v1/outputs/{filename.split('/')[-1]}"
        )
        with open(filename.split('/')[-1], 'wb') as f:
            f.write(file_response.content)
```

## Docker Deployment

### Quick Start with Docker

```bash
# 1. Clone and enter directory
cd QSP-Automation

# 2. Copy environment file
cp env.example .env

# 3. Place input files in inputs/ directory

# 4. Build and start
docker-compose up -d --build

# 5. Check services are running
docker-compose ps
```

### Services

| Service | URL | Description |
|---------|-----|-------------|
| QSP API | http://localhost:8000 | FastAPI application |
| API Docs | http://localhost:8000/docs | Swagger UI |
| n8n | http://localhost:5678 | Workflow automation |

### n8n Workflow Automation

n8n provides a visual workflow interface for running the automation:

1. Open http://localhost:5678
2. Login: `admin` / `qsp2025`
3. Import workflow from `n8n/workflows/`
4. Execute manually or via webhook

**Available Workflows:**
- `qsp_quarterly_report_workflow.json` - Basic manual trigger
- `qsp_full_workflow_with_upload.json` - Advanced with webhook

### Docker Commands

```bash
# View logs
docker-compose logs -f qsp-automation-api
docker-compose logs -f n8n

# Restart services
docker-compose restart

# Stop all
docker-compose down

# Rebuild after code changes
docker-compose up -d --build

# Clean up volumes
docker-compose down -v
```

### File Paths in Docker

When running in Docker, use these paths:
- Input files: `inputs/filename.xlsx`
- Output files: `outputs/filename.xlsx`
- Config files: `config/filename.yaml`

## License

Proprietary - QSP ESS B.V.
