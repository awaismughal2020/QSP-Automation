# Phase 1 — AI-Verified Management Accounts & Compliance

## Overview

Phase 1 takes the client's uploaded financial files, produces a Management
Accounts workbook verified by Claude Opus 4.6, and then builds the Compliance
Certificate from the verified output. Both files are returned for download.

## Endpoint

```
POST /api/v1/generate-accounts
```

## Execution order

```
Upload files
     │
     ▼
Parse BDO / rent roll / sales tracker
     │
     ▼
Build Management Accounts (Python + openpyxl)
  ├─ Copy BDO data into new sheet
  ├─ Insert quarter column with Excel formulas
  ├─ Compute expected numeric values (shadow model)
  └─ Stamp BDO ground truth on LTM rows 19 & 68
     │
     ▼
Save pre-AI audit copy (*_pre_ai.xlsx)
     │
     ▼
AI Verification (Claude Opus 4.6)
  ├─ Extract full workbook context (formulas, BDO maps, rules)
  ├─ Build config-driven prompt (dynamic quarter/columns/rows)
  ├─ Claude returns JSON patches
  ├─ Apply allowlisted patches in-place
  ├─ Reject patches on BDO-anchored LTM cells
  └─ Re-enforce BDO ground truth + re-validate
     │
     ▼
Refresh computed_values from verified file
     │
     ▼
Build Compliance Certificate from post-AI Management Accounts
     │
     ▼
Return both files (+ AI verification status)
```

## Key design rules

1. **Compliance is always built after AI verification** so its values match the
   final Management Accounts file.

2. **BDO-anchored LTM cells (rows 19 and 68)** hold numeric ground truth from
   "Resultaat na belasting" and are never replaced with formulas by AI.
   The AI may only fix upstream detail rows.

3. **All references are dynamic** — quarter, year, column letters, BDO sheet
   name, and the BDO result row are resolved at runtime from the uploaded files
   and `config/accounting_rules.yaml`.

4. **Formulas are preserved.** AI patches fix incorrect formula references but
   never replace a formula with a hard-coded number (except BDO-anchored cells
   which are values by design).

## API response

The response includes:

| Field                        | Description                              |
|------------------------------|------------------------------------------|
| `output_files`               | List of all generated file paths         |
| `management_accounts_path`   | Direct path to the MA workbook           |
| `compliance_certificate_path`| Direct path to the Compliance workbook   |
| `ai_verification`            | Status, patches applied, revalidation    |
| `warnings`                   | Any validation or AI warnings            |

Files can be downloaded via `GET /api/v1/outputs/{filename}`.

## Failure handling

- If Claude errors or returns non-JSON, the pre-AI Management Accounts file is
  used unchanged and a warning is included in the response.
- Rejected patches (e.g. targeting anchored cells) are logged with reasons.
- The `*_pre_ai.xlsx` audit copy is always available for comparison.
