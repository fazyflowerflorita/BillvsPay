# Bill vs Pay Reconciliation

Python web application for reconciling Payroll and Billing Excel exports after they are downloaded from the source systems.

## What It Does

- Upload Payroll and Billing reports.
- Clean blank rows, duplicates, spacing, dates, numeric fields, and employee IDs.
- Apply configurable business rules from `config/business_rules.xlsx`.
- Match transactions by employee, earning/category code, week start, and week end.
- Flag missing records, hours mismatches, amount mismatches, and duplicate invoices.
- Generate a PTL-style Excel workbook with Payroll Report and Billing Report tabs plus summary, audit, match detail, and rules tabs.

## Quick Start

Use the bundled Codex Python runtime because it includes Pandas and OpenPyXL:

```powershell
& 'C:\Users\FazyFlowerFlorita\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' app\server.py
```

Then open:

```text
http://127.0.0.1:8765
```

## Website Deployment

This app can be deployed to a real website URL. See `DEPLOYMENT.md`.

For hosted environments, set:

```text
HOST=0.0.0.0
PORT=<provided by hosting platform>
```

## Input Files

The app expects `.xlsx` reports. It is configured for the PTL Payroll and Billing export structure in the provided samples.

Recommended columns:

- Payroll: `Personnel number`, `Worker`, `Customer name`, `Pay statement`, `Payment date`, `Earning code`, `Earnings date`, `Quantity`, `Rate`, `Amount`, `Pay cycle`, `Pay period.Period start date`, `Pay period.Period end date`, `Position`, `Posted`, `Payee type name`
- Billing: `Personalnumber` or `PayrollEmpID`, `SOP Number`, `Invoice Date`, `Item Number`, `Category Id`, `Item Description`, `Placement ID`, `Week Start`, `Week End`, `QTY`, `Extended Cost`, `Extended Price`, `Unit Cost Price`, `Unit Sales Price`, `Customer Number`, `Customer Name`, `Batch Number`, `SOP Type`, `Payee Type`

Payroll `Amount` is compared to Billing `Extended Cost`. Payroll `Quantity` is compared to Billing `QTY`.

## Business Rules

Edit `config/business_rules.xlsx` to change reconciliation behavior without changing code:

- `column_mapping`: maps source report column headers to standard application fields.
- `included_earnings_codes`: earnings codes to include. If blank, all non-excluded codes are included.
- `excluded_earnings_codes`: earnings codes to exclude, such as ACA, BGC, and FEE.
- `settings`: tolerance and filter settings.

## Output

Generated reports are saved to `output/` and can also be downloaded from the web app.

Report tabs:

- `Summary`
- `Payroll Report`
- `Billing Report`
- `Exception Summary`
- `Match Details`
- `Audit`
- `Business Rules`

The Billing Report tab includes `HrsDifference`, `AmntDifference`, and `Comments`.
The Payroll Report tab also includes these reconciliation status fields so paid-but-not-billed items are easy to identify.

Supported comments:

- `Tallied`
- `Billed but not Paid`
- `Paid But not Billed`
- `Difference in working Hours`
- `Duplicate Invoice`
- `Over Billed`
- `Short Billed`

Color coding:

- Green: `Tallied`
- Red: `Billed but not Paid`
- Orange: `Paid But not Billed`
- Yellow: `Difference in working Hours`
- Purple: `Duplicate Invoice`
- Light orange: amount mismatch such as `Over Billed` or `Short Billed`

Each report also includes a `Color Legend` tab.

## Sample Data

Run this to generate sample Payroll and Billing files:

```powershell
& 'C:\Users\FazyFlowerFlorita\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' samples\make_sample_data.py
```
