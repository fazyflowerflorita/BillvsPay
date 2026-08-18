# Bill vs Pay Reconciliation - Data Issues & Solutions

## Critical Issues Found

### Issue 1: EXPENSES - Missing Employee IDs in Billing
**Status**: 🔴 CRITICAL

- **Payroll**: 391 expense rows with `Earning code = 'Expenses'` + Personnel number
- **Billing**: 613 expense rows with `Category Id = 'Expense'` + ALL have NULL Personalnumber
- **Root Cause**: Expense transactions are not tied to employees in billing system
- **Solution**: Must match expenses by:
  1. Employee NAME (First/Last) in billing instead of ID
  2. Consolidate multiple billing expenses to single payroll expense by:
     - Employee Name
     - Category = 'Expense'
     - Month (from Week Start/Invoice Date)
  3. Use Extended Price (not Extended Cost) for expenses

### Issue 2: SICK LEAVE - Split Transactions
**Status**: 🟠 MEDIUM

- **Payroll**: 360 sick leave rows; 102 employees with MULTIPLE records
  - Example: Employee E001321 has 2 sick leave records (different pay periods)
- **Billing**: 255 sick leave rows; 62 with multiple records
- **Root Cause**: Sick leave can span multiple pay cycles/billing periods
- **Solution**: Consolidate BOTH sides before matching:
  1. Group payroll by: Personnel number + Category + MONTH
  2. Sum: Amount, Quantity
  3. Group billing by: Personalnumber + Category + MONTH
  4. Sum: Extended Cost, QTY
  5. Match consolidated records
  6. Mark as "Tallied - Consolidated Sick Leave"

### Issue 3: ROUNDED OFF DIFFERENCES
**Status**: 🟠 MEDIUM

- **Current**: Amount matching is exact comparison
- **Requirement**: Differences < $1.00 should be flagged as "Rounded Off Difference"
- **Solution**: Add tolerance check before other mismatch flags
  ```python
  if abs(payroll_amount - billing_amount) < 1.00:
      comment = "Rounded Off Difference"
  ```

### Issue 4: WEEKLY vs MONTHLY BILLING
**Status**: 🟡 LOW (Less common in sample data)

- **Billing**: 444 unique week ranges (weekly, monthly, custom)
- **Payroll**: 111 unique pay periods (weekly, bi-weekly, monthly)
- **Root Cause**: Period boundaries don't always align
- **Solution**: 
  1. Detect period type (weekly = 7 days, monthly = ~30 days)
  2. If payroll is weekly and billing is monthly:
     - Sum payroll amounts for matching month
     - Match against single monthly billing row
  3. Mark as "Tallied - Monthly Billing"

## Data Statistics

| Aspect | Payroll | Billing |
|--------|---------|---------|
| Total Rows | 25,269 | 57,842 |
| Regular (RT) | 19,599 | 39,746 |
| Expenses | 391 | 613 |
| Sick Leave | 360 | 255 |
| Overtime | 2,335 | 5,402 |
| PTO | 189 | 82 |
| Salary | 1,470 | N/A |
| Other | 1,325 | 1,356 |

## Matching Strategy (Priority Order)

1. **Direct Match** (No consolidation needed)
   - Same Employee ID
   - Same Category/Earning code
   - Amounts within tolerance OR < $1.00 difference
   - Quantities match
   - → Mark: "Tallied"

2. **Rounded Difference**
   - Same Employee ID, Category
   - Amount difference < $1.00
   - → Mark: "Rounded Off Difference", Color: Light Green

3. **Sick Leave Consolidation**
   - Group payroll: (Employee ID, Category, Month)
   - Group billing: (Employee ID, Category, Month)
   - Match consolidated records
   - → Mark: "Tallied - Consolidated Sick Leave", Color: Green

4. **Expense Consolidation**
   - Payroll: (Employee ID, 'Expenses', Month)
   - Billing: (Employee NAME, 'Expense', Month, Extended Price)
   - Match by: Employee NAME + Month
   - Sum billing Extended Price
   - → Mark: "Tallied - Consolidated Expense", Color: Green

5. **Monthly Billing Consolidation**
   - Detect period types (weekly vs monthly)
   - If payroll=weekly, billing=monthly:
     - Sum payroll for matching month
     - Match against monthly billing
   - → Mark: "Tallied - Monthly Billing", Color: Green

6. **Unmatched Records**
   - "Paid But not Billed" (in payroll, not in billing)
   - "Billed but not Paid" (in billing, not in payroll)
   - Color: Orange, Red

## Implementation Notes

### Column Mapping
```
Payroll                              → Billing
Personnel number                     → Personalnumber
Worker / (First name + Last name)    → (First Name + Last Name)
Earning code                         → Category Id
Amount                               → Extended Cost OR Extended Price (for expenses)
Quantity                             → QTY
Payment date / Earnings date         → Invoice Date
Pay period.Period start date         → Week Start
Pay period.Period end date           → Week End
```

### Special Handling
- **Expenses**: Use Extended Price, match by Employee NAME
- **Sick Leave**: Consolidate by month before matching
- **Amounts < $1**: Treat as rounding difference
- **Missing payroll employee IDs**: These are expenses; use Employee NAME

### Test Cases Needed
1. ✓ Single matching transaction (exact amounts, quantities)
2. ✓ Rounded off difference (<$1.00)
3. ✓ Expense split across multiple billing transactions
4. ✓ Sick leave split across multiple pay periods
5. ✓ Monthly billing vs weekly payroll
6. ✓ Unmatched payroll (paid but not billed)
7. ✓ Unmatched billing (billed but not paid)
