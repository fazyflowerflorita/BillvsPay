from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"


payroll = pd.DataFrame(
    [
        {"Legal Entity": "Pride US", "Employee ID": "1001", "Personal Number": "P1001", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "REG", "Working Hours": 40, "Payroll Amount": 1200},
        {"Legal Entity": "Pride US", "Employee ID": "1002", "Personal Number": "P1002", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "REG", "Working Hours": 40, "Payroll Amount": 1100},
        {"Legal Entity": "Pride US", "Employee ID": "1003", "Personal Number": "P1003", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "OT", "Working Hours": 5, "Payroll Amount": 225},
        {"Legal Entity": "Pride US", "Employee ID": "1004", "Personal Number": "P1004", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "REG", "Working Hours": 36, "Payroll Amount": 950},
        {"Legal Entity": "Pride US", "Employee ID": "1005", "Personal Number": "P1005", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "ACA", "Working Hours": 0, "Payroll Amount": 50},
    ]
)

billing = pd.DataFrame(
    [
        {"Legal Entity": "Pride US", "Employee ID": "1001", "Personal Number": "P1001", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "REG", "Working Hours": 40, "Billing Amount": 1200, "Customer": "Acme", "Invoice Number": "INV-1"},
        {"Legal Entity": "Pride US", "Employee ID": "1002", "Personal Number": "P1002", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "REG", "Working Hours": 38, "Billing Amount": 1100, "Customer": "Acme", "Invoice Number": "INV-2"},
        {"Legal Entity": "Pride US", "Employee ID": "1003", "Personal Number": "P1003", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "OT", "Working Hours": 5, "Billing Amount": 250, "Customer": "Acme", "Invoice Number": "INV-3"},
        {"Legal Entity": "Pride US", "Employee ID": "1006", "Personal Number": "P1006", "Employee Type": "W2", "Pay Period": "2026-07-05", "Earnings Code": "REG", "Working Hours": 40, "Billing Amount": 1000, "Customer": "Acme", "Invoice Number": "INV-4"},
    ]
)


if __name__ == "__main__":
    SAMPLES.mkdir(parents=True, exist_ok=True)
    payroll.to_excel(SAMPLES / "sample_payroll.xlsx", index=False)
    billing.to_excel(SAMPLES / "sample_billing.xlsx", index=False)
    print(f"Created {SAMPLES / 'sample_payroll.xlsx'}")
    print(f"Created {SAMPLES / 'sample_billing.xlsx'}")

