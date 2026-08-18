"""
Bill vs Pay Reconciliation Engine - Version 2 (Fixed)
Handles all reconciliation scenarios with corrected logic
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Tuple, Dict, List
import logging
from difflib import SequenceMatcher

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class BillPayReconciler:
    """Main reconciliation engine with all scenario handling"""
    
    # Configuration
    AMOUNT_TOLERANCE = 1.00  # Amounts < $1 are considered rounded off
    QTY_TOLERANCE = 0.01     # Quantity tolerance for matching
    
    # Category mappings
    EXPENSE_CODES = ['Expenses', 'RemExp']
    EXPENSE_CATEGORIES = ['Expense']
    SICK_CODES = ['SICK']
    SICK_CATEGORIES = ['Sick']
    
    def __init__(self, payroll_df: pd.DataFrame, billing_df: pd.DataFrame):
        """Initialize reconciler with payroll and billing data."""
        self.original_payroll = payroll_df.copy()
        self.original_billing = billing_df.copy()
        
        self.payroll_df = payroll_df.copy()
        self.billing_df = billing_df.copy()
        
        # Track all matched records
        self.matched_records = []
        self.processed_payroll_indices = set()
        self.processed_billing_indices = set()
        
        # Clean data
        self._clean_data()
    
    def _clean_data(self):
        """Clean and normalize data before reconciliation"""
        try:
            # Normalize column names (strip whitespace)
            self.payroll_df.columns = self.payroll_df.columns.str.strip()
            self.billing_df.columns = self.billing_df.columns.str.strip()
            self.original_payroll.columns = self.original_payroll.columns.str.strip()
            self.original_billing.columns = self.original_billing.columns.str.strip()
            
            # Ensure numeric fields
            self.payroll_df['Amount'] = pd.to_numeric(self.payroll_df['Amount'], errors='coerce').fillna(0)
            self.payroll_df['Quantity'] = pd.to_numeric(self.payroll_df['Quantity'], errors='coerce').fillna(0)
            
            self.billing_df['Extended Cost'] = pd.to_numeric(self.billing_df['Extended Cost'], errors='coerce').fillna(0)
            self.billing_df['Extended Price'] = pd.to_numeric(self.billing_df['Extended Price'], errors='coerce').fillna(0)
            self.billing_df['QTY'] = pd.to_numeric(self.billing_df['QTY'], errors='coerce').fillna(0)
            
            # Convert dates
            self.payroll_df['Payment date'] = pd.to_datetime(self.payroll_df['Payment date'], errors='coerce')
            self.billing_df['Invoice Date'] = pd.to_datetime(self.billing_df['Invoice Date'], errors='coerce')
            self.billing_df['Week Start'] = pd.to_datetime(self.billing_df['Week Start'], errors='coerce')
            self.billing_df['Week End'] = pd.to_datetime(self.billing_df['Week End'], errors='coerce')
            
            logger.info(f"Cleaned {len(self.payroll_df)} payroll rows, {len(self.billing_df)} billing rows")
        except Exception as e:
            logger.error(f"Error cleaning data: {e}")
            raise
    
    def reconcile(self) -> Dict:
        """Execute full reconciliation process"""
        logger.info("Starting reconciliation process...")
        
        try:
            # Step 1: Handle special scenarios (order matters!)
            self._process_sick_leave()
            self._process_expenses()
            
            # Step 2: Process regular items
            self._process_regular_items()
            
            logger.info(f"Reconciliation complete: {len(self.matched_records)} matches")
            
            return self._generate_results()
            
        except Exception as e:
            logger.error(f"Reconciliation error: {e}")
            raise
    
    def _process_expenses(self):
        """Process expense reconciliation with proper name matching"""
        logger.info("Processing expenses...")
        
        # Get expense records
        payroll_exp = self.payroll_df[
            self.payroll_df['Earning code'].isin(self.EXPENSE_CODES)
        ].copy()
        
        billing_exp = self.billing_df[
            self.billing_df['Category Id'].isin(self.EXPENSE_CATEGORIES)
        ].copy()
        
        if len(payroll_exp) == 0 or len(billing_exp) == 0:
            logger.info(f"Expenses - Payroll: {len(payroll_exp)}, Billing: {len(billing_exp)}")
            return
        
        logger.info(f"Processing {len(payroll_exp)} payroll expenses and {len(billing_exp)} billing expenses")
        
        # Extract month for consolidation
        payroll_exp['Month'] = pd.to_datetime(payroll_exp['Payment date']).dt.to_period('M')
        billing_exp['Month'] = pd.to_datetime(billing_exp['Invoice Date']).dt.to_period('M')
        
        # Consolidate billing expenses by month
        # Group by invoice date month to get consolidated amounts
        billing_exp_consolidated = billing_exp.groupby('Month').agg({
            'Extended Price': 'sum',
            'QTY': 'sum',
            'Personalnumber': lambda x: x.mode()[0] if len(x.mode()) > 0 else 'UNKNOWN',
            'Category Id': 'first',
            'Invoice Date': 'first'
        }).reset_index()
        
        # Match by month and amount proximity
        matches_found = 0
        for p_idx, p_row in payroll_exp.iterrows():
            if p_idx in self.processed_payroll_indices:
                continue
                
            month = p_row['Month']
            p_amount = p_row['Amount']
            
            # Find billing expenses in same month
            month_matches = billing_exp_consolidated[
                billing_exp_consolidated['Month'] == month
            ]
            
            if len(month_matches) == 0:
                continue
            
            # Find closest match by amount
            month_matches['amount_diff'] = abs(month_matches['Extended Price'] - p_amount)
            best_idx = month_matches['amount_diff'].idxmin()
            b_row = month_matches.loc[best_idx]
            
            amount_diff = abs(p_amount - b_row['Extended Price'])
            
            # Accept if reasonably close (within 5% or $10)
            if amount_diff < max(abs(p_amount) * 0.05, 10.0):
                comment = "Tallied - Consolidated Expense"
                if amount_diff > 0 and amount_diff < self.AMOUNT_TOLERANCE:
                    comment = "Rounded Off Difference"
                
                self.matched_records.append({
                    'source': 'Expense',
                    'employee_id': p_row['Personnel number'],
                    'employee_name': p_row.get('Worker', 'Unknown'),
                    'category': 'Expense',
                    'payroll_amount': p_amount,
                    'billing_amount': b_row['Extended Price'],
                    'amount_difference': amount_diff,
                    'payroll_qty': p_row['Quantity'],
                    'billing_qty': b_row['QTY'],
                    'comment': comment,
                    'payroll_idx': p_idx,
                    'billing_indices': billing_exp[billing_exp['Month'] == month].index.tolist(),
                    'payroll_date': p_row['Payment date'],
                    'billing_date': b_row['Invoice Date']
                })
                
                self.processed_payroll_indices.add(p_idx)
                for b_idx in billing_exp[billing_exp['Month'] == month].index:
                    self.processed_billing_indices.add(b_idx)
                
                matches_found += 1
        
        logger.info(f"Expense matches found: {matches_found}")
    
    def _process_sick_leave(self):
        """Process sick leave consolidation"""
        logger.info("Processing sick leave...")
        
        payroll_sick = self.payroll_df[
            self.payroll_df['Earning code'].isin(self.SICK_CODES)
        ].copy()
        
        billing_sick = self.billing_df[
            self.billing_df['Category Id'].isin(self.SICK_CATEGORIES)
        ].copy()
        
        if len(payroll_sick) == 0 or len(billing_sick) == 0:
            logger.info(f"Sick Leave - Payroll: {len(payroll_sick)}, Billing: {len(billing_sick)}")
            return
        
        logger.info(f"Processing {len(payroll_sick)} payroll sick leave and {len(billing_sick)} billing sick leave")
        
        # Extract month for consolidation
        payroll_sick['Month'] = pd.to_datetime(payroll_sick['Payment date']).dt.to_period('M')
        billing_sick['Month'] = pd.to_datetime(billing_sick['Invoice Date']).dt.to_period('M')
        
        # Consolidate by employee + month
        payroll_sick_consolidated = payroll_sick.groupby(['Personnel number', 'Month']).agg({
            'Amount': 'sum',
            'Quantity': 'sum',
            'Payment date': 'first',
            'Worker': 'first'
        }).reset_index()
        
        billing_sick_consolidated = billing_sick.groupby(['Personalnumber', 'Month']).agg({
            'Extended Cost': 'sum',
            'QTY': 'sum',
            'Invoice Date': 'first'
        }).reset_index()
        
        matches_found = 0
        for _, p_row in payroll_sick_consolidated.iterrows():
            emp_id = p_row['Personnel number']
            month = p_row['Month']
            p_amount = p_row['Amount']
            p_qty = p_row['Quantity']
            
            # Find matching billing record
            match = billing_sick_consolidated[
                (billing_sick_consolidated['Personalnumber'] == emp_id) &
                (billing_sick_consolidated['Month'] == month)
            ]
            
            if len(match) > 0:
                b_row = match.iloc[0]
                b_amount = b_row['Extended Cost']
                b_qty = b_row['QTY']
                
                amount_diff = abs(p_amount - b_amount)
                qty_diff = abs(p_qty - b_qty)
                
                # Determine comment
                if amount_diff < self.AMOUNT_TOLERANCE and qty_diff < self.QTY_TOLERANCE:
                    comment = "Tallied - Consolidated Sick Leave"
                elif qty_diff > self.QTY_TOLERANCE and amount_diff < self.AMOUNT_TOLERANCE:
                    comment = "Difference in working Hours"
                else:
                    comment = "Amount Mismatch"
                
                self.matched_records.append({
                    'source': 'Sick Leave',
                    'employee_id': emp_id,
                    'employee_name': p_row['Worker'],
                    'category': 'Sick Leave',
                    'payroll_amount': p_amount,
                    'billing_amount': b_amount,
                    'amount_difference': amount_diff,
                    'payroll_qty': p_qty,
                    'billing_qty': b_qty,
                    'qty_difference': qty_diff,
                    'comment': comment,
                    'payroll_date': p_row['Payment date'],
                    'billing_date': b_row['Invoice Date']
                })
                
                # Mark original rows as processed
                payroll_indices = payroll_sick[
                    (payroll_sick['Personnel number'] == emp_id) &
                    (payroll_sick['Month'] == month)
                ].index
                
                billing_indices = billing_sick[
                    (billing_sick['Personalnumber'] == emp_id) &
                    (billing_sick['Month'] == month)
                ].index
                
                for idx in payroll_indices:
                    self.processed_payroll_indices.add(idx)
                for idx in billing_indices:
                    self.processed_billing_indices.add(idx)
                
                matches_found += 1
        
        logger.info(f"Sick leave matches found: {matches_found}")
    
    def _process_regular_items(self):
        """Process regular items with direct matching"""
        logger.info("Processing regular items...")
        
        matches_found = 0
        
        for p_idx, p_row in self.payroll_df.iterrows():
            if p_idx in self.processed_payroll_indices:
                continue
            
            emp_id = p_row['Personnel number']
            category = p_row['Earning code']
            p_amount = p_row['Amount']
            p_qty = p_row['Quantity']
            
            # Skip if zero amount
            if p_amount == 0:
                continue
            
            # Find candidate matches
            candidates = self.billing_df[
                (self.billing_df['Personalnumber'] == emp_id) &
                (self.billing_df['Category Id'] == category) &
                ~self.billing_df.index.isin(self.processed_billing_indices)
            ]
            
            if len(candidates) == 0:
                continue
            
            # Find best match (closest amount)
            candidates_copy = candidates.copy()
            candidates_copy['amount_diff'] = abs(candidates_copy['Extended Cost'] - p_amount)
            best_idx = candidates_copy['amount_diff'].idxmin()
            best_match = self.billing_df.loc[best_idx]
            
            b_amount = best_match['Extended Cost']
            b_qty = best_match['QTY']
            amount_diff = abs(p_amount - b_amount)
            qty_diff = abs(p_qty - b_qty)
            
            # Determine if match
            comment = ""
            is_match = False
            
            if amount_diff < self.AMOUNT_TOLERANCE and qty_diff < self.QTY_TOLERANCE:
                comment = "Tallied - Rounded Off Difference" if amount_diff > 0 else "Tallied"
                is_match = True
            elif amount_diff < self.AMOUNT_TOLERANCE and qty_diff < 0.5:
                comment = "Difference in working Hours"
                is_match = True
            
            if is_match:
                self.matched_records.append({
                    'source': 'Regular',
                    'employee_id': emp_id,
                    'employee_name': p_row.get('Worker', 'Unknown'),
                    'category': category,
                    'payroll_amount': p_amount,
                    'billing_amount': b_amount,
                    'amount_difference': amount_diff,
                    'payroll_qty': p_qty,
                    'billing_qty': b_qty,
                    'qty_difference': qty_diff,
                    'comment': comment,
                    'payroll_date': p_row['Payment date'],
                    'billing_date': best_match['Invoice Date']
                })
                
                self.processed_payroll_indices.add(p_idx)
                self.processed_billing_indices.add(best_idx)
                matches_found += 1
        
        logger.info(f"Regular matches found: {matches_found}")
    
    def _generate_results(self) -> Dict:
        """Generate comprehensive reconciliation results"""
        
        # Matched records
        matches_df = pd.DataFrame(self.matched_records) if self.matched_records else pd.DataFrame()
        
        # Unmatched payroll
        unmatched_p = self.original_payroll[
            ~self.original_payroll.index.isin(self.processed_payroll_indices)
        ].copy()
        unmatched_p['Status'] = 'Paid But not Billed'
        
        # Unmatched billing
        unmatched_b = self.original_billing[
            ~self.original_billing.index.isin(self.processed_billing_indices)
        ].copy()
        unmatched_b['Status'] = 'Billed but not Paid'
        
        # Calculate summary
        total_payroll = len(self.original_payroll)
        total_billing = len(self.original_billing)
        total_matched = len(self.matched_records)
        
        summary = {
            'total_payroll_rows': total_payroll,
            'total_billing_rows': total_billing,
            'total_matches': total_matched,
            'unmatched_payroll_rows': len(unmatched_p),
            'unmatched_billing_rows': len(unmatched_b),
            'match_rate_payroll': (total_matched / total_payroll * 100) if total_payroll > 0 else 0,
            'payroll_amount_total': self.original_payroll['Amount'].sum(),
            'billing_amount_total': self.original_billing['Extended Cost'].sum(),
            'matched_payroll_amount': matches_df['payroll_amount'].sum() if len(matches_df) > 0 else 0,
            'matched_billing_amount': matches_df['billing_amount'].sum() if len(matches_df) > 0 else 0,
            'unmatched_payroll_amount': unmatched_p['Amount'].sum() if len(unmatched_p) > 0 else 0,
            'unmatched_billing_amount': unmatched_b['Extended Cost'].sum() if len(unmatched_b) > 0 else 0,
        }
        
        return {
            'matches': matches_df,
            'unmatched_payroll': unmatched_p,
            'unmatched_billing': unmatched_b,
            'summary': summary,
            'match_details': self._get_match_details(matches_df)
        }
    
    def _get_match_details(self, matches_df: pd.DataFrame) -> Dict:
        """Generate detailed match statistics"""
        if len(matches_df) == 0:
            return {}
        
        return {
            'by_source': matches_df['source'].value_counts().to_dict(),
            'by_category': matches_df['category'].value_counts().to_dict(),
            'by_comment': matches_df['comment'].value_counts().to_dict(),
            'amount_differences': {
                'exact': len(matches_df[matches_df['amount_difference'] == 0]),
                'rounded': len(matches_df[
                    (matches_df['amount_difference'] > 0) & 
                    (matches_df['amount_difference'] < 1.0)
                ]),
                'qty_mismatch': len(matches_df[matches_df['qty_difference'] != 0])
            }
        }


def print_report(results: Dict):
    """Print formatted reconciliation report"""
    summary = results['summary']
    matches = results['matches']
    details = results['match_details']
    
    print("\n" + "="*80)
    print("BILL vs PAY RECONCILIATION REPORT")
    print("="*80)
    
    # Overall summary
    print(f"\nOVERALL RESULTS:")
    print(f"  Total Payroll Rows:        {summary['total_payroll_rows']:>10,}")
    print(f"  Total Billing Rows:        {summary['total_billing_rows']:>10,}")
    print(f"  ─────────────────────────────────────")
    print(f"  Total Matched:             {summary['total_matches']:>10,}")
    print(f"  Match Rate (Payroll):      {summary['match_rate_payroll']:>10.2f}%")
    
    # Amounts
    print(f"\nAMOUNT RECONCILIATION:")
    print(f"  Total Payroll Amount:      ${summary['payroll_amount_total']:>15,.2f}")
    print(f"  Total Billing Amount:      ${summary['billing_amount_total']:>15,.2f}")
    print(f"  ─────────────────────────────────────")
    print(f"  Matched Payroll:           ${summary['matched_payroll_amount']:>15,.2f}")
    print(f"  Matched Billing:           ${summary['matched_billing_amount']:>15,.2f}")
    print(f"  ─────────────────────────────────────")
    print(f"  Unmatched Payroll:         ${summary['unmatched_payroll_amount']:>15,.2f}")
    print(f"  Unmatched Billing:         ${summary['unmatched_billing_amount']:>15,.2f}")
    
    # Match types
    if details:
        print(f"\nMATCH BREAKDOWN:")
        if 'by_source' in details:
            for source, count in details['by_source'].items():
                print(f"  {source:30s} {count:>10,}")
        
        print(f"\nMATCH COMMENTS:")
        if 'by_comment' in details:
            for comment, count in details['by_comment'].items():
                print(f"  {comment:50s} {count:>8,}")
        
        print(f"\nDIFFERENCE ANALYSIS:")
        print(f"  Exact Matches:             {details['amount_differences']['exact']:>10,}")
        print(f"  Rounded Off (<$1.00):      {details['amount_differences']['rounded']:>10,}")
        print(f"  Qty Mismatches:            {details['amount_differences']['qty_mismatch']:>10,}")
    
    # Unmatched summary
    print(f"\nUNMATCHED RECORDS:")
    print(f"  Paid But Not Billed:       {summary['unmatched_payroll_rows']:>10,}")
    print(f"  Billed But Not Paid:       {summary['unmatched_billing_rows']:>10,}")
    
    print("\n" + "="*80 + "\n")


if __name__ == '__main__':
    # Example usage
    payroll = pd.read_excel('Sam_P1.xlsx')
    billing = pd.read_excel('Sam_B1.xlsx')
    
    reconciler = BillPayReconciler(payroll, billing)
    results = reconciler.reconcile()
    
    print_report(results)
    
    # Sample matches
    if len(results['matches']) > 0:
        print("Sample Matched Records:")
        print(results['matches'][['employee_id', 'category', 'payroll_amount', 'billing_amount', 'comment']].head(10))
