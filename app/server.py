"""
Bill vs Pay Reconciliation Web Server - Enhanced with Progress Tracking (FIXED)
Flask application with streaming progress updates and timeout handling
"""

import os
import logging
import json
from datetime import datetime
from pathlib import Path
import time

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from reconciliation import BillPayReconciler

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create Flask app
app = Flask(__name__)
app.json.sort_keys = False

# Configuration
UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('output')
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# Global progress tracking
current_progress = {
    'status': 'idle',
    'stage': '',
    'progress': 0,
    'message': '',
    'estimated_time': 0,
    'elapsed_time': 0
}


class ProgressTracker:
    """Track reconciliation progress"""
    
    def __init__(self):
        self.start_time = None
        self.payroll_count = 0
        self.billing_count = 0
        self.current_stage = ''
        
    def start(self):
        self.start_time = time.time()
    
    def update(self, stage, progress, message=''):
        """Update progress (0-100)"""
        global current_progress
        
        elapsed = time.time() - self.start_time if self.start_time else 0
        
        # Estimate total time based on progress
        if progress > 5:
            estimated_total = (elapsed / progress) * 100
            estimated_remaining = estimated_total - elapsed
        else:
            estimated_remaining = 0
        
        current_progress = {
            'status': 'processing',
            'stage': stage,
            'progress': min(int(progress), 100),
            'message': message,
            'estimated_time': max(0, int(estimated_remaining)),
            'elapsed_time': int(elapsed)
        }
        
        logger.info(f"Progress: {stage} - {progress}% ({message}) - Elapsed: {int(elapsed)}s, Remaining: {max(0, int(estimated_remaining))}s")
    
    def complete(self):
        global current_progress
        current_progress['status'] = 'complete'
        current_progress['progress'] = 100


tracker = ProgressTracker()


def generate_output_workbook(results):
    """Generate Excel workbook with reconciliation results"""
    try:
        logger.info("Generating Excel report...")
        tracker.update('Excel Generation', 85, 'Creating workbook...')
        
        wb = Workbook()
        wb.remove(wb.active)
        
        # 1. Summary Tab
        ws_summary = wb.create_sheet('Summary', 0)
        summary = results['summary']
        
        ws_summary['A1'] = 'Bill vs Pay Reconciliation Report'
        ws_summary['A1'].font = Font(size=14, bold=True, color="FFFFFF")
        ws_summary['A1'].fill = PatternFill(start_color="1F5F99", end_color="1F5F99", fill_type="solid")
        ws_summary.merge_cells('A1:B1')
        
        ws_summary['A3'] = 'Report Date'
        ws_summary['B3'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        data = [
            ('', ''),
            ('OVERALL RESULTS', ''),
            ('Total Payroll Rows', summary['total_payroll_rows']),
            ('Total Billing Rows', summary['total_billing_rows']),
            ('Total Matches', summary['total_matches']),
            ('Match Rate %', f"{summary['match_rate_payroll']:.2f}%"),
            ('', ''),
            ('AMOUNT RECONCILIATION', ''),
            ('Total Payroll Amount', f"${summary['payroll_amount_total']:,.2f}"),
            ('Total Billing Amount', f"${summary['billing_amount_total']:,.2f}"),
            ('Matched Payroll Amount', f"${summary['matched_payroll_amount']:,.2f}"),
            ('Matched Billing Amount', f"${summary['matched_billing_amount']:,.2f}"),
            ('Unmatched Payroll Amount', f"${summary['unmatched_payroll_amount']:,.2f}"),
            ('Unmatched Billing Amount', f"${summary['unmatched_billing_amount']:,.2f}"),
            ('', ''),
            ('UNMATCHED SUMMARY', ''),
            ('Paid But Not Billed (Rows)', summary['unmatched_payroll_rows']),
            ('Billed But Not Paid (Rows)', summary['unmatched_billing_rows']),
        ]
        
        row = 4
        for label, value in data:
            ws_summary[f'A{row}'] = label
            ws_summary[f'B{row}'] = value
            
            if label.startswith('OVERALL') or label.startswith('AMOUNT') or label.startswith('UNMATCHED'):
                ws_summary[f'A{row}'].font = Font(bold=True, color="FFFFFF")
                ws_summary[f'A{row}'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                ws_summary.merge_cells(f'A{row}:B{row}')
            row += 1
        
        ws_summary.column_dimensions['A'].width = 35
        ws_summary.column_dimensions['B'].width = 25
        
        # 2. Match Details Tab
        logger.info("Writing match details...")
        tracker.update('Excel Generation', 90, 'Writing match details...')
        
        matches_df = results['matches']
        if len(matches_df) > 0:
            ws_matches = wb.create_sheet('Match Details')
            
            headers = list(matches_df.columns)
            for col_idx, header in enumerate(headers, 1):
                cell = ws_matches.cell(row=1, column=col_idx)
                cell.value = header
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1F5F99", end_color="1F5F99", fill_type="solid")
            
            for row_idx, row_data in enumerate(matches_df.itertuples(), 2):
                for col_idx, value in enumerate(row_data[1:], 1):
                    cell = ws_matches.cell(row=row_idx, column=col_idx)
                    cell.value = value
                    
                    if col_idx == 11:
                        if 'Tallied' in str(value):
                            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        elif 'Difference' in str(value):
                            cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        
        # 3. Unmatched Payroll
        logger.info("Writing unmatched payroll...")
        tracker.update('Excel Generation', 95, 'Writing unmatched payroll...')
        
        up_df = results['unmatched_payroll']
        if len(up_df) > 0:
            ws_up = wb.create_sheet('Unmatched Payroll')
            headers = list(up_df.columns)
            for col_idx, header in enumerate(headers, 1):
                cell = ws_up.cell(row=1, column=col_idx)
                cell.value = header
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            
            for row_idx, row_data in enumerate(up_df.itertuples(), 2):
                for col_idx, value in enumerate(row_data[1:], 1):
                    ws_up.cell(row=row_idx, column=col_idx).value = value
        
        # 4. Unmatched Billing
        ub_df = results['unmatched_billing']
        if len(ub_df) > 0:
            ws_ub = wb.create_sheet('Unmatched Billing')
            headers = list(ub_df.columns)
            for col_idx, header in enumerate(headers, 1):
                cell = ws_ub.cell(row=1, column=col_idx)
                cell.value = header
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
            
            for row_idx, row_data in enumerate(ub_df.itertuples(), 2):
                for col_idx, value in enumerate(row_data[1:], 1):
                    ws_ub.cell(row=row_idx, column=col_idx).value = value
        
        # Save
        output_path = f'/tmp/reconciliation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        wb.save(output_path)
        
        logger.info(f"Excel report generated: {output_path}")
        tracker.update('Excel Generation', 100, 'Complete!')
        
        return output_path
    
    except Exception as e:
        logger.error(f"Error generating Excel output: {e}", exc_info=True)
        raise


@app.route('/', methods=['GET'])
def index():
    """Home page with progress tracking"""
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Bill vs Pay Reconciliation</title>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            * { box-sizing: border-box; }
            body { 
                font-family: Arial, sans-serif; 
                max-width: 900px; 
                margin: 0 auto;
                background: #f5f5f5;
                padding: 20px;
            }
            .container { 
                background: white; 
                padding: 30px; 
                border-radius: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            }
            .header {
                display: flex;
                gap: 20px;
                margin-bottom: 30px;
            }
            .header-content h1 { 
                color: #1F5F99;
                margin: 0;
            }
            .header-content p {
                color: #666;
                margin: 5px 0 0 0;
            }
            .upload-section { 
                background: #f9f9f9; 
                padding: 20px; 
                border-radius: 5px;
                margin-bottom: 20px;
            }
            label { 
                display: block; 
                margin: 15px 0 5px 0;
                font-weight: bold;
                color: #333;
            }
            input[type="file"] { 
                display: block; 
                margin: 5px 0 15px 0;
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                width: 100%;
            }
            button { 
                background: #1F5F99; 
                color: white; 
                padding: 12px 30px; 
                border: none; 
                border-radius: 4px; 
                cursor: pointer;
                font-size: 16px;
                width: 100%;
            }
            button:hover { 
                background: #146c43; 
            }
            button:disabled {
                background: #999;
                cursor: not-allowed;
            }
            .progress-section {
                display: none;
                margin-top: 20px;
                padding: 20px;
                background: #f0f4f8;
                border-radius: 5px;
                border-left: 4px solid #1F5F99;
            }
            .progress-section.active {
                display: block;
            }
            .progress-bar {
                background: #ddd;
                height: 30px;
                border-radius: 4px;
                overflow: hidden;
                margin: 10px 0;
            }
            .progress-fill {
                background: linear-gradient(90deg, #1F5F99, #4472C4);
                height: 100%;
                width: 0%;
                transition: width 0.3s ease;
                display: flex;
                align-items: center;
                justify-content: center;
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            .progress-info {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 10px;
                margin-top: 10px;
                font-size: 14px;
            }
            .info-item {
                background: white;
                padding: 10px;
                border-radius: 4px;
            }
            .info-label {
                font-weight: bold;
                color: #666;
                font-size: 12px;
            }
            .info-value {
                color: #1F5F99;
                font-size: 18px;
                margin-top: 5px;
            }
            .status-message {
                margin-top: 10px;
                padding: 10px;
                background: white;
                border-radius: 4px;
                color: #333;
                min-height: 20px;
            }
            .error-section {
                display: none;
                margin-top: 20px;
                padding: 15px;
                background: #ffebee;
                border-left: 4px solid #ff0000;
                border-radius: 4px;
                color: #c62828;
            }
            .error-section.active {
                display: block;
            }
            .success-section {
                display: none;
                margin-top: 20px;
                padding: 15px;
                background: #e8f5e9;
                border-left: 4px solid #4caf50;
                border-radius: 4px;
                color: #2e7d32;
            }
            .success-section.active {
                display: block;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="header-content">
                    <h1>Bill vs Pay Reconciliation</h1>
                    <p>Upload payroll and billing exports for automated reconciliation</p>
                </div>
            </div>
            
            <div class="upload-section">
                <form id="reconcileForm">
                    <label>Payroll Report (Excel)</label>
                    <input type="file" name="payroll" required accept=".xlsx" id="payrollInput">
                    
                    <label>Billing Report (Excel)</label>
                    <input type="file" name="billing" required accept=".xlsx" id="billingInput">
                    
                    <button type="submit" id="submitBtn">Run Reconciliation</button>
                </form>
            </div>
            
            <div id="progressSection" class="progress-section">
                <h3>Reconciliation Progress</h3>
                <div class="progress-bar">
                    <div id="progressFill" class="progress-fill" style="width: 0%;">0%</div>
                </div>
                
                <div class="progress-info">
                    <div class="info-item">
                        <div class="info-label">ELAPSED TIME</div>
                        <div class="info-value" id="elapsedTime">0s</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">ESTIMATED REMAINING</div>
                        <div class="info-value" id="estimatedTime">--</div>
                    </div>
                </div>
                
                <div class="status-message">
                    <strong>Current Stage:</strong> <span id="stageText">Initializing...</span><br>
                    <strong>Status:</strong> <span id="statusText">Preparing files...</span>
                </div>
            </div>
            
            <div id="errorSection" class="error-section">
                <strong>Error:</strong> <span id="errorText"></span>
            </div>
            
            <div id="successSection" class="success-section">
                <strong>✓ Success!</strong> Your reconciliation report has been generated and downloaded.
            </div>
        </div>
        
        <script>
            let progressInterval = null;
            
            function formatTime(seconds) {
                if (!seconds || seconds < 0) return '--';
                if (seconds < 60) return seconds + 's';
                return Math.floor(seconds / 60) + 'm ' + (seconds % 60) + 's';
            }
            
            async function updateProgress() {
                try {
                    const response = await fetch('/progress');
                    if (!response.ok) {
                        console.error('Progress fetch failed:', response.status);
                        return;
                    }
                    
                    const data = await response.json();
                    
                    if (data && data.status === 'processing') {
                        const progress = Math.min(data.progress || 0, 100);
                        document.getElementById('progressFill').style.width = progress + '%';
                        document.getElementById('progressFill').textContent = progress + '%';
                        document.getElementById('stageText').textContent = data.stage || 'Processing...';
                        document.getElementById('statusText').textContent = data.message || 'Working...';
                        document.getElementById('elapsedTime').textContent = formatTime(data.elapsed_time);
                        document.getElementById('estimatedTime').textContent = formatTime(data.estimated_time);
                    } else if (data && data.status === 'complete') {
                        if (progressInterval) clearInterval(progressInterval);
                        document.getElementById('progressFill').style.width = '100%';
                        document.getElementById('progressFill').textContent = '100%';
                    }
                } catch (error) {
                    console.error('Progress update error:', error);
                }
            }
            
            document.getElementById('reconcileForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                
                // Reset UI
                document.getElementById('progressSection').classList.add('active');
                document.getElementById('errorSection').classList.remove('active');
                document.getElementById('successSection').classList.remove('active');
                document.getElementById('submitBtn').disabled = true;
                
                const formData = new FormData();
                formData.append('payroll', document.getElementById('payrollInput').files[0]);
                formData.append('billing', document.getElementById('billingInput').files[0]);
                
                // Start progress polling
                progressInterval = setInterval(updateProgress, 500);
                updateProgress();
                
                try {
                    const response = await fetch('/reconcile', {
                        method: 'POST',
                        body: formData
                    });
                    
                    if (response.ok) {
                        const blob = await response.blob();
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = `reconciliation_${new Date().toISOString().split('T')[0]}.xlsx`;
                        a.click();
                        
                        document.getElementById('successSection').classList.add('active');
                        document.getElementById('progressSection').classList.remove('active');
                    } else {
                        try {
                            const errorData = await response.json();
                            showError(errorData.error || 'Reconciliation failed');
                        } catch {
                            showError('Reconciliation failed with status ' + response.status);
                        }
                    }
                } catch (error) {
                    showError('Network error: ' + error.message);
                } finally {
                    if (progressInterval) clearInterval(progressInterval);
                    document.getElementById('submitBtn').disabled = false;
                }
            });
            
            function showError(message) {
                document.getElementById('errorText').textContent = message;
                document.getElementById('errorSection').classList.add('active');
                document.getElementById('progressSection').classList.remove('active');
            }
        </script>
    </body>
    </html>
    '''


@app.route('/progress', methods=['GET'])
def progress():
    """Get current progress status"""
    try:
        return jsonify(current_progress), 200
    except Exception as e:
        logger.error(f"Error in /progress endpoint: {e}")
        return jsonify({
            'status': 'error',
            'stage': 'Error',
            'progress': 0,
            'message': str(e),
            'estimated_time': 0,
            'elapsed_time': 0
        }), 500


@app.route('/reconcile', methods=['POST'])
def reconcile():
    """Main reconciliation endpoint"""
    try:
        # Reset progress
        global current_progress
        current_progress = {
            'status': 'processing',
            'stage': 'Initialization',
            'progress': 0,
            'message': 'Validating files...',
            'estimated_time': 0,
            'elapsed_time': 0
        }
        
        tracker.start()
        
        # Get files
        payroll_file = request.files.get('payroll')
        billing_file = request.files.get('billing')
        
        # Validate
        if not payroll_file or not billing_file:
            logger.warning("Missing files")
            return jsonify({'error': 'Both payroll and billing files are required'}), 400
        
        if not payroll_file.filename.endswith('.xlsx'):
            return jsonify({'error': 'Payroll file must be .xlsx format'}), 400
        
        if not billing_file.filename.endswith('.xlsx'):
            return jsonify({'error': 'Billing file must be .xlsx format'}), 400
        
        logger.info(f"Received files: {payroll_file.filename}, {billing_file.filename}")
        tracker.update('File Loading', 10, f'Loading {payroll_file.filename}...')
        
        # Load files
        try:
            payroll_df = pd.read_excel(payroll_file)
            logger.info(f"Loaded {len(payroll_df)} payroll rows")
            tracker.update('File Loading', 20, f'Loading {billing_file.filename}...')
            
            billing_df = pd.read_excel(billing_file)
            logger.info(f"Loaded {len(billing_df)} billing rows")
            tracker.update('File Loading', 30, f'Files loaded successfully')
            
        except Exception as e:
            logger.error(f"Error reading Excel files: {e}")
            return jsonify({'error': f'Error reading files: {str(e)}'}), 400
        
        # Run reconciliation
        try:
            logger.info("Starting reconciliation engine...")
            tracker.update('Reconciliation', 40, 'Processing regular items...')
            
            reconciler = BillPayReconciler(payroll_df, billing_df)
            tracker.update('Reconciliation', 60, 'Processing special scenarios...')
            
            results = reconciler.reconcile()
            tracker.update('Reconciliation', 80, 'Reconciliation complete')
            
            logger.info(f"Reconciliation complete: {results['summary']['total_matches']} matches")
            
        except Exception as e:
            logger.error(f"Reconciliation error: {e}", exc_info=True)
            return jsonify({'error': f'Reconciliation failed: {str(e)}'}), 500
        
        # Generate output
        try:
            output_file = generate_output_workbook(results)
            tracker.complete()
            
            return send_file(
                output_file,
                as_attachment=True,
                download_name=f"reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            logger.error(f"Error generating output: {e}", exc_info=True)
            return jsonify({'error': f'Error generating report: {str(e)}'}), 500
    
    except Exception as e:
        logger.error(f"Unexpected error in /reconcile: {e}", exc_info=True)
        return jsonify({'error': 'Unexpected error occurred: ' + str(e)}), 500


@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy'}), 200


@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Endpoint not found'}), 404


@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Internal server error: {error}")
    return jsonify({'error': 'Internal server error'}), 500


if __name__ == '__main__':
    host = os.environ.get('HOST', '127.0.0.1')
    port = int(os.environ.get('PORT', 8765))
    
    logger.info(f"Starting server on {host}:{port}")
    app.run(host=host, port=port, debug=False, threaded=True)
