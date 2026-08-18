"""
TEST SERVER - Minimal Flask app with NO reconciliation logic
Just to verify the server works and returns proper JSON
"""

import os
import logging
from datetime import datetime

from flask import Flask, request, jsonify

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Simple test progress
progress_data = {'status': 'idle', 'progress': 0, 'stage': '', 'message': '', 'elapsed_time': 0, 'estimated_time': 0}

@app.route('/', methods=['GET'])
def index():
    """Home page"""
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Bill vs Pay Reconciliation</title>
        <style>
            body { font-family: Arial; max-width: 900px; margin: 50px auto; padding: 20px; }
            .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
            h1 { color: #1F5F99; }
            .upload-section { background: #f9f9f9; padding: 20px; border-radius: 5px; margin: 20px 0; }
            label { display: block; margin: 15px 0 5px 0; font-weight: bold; }
            input[type="file"] { display: block; margin: 5px 0 15px 0; padding: 8px; width: 100%; }
            button { background: #1F5F99; color: white; padding: 12px; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-size: 16px; }
            button:hover { background: #146c43; }
            .status { margin: 20px 0; padding: 15px; background: #e3f2fd; border-radius: 4px; }
            .error { background: #ffebee; color: #c62828; }
            .success { background: #e8f5e9; color: #2e7d32; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Bill vs Pay Reconciliation</h1>
            <p>Upload payroll and billing exports</p>
            
            <div class="upload-section">
                <form id="testForm">
                    <label>Payroll Report (Excel)</label>
                    <input type="file" name="payroll" required accept=".xlsx">
                    
                    <label>Billing Report (Excel)</label>
                    <input type="file" name="billing" required accept=".xlsx">
                    
                    <button type="submit">Test Upload (No Reconciliation)</button>
                </form>
            </div>
            
            <div id="status" class="status" style="display:none;"></div>
        </div>
        
        <script>
            document.getElementById('testForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const status = document.getElementById('status');
                status.textContent = 'Testing server... Please wait...';
                status.style.display = 'block';
                status.className = 'status';
                
                const formData = new FormData();
                formData.append('payroll', document.querySelector('input[name="payroll"]').files[0]);
                formData.append('billing', document.querySelector('input[name="billing"]').files[0]);
                
                try {
                    const response = await fetch('/test-upload', {
                        method: 'POST',
                        body: formData
                    });
                    
                    console.log('Response status:', response.status);
                    console.log('Response ok:', response.ok);
                    
                    const text = await response.text();
                    console.log('Response text:', text);
                    
                    if (response.ok) {
                        status.textContent = '✓ SUCCESS! Server is working. Files received: ' + text;
                        status.className = 'status success';
                    } else {
                        status.textContent = '✗ Server returned ' + response.status + ': ' + text;
                        status.className = 'status error';
                    }
                } catch (error) {
                    status.textContent = '✗ Network error: ' + error.message;
                    status.className = 'status error';
                    console.error('Error:', error);
                }
            });
        </script>
    </body>
    </html>
    '''

@app.route('/progress', methods=['GET'])
def progress():
    """Get progress"""
    try:
        return jsonify({
            'status': 'idle',
            'progress': 0,
            'stage': 'Test Server',
            'message': 'Server is responding correctly',
            'elapsed_time': 0,
            'estimated_time': 0
        }), 200
    except Exception as e:
        logger.error(f"Progress error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/test-upload', methods=['POST'])
def test_upload():
    """Just test that uploads work - no reconciliation"""
    try:
        payroll_file = request.files.get('payroll')
        billing_file = request.files.get('billing')
        
        if not payroll_file or not billing_file:
            return 'Missing files', 400
        
        logger.info(f"Test received: {payroll_file.filename}, {billing_file.filename}")
        
        # Just return success
        return f'Server working! Received {payroll_file.filename} and {billing_file.filename}', 200
    
    except Exception as e:
        logger.error(f"Test upload error: {e}", exc_info=True)
        return f'Error: {str(e)}', 500

@app.route('/health', methods=['GET'])
def health():
    """Health check"""
    return jsonify({'status': 'healthy'}), 200

@app.errorhandler(404)
def not_found(error):
    return 'Not found', 404

@app.errorhandler(500)
def server_error(error):
    logger.error(f"Server error: {error}")
    return 'Server error: ' + str(error), 500

if __name__ == '__main__':
    host = os.environ.get('HOST', '127.0.0.1')
    port = int(os.environ.get('PORT', 8765))
    
    logger.info(f"TEST SERVER starting on {host}:{port}")
    logger.info("This server has NO reconciliation logic - just testing if Flask works")
    app.run(host=host, port=port, debug=False, threaded=True)
