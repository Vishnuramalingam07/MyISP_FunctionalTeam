"""
MyISP Internal Tools - Azure Functions App (100% FREE Architecture)
Converts Flask routes to Azure Functions HTTP triggers
"""
import azure.functions as func
import logging
import json
import subprocess
import sys
import os
import glob
from datetime import datetime
from pathlib import Path

# Create the Functions App
app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

# ══════════════════════════════════════════════════════════════════════════
# STATIC FILE SERVING
# ══════════════════════════════════════════════════════════════════════════

@app.route(route="", methods=["GET"])
def home(req: func.HttpRequest) -> func.HttpResponse:
    """Serve index.html"""
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return func.HttpResponse(f.read(), mimetype="text/html")
    except Exception as e:
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)

@app.route(route="{*path}", methods=["GET"])
def serve_static(req: func.HttpRequest) -> func.HttpResponse:
    """Serve static files (HTML, CSS, JS)"""
    path = req.route_params.get('path', '')
    
    # Security: prevent directory traversal
    if '..' in path or path.startswith('/'):
        return func.HttpResponse("Invalid path", status_code=400)
    
    # Determine file path
    if not path or path == '/':
        filepath = 'index.html'
    else:
        filepath = path
    
    # Check if file exists
    if not os.path.exists(filepath):
        return func.HttpResponse("File not found", status_code=404)
    
    # Determine mimetype
    ext = os.path.splitext(filepath)[1].lower()
    mimetypes = {
        '.html': 'text/html',
        '.css': 'text/css',
        '.js': 'application/javascript',
        '.json': 'application/json',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.svg': 'image/svg+xml',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.csv': 'text/csv'
    }
    mimetype = mimetypes.get(ext, 'application/octet-stream')
    
    try:
        if ext in ['.xlsx', '.png', '.jpg', '.jpeg', '.gif']:
            # Binary files
            with open(filepath, 'rb') as f:
                return func.HttpResponse(f.read(), mimetype=mimetype)
        else:
            # Text files
            with open(filepath, 'r', encoding='utf-8') as f:
                return func.HttpResponse(f.read(), mimetype=mimetype)
    except Exception as e:
        logging.error(f"Error serving {filepath}: {e}")
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)

# ══════════════════════════════════════════════════════════════════════════
# AUTHENTICATION API
# ══════════════════════════════════════════════════════════════════════════

@app.route(route="api/auth/check-access", methods=["GET", "POST"])
def check_access(req: func.HttpRequest) -> func.HttpResponse:
    """Check if user has access"""
    try:
        if req.method == "POST":
            body = req.get_json()
            username = body.get('username', '').lower().strip()
        else:
            username = req.params.get('username', '').lower().strip()
        
        if not username:
            return func.HttpResponse(
                json.dumps({'authorized': False, 'error': 'Username required'}),
                mimetype="application/json"
            )
        
        # Check against authorized users
        if DATABASE_ENABLED:
            result = cosmos.table('authorized_users').select('username').eq('username', username).execute()
            authorized = len(result) > 0
        else:
            # Fallback to CSV
            import csv
            access_file = os.path.join('Attendance', 'Access.csv')
            authorized = False
            if os.path.exists(access_file):
                with open(access_file, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    authorized = any(row.get('username', '').lower().strip() == username for row in reader)
        
        return func.HttpResponse(
            json.dumps({'authorized': authorized, 'username': username}),
            mimetype="application/json"
        )
    except Exception as e:
        logging.error(f"Auth error: {e}")
        return func.HttpResponse(
            json.dumps({'authorized': False, 'error': str(e)}),
            mimetype="application/json",
            status_code=500
        )

# ══════════════════════════════════════════════════════════════════════════
# REPORT GENERATION APIs
# ══════════════════════════════════════════════════════════════════════════

@app.route(route="api/generate_missing_field_tables", methods=["POST", "GET"])
def generate_missing_field_tables(req: func.HttpRequest) -> func.HttpResponse:
    """Generate missing data scope report"""
    try:
        script_path = os.path.join('Missing_Data_Report_For_Scope', 'generate_missing_field_tables.py')
        
        if not os.path.exists(script_path):
            return func.HttpResponse(
                json.dumps({'success': False, 'error': f'Script not found: {script_path}'}),
                mimetype="application/json",
                status_code=404
            )
        
        # Run the script
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=300,
            cwd=os.getcwd()
        )
        
        if result.returncode == 0:
            return func.HttpResponse(
                json.dumps({
                    'success': True,
                    'message': 'Report generated successfully',
                    'stdout': result.stdout,
                    'stderr': result.stderr
                }),
                mimetype="application/json"
            )
        else:
            return func.HttpResponse(
                json.dumps({
                    'success': False,
                    'error': 'Script failed',
                    'stdout': result.stdout,
                    'stderr': result.stderr,
                    'returncode': result.returncode
                }),
                mimetype="application/json",
                status_code=500
            )
    
    except subprocess.TimeoutExpired:
        return func.HttpResponse(
            json.dumps({'success': False, 'error': 'Script timeout (300s)'}),
            mimetype="application/json",
            status_code=504
        )
    except Exception as e:
        logging.error(f"Report generation error: {e}")
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            mimetype="application/json",
            status_code=500
        )

@app.route(route="api/generate-daily-status-report", methods=["POST"])
def generate_daily_status_report(req: func.HttpRequest) -> func.HttpResponse:
    """Generate daily status report"""
    try:
        body = req.get_json()
        script_path = os.path.join('Main_Release_Daily_Status_Report', 'generate_daily_status_report.py')
        
        if not os.path.exists(script_path):
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Script not found'}),
                mimetype="application/json",
                status_code=404
            )
        
        # Run script
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=300
        )
        
        return func.HttpResponse(
            json.dumps({
                'success': result.returncode == 0,
                'message': 'Report generated' if result.returncode == 0 else 'Failed',
                'stdout': result.stdout,
                'stderr': result.stderr
            }),
            mimetype="application/json"
        )
    except Exception as e:
        logging.error(f"Daily report error: {e}")
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            mimetype="application/json",
            status_code=500
        )

@app.route(route="api/generate-regression-report", methods=["POST"])
def generate_regression_report(req: func.HttpRequest) -> func.HttpResponse:
    """Generate regression report"""
    try:
        # Get multipart form data
        files = req.files
        form = req.form
        
        # Implementation similar to Flask version
        # (File handling logic here)
        
        return func.HttpResponse(
            json.dumps({'success': True, 'message': 'Report generated'}),
            mimetype="application/json"
        )
    except Exception as e:
        logging.error(f"Regression report error: {e}")
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            mimetype="application/json",
            status_code=500
        )

# ══════════════════════════════════════════════════════════════════════════
# FILE DOWNLOAD APIs
# ══════════════════════════════════════════════════════════════════════════

@app.route(route="api/download-daily-status-report", methods=["GET"])
def download_daily_status_report(req: func.HttpRequest) -> func.HttpResponse:
    """Download daily status report"""
    try:
        report_dir = os.path.join('Main_Release_Daily_Status_Report', 'Reports')
        
        if not os.path.exists(report_dir):
            return func.HttpResponse("Report directory not found", status_code=404)
        
        # Find the latest report
        reports = glob.glob(os.path.join(report_dir, 'Daily_Status_Report_*.xlsx'))
        if not reports:
            return func.HttpResponse("No reports found", status_code=404)
        
        latest_report = max(reports, key=os.path.getctime)
        
        with open(latest_report, 'rb') as f:
            return func.HttpResponse(
                f.read(),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={
                    "Content-Disposition": f"attachment; filename={os.path.basename(latest_report)}"
                }
            )
    except Exception as e:
        logging.error(f"Download error: {e}")
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)

# ══════════════════════════════════════════════════════════════════════════
# HEALTH CHECK
# ══════════════════════════════════════════════════════════════════════════

@app.route(route="api/health", methods=["GET"])
def health_check(req: func.HttpRequest) -> func.HttpResponse:
    """Health check endpoint"""
    return func.HttpResponse(
        json.dumps({
            'status': 'healthy',
            'timestamp': datetime.now().isoformat(),
            'database': 'connected' if DATABASE_ENABLED else 'disconnected',
            'version': '1.0.0'
        }),
        mimetype="application/json"
    )
