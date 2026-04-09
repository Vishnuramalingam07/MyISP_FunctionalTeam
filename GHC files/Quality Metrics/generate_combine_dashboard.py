from pathlib import Path
from datetime import datetime
import base64
import html

# File paths
base_path = Path(r'C:\Users\d.sampathkumar\GHC files\Quality Metrics')

# Read the source files
with open(base_path / 'Dashboard_Complete_Quality_Metrics.html', 'r', encoding='utf-8') as f:
    overview_html = f.read()

with open(base_path / 'Dashboard_Closure_Reopen_Analysis.html', 'r', encoding='utf-8') as f:
    closure_html = f.read()

with open(base_path / 'Dashboard_RCA_Analysis.html', 'r', encoding='utf-8') as f:
    rca_html = f.read()

# Function to convert HTML to data URI
def html_to_data_uri(html_content):
    """Convert HTML content to a data URI for embedding in iframe"""
    # URL encode the HTML content (better than base64 for HTML)
    # Using percent encoding
    import urllib.parse
    encoded = urllib.parse.quote(html_content)
    return f"data:text/html;charset=utf-8,{encoded}"

# Hide h1 titles in embedded versions
def hide_h1_titles(html_content):
    """Add CSS to hide h1 titles when embedded"""
    style_to_add = """
    <style>
        .header h1, .main-header h1 { display: none !important; }
        .header { padding: 10px 30px !important; }
        .header p, .header .subtitle { font-size: 0.85em !important; margin: 0 !important; }
        body { padding: 0 !important; margin: 0 !important; }
        .container, .main-container { margin: 0 !important; border-radius: 0 !important; box-shadow: none !important; }
    </style>
    """
    return html_content.replace('</head>', style_to_add + '\n</head>')

# Prepare embedded versions
overview_embedded = hide_h1_titles(overview_html)
closure_embedded = hide_h1_titles(closure_html)
rca_embedded = hide_h1_titles(rca_html)

# Create data URIs
overview_uri = html_to_data_uri(overview_embedded)
closure_uri = html_to_data_uri(closure_embedded)
rca_uri = html_to_data_uri(rca_embedded)

# Create the single combined file
output_file = base_path / 'Dashboard_quality_metrics_FINAL.html'

html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Quality Metrics Dashboard - Integrated View</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 5px;
            overflow: hidden;
        }}

        .main-container {{
            max-width: 100%;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
            height: calc(100vh - 10px);
            display: flex;
            flex-direction: column;
        }}

        .main-header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 30px;
            text-align: center;
        }}

        .main-header h1 {{
            font-size: 0.95em;
            margin-bottom: 3px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
            font-weight: 700;
        }}

        .main-header p {{
            font-size: 0.95em;
            opacity: 0.95;
            margin-top: 0px;
        }}

        .main-tabs {{
            display: flex;
            background: #f8f9fa;
            border-bottom: 4px solid #667eea;
            flex-wrap: wrap;
            padding: 0;
            position: sticky;
            top: 0;
            z-index: 1000;
        }}

        .main-tab {{
            flex: 1;
            min-width: 200px;
            padding: 12px 20px;
            background: linear-gradient(to bottom, #e9ecef 0%, #dee2e6 100%);
            border: none;
            cursor: pointer;
            font-size: 13px;
            font-weight: 700;
            transition: all 0.3s ease;
            border-right: 2px solid #adb5bd;
            position: relative;
            color: #495057;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .main-tab:last-child {{
            border-right: none;
        }}

        .main-tab:hover {{
            background: linear-gradient(to bottom, #dee2e6 0%, #ced4da 100%);
            transform: translateY(-2px);
        }}

        .main-tab.active {{
            background: white;
            color: #667eea;
            font-weight: 800;
        }}

        .main-tab.active::after {{
            content: '';
            position: absolute;
            bottom: -4px;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(to right, #667eea 0%, #764ba2 100%);
        }}

        .dashboard-container {{
            width: 100%;
            position: relative;
            display: none;
            background: white;
            flex: 1;
            overflow: hidden;
        }}

        .dashboard-container.active {{
            display: flex;
            flex-direction: column;
        }}

        .dashboard-container iframe {{
            width: 100%;
            height: 100%;
            border: none;
            display: block;
            flex: 1;
        }}

        .loading-overlay {{
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: white;
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10;
        }}

        .loading-overlay.hidden {{
            display: none;
        }}

        .loading-spinner {{
            text-align: center;
            color: #667eea;
        }}

        .loading-spinner h3 {{
            font-size: 20px;
            margin-bottom: 20px;
            color: #495057;
        }}

        .spinner {{
            width: 60px;
            height: 60px;
            margin: 0 auto;
            border: 6px solid #f3f3f3;
            border-top: 6px solid #667eea;
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }}

        @keyframes spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
    </style>
</head>
<body>
    <div class="main-container">
        <div class="main-header">
            <h1>📊 Quality Metrics Dashboard</h1>
            <p>Comprehensive Release Quality Analysis & Performance Tracking</p>
        </div>

        <div class="main-tabs">
            <button class="main-tab active" onclick="showDashboard('overview')">
                Overview & Metrics
            </button>
            <button class="main-tab" onclick="showDashboard('closure')">
                Closure & Reopen Analysis
            </button>
            <button class="main-tab" onclick="showDashboard('rca')">
                RCA Analysis
            </button>
        </div>

        <div id="overview" class="dashboard-container active">
            <div class="loading-overlay" id="overview-loading">
                <div class="loading-spinner">
                    <h3>Loading Overview Dashboard...</h3>
                    <div class="spinner"></div>
                </div>
            </div>
            <iframe src="{overview_uri}" onload="hideLoading('overview')"></iframe>
        </div>

        <div id="closure" class="dashboard-container">
            <div class="loading-overlay" id="closure-loading">
                <div class="loading-spinner">
                    <h3>Loading Closure & Reopen Analysis...</h3>
                    <div class="spinner"></div>
                </div>
            </div>
            <iframe src="{closure_uri}" onload="hideLoading('closure')"></iframe>
        </div>

        <div id="rca" class="dashboard-container">
            <div class="loading-overlay" id="rca-loading">
                <div class="loading-spinner">
                    <h3>Loading RCA Analysis...</h3>
                    <div class="spinner"></div>
                </div>
            </div>
            <iframe src="{rca_uri}" onload="hideLoading('rca')"></iframe>
        </div>
    </div>

    <script>
        function showDashboard(dashboardId) {{
            var containers = document.getElementsByClassName('dashboard-container');
            for (var i = 0; i < containers.length; i++) {{
                containers[i].classList.remove('active');
            }}

            var tabs = document.getElementsByClassName('main-tab');
            for (var i = 0; i < tabs.length; i++) {{
                tabs[i].classList.remove('active');
            }}

            document.getElementById(dashboardId).classList.add('active');
            event.target.classList.add('active');
        }}

        function hideLoading(dashboardId) {{
            var loading = document.getElementById(dashboardId + '-loading');
            if (loading) {{
                loading.classList.add('hidden');
            }}
        }}
    </script>
</body>
</html>'''

# Write the combined HTML
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

file_size_mb = output_file.stat().st_size / (1024 * 1024)

print(f"✓ Single-file dashboard with embedded data URIs created successfully!")
print(f"\nOutput file:")
print(f"  {output_file}")
print(f"  Size: {file_size_mb:.2f} MB")
print(f"\n✅ This is a TRUE STANDALONE file!")
print(f"✅ All dashboards embedded as data URIs (no external files needed)")
print(f"✅ All tabs and sub-tabs will work correctly with full isolation")
print(f"✅ Ready to email as a single attachment")
print(f"\nNote: File is larger because entire dashboards are embedded, but it's truly standalone.")
