import pandas as pd
import numpy as np
import re

# Read the Excel file
input_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\quality metrcis input_RCA Analysis.xlsx'

# Load all sheets dynamically
xl = pd.ExcelFile(input_file)
all_data = []

# Get all sheet names from the Excel file
releases = xl.sheet_names
print(f"Found {len(releases)} sheets in Excel file:")
for sheet in releases:
    print(f"  - {sheet}")

# Read all sheets dynamically
for release in releases:
    df = pd.read_excel(xl, release)
    df['Release'] = release
    all_data.append(df)
    print(f"  Loaded {len(df)} records from '{release}'")

# Combine all data
combined_df = pd.concat(all_data, ignore_index=True)

print(f"\nTotal loaded: {len(combined_df)} rows from {len(releases)} sheets")

# Clean data
combined_df['Count'] = pd.to_numeric(combined_df['Count'], errors='coerce').fillna(0).astype(int)

# Categorize RCA Type
combined_df['RCA Category'] = combined_df['RCA Type'].apply(
    lambda x: 'Dev Countable' if pd.notna(x) and 'DEV Countable' in str(x) else 'Dev NOT Countable'
)

print(f"Dev Countable: {len(combined_df[combined_df['RCA Category'] == 'Dev Countable'])}")
print(f"Dev NOT Countable: {len(combined_df[combined_df['RCA Category'] == 'Dev NOT Countable'])}")

# Create summaries by different dimensions
def create_summary(group_cols):
    summary = combined_df.groupby(group_cols + ['RCA Category']).agg({'Count': 'sum'}).reset_index()
    # Pivot to get Dev and Non-Dev as separate columns
    pivot = summary.pivot_table(
        index=group_cols, 
        columns='RCA Category', 
        values='Count', 
        fill_value=0
    ).reset_index()
    pivot.columns.name = None
    pivot['Total'] = pivot.get('Valid Defects', 0) + pivot.get('Invalid Defects', 0)
    pivot['Dev %'] = ((pivot.get('Valid Defects', 0) / pivot['Total']) * 100).fillna(0).round(1)
    return pivot

# Prepare data for JavaScript - create hierarchical structure
import json

# Build hierarchical data structure for filters
hierarchy_data = []
for _, row in combined_df.iterrows():
    hierarchy_data.append({
        'release': row['Release'],
        'adPoc': row['AD POC'],
        'smPoc': row['SM POC'],
        'mPoc': row['M POC'],
        'nodeName': row['Node Name'],
        'rcaType': row['RCA Type'],
        'rcaCategory': row['RCA Category'],
        'myspRca': row['mySP RCA'],
        'count': int(row['Count'])
    })

# Calculate overall metrics
total_rcas = int(combined_df['Count'].sum())
dev_countable = int(combined_df[combined_df['RCA Category'] == 'Dev Countable']['Count'].sum())
dev_not_countable = int(combined_df[combined_df['RCA Category'] == 'Dev NOT Countable']['Count'].sum())
dev_pct = round((dev_countable / total_rcas * 100), 1) if total_rcas > 0 else 0

# Generate dynamic table headers for releases
release_headers = ''.join([f'<th>{release}</th>' for release in releases])
release_js_array = json.dumps(releases)

# Start HTML generation
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>RCA Analysis Dashboard</title>
    <style>
        * {{margin:0;padding:0;box-sizing:border-box}}
        body {{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:10px;min-height:100vh}}
        .container {{max-width:1600px;margin:0 auto;background:white;border-radius:15px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}}
        .header {{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:15px;text-align:center}}
        .header h1 {{font-size:2.2em;margin-bottom:5px;text-shadow:2px 2px 4px rgba(0,0,0,0.2)}}
        .header p {{font-size:1.1em;opacity:0.9}}
        .content {{padding:20px}}
        .filter-section {{margin-bottom:15px;padding:15px;background:#f8f9fa;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}}
        .filter-title {{font-size:1.1em;color:#333;margin-bottom:10px;font-weight:600}}
        .filters {{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}}
        .filter-group {{display:flex;flex-direction:column}}
        .filter-group label {{font-weight:600;color:#555;margin-bottom:4px;font-size:0.8em}}
        .filter-group select {{padding:6px 8px;border:2px solid #667eea;border-radius:6px;font-size:0.85em;background:white;cursor:pointer;transition:all 0.3s}}
        .filter-group select:disabled {{background:#e9ecef;cursor:not-allowed;border-color:#ddd}}
        .filter-group select:enabled:hover {{border-color:#764ba2;box-shadow:0 2px 8px rgba(102,126,234,0.2)}}
        .section-title {{font-size:1.5em;color:#333;margin:15px 0 10px 0;padding-bottom:8px;border-bottom:3px solid #667eea}}
        .metrics-grid {{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:20px;margin:20px 0}}
        .metric-card {{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:20px;border-radius:10px;box-shadow:0 4px 12px rgba(0,0,0,0.1)}}
        .metric-card.green {{background:linear-gradient(135deg,#11998e 0%,#38ef7d 100%)}}
        .metric-card.orange {{background:linear-gradient(135deg,#f093fb 0%,#f5576c 100%)}}
        .metric-label {{font-size:0.85em;opacity:0.9;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px}}
        .metric-value {{font-size:2.2em;font-weight:bold}}
        .table-wrapper {{margin:10px 0;border:2px solid #667eea;border-radius:8px;overflow:hidden}}
        .data-table {{width:100%;border-collapse:collapse;border:1px solid #ddd;table-layout:fixed;font-size:0.85em}}
        .data-table th {{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:10px;text-align:left;font-weight:600;border:1px solid #555}}
        .data-table td {{padding:10px;border:1px solid #ddd}}
        .data-table tr:hover {{background:#f5f5f5}}
        .data-table th:first-child, .data-table td:first-child {{width:30%;word-wrap:break-word}}
        .data-table th:not(:first-child), .data-table td:not(:first-child) {{width:23.33%;text-align:center;white-space:nowrap}}
        #categoryTab .data-table {{table-layout:auto}}
        #categoryTab .data-table th:nth-child(1), #categoryTab .data-table td:nth-child(1) {{width:12%;text-align:left;word-wrap:break-word}}
        #categoryTab .data-table th:nth-child(2), #categoryTab .data-table td:nth-child(2) {{width:12%;text-align:left;word-wrap:break-word}}
        #categoryTab .data-table th:nth-child(3), #categoryTab .data-table td:nth-child(3) {{width:12%;text-align:left;word-wrap:break-word}}
        #categoryTab .data-table th:nth-child(4), #categoryTab .data-table td:nth-child(4) {{width:28%;text-align:left;word-wrap:break-word}}
        #categoryTab .data-table th:nth-child(5), #categoryTab .data-table td:nth-child(5) {{width:9%;text-align:center}}
        #categoryTab .data-table th:nth-child(6), #categoryTab .data-table td:nth-child(6) {{width:9%;text-align:center}}
        #categoryTab .data-table th:nth-child(7), #categoryTab .data-table td:nth-child(7) {{width:9%;text-align:center}}
        #categoryTab .data-table th:nth-child(8), #categoryTab .data-table td:nth-child(8) {{width:9%;text-align:center}}
        .summary-section {{margin:10px 0;padding:15px;background:#f8f9fa;border-radius:8px}}
        .no-data {{text-align:center;padding:20px;color:#999;font-size:1.1em}}
        .tabs {{display:flex;border-bottom:3px solid #667eea;margin-bottom:15px;background:#f8f9fa;border-radius:10px 10px 0 0}}
        .tab {{flex:1;padding:15px 20px;text-align:center;cursor:pointer;font-weight:600;font-size:1.1em;color:#555;transition:all 0.3s;border-bottom:3px solid transparent;margin-bottom:-3px}}
        .tab:hover {{background:#e9ecef}}
        .tab.active {{color:#667eea;border-bottom-color:#667eea;background:white}}
        .tab-content {{display:none}}
        .tab-content.active {{display:block}}
        .category-section {{margin:12px 0;padding:15px;background:#f8f9fa;border-radius:8px;border-left:5px solid #667eea}}
        .category-title {{font-size:1.3em;color:#667eea;margin-bottom:12px;font-weight:600}}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 RCA Analysis Dashboard</h1>
            <p>Root Cause Analysis - Valid vs InValid Defects</p>
        </div>
        <div class="content">
            <div class="tabs">
                <div class="tab active" onclick="switchTab('detailedTab')">Detailed RCA Breakdown</div>
                <div class="tab" onclick="switchTab('categoryTab')">Category Comparison</div>
            </div>
            
            <div id="detailedTab" class="tab-content active">
            <div class="filter-section">
                <div class="filter-title">📋 Filters</div>
                <div class="filters">
                    <div class="filter-group">
                        <label for="adPocFilter">AD POC</label>
                        <select id="adPocFilter" onchange="handleAdPocChange()">
                            <option value="">All AD POCs</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="smPocFilter">SM POC</label>
                        <select id="smPocFilter" disabled onchange="handleSmPocChange()">
                            <option value="">All SM POCs</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="mPocFilter">M POC</label>
                        <select id="mPocFilter" disabled onchange="handleMPocChange()">
                            <option value="">All M POCs</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label for="nodeNameFilter">Node Name</label>
                        <select id="nodeNameFilter" disabled onchange="handleNodeNameChange()">
                            <option value="">All Nodes</option>
                        </select>
                    </div>
                </div>
            </div>
            
            <div id="summaryContent">
                <div class="summary-section">
                    <h2 class="section-title">📊 RCA Summary</h2>
                    <div class="table-wrapper">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Metric</th>
                                    {release_headers}
                                </tr>
                            </thead>
                            <tbody id="rcaSummaryTable">
                                <tr><td colspan="4" class="no-data">Loading...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
                
                <div class="summary-section">
                    <h2 class="section-title">✅ Valid Defect Summary</h2>
                    <div class="table-wrapper">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>mySP RCA</th>
                                    {release_headers}
                                </tr>
                            </thead>
                            <tbody id="devCountableTable">
                                <tr><td colspan="4" class="no-data">Loading...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
                
                <div class="summary-section">
                    <h2 class="section-title">⚠️ Invalid Defect Summary</h2>
                    <div class="table-wrapper">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>mySP RCA</th>
                                    {release_headers}
                                </tr>
                            </thead>
                            <tbody id="devNotCountableTable">
                                <tr><td colspan="4" class="no-data">Loading...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            </div>
            
            <div id="categoryTab" class="tab-content">
                <div class="summary-section">
                    <h2 class="section-title">🏆 Top 5 Valid Defect Categories - Node Analysis</h2>
                    <p style="color:#666;margin-bottom:20px">Top 10 nodes for each of the top 5 Valid Defect mySP RCA categories across releases</p>
                </div>
                <div id="categoryComparisonContent">
                    <div class="no-data">Loading...</div>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        const rawData = {json.dumps(hierarchy_data)};
        const releases = {release_js_array};
        
        function switchTab(tabId) {{
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
            
            // Show selected tab
            document.getElementById(tabId).classList.add('active');
            event.target.classList.add('active');
            
            // Load category comparison data when switching to that tab
            if (tabId === 'categoryTab') {{
                loadCategoryComparison();
            }}
        }}
        
        function handleAdPocChange() {{
            const adPoc = document.getElementById('adPocFilter').value;
            const smPocSelect = document.getElementById('smPocFilter');
            const mPocSelect = document.getElementById('mPocFilter');
            const nodeNameSelect = document.getElementById('nodeNameFilter');
            
            if (adPoc) {{
                smPocSelect.disabled = false;
                populateSmPocFilter();
            }} else {{
                smPocSelect.disabled = true;
                smPocSelect.value = '';
                mPocSelect.disabled = true;
                mPocSelect.value = '';
                nodeNameSelect.disabled = true;
                nodeNameSelect.value = '';
            }}
            updateSummary();
        }}
        
        function handleSmPocChange() {{
            const smPoc = document.getElementById('smPocFilter').value;
            const mPocSelect = document.getElementById('mPocFilter');
            const nodeNameSelect = document.getElementById('nodeNameFilter');
            
            if (smPoc) {{
                mPocSelect.disabled = false;
                populateMPocFilter();
            }} else {{
                mPocSelect.disabled = true;
                mPocSelect.value = '';
                nodeNameSelect.disabled = true;
                nodeNameSelect.value = '';
            }}
            updateSummary();
        }}
        
        function handleMPocChange() {{
            const mPoc = document.getElementById('mPocFilter').value;
            const nodeNameSelect = document.getElementById('nodeNameFilter');
            
            if (mPoc) {{
                nodeNameSelect.disabled = false;
                populateNodeNameFilter();
            }} else {{
                nodeNameSelect.disabled = true;
                nodeNameSelect.value = '';
            }}
            updateSummary();
        }}
        
        function handleNodeNameChange() {{
            updateSummary();
        }}
        
        function populateAdPocFilter() {{
            const adPocSelect = document.getElementById('adPocFilter');
            const smPocSelect = document.getElementById('smPocFilter');
            const mPocSelect = document.getElementById('mPocFilter');
            const nodeNameSelect = document.getElementById('nodeNameFilter');
            
            const adPocs = [...new Set(rawData.map(d => d.adPoc))].sort();
            adPocSelect.innerHTML = '<option value="">All AD POCs</option>' + 
                adPocs.map(poc => `<option value="${{poc}}">${{poc}}</option>`).join('');
            
            smPocSelect.value = '';
            smPocSelect.disabled = true;
            mPocSelect.value = '';
            mPocSelect.disabled = true;
            nodeNameSelect.value = '';
            nodeNameSelect.disabled = true;
        }}
        
        function populateSmPocFilter() {{
            const adPoc = document.getElementById('adPocFilter').value;
            const smPocSelect = document.getElementById('smPocFilter');
            const mPocSelect = document.getElementById('mPocFilter');
            const nodeNameSelect = document.getElementById('nodeNameFilter');
            
            let filteredData = rawData;
            if (adPoc) filteredData = filteredData.filter(d => d.adPoc === adPoc);
            
            const smPocs = [...new Set(filteredData.map(d => d.smPoc))].sort();
            smPocSelect.innerHTML = '<option value="">All SM POCs</option>' + 
                smPocs.map(poc => `<option value="${{poc}}">${{poc}}</option>`).join('');
            
            mPocSelect.value = '';
            mPocSelect.disabled = true;
            nodeNameSelect.value = '';
            nodeNameSelect.disabled = true;
        }}
        
        function populateMPocFilter() {{
            const adPoc = document.getElementById('adPocFilter').value;
            const smPoc = document.getElementById('smPocFilter').value;
            const mPocSelect = document.getElementById('mPocFilter');
            
            let filteredData = rawData;
            if (adPoc) filteredData = filteredData.filter(d => d.adPoc === adPoc);
            if (smPoc) filteredData = filteredData.filter(d => d.smPoc === smPoc);
            
            const mPocs = [...new Set(filteredData.map(d => d.mPoc))].sort();
            mPocSelect.innerHTML = '<option value="">All M POCs</option>' + 
                mPocs.map(poc => `<option value="${{poc}}">${{poc}}</option>`).join('');
        }}
        
        function populateNodeNameFilter() {{
            const adPoc = document.getElementById('adPocFilter').value;
            const smPoc = document.getElementById('smPocFilter').value;
            const mPoc = document.getElementById('mPocFilter').value;
            const nodeNameSelect = document.getElementById('nodeNameFilter');
            
            let filteredData = rawData;
            if (adPoc) filteredData = filteredData.filter(d => d.adPoc === adPoc);
            if (smPoc) filteredData = filteredData.filter(d => d.smPoc === smPoc);
            if (mPoc) filteredData = filteredData.filter(d => d.mPoc === mPoc);
            
            const nodeNames = [...new Set(filteredData.map(d => d.nodeName))].sort();
            nodeNameSelect.innerHTML = '<option value="">All Nodes</option>' + 
                nodeNames.map(node => `<option value="${{node}}">${{node}}</option>`).join('');
        }}
        
        function updateSummary() {{
            const adPoc = document.getElementById('adPocFilter').value;
            const smPoc = document.getElementById('smPocFilter').value;
            const mPoc = document.getElementById('mPocFilter').value;
            const nodeName = document.getElementById('nodeNameFilter').value;
            
            // Get base filtered data (without release filter)
            let baseFilteredData = rawData;
            if (adPoc) baseFilteredData = baseFilteredData.filter(d => d.adPoc === adPoc);
            if (smPoc) baseFilteredData = baseFilteredData.filter(d => d.smPoc === smPoc);
            if (mPoc) baseFilteredData = baseFilteredData.filter(d => d.mPoc === mPoc);
            if (nodeName) baseFilteredData = baseFilteredData.filter(d => d.nodeName === nodeName);
            
            // Calculate RCA Summary for each release
            const rcaSummaryTable = document.getElementById('rcaSummaryTable');
            let summaryRows = '';
            
            const metrics = ['Total Count', 'Valid Defects', 'Invalid Defects', '% Valid Defects', '% InValid Defects'];
            metrics.forEach(metric => {{
                let row = `<tr><td><strong>${{metric}}</strong></td>`;
                releases.forEach(release => {{
                    const releaseData = baseFilteredData.filter(d => d.release === release);
                    const totalCount = releaseData.reduce((sum, d) => sum + d.count, 0);
                    const devCountable = releaseData.filter(d => d.rcaCategory === 'Dev Countable').reduce((sum, d) => sum + d.count, 0);
                    const devNotCountable = releaseData.filter(d => d.rcaCategory === 'Dev NOT Countable').reduce((sum, d) => sum + d.count, 0);
                    const devPct = totalCount > 0 ? ((devCountable / totalCount) * 100).toFixed(1) : 0;
                    const devNotPct = totalCount > 0 ? ((devNotCountable / totalCount) * 100).toFixed(1) : 0;
                    
                    let value = '';
                    if (metric === 'Total Count') value = totalCount;
                    else if (metric === 'Valid Defects') value = devCountable;
                    else if (metric === 'Invalid Defects') value = devNotCountable;
                    else if (metric === '% Valid Defects') value = devPct + '%';
                    else if (metric === '% InValid Defects') value = devNotPct + '%';
                    
                    row += `<td>${{value}}</td>`;
                }});
                row += '</tr>';
                summaryRows += row;
            }});
            rcaSummaryTable.innerHTML = summaryRows;
            
            // Dev Countable Summary - group by mySP RCA across releases
            const devCountableByMySp = {{}};
            releases.forEach(release => {{
                const releaseData = baseFilteredData.filter(d => d.release === release && d.rcaCategory === 'Dev Countable');
                releaseData.forEach(d => {{
                    if (!devCountableByMySp[d.myspRca]) {{
                        devCountableByMySp[d.myspRca] = {{}};
                        releases.forEach(r => devCountableByMySp[d.myspRca][r] = 0);
                    }}
                    devCountableByMySp[d.myspRca][release] = (devCountableByMySp[d.myspRca][release] || 0) + d.count;
                }});
            }});
            
            const devCountableTable = document.getElementById('devCountableTable');
            if (Object.keys(devCountableByMySp).length === 0) {{
                devCountableTable.innerHTML = '<tr><td colspan="4" class="no-data">No data available</td></tr>';
            }} else {{
                // Sort by total count across all releases
                const sortedMySp = Object.entries(devCountableByMySp)
                    .map(([mysp, releaseCounts]) => {{
                        const total = Object.values(releaseCounts).reduce((sum, val) => sum + val, 0);
                        return [mysp, releaseCounts, total];
                    }})
                    .sort((a, b) => b[2] - a[2]);
                
                const rows = sortedMySp.map(([mysp, releaseCounts]) => {{
                    let row = `<tr><td>${{mysp}}</td>`;
                    releases.forEach(release => {{
                        const count = releaseCounts[release] || 0;
                        // Calculate total defects for this release (both Dev Countable and Dev NOT Countable)
                        const releaseTotal = baseFilteredData
                            .filter(d => d.release === release)
                            .reduce((sum, d) => sum + d.count, 0);
                        const pct = releaseTotal > 0 ? ((count / releaseTotal) * 100).toFixed(1) : 0;
                        row += `<td>${{count}} (${{pct}}%)</td>`;
                    }});
                    row += '</tr>';
                    return row;
                }}).join('');
                devCountableTable.innerHTML = rows;
            }}
            
            // Dev NOT Countable Summary - group by mySP RCA across releases
            const devNotCountableByMySp = {{}};
            releases.forEach(release => {{
                const releaseData = baseFilteredData.filter(d => d.release === release && d.rcaCategory === 'Dev NOT Countable');
                releaseData.forEach(d => {{
                    if (!devNotCountableByMySp[d.myspRca]) {{
                        devNotCountableByMySp[d.myspRca] = {{}};
                        releases.forEach(r => devNotCountableByMySp[d.myspRca][r] = 0);
                    }}
                    devNotCountableByMySp[d.myspRca][release] = (devNotCountableByMySp[d.myspRca][release] || 0) + d.count;
                }});
            }});
            
            const devNotCountableTable = document.getElementById('devNotCountableTable');
            if (Object.keys(devNotCountableByMySp).length === 0) {{
                devNotCountableTable.innerHTML = '<tr><td colspan="4" class="no-data">No data available</td></tr>';
            }} else {{
                // Sort by total count across all releases
                const sortedMySp = Object.entries(devNotCountableByMySp)
                    .map(([mysp, releaseCounts]) => {{
                        const total = Object.values(releaseCounts).reduce((sum, val) => sum + val, 0);
                        return [mysp, releaseCounts, total];
                    }})
                    .sort((a, b) => b[2] - a[2]);
                
                const rows = sortedMySp.map(([mysp, releaseCounts]) => {{
                    let row = `<tr><td>${{mysp}}</td>`;
                    releases.forEach(release => {{
                        const count = releaseCounts[release] || 0;
                        // Calculate total defects for this release (both Dev Countable and Dev NOT Countable)
                        const releaseTotal = baseFilteredData
                            .filter(d => d.release === release)
                            .reduce((sum, d) => sum + d.count, 0);
                        const pct = releaseTotal > 0 ? ((count / releaseTotal) * 100).toFixed(1) : 0;
                        row += `<td>${{count}} (${{pct}}%)</td>`;
                    }});
                    row += '</tr>';
                    return row;
                }}).join('');
                devNotCountableTable.innerHTML = rows;
            }}
        }}
        
        function loadCategoryComparison() {{
            const categoryContent = document.getElementById('categoryComparisonContent');
            
            // Get Dev Countable data
            const devCountableData = rawData.filter(d => d.rcaCategory === 'Dev Countable');
            
            // Calculate top 5 mySP RCA categories by total count
            const categoryTotals = {{}};
            devCountableData.forEach(d => {{
                if (!categoryTotals[d.myspRca]) categoryTotals[d.myspRca] = 0;
                categoryTotals[d.myspRca] += d.count;
            }});
            
            const top5Categories = Object.entries(categoryTotals)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 5)
                .map(entry => entry[0]);
            
            let html = '';
            
            // For each top category, create a section with top 10 nodes
            top5Categories.forEach((category, index) => {{
                const categoryData = devCountableData.filter(d => d.myspRca === category);
                
                // Calculate node totals across all releases
                const nodeTotals = {{}};
                categoryData.forEach(d => {{
                    if (!nodeTotals[d.nodeName]) nodeTotals[d.nodeName] = 0;
                    nodeTotals[d.nodeName] += d.count;
                }});
                
                const top10Nodes = Object.entries(nodeTotals)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 10)
                    .map(entry => entry[0]);
                
                html += `<div class="category-section">
                    <h3 class="category-title">${{index + 1}}. ${{category}}</h3>
                    <div class="table-wrapper">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>AD POC</th>
                                    <th>SM POC</th>
                                    <th>M POC</th>
                                    <th>Node Name</th>
                                    ${{releases.map(r => '<th>' + r + '</th>').join('')}}
                                    <th>Total</th>
                                </tr>
                            </thead>
                            <tbody>`;
                
                top10Nodes.forEach(node => {{
                    const nodeData = categoryData.filter(d => d.nodeName === node);
                    
                    // Get POC information (assuming all records for this node have same POC values)
                    const firstRecord = nodeData[0];
                    const adPoc = firstRecord.adPoc || '';
                    const smPoc = firstRecord.smPoc || '';
                    const mPoc = firstRecord.mPoc || '';
                    
                    let row = `<tr><td>${{adPoc}}</td><td>${{smPoc}}</td><td>${{mPoc}}</td><td>${{node}}</td>`;
                    let total = 0;
                    
                    releases.forEach(release => {{
                        const releaseCount = nodeData
                            .filter(d => d.release === release)
                            .reduce((sum, d) => sum + d.count, 0);
                        row += `<td>${{releaseCount}}</td>`;
                        total += releaseCount;
                    }});
                    
                    row += `<td><strong>${{total}}</strong></td></tr>`;
                    html += row;
                }});
                
                html += `</tbody>
                        </table>
                    </div>
                </div>`;
            }});
            
            categoryContent.innerHTML = html || '<div class="no-data">No data available</div>';
        }}
        
        // Initialize on page load
        window.addEventListener('DOMContentLoaded', () => {{
            populateAdPocFilter();
            updateSummary();
        }});
    </script>
</body>
</html>"""

# Write output
output_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Dashboard_RCA_Analysis.html'

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\n✅ Dashboard generated successfully!")
print(f"📊 Output file: {output_file}")
print(f"\n📈 Summary:")
print(f"  Total RCAs: {total_rcas}")
print(f"  Valid Defects: {dev_countable} ({dev_pct}%)")
print(f"  Invalid Defects: {dev_not_countable} ({round(100-dev_pct, 1)}%)")
print(f"  Releases covered: {', '.join(releases)}")

