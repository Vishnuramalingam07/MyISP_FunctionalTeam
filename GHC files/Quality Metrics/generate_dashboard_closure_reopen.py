import pandas as pd
import json
from datetime import datetime

# Read the Excel file
excel_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\quality_metrics_input_closure_reopen.xlsx'

# Read all sheets dynamically
excel_data = pd.read_excel(excel_file, sheet_name=None)  # Read all sheets

# Get all sheet names
sheet_names = list(excel_data.keys())
print(f"Found {len(sheet_names)} sheets in Excel file:")
for sheet in sheet_names:
    print(f"  - {sheet}")

# Combine all sheets and add Release column
dataframes = []
for sheet_name in sheet_names:
    df = excel_data[sheet_name]
    df['Release'] = sheet_name  # Use sheet name as release name
    dataframes.append(df)
    print(f"  Loaded {len(df)} records from '{sheet_name}'")

# Combine all dataframes
combined_df = pd.concat(dataframes, ignore_index=True)

print(f"Total records: {len(combined_df)}")
for sheet in sheet_names:
    print(f"{sheet}: {len(dataframes[sheet_names.index(sheet)])} records")
print(f"\nColumn names in Excel:")
for i, col in enumerate(combined_df.columns, 1):
    print(f"  {i}. {col}")

print(f"\nSample SLA values:")
print(combined_df['SLA'].value_counts())
print(f"\nSample Severity values:")
print(combined_df['Severity'].value_counts())

# Get unique values for filters (using dynamic sheet names)
releases = sheet_names  # Use actual sheet names from Excel
ad_pocs = sorted(combined_df['AD POC'].dropna().unique().tolist())
sm_pocs = sorted(combined_df['SM POC'].dropna().unique().tolist())
m_pocs = sorted(combined_df['M POC'].dropna().unique().tolist())

# Create POC hierarchy
poc_hierarchy = []
for _, row in combined_df[['AD POC', 'SM POC', 'M POC', 'Node Name']].drop_duplicates().iterrows():
    poc_hierarchy.append({
        'adPoc': row['AD POC'],
        'smPoc': row['SM POC'],
        'mPoc': row['M POC'],
        'nodeName': row['Node Name']
    })

# Calculate release metrics
release_metrics = []
for release in releases:
    release_data = combined_df[combined_df['Release'] == release]
    total = len(release_data)
    closed = release_data['Closure Trend (Days)'].notna().sum()
    closure_rate = round((closed / total * 100), 1) if total > 0 else 0
    avg_closure = round(release_data['Closure Trend (Days)'].mean(), 1) if closed > 0 else 0
    reopen_rate = round(((release_data['Re open Count'] > 0).sum() / total * 100), 1) if total > 0 else 0
    
    release_metrics.append({
        'release': release,
        'total': int(total),
        'closed': int(closed),
        'closureRate': float(closure_rate),
        'avgClosure': float(avg_closure),
        'reopenRate': float(reopen_rate)
    })

# Calculate detailed metrics by release, POC hierarchy and node
release_metrics_detail = []
for release in releases:
    release_data = combined_df[combined_df['Release'] == release]
    for _, poc_row in release_data[['AD POC', 'SM POC', 'M POC', 'Node Name']].drop_duplicates().iterrows():
        ad_poc = poc_row['AD POC']
        sm_poc = poc_row['SM POC']
        m_poc = poc_row['M POC']
        node_name = poc_row['Node Name']
        
        # Filter to this exact combination
        node_data = release_data[
            (release_data['AD POC'] == ad_poc) &
            (release_data['SM POC'] == sm_poc) &
            (release_data['M POC'] == m_poc) &
            (release_data['Node Name'] == node_name)
        ]
        
        total = len(node_data)
        if total == 0:
            continue
            
        closed = node_data['Closure Trend (Days)'].notna().sum()
        closure_rate = round((closed / total * 100), 1) if total > 0 else 0
        avg_closure = round(node_data['Closure Trend (Days)'].mean(), 1) if closed > 0 else 0
        
        reopened = (node_data['Re open Count'] > 0).sum()
        total_reopens = int(node_data['Re open Count'].sum())
        reopen_rate = round((reopened / total * 100), 1) if total > 0 else 0
        
        met_sla = int((node_data['SLA'] == 'Met SLA').sum()) if 'SLA' in node_data.columns else 0
        not_met_sla = int((node_data['SLA'] == 'Not Met SLA').sum()) if 'SLA' in node_data.columns else 0
        
        # Closure time buckets
        closed_data = node_data[node_data['Closure Trend (Days)'].notna()]
        closed_0to3 = int((closed_data['Closure Trend (Days)'] <= 3).sum())
        closed_4to10 = int(((closed_data['Closure Trend (Days)'] > 3) & (closed_data['Closure Trend (Days)'] <= 10)).sum())
        closed_greater10 = int((closed_data['Closure Trend (Days)'] > 10).sum())
        
        backlog = total - closed
        
        release_metrics_detail.append({
            'release': release,
            'adPoc': ad_poc,
            'smPoc': sm_poc,
            'mPoc': m_poc,
            'nodeName': node_name,
            'total': int(total),
            'closed': int(closed),
            'closureRate': float(closure_rate),
            'avgClosure': float(avg_closure),
            'reopened': int(reopened),
            'totalReopens': int(total_reopens),
            'reopenRate': float(reopen_rate),
            'metSLA': int(met_sla),
            'notMetSLA': int(not_met_sla),
            'closed0to3': int(closed_0to3),
            'closed4to10': int(closed_4to10),
            'closedGreater10': int(closed_greater10),
            'backlog': int(backlog)
        })

# Calculate closure by release (for table)
closure_by_release = []
for release in releases:
    release_data = combined_df[combined_df['Release'] == release]
    total = len(release_data)
    closed = release_data['Closure Trend (Days)'].notna().sum()
    closure_rate = round((closed / total * 100), 1) if total > 0 else 0
    avg_closure = round(release_data['Closure Trend (Days)'].mean(), 1) if closed > 0 else 0
    
    closed_data = release_data[release_data['Closure Trend (Days)'].notna()]
    closed_0to3 = int((closed_data['Closure Trend (Days)'] <= 3).sum())
    closed_4to10 = int(((closed_data['Closure Trend (Days)'] > 3) & (closed_data['Closure Trend (Days)'] <= 10)).sum())
    closed_greater10 = int((closed_data['Closure Trend (Days)'] > 10).sum())
    backlog = total - closed
    
    closure_by_release.append({
        'release': release,
        'total': int(total),
        'closed': int(closed),
        'closureRate': float(closure_rate),
        'avgClosure': float(avg_closure),
        'closed0to3': int(closed_0to3),
        'closed4to10': int(closed_4to10),
        'closedGreater10': int(closed_greater10),
        'backlog': int(backlog)
    })

# Calculate reopen by release
reopen_by_release = []
for release in releases:
    release_data = combined_df[combined_df['Release'] == release]
    for _, poc_row in release_data[['AD POC', 'SM POC', 'M POC', 'Node Name']].drop_duplicates().iterrows():
        ad_poc = poc_row['AD POC']
        sm_poc = poc_row['SM POC']
        m_poc = poc_row['M POC']
        node_name = poc_row['Node Name']
        
        node_data = release_data[
            (release_data['AD POC'] == ad_poc) &
            (release_data['SM POC'] == sm_poc) &
            (release_data['M POC'] == m_poc) &
            (release_data['Node Name'] == node_name)
        ]
        
        total = len(node_data)
        if total == 0:
            continue
            
        reopened_once = int((node_data['Re open Count'] == 1).sum())
        reopened_twice = int((node_data['Re open Count'] == 2).sum())
        reopened_more = int((node_data['Re open Count'] > 2).sum())
        total_reopened = reopened_once + reopened_twice + reopened_more
        reopen_rate = round((total_reopened / total * 100), 1) if total > 0 else 0
        total_reopens = int(node_data['Re open Count'].sum())
        
        reopen_by_release.append({
            'release': release,
            'adPoc': ad_poc,
            'smPoc': sm_poc,
            'mPoc': m_poc,
            'nodeName': node_name,
            'total': int(total),
            'reopenedOnce': int(reopened_once),
            'reopenedTwice': int(reopened_twice),
            'reopenedMore': int(reopened_more),
            'totalReopened': int(total_reopened),
            'reopenRate': float(reopen_rate),
            'totalReopens': int(total_reopens)
        })

# Calculate SM POC reopen stats
sm_poc_reopen = []
for ad_poc in ad_pocs:
    ad_data = combined_df[combined_df['AD POC'] == ad_poc]
    for sm_poc in ad_data['SM POC'].dropna().unique():
        sm_data = ad_data[ad_data['SM POC'] == sm_poc]
        total = len(sm_data)
        reopened = (sm_data['Re open Count'] > 0).sum()
        reopen_rate = round((reopened / total * 100), 1) if total > 0 else 0
        total_reopens = int(sm_data['Re open Count'].sum())
        
        sm_poc_reopen.append({
            'adPoc': ad_poc,
            'smPoc': sm_poc,
            'total': int(total),
            'reopened': int(reopened),
            'reopenRate': float(reopen_rate),
            'totalReopens': int(total_reopens)
        })

# Calculate M POC detail
m_poc_detail = []
for ad_poc in ad_pocs:
    ad_data = combined_df[combined_df['AD POC'] == ad_poc]
    for sm_poc in ad_data['SM POC'].dropna().unique():
        sm_data = ad_data[ad_data['SM POC'] == sm_poc]
        for m_poc in sm_data['M POC'].dropna().unique():
            m_data = sm_data[sm_data['M POC'] == m_poc]
            for node_name in m_data['Node Name'].dropna().unique():
                node_data = m_data[m_data['Node Name'] == node_name]
                
                total = len(node_data)
                closed = node_data['Closure Trend (Days)'].notna().sum()
                avg_closure = round(node_data[node_data['Closure Trend (Days)'].notna()]['Closure Trend (Days)'].mean(), 1) if closed > 0 else 0
                reopened = (node_data['Re open Count'] > 0).sum()
                
                m_poc_detail.append({
                    'adPoc': ad_poc,
                    'smPoc': sm_poc,
                    'mPoc': m_poc,
                    'nodeName': node_name,
                    'total': int(total),
                    'closed': int(closed),
                    'avgClosure': float(avg_closure),
                    'reopened': int(reopened)
                })

# Calculate M POC SLA detail by severity and release
m_poc_sla_detail = []
severity_mapping = {
    '1 - Critical': 'Critical',
    '2 - High': 'High',
    '3 - Medium': 'Medium',
    '4 - Low': 'Low'
}
severities = ['1 - Critical', '2 - High', '3 - Medium', '4 - Low']

for release in releases:
    release_data = combined_df[combined_df['Release'] == release]
    for ad_poc in ad_pocs:
        ad_data = release_data[release_data['AD POC'] == ad_poc]
        for sm_poc in ad_data['SM POC'].dropna().unique():
            sm_data = ad_data[ad_data['SM POC'] == sm_poc]
            for m_poc in sm_data['M POC'].dropna().unique():
                m_data = sm_data[sm_data['M POC'] == m_poc]
                
                # Calculate SLA stats by severity
                sla_stats = {}
                for severity in severities:
                    severity_data = m_data[m_data['Severity'] == severity]
                    met_sla = int((severity_data['SLA'] == 'Met SLA').sum())
                    not_met_sla = int((severity_data['SLA'] == 'Not Met SLA').sum())
                    clean_name = severity_mapping[severity].lower()
                    sla_stats[clean_name] = {
                        'met': met_sla,
                        'notMet': not_met_sla,
                        'total': met_sla + not_met_sla
                    }
                
                # Only add if there's data
                total_defects = sum([sla_stats[severity_mapping[s].lower()]['total'] for s in severities])
                if total_defects > 0:
                    m_poc_sla_detail.append({
                        'release': release,
                        'adPoc': ad_poc,
                        'smPoc': sm_poc,
                        'mPoc': m_poc,
                        'critical': sla_stats['critical'],
                        'high': sla_stats['high'],
                        'medium': sla_stats['medium'],
                        'low': sla_stats['low'],
                        'totalDefects': int(total_defects)
                    })

# Get top reopen nodes with release-wise details
top_reopen_nodes_detailed = []
for _, node_group in combined_df.groupby('Node Name'):
    node_name = node_group['Node Name'].iloc[0]
    ad_poc = node_group['AD POC'].iloc[0]
    sm_poc = node_group['SM POC'].iloc[0]
    m_poc = node_group['M POC'].iloc[0]
    
    # Calculate reopens for each release dynamically
    release_reopens = {}
    total_reopens = 0
    for release in releases:
        reopens = int(node_group[node_group['Release'] == release]['Re open Count'].sum())
        release_reopens[release] = reopens
        total_reopens += reopens
    
    if total_reopens > 0:
        node_entry = {
            'node': node_name,
            'adPoc': ad_poc,
            'smPoc': sm_poc,
            'mPoc': m_poc,
            'total': total_reopens
        }
        # Add each release as a separate key
        for i, release in enumerate(releases):
            node_entry[f'release{i}'] = release_reopens[release]
        top_reopen_nodes_detailed.append(node_entry)

# Sort by total reopens descending and take top 10
top_reopen_nodes_detailed.sort(key=lambda x: x['total'], reverse=True)
top_reopen_nodes_detailed = top_reopen_nodes_detailed[:10]

# Get top delayed nodes (closure > 3 days) with release-wise details
top_delayed_nodes_detailed = []
for _, node_group in combined_df.groupby('Node Name'):
    node_name = node_group['Node Name'].iloc[0]
    ad_poc = node_group['AD POC'].iloc[0]
    sm_poc = node_group['SM POC'].iloc[0]
    m_poc = node_group['M POC'].iloc[0]
    
    # Count bugs closed > 3 days for each release dynamically
    release_delayed = {}
    total_delayed = 0
    for release in releases:
        release_data = node_group[node_group['Release'] == release]
        delayed = int((release_data['Closure Trend (Days)'] > 3).sum())
        release_delayed[release] = delayed
        total_delayed += delayed
    
    if total_delayed > 0:
        node_entry = {
            'node': node_name,
            'adPoc': ad_poc,
            'smPoc': sm_poc,
            'mPoc': m_poc,
            'total': total_delayed
        }
        # Add each release as a separate key
        for i, release in enumerate(releases):
            node_entry[f'release{i}'] = release_delayed[release]
        top_delayed_nodes_detailed.append(node_entry)

# Sort by total delayed descending and take top 10
top_delayed_nodes_detailed.sort(key=lambda x: x['total'], reverse=True)
top_delayed_nodes_detailed = top_delayed_nodes_detailed[:10]

print(f"\n📊 Data processing completed:")
print(f"  - Release metrics: {len(release_metrics)}")
print(f"  - Detailed metrics: {len(release_metrics_detail)}")
print(f"  - Reopen data: {len(reopen_by_release)}")
print(f"  - SM POC stats: {len(sm_poc_reopen)}")
print(f"  - M POC details: {len(m_poc_detail)}")
print(f"  - M POC SLA details: {len(m_poc_sla_detail)}")

# Generate HTML
output_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Dashboard_Closure_Reopen_Analysis.html'

html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Closure & Reopen Analysis Dashboard</title>

    <style>
        * {{margin:0;padding:0;box-sizing:border-box}}
        body {{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:10px;min-height:100vh}}
        .container {{max-width:1800px;margin:0 auto;background:white;border-radius:15px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}}
        .header {{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:15px;text-align:center}}
        .header h1 {{font-size:2.2em;margin-bottom:5px;text-shadow:2px 2px 4px rgba(0,0,0,0.2)}}
        .header p {{font-size:1.1em;opacity:0.9}}
        .content {{padding:20px}}
        
        .metrics-table-section {{margin:10px 0;padding:15px;background:#f8f9fa;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}}
        .metrics-table-title {{font-size:1.3em;color:#333;margin-bottom:12px;font-weight:600}}
        .metrics-table {{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);border:2px solid #667eea}}
        .metrics-table th {{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:12px;text-align:left;font-weight:600;font-size:0.85em;border:1px solid rgba(255,255,255,0.3)}}
        .metrics-table td {{padding:12px;border:1px solid #d0d0d0;font-size:0.85em}}
        .metrics-table tr:last-child td {{border-bottom:1px solid #d0d0d0}}
        .metrics-table tr:hover {{background:#f5f5f5}}
        .metrics-table td:first-child {{font-weight:600;color:#667eea}}
        
        .filter-section {{margin-bottom:15px;padding:15px;background:#f8f9fa;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}}
        .filter-title {{font-size:1.1em;color:#333;margin-bottom:10px;font-weight:600}}
        .filters {{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}}
        .filter-group {{display:flex;flex-direction:column}}
        .filter-group label {{font-weight:600;color:#555;margin-bottom:4px;font-size:0.8em}}
        .filter-group select {{padding:6px 8px;border:2px solid #667eea;border-radius:6px;font-size:0.85em;background:white;cursor:pointer;transition:all 0.3s}}
        .filter-group select:hover {{border-color:#764ba2;box-shadow:0 2px 8px rgba(102,126,234,0.2)}}
        
        .charts-section {{margin:15px 0}}
        .section-title {{font-size:1.5em;color:#333;margin:15px 0 10px 0;padding-bottom:8px;border-bottom:3px solid #667eea}}
        .chart-grid {{display:grid;grid-template-columns:repeat(auto-fit,minmax(500px,1fr));gap:15px;margin:15px 0}}
        .chart-container {{background:white;padding:15px;border-radius:12px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border:1px solid #e0e0e0}}
        .chart-title {{font-size:1.2em;color:#333;margin-bottom:12px;font-weight:600;text-align:center}}
        canvas {{max-height:400px}}
        
        .table-wrapper {{margin:10px 0;border:3px solid #667eea;border-radius:8px;overflow:hidden;overflow-x:auto;box-shadow:0 4px 12px rgba(0,0,0,0.15)}}
        .data-table {{width:100%;border-collapse:collapse;font-size:0.9em;min-width:800px}}
        .data-table th {{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:12px;text-align:left;font-weight:600;position:sticky;top:0;z-index:10;border:1px solid rgba(255,255,255,0.3)}}
        .data-table td {{padding:12px;border:1px solid #d0d0d0}}
        .data-table tr:hover {{background:#f5f5f5}}
        .data-table tr:nth-child(even) {{background:#f9f9f9}}
        .data-table tr:nth-child(even):hover {{background:#f0f0f0}}
        
        .tabs {{display:flex;border-bottom:3px solid #667eea;margin-bottom:15px;background:#f8f9fa;border-radius:10px 10px 0 0}}
        .tab {{flex:1;padding:15px 20px;text-align:center;cursor:pointer;font-weight:600;font-size:1.1em;color:#555;transition:all 0.3s;border-bottom:3px solid transparent;margin-bottom:-3px}}
        .tab:hover {{background:#e9ecef}}
        .tab.active {{color:#667eea;border-bottom-color:#667eea;background:white}}
        .tab-content {{display:none}}
        .tab-content.active {{display:block}}
        
        .sub-tabs {{display:flex;border-bottom:2px solid #764ba2;margin:10px 0;background:#fafafa;border-radius:8px 8px 0 0}}
        .sub-tab {{flex:1;padding:12px 18px;text-align:center;cursor:pointer;font-weight:500;font-size:1em;color:#666;transition:all 0.3s;border-bottom:2px solid transparent;margin-bottom:-2px}}
        .sub-tab:hover {{background:#f0f0f0}}
        .sub-tab.active {{color:#764ba2;border-bottom-color:#764ba2;background:white}}
        .sub-tab-content {{display:none}}
        .sub-tab-content.active {{display:block}}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Defect Closure & Reopen Analysis Dashboard</h1>
            <p>Track defect closure efficiency and reopen trends across releases</p>
        </div>
        
        <div class="content">
            <!-- Main Tabs -->
            <div class="tabs">
                <div class="tab active" onclick="switchMainTab('detailedTab')">Detailed Analysis By Node Name</div>
                <div class="tab" onclick="switchMainTab('comparativeTab')">Comparative Analysis</div>
            </div>
            
            <!-- Main Tab 1: Detailed Analysis -->
            <div id="detailedTab" class="tab-content active">
                <!-- Filters -->
                <div class="filter-section">
                    <div class="filter-title">🔍 Filters</div>
                    <div class="filters">
                        <div class="filter-group">
                            <label for="adPocFilter">AD POC</label>
                            <select id="adPocFilter" onchange="onAdPocChange()">
                                <option value="">Select AD POC</option>
                            </select>
                        </div>
                        <div class="filter-group">
                            <label for="smPocFilter">SM POC</label>
                            <select id="smPocFilter" onchange="onSmPocChange()" disabled>
                                <option value="">Select SM POC</option>
                            </select>
                        </div>
                        <div class="filter-group">
                            <label for="mPocFilter">M POC</label>
                            <select id="mPocFilter" onchange="onMPocChange()" disabled>
                                <option value="">Select M POC</option>
                            </select>
                        </div>
                        <div class="filter-group">
                            <label for="nodeNameFilter">Node Name</label>
                            <select id="nodeNameFilter" onchange="onNodeNameChange()" disabled>
                                <option value="">Select Node Name</option>
                            </select>
                        </div>
                    </div>
                </div>
                
                <!-- Metrics Table -->
                <div class="metrics-table-section">
                    <div class="metrics-table-title">📊 Release Metrics Summary</div>
                    <table class="metrics-table" id="releaseMetricsTable">
                        <thead>
                            <tr>
                                <th>Release</th>
                                <th>Closed Defects</th>
                                <th>Met SLA</th>
                                <th>Not Met SLA</th>
                                <th>Avg Closure Time (Days)</th>
                                <th>Reopened Defects</th>
                                <th>Total Reopens</th>
                                <th>Reopen Rate (%)</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>
                
                <!-- Sub Tabs -->
                <div class="sub-tabs">
                    <div class="sub-tab active" onclick="switchSubTab('closureSubTab')">Closure Analysis</div>
                    <div class="sub-tab" onclick="switchSubTab('reopenSubTab')">Reopen Analysis</div>
                </div>
                
                <!-- Sub Tab 1: Closure Analysis -->
                <div id="closureSubTab" class="sub-tab-content active">
                    <h2 class="section-title">Closure Performance by Release</h2>
                    <div class="table-wrapper">
                        <table class="data-table" id="closureByReleaseTable">
                            <thead>
                                <tr>
                                    <th>Metric</th>
                                    {''.join([f'<th>{r}</th>' for r in releases])}
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    
                    <h2 class="section-title" style="margin-top: 40px;">M POC - Met SLA by Severity & Release</h2>
                    <div class="table-wrapper">
                        <table class="data-table" id="mPocMetSlaTable">
                            <thead>
                                <tr>
                                    <th>M POC</th>
                                    <th>Release</th>
                                    <th style="background: #e74c3c;">Critical</th>
                                    <th style="background: #e67e22;">High</th>
                                    <th style="background: #f39c12;">Medium</th>
                                    <th style="background: #3498db;">Low</th>
                                    <th style="background: #4CAF50;">Total Met SLA</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    
                    <h2 class="section-title" style="margin-top: 40px;">M POC - Not Met SLA by Severity & Release</h2>
                    <div class="table-wrapper">
                        <table class="data-table" id="mPocNotMetSlaTable">
                            <thead>
                                <tr>
                                    <th>M POC</th>
                                    <th>Release</th>
                                    <th style="background: #e74c3c;">Critical</th>
                                    <th style="background: #e67e22;">High</th>
                                    <th style="background: #f39c12;">Medium</th>
                                    <th style="background: #3498db;">Low</th>
                                    <th style="background: #f44336;">Total Not Met SLA</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Sub Tab 2: Reopen Analysis -->
                <div id="reopenSubTab" class="sub-tab-content">
                    <h2 class="section-title">Reopen Performance by Release</h2>
                    <div class="table-wrapper">
                        <table class="data-table" id="reopenByReleaseTable">
                            <thead>
                                <tr>
                                    <th>Metric</th>
                                    {''.join([f'<th>{r}</th>' for r in releases])}
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <!-- Main Tab 2: Comparative Analysis -->
            <div id="comparativeTab" class="tab-content">
                <h2 class="section-title">Nodes with Most Reopened Bugs</h2>
                <div class="table-wrapper">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Rank</th>
                                <th>Node Name</th>
                                <th>AD POC</th>
                                <th>SM POC</th>
                                <th>M POC</th>
                                {''.join([f'<th>{r}</th>' for r in releases])}
                                <th>Total Reopens</th>
                            </tr>
                        </thead>
                        <tbody id="topReopenNodesTbody"></tbody>
                    </table>
                </div>
                
                <h2 class="section-title" style="margin-top: 40px;">Nodes with Delayed Closure of Bugs</h2>
                <div class="table-wrapper">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Rank</th>
                                <th>Node Name</th>
                                <th>AD POC</th>
                                <th>SM POC</th>
                                <th>M POC</th>
                                {''.join([f'<th>{r}</th>' for r in releases])}
                                <th>Total</th>
                            </tr>
                        </thead>
                        <tbody id="topDelayedNodesTbody"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
        var releases = {json.dumps(releases)};
        var adPocs = {json.dumps(ad_pocs)};
        var smPocs = {json.dumps(sm_pocs)};
        var mPocs = {json.dumps(m_pocs)};
        var pocHierarchy = {json.dumps(poc_hierarchy)};
        var releaseMetrics = {json.dumps(release_metrics)};
        var releaseMetricsDetail = {json.dumps(release_metrics_detail)};
        var closureByRelease = {json.dumps(closure_by_release)};
        var reopenByRelease = {json.dumps(reopen_by_release)};
        var smPocReopen = {json.dumps(sm_poc_reopen)};
        var mPocDetail = {json.dumps(m_poc_detail)};
        var mPocSlaDetail = {json.dumps(m_poc_sla_detail)};
        var topReopenNodesDetailed = {json.dumps(top_reopen_nodes_detailed)};
        var topDelayedNodesDetailed = {json.dumps(top_delayed_nodes_detailed)};
        
        // Main tab switching
        function switchMainTab(tabId) {{
            var tabs = document.querySelectorAll('.tab-content');
            var tabButtons = document.querySelectorAll('.tabs > .tab');
            
            for (var i = 0; i < tabs.length; i++) {{
                tabs[i].classList.remove('active');
            }}
            for (var i = 0; i < tabButtons.length; i++) {{
                tabButtons[i].classList.remove('active');
            }}
            
            document.getElementById(tabId).classList.add('active');
            event.target.classList.add('active');
        }}
        
        // Sub-tab switching within Detailed Analysis
        function switchSubTab(subTabId) {{
            var tabs = document.querySelectorAll('.sub-tab-content');
            var tabButtons = document.querySelectorAll('.sub-tabs > .sub-tab');
            
            for (var i = 0; i < tabs.length; i++) {{
                tabs[i].classList.remove('active');
            }}
            for (var i = 0; i < tabButtons.length; i++) {{
                tabButtons[i].classList.remove('active');
            }}
            
            document.getElementById(subTabId).classList.add('active');
            event.target.classList.add('active');
        }}
        
        // Populate AD POC filter
        function populateAdPocFilter() {{
            var select = document.getElementById('adPocFilter');
            select.innerHTML = '<option value="">Select AD POC</option>';
            for (var i = 0; i < adPocs.length; i++) {{
                var option = document.createElement('option');
                option.value = adPocs[i];
                option.text = adPocs[i];
                select.appendChild(option);
            }}
        }}
        
        // AD POC change handler
        function onAdPocChange() {{
            var adPoc = document.getElementById('adPocFilter').value;
            var smSelect = document.getElementById('smPocFilter');
            var mSelect = document.getElementById('mPocFilter');
            var nodeSelect = document.getElementById('nodeNameFilter');
            
            // Reset dependent dropdowns
            smSelect.innerHTML = '<option value="">Select SM POC</option>';
            mSelect.innerHTML = '<option value="">Select M POC</option>';
            nodeSelect.innerHTML = '<option value="">Select Node Name</option>';
            
            if (adPoc) {{
                // Get unique SM POCs for this AD POC
                var smPocList = [];
                for (var i = 0; i < pocHierarchy.length; i++) {{
                    if (pocHierarchy[i].adPoc === adPoc && smPocList.indexOf(pocHierarchy[i].smPoc) === -1) {{
                        smPocList.push(pocHierarchy[i].smPoc);
                    }}
                }}
                smPocList.sort();
                for (var i = 0; i < smPocList.length; i++) {{
                    var option = document.createElement('option');
                    option.value = smPocList[i];
                    option.text = smPocList[i];
                    smSelect.appendChild(option);
                }}
                smSelect.disabled = false;
            }} else {{
                smSelect.disabled = true;
                mSelect.disabled = true;
                nodeSelect.disabled = true;
            }}
            
            applyFilters();
        }}
        
        // SM POC change handler
        function onSmPocChange() {{
            var adPoc = document.getElementById('adPocFilter').value;
            var smPoc = document.getElementById('smPocFilter').value;
            var mSelect = document.getElementById('mPocFilter');
            var nodeSelect = document.getElementById('nodeNameFilter');
            
            mSelect.innerHTML = '<option value="">Select M POC</option>';
            nodeSelect.innerHTML = '<option value="">Select Node Name</option>';
            
            if (smPoc) {{
                var mPocList = [];
                for (var i = 0; i < pocHierarchy.length; i++) {{
                    if (pocHierarchy[i].adPoc === adPoc && pocHierarchy[i].smPoc === smPoc && mPocList.indexOf(pocHierarchy[i].mPoc) === -1) {{
                        mPocList.push(pocHierarchy[i].mPoc);
                    }}
                }}
                mPocList.sort();
                for (var i = 0; i < mPocList.length; i++) {{
                    var option = document.createElement('option');
                    option.value = mPocList[i];
                    option.text = mPocList[i];
                    mSelect.appendChild(option);
                }}
                mSelect.disabled = false;
            }} else {{
                mSelect.disabled = true;
                nodeSelect.disabled = true;
            }}
            
            applyFilters();
        }}
        
        // M POC change handler
        function onMPocChange() {{
            var adPoc = document.getElementById('adPocFilter').value;
            var smPoc = document.getElementById('smPocFilter').value;
            var mPoc = document.getElementById('mPocFilter').value;
            var nodeSelect = document.getElementById('nodeNameFilter');
            
            nodeSelect.innerHTML = '<option value="">Select Node Name</option>';
            
            if (mPoc) {{
                var nodeList = [];
                for (var i = 0; i < pocHierarchy.length; i++) {{
                    if (pocHierarchy[i].adPoc === adPoc && pocHierarchy[i].smPoc === smPoc && pocHierarchy[i].mPoc === mPoc) {{
                        nodeList.push(pocHierarchy[i].nodeName);
                    }}
                }}
                nodeList.sort();
                for (var i = 0; i < nodeList.length; i++) {{
                    var option = document.createElement('option');
                    option.value = nodeList[i];
                    option.text = nodeList[i];
                    nodeSelect.appendChild(option);
                }}
                nodeSelect.disabled = false;
            }} else {{
                nodeSelect.disabled = true;
            }}
            
            applyFilters();
        }}
        
        // Node Name change handler
        function onNodeNameChange() {{
            applyFilters();
        }}
        
        // Apply filters to all tables
        function applyFilters() {{
            var adPoc = document.getElementById('adPocFilter').value;
            var smPoc = document.getElementById('smPocFilter').value;
            var mPoc = document.getElementById('mPocFilter').value;
            var nodeName = document.getElementById('nodeNameFilter').value;
            
            updateReleaseMetricsTable(adPoc, smPoc, mPoc, nodeName);
            updateClosureByReleaseTable(adPoc, smPoc, mPoc, nodeName);
            updateReopenByReleaseTable(adPoc, smPoc, mPoc, nodeName);
            updateMPocMetSlaTable(adPoc, smPoc, mPoc, nodeName);
            updateMPocNotMetSlaTable(adPoc, smPoc, mPoc, nodeName);
        }}
        
        // Update Release Metrics Table
        function updateReleaseMetricsTable(adPoc, smPoc, mPoc, nodeName) {{
            var tbody = document.querySelector('#releaseMetricsTable tbody');
            var filtered = releaseMetricsDetail;
            
            if (adPoc) filtered = filtered.filter(function(d) {{ return d.adPoc === adPoc; }});
            if (smPoc) filtered = filtered.filter(function(d) {{ return d.smPoc === smPoc; }});
            if (mPoc) filtered = filtered.filter(function(d) {{ return d.mPoc === mPoc; }});
            if (nodeName) filtered = filtered.filter(function(d) {{ return d.nodeName === nodeName; }});
            
            var releaseStats = {{}};
            for (var i = 0; i < releases.length; i++) {{
                releaseStats[releases[i]] = {{
                    closed: 0,
                    metSLA: 0,
                    notMetSLA: 0,
                    avgClosure: [],
                    reopened: 0,
                    totalReopens: 0,
                    total: 0
                }};
            }}
            
            for (var i = 0; i < filtered.length; i++) {{
                var d = filtered[i];
                releaseStats[d.release].closed += d.closed;
                releaseStats[d.release].metSLA += d.metSLA;
                releaseStats[d.release].notMetSLA += d.notMetSLA;
                if (d.closed > 0) {{
                    for (var j = 0; j < d.closed; j++) {{
                        releaseStats[d.release].avgClosure.push(d.avgClosure);
                    }}
                }}
                releaseStats[d.release].reopened += d.reopened;
                releaseStats[d.release].totalReopens += d.totalReopens;
                releaseStats[d.release].total += d.total;
            }}
            
            var html = '';
            for (var i = 0; i < releases.length; i++) {{
                var release = releases[i];
                var stats = releaseStats[release];
                var avgClosure = stats.avgClosure.length > 0 ? 
                    (stats.avgClosure.reduce(function(a,b){{return a+b}}, 0) / stats.avgClosure.length).toFixed(1) : '0.0';
                var reopenRate = stats.total > 0 ? ((stats.reopened / stats.total * 100).toFixed(1)) : '0.0';
                
                var metSLAPct = stats.closed > 0 ? ((stats.metSLA / stats.closed * 100).toFixed(1)) : '0.0';
                var notMetSLAPct = stats.closed > 0 ? ((stats.notMetSLA / stats.closed * 100).toFixed(1)) : '0.0';
                
                html += '<tr>';
                html += '<td>' + release + '</td>';
                html += '<td>' + stats.closed + '</td>';
                html += '<td>' + stats.metSLA + ' (' + metSLAPct + '%)</td>';
                html += '<td>' + stats.notMetSLA + ' (' + notMetSLAPct + '%)</td>';
                html += '<td>' + avgClosure + '</td>';
                html += '<td>' + stats.reopened + '</td>';
                html += '<td>' + stats.totalReopens + '</td>';
                html += '<td>' + reopenRate + '%</td>';
                html += '</tr>';
            }}
            
            tbody.innerHTML = html;
        }}
        
        // Update Closure by Release Table
        function updateClosureByReleaseTable(adPoc, smPoc, mPoc, nodeName) {{
            var tbody = document.querySelector('#closureByReleaseTable tbody');
            var filtered = releaseMetricsDetail;
            
            if (adPoc) filtered = filtered.filter(function(d) {{ return d.adPoc === adPoc; }});
            if (smPoc) filtered = filtered.filter(function(d) {{ return d.smPoc === smPoc; }});
            if (mPoc) filtered = filtered.filter(function(d) {{ return d.mPoc === mPoc; }});
            if (nodeName) filtered = filtered.filter(function(d) {{ return d.nodeName === nodeName; }});
            
            var releaseStats = {{}};
            for (var i = 0; i < releases.length; i++) {{
                releaseStats[releases[i]] = {{
                    closed: 0,
                    avgClosureDays: [],
                    closed0to3: 0,
                    closed4to10: 0,
                    closedGreater10: 0
                }};
            }}
            
            for (var i = 0; i < filtered.length; i++) {{
                var d = filtered[i];
                releaseStats[d.release].closed += d.closed;
                if (d.closed > 0) {{
                    for (var j = 0; j < d.closed; j++) {{
                        releaseStats[d.release].avgClosureDays.push(d.avgClosure);
                    }}
                }}
                releaseStats[d.release].closed0to3 += d.closed0to3;
                releaseStats[d.release].closed4to10 += d.closed4to10;
                releaseStats[d.release].closedGreater10 += d.closedGreater10;
            }}
            
            var metrics = [
                'Closed Defects',
                'Avg Closure Time (Days)',
                'Defects closed within 0 to 3 days',
                'Defects closed within 4 to 10 days',
                'Defects closed greater than 10 days'
            ];
            
            var html = '';
            for (var i = 0; i < metrics.length; i++) {{
                var metric = metrics[i];
                html += '<tr><td><strong>' + metric + '</strong></td>';
                
                for (var j = 0; j < releases.length; j++) {{
                    var stats = releaseStats[releases[j]];
                    var value;
                    
                    if (metric === 'Closed Defects') {{
                        value = stats.closed;
                    }}
                    else if (metric === 'Avg Closure Time (Days)') {{
                        var avgClosure = stats.avgClosureDays.length > 0 ? 
                            (stats.avgClosureDays.reduce(function(a,b){{return a+b}}, 0) / stats.avgClosureDays.length).toFixed(1) : '0.0';
                        value = avgClosure;
                    }}
                    else if (metric === 'Defects closed within 0 to 3 days') {{
                        var pct = stats.closed > 0 ? ((stats.closed0to3 / stats.closed * 100).toFixed(1)) : '0.0';
                        value = stats.closed0to3 + ' (' + pct + '%)';
                    }}
                    else if (metric === 'Defects closed within 4 to 10 days') {{
                        var pct = stats.closed > 0 ? ((stats.closed4to10 / stats.closed * 100).toFixed(1)) : '0.0';
                        value = stats.closed4to10 + ' (' + pct + '%)';
                    }}
                    else if (metric === 'Defects closed greater than 10 days') {{
                        var pct = stats.closed > 0 ? ((stats.closedGreater10 / stats.closed * 100).toFixed(1)) : '0.0';
                        value = stats.closedGreater10 + ' (' + pct + '%)';
                    }}
                    
                    html += '<td>' + value + '</td>';
                }}
                
                html += '</tr>';
            }}
            
            tbody.innerHTML = html;
        }}
        
        // Update Reopen by Release Table
        function updateReopenByReleaseTable(adPoc, smPoc, mPoc, nodeName) {{
            var tbody = document.querySelector('#reopenByReleaseTable tbody');
            var filtered = reopenByRelease;
            
            if (adPoc) filtered = filtered.filter(function(d) {{ return d.adPoc === adPoc; }});
            if (smPoc) filtered = filtered.filter(function(d) {{ return d.smPoc === smPoc; }});
            if (mPoc) filtered = filtered.filter(function(d) {{ return d.mPoc === mPoc; }});
            if (nodeName) filtered = filtered.filter(function(d) {{ return d.nodeName === nodeName; }});
            
            var releaseStats = {{}};
            for (var i = 0; i < releases.length; i++) {{
                releaseStats[releases[i]] = {{
                    total: 0,
                    reopenedOnce: 0,
                    reopenedTwice: 0,
                    reopenedMore: 0,
                    totalReopens: 0
                }};
            }}
            
            for (var i = 0; i < filtered.length; i++) {{
                var d = filtered[i];
                releaseStats[d.release].total += d.total;
                releaseStats[d.release].reopenedOnce += d.reopenedOnce;
                releaseStats[d.release].reopenedTwice += d.reopenedTwice;
                releaseStats[d.release].reopenedMore += d.reopenedMore;
                releaseStats[d.release].totalReopens += d.totalReopens;
            }}
            
            var metrics = [
                'Total Defects',
                'Defects Reopened Once',
                'Defects Reopened twice',
                'Defects Reopened more than twice',
                'Reopen%',
                'Total reopens'
            ];
            
            var html = '';
            for (var i = 0; i < metrics.length; i++) {{
                var metric = metrics[i];
                html += '<tr><td><strong>' + metric + '</strong></td>';
                
                for (var j = 0; j < releases.length; j++) {{
                    var stats = releaseStats[releases[j]];
                    var value;
                    
                    if (metric === 'Total Defects') value = stats.total;
                    else if (metric === 'Defects Reopened Once') value = stats.reopenedOnce;
                    else if (metric === 'Defects Reopened twice') value = stats.reopenedTwice;
                    else if (metric === 'Defects Reopened more than twice') value = stats.reopenedMore;
                    else if (metric === 'Reopen%') {{
                        var totalReopened = stats.reopenedOnce + stats.reopenedTwice + stats.reopenedMore;
                        value = stats.total > 0 ? ((totalReopened / stats.total * 100).toFixed(1) + '%') : '0.0%';
                    }}
                    else if (metric === 'Total reopens') value = stats.totalReopens;
                    
                    html += '<td>' + value + '</td>';
                }}
                
                html += '</tr>';
            }}
            
            tbody.innerHTML = html;
        }}
        
        // Update M POC Detail Table
        function updateMPocDetailTable(adPoc, smPoc, mPoc, nodeName) {{
            var tbody = document.querySelector('#mPocDetailTable tbody');
            var filtered = mPocDetail;
            
            if (adPoc) filtered = filtered.filter(function(d) {{ return d.adPoc === adPoc; }});
            if (smPoc) filtered = filtered.filter(function(d) {{ return d.smPoc === smPoc; }});
            if (mPoc) filtered = filtered.filter(function(d) {{ return d.mPoc === mPoc; }});
            if (nodeName) filtered = filtered.filter(function(d) {{ return d.nodeName === nodeName; }});
            
            // Aggregate by M POC
            var mPocAgg = {{}};
            for (var i = 0; i < filtered.length; i++) {{
                var d = filtered[i];
                if (!mPocAgg[d.mPoc]) {{
                    mPocAgg[d.mPoc] = {{total: 0, closed: 0, closureDays: [], reopened: 0}};
                }}
                mPocAgg[d.mPoc].total += d.total;
                mPocAgg[d.mPoc].closed += d.closed;
                for (var j = 0; j < d.closed; j++) {{
                    mPocAgg[d.mPoc].closureDays.push(d.avgClosure);
                }}
                mPocAgg[d.mPoc].reopened += d.reopened;
            }}
            
            var rows = [];
            for (var poc in mPocAgg) {{
                var stats = mPocAgg[poc];
                var avgClosure = stats.closureDays.length > 0 ?
                    (stats.closureDays.reduce(function(a,b){{return a+b}}, 0) / stats.closureDays.length).toFixed(1) : '0.0';
                rows.push({{
                    poc: poc,
                    total: stats.total,
                    closed: stats.closed,
                    closureRate: (stats.closed / stats.total * 100).toFixed(1),
                    avgClosure: avgClosure,
                    reopened: stats.reopened,
                    reopenRate: (stats.reopened / stats.total * 100).toFixed(1)
                }});
            }}
            
            rows.sort(function(a, b) {{ return b.closureRate - a.closureRate; }});
            
            var html = '';
            if (rows.length === 0) {{
                html = '<tr><td colspan="7" style="text-align:center;">No data available</td></tr>';
            }} else {{
                for (var i = 0; i < rows.length; i++) {{
                    var r = rows[i];
                    html += '<tr>';
                    html += '<td><strong>' + r.poc + '</strong></td>';
                    html += '<td>' + r.total + '</td>';
                    html += '<td>' + r.closed + '</td>';
                    html += '<td>' + r.closureRate + '%</td>';
                    html += '<td>' + r.avgClosure + '</td>';
                    html += '<td>' + r.reopened + '</td>';
                    html += '<td>' + r.reopenRate + '%</td>';
                    html += '</tr>';
                }}
            }}
            
            tbody.innerHTML = html;
        }}
        
        // Update M POC Met SLA Table
        function updateMPocMetSlaTable(adPoc, smPoc, mPoc, nodeName) {{
            var tbody = document.querySelector('#mPocMetSlaTable tbody');
            var filtered = mPocSlaDetail;
            
            if (adPoc) filtered = filtered.filter(function(d) {{ return d.adPoc === adPoc; }});
            if (smPoc) filtered = filtered.filter(function(d) {{ return d.smPoc === smPoc; }});
            if (mPoc) filtered = filtered.filter(function(d) {{ return d.mPoc === mPoc; }});
            
            // Sort by M POC and Release
            filtered.sort(function(a, b) {{
                if (a.mPoc !== b.mPoc) return a.mPoc.localeCompare(b.mPoc);
                var releaseOrder = {{'Oct 11': 1, 'Nov 8': 2, 'Dec 13': 3}};
                return releaseOrder[a.release] - releaseOrder[b.release];
            }});
            
            var html = '';
            if (filtered.length === 0) {{
                html = '<tr><td colspan="7" style="text-align:center;">No data available</td></tr>';
            }} else {{
                var prevMPoc = null;
                var mPocRowspan = {{}};
                
                // Calculate rowspan for each M POC
                for (var i = 0; i < filtered.length; i++) {{
                    var currentMPoc = filtered[i].mPoc;
                    if (!mPocRowspan[currentMPoc]) {{
                        mPocRowspan[currentMPoc] = 0;
                    }}
                    mPocRowspan[currentMPoc]++;
                }}
                
                var mPocFirstRow = {{}};
                for (var i = 0; i < filtered.length; i++) {{
                    var d = filtered[i];
                    var totalMet = d.critical.met + d.high.met + d.medium.met + d.low.met;
                    var isFirstRowForMPoc = (prevMPoc !== d.mPoc);
                    
                    html += '<tr>';
                    
                    // Only add M POC cell for first row of each M POC
                    if (isFirstRowForMPoc) {{
                        html += '<td rowspan="' + mPocRowspan[d.mPoc] + '" style="vertical-align: middle;"><strong>' + d.mPoc + '</strong></td>';
                    }}
                    
                    html += '<td>' + d.release + '</td>';
                    html += '<td style="background: #d5f4e6;">' + d.critical.met + '</td>';
                    html += '<td style="background: #d5f4e6;">' + d.high.met + '</td>';
                    html += '<td style="background: #d5f4e6;">' + d.medium.met + '</td>';
                    html += '<td style="background: #d5f4e6;">' + d.low.met + '</td>';
                    html += '<td style="background: #a8e6cf; font-weight: bold;">' + totalMet + '</td>';
                    html += '</tr>';
                    
                    prevMPoc = d.mPoc;
                }}
            }}
            
            tbody.innerHTML = html;
        }}
        
        // Update M POC Not Met SLA Table
        function updateMPocNotMetSlaTable(adPoc, smPoc, mPoc, nodeName) {{
            var tbody = document.querySelector('#mPocNotMetSlaTable tbody');
            var filtered = mPocSlaDetail;
            
            if (adPoc) filtered = filtered.filter(function(d) {{ return d.adPoc === adPoc; }});
            if (smPoc) filtered = filtered.filter(function(d) {{ return d.smPoc === smPoc; }});
            if (mPoc) filtered = filtered.filter(function(d) {{ return d.mPoc === mPoc; }});
            
            // Sort by M POC and Release
            filtered.sort(function(a, b) {{
                if (a.mPoc !== b.mPoc) return a.mPoc.localeCompare(b.mPoc);
                var releaseOrder = {{'Oct 11': 1, 'Nov 8': 2, 'Dec 13': 3}};
                return releaseOrder[a.release] - releaseOrder[b.release];
            }});
            
            var html = '';
            if (filtered.length === 0) {{
                html = '<tr><td colspan="7" style="text-align:center;">No data available</td></tr>';
            }} else {{
                var prevMPoc = null;
                var mPocRowspan = {{}};
                
                // Calculate rowspan for each M POC
                for (var i = 0; i < filtered.length; i++) {{
                    var currentMPoc = filtered[i].mPoc;
                    if (!mPocRowspan[currentMPoc]) {{
                        mPocRowspan[currentMPoc] = 0;
                    }}
                    mPocRowspan[currentMPoc]++;
                }}
                
                var mPocFirstRow = {{}};
                for (var i = 0; i < filtered.length; i++) {{
                    var d = filtered[i];
                    var totalNotMet = d.critical.notMet + d.high.notMet + d.medium.notMet + d.low.notMet;
                    var isFirstRowForMPoc = (prevMPoc !== d.mPoc);
                    
                    html += '<tr>';
                    
                    // Only add M POC cell for first row of each M POC
                    if (isFirstRowForMPoc) {{
                        html += '<td rowspan="' + mPocRowspan[d.mPoc] + '" style="vertical-align: middle;"><strong>' + d.mPoc + '</strong></td>';
                    }}
                    
                    html += '<td>' + d.release + '</td>';
                    html += '<td style="background: #fadbd8;">' + d.critical.notMet + '</td>';
                    html += '<td style="background: #fadbd8;">' + d.high.notMet + '</td>';
                    html += '<td style="background: #fadbd8;">' + d.medium.notMet + '</td>';
                    html += '<td style="background: #fadbd8;">' + d.low.notMet + '</td>';
                    html += '<td style="background: #ffb3ba; font-weight: bold;">' + totalNotMet + '</td>';
                    html += '</tr>';
                    
                    prevMPoc = d.mPoc;
                }}
            }}
            
            tbody.innerHTML = html;
        }}
        
        // Populate Comparative Analysis
        function populateComparativeTables() {{
            // Top Reopen Nodes
            var tbody = document.getElementById('topReopenNodesTbody');
            var html = '';
            for (var i = 0; i < topReopenNodesDetailed.length; i++) {{
                var n = topReopenNodesDetailed[i];
                html += '<tr>';
                html += '<td><strong>' + (i + 1) + '</strong></td>';
                html += '<td>' + n.node + '</td>';
                html += '<td>' + n.adPoc + '</td>';
                html += '<td>' + n.smPoc + '</td>';
                html += '<td>' + n.mPoc + '</td>';
                // Add release columns dynamically
                for (var r = 0; r < releases.length; r++) {{
                    html += '<td>' + (n['release' + r] || 0) + '</td>';
                }}
                html += '<td><strong>' + n.total + '</strong></td>';
                html += '</tr>';
            }}
            tbody.innerHTML = html;
            
            // Top Delayed Nodes
            tbody = document.getElementById('topDelayedNodesTbody');
            html = '';
            for (var i = 0; i < topDelayedNodesDetailed.length; i++) {{
                var n = topDelayedNodesDetailed[i];
                html += '<tr>';
                html += '<td><strong>' + (i + 1) + '</strong></td>';
                html += '<td>' + n.node + '</td>';
                html += '<td>' + n.adPoc + '</td>';
                html += '<td>' + n.smPoc + '</td>';
                html += '<td>' + n.mPoc + '</td>';
                // Add release columns dynamically
                for (var r = 0; r < releases.length; r++) {{
                    html += '<td>' + (n['release' + r] || 0) + '</td>';
                }}
                html += '<td><strong>' + n.total + '</strong></td>';
                html += '</tr>';
            }}
            tbody.innerHTML = html;
        }}
        
        // Initialize
        window.addEventListener('DOMContentLoaded', function() {{
            populateAdPocFilter();
            updateReleaseMetricsTable('', '', '', '');
            updateClosureByReleaseTable('', '', '', '');
            updateReopenByReleaseTable('', '', '', '');
            updateMPocMetSlaTable('', '', '', '');
            updateMPocNotMetSlaTable('', '', '', '');
            populateComparativeTables();
        }});
    </script>
</body>
</html>
"""

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\n✅ Dashboard generated successfully!")
print(f"Output: {output_file}")
