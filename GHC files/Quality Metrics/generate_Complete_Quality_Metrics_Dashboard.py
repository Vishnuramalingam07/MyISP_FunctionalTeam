import pandas as pd
import numpy as np
import json

# Read the Excel files
input_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Quality Metrics Complete Input file for stories and bugs.xlsx'
stories_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Stories score summary.xlsx'

# Load all sheets from main input file
xl = pd.ExcelFile(input_file)
all_data = []

# Automatically get all sheet names from the Excel file
releases = xl.sheet_names
print(f"Found sheets in Excel file: {releases}")

for release in releases:
    df = pd.read_excel(xl, release)
    df['Release'] = release
    all_data.append(df)

# Load Stories score summary data
print(f"\nLoading Stories score summary data...")
stories_xl = pd.ExcelFile(stories_file)
stories_data = []
for release in stories_xl.sheet_names:
    stories_df = pd.read_excel(stories_xl, release)
    stories_df['Release'] = release
    stories_data.append(stories_df)

combined_stories_df = pd.concat(stories_data, ignore_index=True)
print(f"Loaded {len(combined_stories_df)} stories from {len(stories_xl.sheet_names)} sheets")

# Normalize Node Name for case-insensitive matching
combined_stories_df['Node Name Lower'] = combined_stories_df['Node Name'].str.lower()

# Combine all data
combined_df = pd.concat(all_data, ignore_index=True)

# Normalize Node Name for case-insensitive matching
combined_df['Node Name Lower'] = combined_df['Node Name'].str.lower()

print(f"Loaded {len(combined_df)} rows from {len(releases)} sheets")

# Clean data - handle NaN values for all numeric columns (columns E to V)
# This ensures blank cells are treated as zero to avoid NaN conversion errors
numeric_columns = [
    'Testing Not Applicable Stories',
    'Testable stories',
    'Total story Points',
    'Total Bugs',
    'Valid bugs',
    'Invalid bugs',
    'Total Critical',
    'Total High',
    'Total Medium',
    'Total Low',
    'Total PT Bugs',
    'PT Critical',
    'PT High',
    'PT Medium',
    'PT Low',
    'Total UAT Bugs',
    'UAT Critical',
    'UAT High',
    'UAT Medium',
    'UAT Low',
    'Valid Critical',
    'Valid High',
    'Valid Medium',
    'Valid Low',
    'Invalid Critical',
    'Invalid High',
    'Invalid Medium',
    'Invalid Low',
    'PT Valid Critical',
    'PT Valid High',
    'PT Valid Medium',
    'PT Valid Low',
    'PT Invalid Critical',
    'PT Invalid High',
    'PT Invalid Medium',
    'PT Invalid Low',
    'UAT Valid Critical',
    'UAT Valid High',
    'UAT Valid Medium',
    'UAT Valid Low',
    'UAT Invalid Critical',
    'UAT Invalid High',
    'UAT Invalid Medium',
    'UAT Invalid Low'
]

for col in numeric_columns:
    if col in combined_df.columns:
        combined_df[col] = combined_df[col].fillna(0)

# Get unique AD POCs and SM POCs
ad_pocs = combined_df['AD POC'].unique()
sm_pocs = combined_df['SM POC'].unique()

# Create AD POC summary
ad_summary = combined_df.groupby(['Release', 'AD POC']).agg({
    'Node Name': 'count',
    'Testable stories': 'sum',
    'Testing Not Applicable Stories': 'sum',
    'Total story Points': 'sum',
    'Total Bugs': 'sum',
    'Valid bugs': 'sum',
    'Total Critical': 'sum',
    'Total High': 'sum',
    'Total Medium': 'sum',
    'Total Low': 'sum',
    'Total PT Bugs': 'sum',
    'PT Critical': 'sum',
    'PT High': 'sum',
    'PT Medium': 'sum',
    'PT Low': 'sum',
    'Total UAT Bugs': 'sum',
    'UAT Critical': 'sum',
    'UAT High': 'sum',
    'UAT Medium': 'sum',
    'UAT Low': 'sum'
}).reset_index()

ad_summary.columns = ['Release', 'AD POC', 'Scrums', 'Testable Stories', 'Testing NA Stories', 'Story Points', 
                      'Total Bugs', 'Valid bugs', 'Critical', 'High', 'Medium', 'Low',
                      'PT Bugs', 'PT Critical', 'PT High', 'PT Medium', 'PT Low',
                      'UAT Bugs', 'UAT Critical', 'UAT High', 'UAT Medium', 'UAT Low']
ad_summary['Total Stories'] = ad_summary['Testable Stories'] + ad_summary['Testing NA Stories']
ad_summary['Avg Bugs/Story'] = (ad_summary['Valid bugs'] / ad_summary['Total Stories']).replace([np.inf, -np.inf], 0).fillna(0).round(2)
ad_summary['Avg Bugs/Story Point'] = (ad_summary['Valid bugs'] / ad_summary['Story Points']).replace([np.inf, -np.inf], 0).fillna(0).round(2)

# Create SM POC summary
sm_summary = combined_df.groupby(['Release', 'SM POC']).agg({
    'Node Name': 'count',
    'Testable stories': 'sum',
    'Testing Not Applicable Stories': 'sum',
    'Total story Points': 'sum',
    'Total Bugs': 'sum',
    'Valid bugs': 'sum',
    'Total Critical': 'sum',
    'Total High': 'sum',
    'Total Medium': 'sum',
    'Total Low': 'sum',
    'Total PT Bugs': 'sum',
    'PT Critical': 'sum',
    'PT High': 'sum',
    'PT Medium': 'sum',
    'PT Low': 'sum',
    'Total UAT Bugs': 'sum',
    'UAT Critical': 'sum',
    'UAT High': 'sum',
    'UAT Medium': 'sum',
    'UAT Low': 'sum'
}).reset_index()

sm_summary.columns = ['Release', 'SM POC', 'Scrums', 'Testable Stories', 'Testing NA Stories', 'Story Points', 
                      'Total Bugs', 'Valid bugs', 'Critical', 'High', 'Medium', 'Low',
                      'PT Bugs', 'PT Critical', 'PT High', 'PT Medium', 'PT Low',
                      'UAT Bugs', 'UAT Critical', 'UAT High', 'UAT Medium', 'UAT Low']
sm_summary['Total Stories'] = sm_summary['Testable Stories'] + sm_summary['Testing NA Stories']
sm_summary['Avg Bugs/Story'] = (sm_summary['Valid bugs'] / sm_summary['Total Stories']).replace([np.inf, -np.inf], 0).fillna(0).round(2)
sm_summary['Avg Bugs/Story Point'] = (sm_summary['Valid bugs'] / sm_summary['Story Points']).replace([np.inf, -np.inf], 0).fillna(0).round(2)

# Create M POC summary
m_summary = combined_df.groupby(['Release', 'M POC']).agg({
    'Node Name': 'count',
    'Testable stories': 'sum',
    'Testing Not Applicable Stories': 'sum',
    'Total story Points': 'sum',
    'Total Bugs': 'sum',
    'Valid bugs': 'sum',
    'Total Critical': 'sum',
    'Total High': 'sum',
    'Total Medium': 'sum',
    'Total Low': 'sum',
    'Total PT Bugs': 'sum',
    'PT Critical': 'sum',
    'PT High': 'sum',
    'PT Medium': 'sum',
    'PT Low': 'sum',
    'Total UAT Bugs': 'sum',
    'UAT Critical': 'sum',
    'UAT High': 'sum',
    'UAT Medium': 'sum',
    'UAT Low': 'sum'
}).reset_index()

m_summary.columns = ['Release', 'M POC', 'Scrums', 'Testable Stories', 'Testing NA Stories', 'Story Points', 
                      'Total Bugs', 'Valid bugs', 'Critical', 'High', 'Medium', 'Low',
                      'PT Bugs', 'PT Critical', 'PT High', 'PT Medium', 'PT Low',
                      'UAT Bugs', 'UAT Critical', 'UAT High', 'UAT Medium', 'UAT Low']
m_summary['Total Stories'] = m_summary['Testable Stories'] + m_summary['Testing NA Stories']
m_summary['Avg Bugs/Story'] = (m_summary['Valid bugs'] / m_summary['Total Stories']).replace([np.inf, -np.inf], 0).fillna(0).round(2)
m_summary['Avg Bugs/Story Point'] = (m_summary['Valid bugs'] / m_summary['Story Points']).replace([np.inf, -np.inf], 0).fillna(0).round(2)

# Get top POCs based on total bugs
ad_totals = ad_summary.groupby('AD POC')['Total Bugs'].sum().sort_values(ascending=False)
sm_totals = sm_summary.groupby('SM POC')['Total Bugs'].sum().sort_values(ascending=False)
m_totals = m_summary.groupby('M POC')['Total Bugs'].sum().sort_values(ascending=False)

top_ad_pocs = ad_totals.head(10).index.tolist()
top_sm_pocs = sm_totals.head(10).index.tolist()
# Use all M POCs instead of just top 10 to ensure complete hierarchy
top_m_pocs = m_totals.index.tolist()

# Create M POC to SM POC mapping to find all SM POCs referenced by M POCs
m_to_sm_temp_mapping = combined_df.groupby('M POC')['SM POC'].first().to_dict()
# Add any SM POCs that are parents of M POCs but not in top 10
sm_pocs_from_m = set(m_to_sm_temp_mapping.values())
for sm_poc in sm_pocs_from_m:
    if sm_poc not in top_sm_pocs:
        top_sm_pocs.append(sm_poc)

# Create SM POC to AD POC mapping and reorder SM POCs by AD POC
sm_to_ad_mapping = combined_df.groupby('SM POC')['AD POC'].first().to_dict()

# Group SM POCs by their AD POC and maintain order
sm_pocs_by_ad = {}
for sm_poc in top_sm_pocs:
    ad_poc = sm_to_ad_mapping.get(sm_poc, 'Unknown')
    if ad_poc not in sm_pocs_by_ad:
        sm_pocs_by_ad[ad_poc] = []
    sm_pocs_by_ad[ad_poc].append(sm_poc)

# Reorder top_sm_pocs to group by AD POC
top_sm_pocs_ordered = []
for ad_poc in sorted(sm_pocs_by_ad.keys()):
    top_sm_pocs_ordered.extend(sm_pocs_by_ad[ad_poc])

top_sm_pocs = top_sm_pocs_ordered

# Create M POC to AD POC mapping
m_to_ad_mapping = combined_df.groupby('M POC')['AD POC'].first().to_dict()
m_to_sm_mapping = combined_df.groupby('M POC')['SM POC'].first().to_dict()

# Group M POCs by their AD POC and maintain order
m_pocs_by_ad = {}
for m_poc in top_m_pocs:
    ad_poc = m_to_ad_mapping.get(m_poc, 'Unknown')
    if ad_poc not in m_pocs_by_ad:
        m_pocs_by_ad[ad_poc] = []
    m_pocs_by_ad[ad_poc].append(m_poc)

# Build M POC hierarchy by AD -> SM -> M (using the corrected AD POC mapping)
m_pocs_hierarchy = {}
for m_poc in top_m_pocs:
    ad_poc = m_to_ad_mapping.get(m_poc, 'Unknown')
    sm_poc = m_to_sm_mapping.get(m_poc, 'Unknown')
    if ad_poc not in m_pocs_hierarchy:
        m_pocs_hierarchy[ad_poc] = {}
    if sm_poc not in m_pocs_hierarchy[ad_poc]:
        m_pocs_hierarchy[ad_poc][sm_poc] = []
    m_pocs_hierarchy[ad_poc][sm_poc].append(m_poc)

# Reorder top_m_pocs to group by AD POC
top_m_pocs_ordered = []
for ad_poc in sorted(m_pocs_by_ad.keys()):
    top_m_pocs_ordered.extend(m_pocs_by_ad[ad_poc])

top_m_pocs = top_m_pocs_ordered

print(f"AD POCs: {len(top_ad_pocs)}")
print(f"SM POCs: {len(top_sm_pocs)}")
print(f"M POCs: {len(top_m_pocs)}")

# Create hierarchy data for combined POC filtering
# Build complete hierarchy mapping for JavaScript
poc_hierarchy = {}
node_to_mpoc_mapping = combined_df.groupby('Node Name')['M POC'].first().to_dict()
mpoc_to_smpoc_mapping = combined_df.groupby('M POC')['SM POC'].first().to_dict()
smpoc_to_adpoc_mapping = combined_df.groupby('SM POC')['AD POC'].first().to_dict()

for ad_poc in combined_df['AD POC'].unique():
    poc_hierarchy[ad_poc] = {}
    ad_data = combined_df[combined_df['AD POC'] == ad_poc]
    for sm_poc in ad_data['SM POC'].unique():
        poc_hierarchy[ad_poc][sm_poc] = {}
        sm_data = ad_data[ad_data['SM POC'] == sm_poc]
        for m_poc in sm_data['M POC'].unique():
            m_data = sm_data[sm_data['M POC'] == m_poc]
            poc_hierarchy[ad_poc][sm_poc][m_poc] = m_data['Node Name'].unique().tolist()

# Create combined POC data with all levels (Node Name level data)
combined_poc_data = []
for _, row in combined_df.iterrows():
    release = row['Release']
    ad_poc = row['AD POC']
    sm_poc = row['SM POC']
    m_poc = row['M POC']
    node_name = row['Node Name']
    
    # Get story metrics for this POC/Node combination
    story_filter = (
        (combined_stories_df['Release'] == release) &
        (combined_stories_df['AD POC'] == ad_poc) &
        (combined_stories_df['SM POC'] == sm_poc) &
        (combined_stories_df['M POC'] == m_poc) &
        (combined_stories_df['Node Name Lower'] == node_name.lower())
    )
    story_data = combined_stories_df[story_filter]
    
    total_testable_stories = len(story_data)
    agent_no_count = (story_data['Agent Augmented delivery_Development'] == 'No').sum()
    delayed_yes_count = (story_data['Delayed Story Delivery'] == 'Yes').sum()
    
    combined_poc_data.append({
        'Release': release,
        'AD POC': ad_poc,
        'SM POC': sm_poc,
        'M POC': m_poc,
        'Node Name': node_name,
        'Total Stories': int(row['Testable stories']) + int(row['Testing Not Applicable Stories']),
        'Testable Stories': int(row['Testable stories']),
        'Testing NA Stories': int(row['Testing Not Applicable Stories']),
        'Story Points': int(row['Total story Points']),
        'Total Bugs': int(row['Total Bugs']),
        'Valid bugs': int(row['Valid bugs']),
        'Invalid bugs': int(row['Invalid bugs']),
        'Critical': int(row['Total Critical']),
        'High': int(row['Total High']),
        'Medium': int(row['Total Medium']),
        'Low': int(row['Total Low']),
        'PT Bugs': int(row['Total PT Bugs']),
        'PT Critical': int(row['PT Critical']),
        'PT High': int(row['PT High']),
        'PT Medium': int(row['PT Medium']),
        'PT Low': int(row['PT Low']),
        'UAT Bugs': int(row['Total UAT Bugs']),
        'UAT Critical': int(row['UAT Critical']),
        'UAT High': int(row['UAT High']),
        'UAT Medium': int(row['UAT Medium']),
        'UAT Low': int(row['UAT Low']),
        'Valid Critical': int(row['Valid Critical']),
        'Valid High': int(row['Valid High']),
        'Valid Medium': int(row['Valid Medium']),
        'Valid Low': int(row['Valid Low']),
        'Invalid Critical': int(row['Invalid Critical']),
        'Invalid High': int(row['Invalid High']),
        'Invalid Medium': int(row['Invalid Medium']),
        'Invalid Low': int(row['Invalid Low']),
        'PT Valid Critical': int(row['PT Valid Critical']),
        'PT Valid High': int(row['PT Valid High']),
        'PT Valid Medium': int(row['PT Valid Medium']),
        'PT Valid Low': int(row['PT Valid Low']),
        'PT Invalid Critical': int(row['PT Invalid Critical']),
        'PT Invalid High': int(row['PT Invalid High']),
        'PT Invalid Medium': int(row['PT Invalid Medium']),
        'PT Invalid Low': int(row['PT Invalid Low']),
        'UAT Valid Critical': int(row['UAT Valid Critical']),
        'UAT Valid High': int(row['UAT Valid High']),
        'UAT Valid Medium': int(row['UAT Valid Medium']),
        'UAT Valid Low': int(row['UAT Valid Low']),
        'UAT Invalid Critical': int(row['UAT Invalid Critical']),
        'UAT Invalid High': int(row['UAT Invalid High']),
        'UAT Invalid Medium': int(row['UAT Invalid Medium']),
        'UAT Invalid Low': int(row['UAT Invalid Low']),
        'Total Testable Stories Count': int(total_testable_stories),
        'Agent Augmented No Count': int(agent_no_count),
        'Delayed Delivery Yes Count': int(delayed_yes_count)
    })

print("")
print("=== M POC Hierarchy ===")
for ad_poc in sorted(m_pocs_hierarchy.keys()):
    for sm_poc in sorted(m_pocs_hierarchy[ad_poc].keys()):
        for m_poc in m_pocs_hierarchy[ad_poc][sm_poc]:
            print(f"M POC: {m_poc} -> SM POC: {sm_poc} -> AD POC: {ad_poc}")

# Calculate overall statistics
total_scrums = combined_df['Node Name'].nunique()
total_stories = int(combined_df['Testable stories'].sum())
total_story_points = int(combined_df['Total story Points'].sum())
total_bugs = int(combined_df['Total Bugs'].sum())
total_critical = int(combined_df['Total Critical'].sum())
total_high = int(combined_df['Total High'].sum())
total_medium = int(combined_df['Total Medium'].sum())
total_low = int(combined_df['Total Low'].sum())

# Calculate release-wise statistics including PT and UAT
release_stats = {}
for release in releases:
    release_df = combined_df[combined_df['Release'] == release]
    testable_stories = int(release_df['Testable stories'].sum())
    testing_na_stories = int(release_df['Testing Not Applicable Stories'].sum())
    total_stories = testable_stories + testing_na_stories
    
    release_stats[release] = {
        'scrums': len(release_df),
        'total_stories': total_stories,
        'testable_stories': testable_stories,
        'testing_na_stories': testing_na_stories,
        'story_points': int(release_df['Total story Points'].sum()),
        'total_bugs': int(release_df['Total Bugs'].sum()),
        'critical': int(release_df['Total Critical'].sum()),
        'high': int(release_df['Total High'].sum()),
        'medium': int(release_df['Total Medium'].sum()),
        'low': int(release_df['Total Low'].sum()),
        # PT metrics
        'pt_bugs': int(release_df['Total PT Bugs'].sum()),
        'pt_critical': int(release_df['PT Critical'].sum()),
        'pt_high': int(release_df['PT High'].sum()),
        'pt_medium': int(release_df['PT Medium'].sum()),
        'pt_low': int(release_df['PT Low'].sum()),
        # UAT metrics
        'uat_bugs': int(release_df['Total UAT Bugs'].sum()),
        'uat_critical': int(release_df['UAT Critical'].sum()),
        'uat_high': int(release_df['UAT High'].sum()),
        'uat_medium': int(release_df['UAT Medium'].sum()),
        'uat_low': int(release_df['UAT Low'].sum())
    }

# Generate HTML content
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Complete Quality Metrics Dashboard</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 10px;
            min-height: 100vh;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}

        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            text-align: center;
        }}

        .header h1 {{
            font-size: 2.2em;
            margin-bottom: 5px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}

        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}

        .tabs {{
            display: flex;
            background: #f8f9fa;
            border-bottom: 2px solid #dee2e6;
            overflow-x: auto;
        }}

        .tab {{
            padding: 15px 30px;
            cursor: pointer;
            border: none;
            background: none;
            font-size: 1em;
            font-weight: 600;
            color: #6c757d;
            transition: all 0.3s ease;
            border-bottom: 3px solid transparent;
            white-space: nowrap;
        }}

        .tab:hover {{
            background: #e9ecef;
            color: #495057;
        }}

        .tab.active {{
            color: #667eea;
            border-bottom-color: #667eea;
            background: white;
        }}

        .tab-content {{
            display: none;
            padding: 20px;
            animation: fadeIn 0.5s;
        }}

        .tab-content.active {{
            display: block;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .section-title {{
            font-size: 1.6em;
            color: #2c3e50;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
        }}

        .release-section {{
            margin-bottom: 25px;
        }}

        .release-header {{
            font-size: 1.5em;
            font-weight: bold;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            margin-bottom: 15px;
            text-align: center;
        }}

        .release-sep {{
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }}

        .release-oct {{
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }}

        .release-nov {{
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }}

        .release-dec {{
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }}

        .release-jan {{
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }}

        .overview-grid {{
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 10px;
            margin-bottom: 20px;
        }}

        .metric-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 15px;
            border-radius: 12px;
            color: white;
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.3);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}

        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 15px 40px rgba(102, 126, 234, 0.4);
        }}

        .metric-card h3 {{
            font-size: 0.85em;
            margin-bottom: 8px;
            opacity: 0.9;
        }}

        .metric-card .value {{
            font-size: 1.6em;
            font-weight: bold;
            margin-bottom: 3px;
        }}

        .metric-card .sub-value {{
            font-size: 0.75em;
            opacity: 0.8;
        }}

        .subsection-title {{
            font-size: 1.3em;
            font-weight: bold;
            color: #2c3e50;
            margin: 15px 0 10px 0;
            padding-left: 12px;
            border-left: 4px solid #667eea;
        }}

        .severity-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            margin-bottom: 30px;
        }}

        .severity-card {{
            padding: 12px;
            border-radius: 10px;
            text-align: center;
            color: white;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}

        .severity-critical {{
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
        }}

        .severity-high {{
            background: linear-gradient(135deg, #fd7e14 0%, #e8590c 100%);
        }}

        .severity-medium {{
            background: linear-gradient(135deg, #ffc107 0%, #e0a800 100%);
        }}

        .severity-low {{
            background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%);
        }}

        .severity-card h4 {{
            font-size: 0.85em;
            margin-bottom: 6px;
            opacity: 0.9;
        }}

        .severity-card .value {{
            font-size: 1.4em;
            font-weight: bold;
        }}

        .poc-section {{
            background: #f8f9fa;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }}

        .poc-header {{
            font-size: 1.5em;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 2px solid #dee2e6;
        }}

        .release-comparison {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
        }}

        .release-card {{
            background: white;
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}

        .release-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 12px rgba(0,0,0,0.15);
        }}

        .release-label {{
            font-size: 1.2em;
            font-weight: bold;
            color: white;
            padding: 8px 16px;
            border-radius: 8px;
            margin-bottom: 20px;
            text-align: center;
        }}

        .metrics-list {{
            list-style: none;
        }}

        .metrics-list li {{
            padding: 12px;
            margin-bottom: 10px;
            background: #f8f9fa;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .metrics-list li span:first-child {{
            font-weight: 600;
            color: #495057;
        }}

        .metrics-list li span:last-child {{
            font-weight: bold;
            color: #667eea;
            font-size: 1.1em;
        }}

        .severity-breakdown {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
            margin-top: 15px;
        }}

        .severity-item {{
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 0.9em;
            display: flex;
            justify-content: space-between;
            color: white;
        }}

        table {{
            width: 98%;
            border-collapse: collapse;
            margin: 20px auto;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            font-size: 0.95em;
        }}

        th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 8px 6px;
            text-align: left;
            font-weight: 600;
            font-size: 0.75em;
        }}

        td {{
            padding: 6px 6px;
            border-bottom: 1px solid #dee2e6;
            font-size: 0.8em;
        }}

        tr:hover {{
            background: #f8f9fa;
        }}

        tr:last-child td {{
            border-bottom: none;
        }}

        .quality-badge {{
            padding: 3px 6px;
            border-radius: 4px;
            font-size: 0.7em;
            font-weight: bold;
            color: white;
            display: inline-block;
            white-space: nowrap;
        }}

        .quality-excellent {{
            background: #28a745;
        }}

        .quality-good {{
            background: #ffc107;
            color: #333;
        }}

        .quality-attention {{
            background: #dc3545;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Complete Quality Metrics Dashboard</h1>
            <p>Stories, Bugs, PT & UAT Defects Analysis across All Releases</p>
        </div>

        <div class="tabs">
            <button class="tab active" onclick="showTab('overview')">Overview</button>
            <button class="tab" onclick="showTab('poc-metrics')">POC wise Metrics</button>
            <button class="tab" onclick="showTab('detailed')">Detailed View</button>
        </div>

        <div id="overview" class="tab-content active">
            <h2 class="section-title">Release-wise Quality Overview</h2>
"""

# Generate release-wise sections
for idx, release in enumerate(releases):
    stats = release_stats[release]
    
    # Determine release class dynamically based on index
    release_classes = ['release-sep', 'release-oct', 'release-nov', 'release-dec', 'release-jan']
    release_class = release_classes[idx % len(release_classes)]
    
    html_content += f"""
            <div class="release-section">
                <div class="release-header {release_class}">{release} Release</div>
                
                <div class="overview-grid" style="grid-template-columns: repeat(5, 1fr);">
                    <div class="metric-card">
                        <h3>Scrums</h3>
                        <div class="value">{stats['scrums']}</div>
                    </div>
                    <div class="metric-card">
                        <h3>Total Stories</h3>
                        <div class="value">{stats['total_stories']}</div>
                    </div>
                    <div class="metric-card">
                        <h3>Testable Stories</h3>
                        <div class="value">{stats['testable_stories']}</div>
                    </div>
                    <div class="metric-card">
                        <h3>Testing NA Stories</h3>
                        <div class="value">{stats['testing_na_stories']}</div>
                    </div>
                    <div class="metric-card">
                        <h3>Story Points</h3>
                        <div class="value">{stats['story_points']}</div>
                    </div>
                </div>

                <div class="subsection-title">Defects Breakdown</div>
                <div style="overflow-x: auto; margin-top: 20px;">
                    <table style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                        <thead>
                            <tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
                                <th style="padding: 15px; text-align: left; font-weight: 600; border: 1px solid #ddd;">Defect Type</th>
                                <th style="padding: 15px; text-align: center; font-weight: 600; border: 1px solid #ddd;">Total</th>
                                <th style="padding: 15px; text-align: center; font-weight: 600; border: 1px solid #ddd;">Critical</th>
                                <th style="padding: 15px; text-align: center; font-weight: 600; border: 1px solid #ddd;">High</th>
                                <th style="padding: 15px; text-align: center; font-weight: 600; border: 1px solid #ddd;">Medium</th>
                                <th style="padding: 15px; text-align: center; font-weight: 600; border: 1px solid #ddd;">Low</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr style="background: #f8f9fa;">
                                <td style="padding: 12px 15px; font-weight: 600; border: 1px solid #ddd;">Total Defects</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 600; color: #333; border: 1px solid #ddd;">{stats['total_bugs']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 600; color: #dc3545; border: 1px solid #ddd;">{stats['critical']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 600; color: #fd7e14; border: 1px solid #ddd;">{stats['high']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 600; color: #ffc107; border: 1px solid #ddd;">{stats['medium']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 600; color: #28a745; border: 1px solid #ddd;">{stats['low']}</td>
                            </tr>
                            <tr style="background: white;">
                                <td style="padding: 12px 15px; font-weight: 500; border: 1px solid #ddd;">PT Defects</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #333; border: 1px solid #ddd;">{stats['pt_bugs']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #dc3545; border: 1px solid #ddd;">{stats['pt_critical']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #fd7e14; border: 1px solid #ddd;">{stats['pt_high']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #ffc107; border: 1px solid #ddd;">{stats['pt_medium']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #28a745; border: 1px solid #ddd;">{stats['pt_low']}</td>
                            </tr>
                            <tr style="background: #f8f9fa;">
                                <td style="padding: 12px 15px; font-weight: 500; border: 1px solid #ddd;">UAT Defects</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #333; border: 1px solid #ddd;">{stats['uat_bugs']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #dc3545; border: 1px solid #ddd;">{stats['uat_critical']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #fd7e14; border: 1px solid #ddd;">{stats['uat_high']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #ffc107; border: 1px solid #ddd;">{stats['uat_medium']}</td>
                                <td style="padding: 12px 15px; text-align: center; font-weight: 500; color: #28a745; border: 1px solid #ddd;">{stats['uat_low']}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
"""

html_content += """
        </div>

        <div id="poc-metrics" class="tab-content">
            <!-- Hierarchical Filters -->
            <div style="margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                <div class="filter-title" style="font-size: 1.1em; color: #333; margin-bottom: 12px; font-weight: 600;">🔍 Filter Hierarchy</div>
                <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; align-items: end;">
                    <div>
                        <label for="poc-ad-filter" style="font-weight: 600; color: #555; margin-bottom: 4px; display: block; font-size: 0.8em;">AD POC:</label>
                        <select id="poc-ad-filter" onchange="updatePOCFilters('ad')" style="width: 100%; padding: 6px 8px; border: 2px solid #667eea; border-radius: 6px; font-size: 0.85em; background: white; cursor: pointer;">
                            <option value="all">All AD POCs</option>
"""

# Add AD POC options
for ad_poc in sorted(poc_hierarchy.keys()):
    html_content += f"""                            <option value="{ad_poc}">{ad_poc}</option>
"""

html_content += """                        </select>
                    </div>
                    <div>
                        <label for="poc-sm-filter" style="font-weight: 600; color: #555; margin-bottom: 4px; display: block; font-size: 0.8em;">SM POC:</label>
                        <select id="poc-sm-filter" onchange="updatePOCFilters('sm')" disabled style="width: 100%; padding: 6px 8px; border: 2px solid #ddd; border-radius: 6px; font-size: 0.85em; background: #e9ecef; cursor: not-allowed;">
                            <option value="all">All SM POCs</option>
                        </select>
                    </div>
                    <div>
                        <label for="poc-m-filter" style="font-weight: 600; color: #555; margin-bottom: 4px; display: block; font-size: 0.8em;">M POC:</label>
                        <select id="poc-m-filter" onchange="updatePOCFilters('m')" disabled style="width: 100%; padding: 6px 8px; border: 2px solid #ddd; border-radius: 6px; font-size: 0.85em; background: #e9ecef; cursor: not-allowed;">
                            <option value="all">All M POCs</option>
                        </select>
                    </div>
                    <div>
                        <label for="poc-node-filter" style="font-weight: 600; color: #555; margin-bottom: 4px; display: block; font-size: 0.8em;">Node Name:</label>
                        <select id="poc-node-filter" onchange="updatePOCFilters('node')" disabled style="width: 100%; padding: 6px 8px; border: 2px solid #ddd; border-radius: 6px; font-size: 0.85em; background: #e9ecef; cursor: not-allowed;">
                            <option value="all">All Nodes</option>
                        </select>
                    </div>
                    <div>
                        <button onclick="resetPOCFilters()" style="width: 100%; padding: 6px 12px; border-radius: 6px; border: none; background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; cursor: pointer; font-weight: 600; font-size: 0.8em; transition: all 0.3s; box-shadow: 0 2px 8px rgba(245,87,108,0.3);" onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 12px rgba(245,87,108,0.4)'" onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(245,87,108,0.3)'">
                            🔄 Reset Filters
                        </button>
                    </div>
                </div>
            </div>

            <!-- POC Data Display -->
            <div id="poc-data-container">
                <!-- Data will be dynamically populated by JavaScript -->
            </div>
            
            <div style="margin-top: 40px; padding: 20px; background: #fff3cd; border-left: 5px solid #ffc107; border-radius: 8px;">
                <p style="margin: 0; color: #856404; font-size: 0.95em;"><strong>Note:</strong> *Avg Bugs/Story is calculated based on the total bugs raised for the entire release under any scrum. Use the hierarchical filters above to drill down from AD POC → SM POC → M POC → Node Name.</p>
            </div>
        </div>

        <div id="detailed" class="tab-content">
            <h2 class="section-title">Detailed Scrum-Level Analysis</h2>
"""


# Prepare detailed view data
combined_df['Total Stories'] = combined_df['Testable stories'] + combined_df['Testing Not Applicable Stories']
combined_df['Avg Bugs/Story'] = (combined_df['Valid bugs'] / combined_df['Total Stories']).replace([np.inf, -np.inf], 0).fillna(0).round(2)
combined_df['Avg Bugs/Story Point'] = (combined_df['Valid bugs'] / combined_df['Total story Points']).replace([np.inf, -np.inf], 0).fillna(0).round(2)

def get_quality_badge(avg_bugs):
    if avg_bugs < 3:
        return '<span class="quality-badge quality-excellent">Excellent</span>'
    elif avg_bugs <= 5:
        return '<span class="quality-badge quality-good">Good</span>'
    else:
        return '<span class="quality-badge quality-attention">Needs Attention</span>'

# Generate detailed table for each release
for release in releases:
    release_data = combined_df[combined_df['Release'] == release].copy()
    
    # Sort by Avg Bugs/Story in descending order (highest bugs per story first)
    release_data = release_data.sort_values('Avg Bugs/Story', ascending=False)
    
    html_content += f"""
            <div class="poc-section">
                <div class="poc-header">{release} - Detailed Breakdown</div>
                <table style="font-size: 0.85em;">
                    <thead>
                        <tr>
                            <th style="width: 8%;">AD POC</th>
                            <th style="width: 8%;">SM POC</th>
                            <th style="width: 8%;">M POC</th>
                            <th style="width: 12%;">Node Name</th>
                            <th style="width: 6%;">Total Stories</th>
                            <th style="width: 6%;">Story Points</th>
                            <th style="width: 6%;">Total Bugs</th>
                            <th style="width: 6%;">Valid Bugs</th>
                            <th style="width: 6%;">Invalid Bugs</th>
                            <th style="width: 8%;">Avg Bugs/Story Point</th>
                            <th style="width: 8%;">Avg Bugs/Story</th>
                            <th style="width: 10%;">Quality</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    for _, row in release_data.iterrows():
        quality_badge = get_quality_badge(row['Avg Bugs/Story'])
        html_content += f"""
                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td><strong>{row['Node Name']}</strong></td>
                            <td>{int(row['Total Stories'])}</td>
                            <td>{int(row['Total story Points'])}</td>
                            <td><strong>{int(row['Total Bugs'])}</strong></td>
                            <td>{int(row['Valid bugs'])}</td>
                            <td>{int(row['Invalid bugs'])}</td>
                            <td>{row['Avg Bugs/Story Point']:.2f}</td>
                            <td>{row['Avg Bugs/Story']:.2f}</td>
                            <td>{quality_badge}</td>
                        </tr>
"""
    
    html_content += """
                    </tbody>
                </table>
            </div>
"""

html_content += """
            <div style="margin-top: 40px; padding: 20px; background: #fff3cd; border-left: 5px solid #ffc107; border-radius: 8px;">
                <p style="margin: 0; color: #856404; font-size: 0.95em;"><strong>Note:</strong> *Avg Bugs/Story is calculated based on the total bugs raised for the entire release under any scrum.</p>
            </div>
        </div>
    </div>

    <script>
        // Use Python-generated hierarchies for consistent filtering
"""
html_content += f"        const smPocHierarchy = {json.dumps(sm_pocs_by_ad)};\n"
html_content += f"        const mPocHierarchy = {json.dumps(m_pocs_hierarchy)};\n"
html_content += f"        const pocHierarchy = {json.dumps(poc_hierarchy)};\n"
html_content += f"        const combinedPocData = {json.dumps(combined_poc_data)};\n"
html_content += f"        const releases = {json.dumps(releases)};\n"
html_content += """
        // POC wise metrics filtering functions
        function updatePOCFilters(level) {
            const adFilter = document.getElementById('poc-ad-filter');
            const smFilter = document.getElementById('poc-sm-filter');
            const mFilter = document.getElementById('poc-m-filter');
            const nodeFilter = document.getElementById('poc-node-filter');
            
            const selectedAd = adFilter.value;
            const selectedSm = smFilter.value;
            const selectedM = mFilter.value;
            
            // Update SM POC filter based on AD selection
            if (selectedAd !== 'all') {
                const smPocs = Object.keys(pocHierarchy[selectedAd] || {});
                const currentSmValue = smFilter.value;
                smFilter.innerHTML = '<option value="all">All SM POCs</option>' +
                    smPocs.map(sm => `<option value="${sm}"${sm === currentSmValue ? ' selected' : ''}>${sm}</option>`).join('');
                smFilter.disabled = false;
                smFilter.style.background = 'white';
                smFilter.style.cursor = 'pointer';
                smFilter.style.borderColor = '#667eea';
                
                // If we just changed AD, reset dependent filters
                if (level === 'ad') {
                    mFilter.value = 'all';
                    mFilter.innerHTML = '<option value="all">All M POCs</option>';
                    mFilter.disabled = true;
                    mFilter.style.background = '#e9ecef';
                    mFilter.style.cursor = 'not-allowed';
                    mFilter.style.borderColor = '#ddd';
                    nodeFilter.value = 'all';
                    nodeFilter.innerHTML = '<option value="all">All Nodes</option>';
                    nodeFilter.disabled = true;
                    nodeFilter.style.background = '#e9ecef';
                    nodeFilter.style.cursor = 'not-allowed';
                    nodeFilter.style.borderColor = '#ddd';
                }
            } else {
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                smFilter.disabled = true;
                smFilter.style.background = '#e9ecef';
                smFilter.style.cursor = 'not-allowed';
                smFilter.style.borderColor = '#ddd';
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                mFilter.disabled = true;
                mFilter.style.background = '#e9ecef';
                mFilter.style.cursor = 'not-allowed';
                mFilter.style.borderColor = '#ddd';
                nodeFilter.innerHTML = '<option value="all">All Nodes</option>';
                nodeFilter.disabled = true;
                nodeFilter.style.background = '#e9ecef';
                nodeFilter.style.cursor = 'not-allowed';
                nodeFilter.style.borderColor = '#ddd';
            }
            
            // Update M POC filter based on SM selection
            if (smFilter.value !== 'all' && selectedAd !== 'all') {
                const mPocs = Object.keys(pocHierarchy[selectedAd][smFilter.value] || {});
                const currentMValue = mFilter.value;
                mFilter.innerHTML = '<option value="all">All M POCs</option>' +
                    mPocs.map(m => `<option value="${m}"${m === currentMValue ? ' selected' : ''}>${m}</option>`).join('');
                mFilter.disabled = false;
                mFilter.style.background = 'white';
                mFilter.style.cursor = 'pointer';
                mFilter.style.borderColor = '#667eea';
                
                // If we just changed SM, reset dependent filter
                if (level === 'sm') {
                    nodeFilter.value = 'all';
                    nodeFilter.innerHTML = '<option value="all">All Nodes</option>';
                    nodeFilter.disabled = true;
                    nodeFilter.style.background = '#e9ecef';
                    nodeFilter.style.cursor = 'not-allowed';
                    nodeFilter.style.borderColor = '#ddd';
                }
            } else if (smFilter.value === 'all' && selectedAd !== 'all') {
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                mFilter.disabled = true;
                mFilter.style.background = '#e9ecef';
                mFilter.style.cursor = 'not-allowed';
                mFilter.style.borderColor = '#ddd';
                nodeFilter.innerHTML = '<option value="all">All Nodes</option>';
                nodeFilter.disabled = true;
                nodeFilter.style.background = '#e9ecef';
                nodeFilter.style.cursor = 'not-allowed';
                nodeFilter.style.borderColor = '#ddd';
            }
            
            // Update Node filter based on M selection
            if (mFilter.value !== 'all' && smFilter.value !== 'all' && selectedAd !== 'all') {
                const nodes = pocHierarchy[selectedAd][smFilter.value][mFilter.value] || [];
                const currentNodeValue = nodeFilter.value;
                nodeFilter.innerHTML = '<option value="all">All Nodes</option>' +
                    nodes.map(node => `<option value="${node}"${node === currentNodeValue ? ' selected' : ''}>${node}</option>`).join('');
                nodeFilter.disabled = false;
                nodeFilter.style.background = 'white';
                nodeFilter.style.cursor = 'pointer';
                nodeFilter.style.borderColor = '#667eea';
            } else {
                // Reset Node filter when M is 'all' or any parent is 'all'
                nodeFilter.value = 'all';
                nodeFilter.innerHTML = '<option value="all">All Nodes</option>';
                nodeFilter.disabled = true;
                nodeFilter.style.background = '#e9ecef';
                nodeFilter.style.cursor = 'not-allowed';
                nodeFilter.style.borderColor = '#ddd';
            }
            
            // Render filtered data
            renderPOCData();
        }
        
        function renderPOCData() {
            const adFilter = document.getElementById('poc-ad-filter').value;
            const smFilter = document.getElementById('poc-sm-filter').value;
            const mFilter = document.getElementById('poc-m-filter').value;
            const nodeFilter = document.getElementById('poc-node-filter').value;
            
            // Filter data based on current selections
            let filteredData = combinedPocData.filter(item => {
                return (adFilter === 'all' || item['AD POC'] === adFilter) &&
                       (smFilter === 'all' || item['SM POC'] === smFilter) &&
                       (mFilter === 'all' || item['M POC'] === mFilter) &&
                       (nodeFilter === 'all' || item['Node Name'] === nodeFilter);
            });
            
            // Group by release and aggregate metrics
            const aggregatedByRelease = {};
            
            releases.forEach(release => {
                const releaseData = filteredData.filter(item => item['Release'] === release);
                
                if (releaseData.length > 0) {
                    const totalStories = releaseData.reduce((sum, item) => sum + (item['Total Stories'] || 0), 0);
                    const testableStories = releaseData.reduce((sum, item) => sum + (item['Testable Stories'] || 0), 0);
                    const testingNAStories = releaseData.reduce((sum, item) => sum + (item['Testing NA Stories'] || 0), 0);
                    const storyPoints = releaseData.reduce((sum, item) => sum + (item['Story Points'] || 0), 0);
                    const totalBugs = releaseData.reduce((sum, item) => sum + (item['Total Bugs'] || 0), 0);
                    const validBugs = releaseData.reduce((sum, item) => sum + (item['Valid bugs'] || 0), 0);
                    const invalidBugs = releaseData.reduce((sum, item) => sum + (item['Invalid bugs'] || 0), 0);
                    const critical = releaseData.reduce((sum, item) => sum + (item['Critical'] || 0), 0);
                    const high = releaseData.reduce((sum, item) => sum + (item['High'] || 0), 0);
                    const medium = releaseData.reduce((sum, item) => sum + (item['Medium'] || 0), 0);
                    const low = releaseData.reduce((sum, item) => sum + (item['Low'] || 0), 0);
                    const ptBugs = releaseData.reduce((sum, item) => sum + (item['PT Bugs'] || 0), 0);
                    const ptCritical = releaseData.reduce((sum, item) => sum + (item['PT Critical'] || 0), 0);
                    const ptHigh = releaseData.reduce((sum, item) => sum + (item['PT High'] || 0), 0);
                    const ptMedium = releaseData.reduce((sum, item) => sum + (item['PT Medium'] || 0), 0);
                    const ptLow = releaseData.reduce((sum, item) => sum + (item['PT Low'] || 0), 0);
                    const uatBugs = releaseData.reduce((sum, item) => sum + (item['UAT Bugs'] || 0), 0);
                    const uatCritical = releaseData.reduce((sum, item) => sum + (item['UAT Critical'] || 0), 0);
                    const uatHigh = releaseData.reduce((sum, item) => sum + (item['UAT High'] || 0), 0);
                    const uatMedium = releaseData.reduce((sum, item) => sum + (item['UAT Medium'] || 0), 0);
                    const uatLow = releaseData.reduce((sum, item) => sum + (item['UAT Low'] || 0), 0);
                    
                    // Valid bugs by severity
                    const validCritical = releaseData.reduce((sum, item) => sum + (item['Valid Critical'] || 0), 0);
                    const validHigh = releaseData.reduce((sum, item) => sum + (item['Valid High'] || 0), 0);
                    const validMedium = releaseData.reduce((sum, item) => sum + (item['Valid Medium'] || 0), 0);
                    const validLow = releaseData.reduce((sum, item) => sum + (item['Valid Low'] || 0), 0);
                    
                    // Invalid bugs by severity
                    const invalidCritical = releaseData.reduce((sum, item) => sum + (item['Invalid Critical'] || 0), 0);
                    const invalidHigh = releaseData.reduce((sum, item) => sum + (item['Invalid High'] || 0), 0);
                    const invalidMedium = releaseData.reduce((sum, item) => sum + (item['Invalid Medium'] || 0), 0);
                    const invalidLow = releaseData.reduce((sum, item) => sum + (item['Invalid Low'] || 0), 0);
                    
                    // PT Valid bugs by severity
                    const ptValidCritical = releaseData.reduce((sum, item) => sum + (item['PT Valid Critical'] || 0), 0);
                    const ptValidHigh = releaseData.reduce((sum, item) => sum + (item['PT Valid High'] || 0), 0);
                    const ptValidMedium = releaseData.reduce((sum, item) => sum + (item['PT Valid Medium'] || 0), 0);
                    const ptValidLow = releaseData.reduce((sum, item) => sum + (item['PT Valid Low'] || 0), 0);
                    
                    // PT Invalid bugs by severity
                    const ptInvalidCritical = releaseData.reduce((sum, item) => sum + (item['PT Invalid Critical'] || 0), 0);
                    const ptInvalidHigh = releaseData.reduce((sum, item) => sum + (item['PT Invalid High'] || 0), 0);
                    const ptInvalidMedium = releaseData.reduce((sum, item) => sum + (item['PT Invalid Medium'] || 0), 0);
                    const ptInvalidLow = releaseData.reduce((sum, item) => sum + (item['PT Invalid Low'] || 0), 0);
                    
                    // UAT Valid bugs by severity
                    const uatValidCritical = releaseData.reduce((sum, item) => sum + (item['UAT Valid Critical'] || 0), 0);
                    const uatValidHigh = releaseData.reduce((sum, item) => sum + (item['UAT Valid High'] || 0), 0);
                    const uatValidMedium = releaseData.reduce((sum, item) => sum + (item['UAT Valid Medium'] || 0), 0);
                    const uatValidLow = releaseData.reduce((sum, item) => sum + (item['UAT Valid Low'] || 0), 0);
                    
                    // UAT Invalid bugs by severity
                    const uatInvalidCritical = releaseData.reduce((sum, item) => sum + (item['UAT Invalid Critical'] || 0), 0);
                    const uatInvalidHigh = releaseData.reduce((sum, item) => sum + (item['UAT Invalid High'] || 0), 0);
                    const uatInvalidMedium = releaseData.reduce((sum, item) => sum + (item['UAT Invalid Medium'] || 0), 0);
                    const uatInvalidLow = releaseData.reduce((sum, item) => sum + (item['UAT Invalid Low'] || 0), 0);
                    
                    // Story metrics
                    const totalTestableStoriesCount = releaseData.reduce((sum, item) => sum + (item['Total Testable Stories Count'] || 0), 0);
                    const agentNoCount = releaseData.reduce((sum, item) => sum + (item['Agent Augmented No Count'] || 0), 0);
                    const delayedYesCount = releaseData.reduce((sum, item) => sum + (item['Delayed Delivery Yes Count'] || 0), 0);
                    
                    const avgBugsPerStory = totalStories > 0 ? (validBugs / totalStories).toFixed(2) : '0.00';
                    const avgBugsPerPoint = storyPoints > 0 ? (validBugs / storyPoints).toFixed(2) : '0.00';
                    
                    const agentNoPercentage = totalTestableStoriesCount > 0 ? ((agentNoCount / totalTestableStoriesCount) * 100).toFixed(1) : '0.0';
                    const delayedYesPercentage = totalTestableStoriesCount > 0 ? ((delayedYesCount / totalTestableStoriesCount) * 100).toFixed(1) : '0.0';
                    
                    aggregatedByRelease[release] = {
                        totalStories, testableStories, testingNAStories, storyPoints, totalBugs, validBugs, invalidBugs,
                        critical, high, medium, low,
                        ptBugs, ptCritical, ptHigh, ptMedium, ptLow,
                        uatBugs, uatCritical, uatHigh, uatMedium, uatLow,
                        validCritical, validHigh, validMedium, validLow,
                        invalidCritical, invalidHigh, invalidMedium, invalidLow,
                        ptValidCritical, ptValidHigh, ptValidMedium, ptValidLow,
                        ptInvalidCritical, ptInvalidHigh, ptInvalidMedium, ptInvalidLow,
                        uatValidCritical, uatValidHigh, uatValidMedium, uatValidLow,
                        uatInvalidCritical, uatInvalidHigh, uatInvalidMedium, uatInvalidLow,
                        avgBugsPerStory, avgBugsPerPoint,
                        totalTestableStoriesCount, agentNoCount, delayedYesCount,
                        agentNoPercentage, delayedYesPercentage,
                        scrumCount: new Set(releaseData.map(item => item['Node Name'])).size
                    };
                }
            });
            
            // Generate HTML for the POC data display
            let html = '<div class="release-comparison">';
            
            releases.forEach((release, idx) => {
                const data = aggregatedByRelease[release];
                if (data) {
                    // Dynamically determine release class based on index
                    const releaseClasses = ['release-sep', 'release-oct', 'release-nov', 'release-dec', 'release-jan'];
                    const releaseClass = releaseClasses[idx % releaseClasses.length];
                    
                    html += `
                        <div class="release-card">
                            <div class="release-label ${releaseClass}">${release}</div>
                            <ul class="metrics-list">
                                <li><span>Scrums:</span> <span>${data.scrumCount}</span></li>
                                <li><span>Total Stories:</span> <span>${data.totalStories}</span></li>
                                <li><span>Testable Stories:</span> <span>${data.testableStories}</span></li>
                                <li><span>Testing NA Stories:</span> <span>${data.testingNAStories}</span></li>
                                <li><span>Story Points:</span> <span>${data.storyPoints}</span></li>
                                <li><span>Bugs Raised:</span> <span>${data.totalBugs}</span></li>
                                <li><span>Valid bugs:</span> <span>${data.validBugs}</span></li>
                                <li><span>Invalid bugs:</span> <span>${data.invalidBugs}</span></li>
                                <li><span>Avg Bugs/Story:</span> <span>${data.avgBugsPerStory}</span></li>
                                <li><span>Avg Bugs/Story Point:</span> <span>${data.avgBugsPerPoint}</span></li>
                            </ul>
                            
                            <!-- Total Bugs Section -->
                            <div style="margin-top: 20px; padding-top: 15px; border-top: 2px solid #dee2e6;">
                                <h4 style="color: #495057; margin-bottom: 12px;">Total Bugs</h4>
                                <table style="width: 100%; border-collapse: collapse; font-size: 0.9em;">
                                    <thead>
                                        <tr style="background-color: #f8f9fa;">
                                            <th style="padding: 8px; text-align: left; border: 1px solid #dee2e6; font-weight: 600;">Bug Type</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #f8d7da;">C</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #fff3cd;">H</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #d1ecf1;">M</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #d4edda;">L</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">Total Bugs (${data.totalBugs})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.critical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.high}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.medium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.low}</td>
                                        </tr>
                                        <tr style="background-color: #f8f9fa;">
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">PT Bugs (${data.ptBugs})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptLow}</td>
                                        </tr>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">UAT Bugs (${data.uatBugs})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatLow}</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            
                            <!-- Valid Bugs Section -->
                            <div style="margin-top: 20px; padding-top: 15px; border-top: 2px solid #dee2e6;">
                                <h4 style="color: #495057; margin-bottom: 12px;">Valid Bugs</h4>
                                <table style="width: 100%; border-collapse: collapse; font-size: 0.9em;">
                                    <thead>
                                        <tr style="background-color: #f8f9fa;">
                                            <th style="padding: 8px; text-align: left; border: 1px solid #dee2e6; font-weight: 600;">Bug Type</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #f8d7da;">C</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #fff3cd;">H</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #d1ecf1;">M</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #d4edda;">L</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">Total Valid (${data.validBugs})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.validCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.validHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.validMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.validLow}</td>
                                        </tr>
                                        <tr style="background-color: #f8f9fa;">
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">PT Valid (${data.ptValidCritical + data.ptValidHigh + data.ptValidMedium + data.ptValidLow})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptValidCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptValidHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptValidMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptValidLow}</td>
                                        </tr>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">UAT Valid (${data.uatValidCritical + data.uatValidHigh + data.uatValidMedium + data.uatValidLow})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatValidCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatValidHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatValidMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatValidLow}</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            
                            <!-- Invalid Bugs Section -->
                            <div style="margin-top: 20px; padding-top: 15px; border-top: 2px solid #dee2e6;">
                                <h4 style="color: #495057; margin-bottom: 12px;">Invalid Bugs</h4>
                                <table style="width: 100%; border-collapse: collapse; font-size: 0.9em;">
                                    <thead>
                                        <tr style="background-color: #f8f9fa;">
                                            <th style="padding: 8px; text-align: left; border: 1px solid #dee2e6; font-weight: 600;">Bug Type</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #f8d7da;">C</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #fff3cd;">H</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #d1ecf1;">M</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600; background-color: #d4edda;">L</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">Total Invalid (${data.invalidBugs})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.invalidCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.invalidHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.invalidMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.invalidLow}</td>
                                        </tr>
                                        <tr style="background-color: #f8f9fa;">
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">PT Invalid (${data.ptInvalidCritical + data.ptInvalidHigh + data.ptInvalidMedium + data.ptInvalidLow})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptInvalidCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptInvalidHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptInvalidMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.ptInvalidLow}</td>
                                        </tr>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">UAT Invalid (${data.uatInvalidCritical + data.uatInvalidHigh + data.uatInvalidMedium + data.uatInvalidLow})</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatInvalidCritical}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatInvalidHigh}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatInvalidMedium}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.uatInvalidLow}</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            
                            <!-- Story Delivery Metrics Section -->
                            <div style="margin-top: 20px; padding-top: 15px; border-top: 2px solid #dee2e6;">
                                <h4 style="color: #495057; margin-bottom: 12px;">Story Delivery Metrics</h4>
                                <table style="width: 100%; border-collapse: collapse; font-size: 0.9em;">
                                    <thead>
                                        <tr style="background-color: #f8f9fa;">
                                            <th style="padding: 8px; text-align: left; border: 1px solid #dee2e6; font-weight: 600;">Metric</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600;">Count</th>
                                            <th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: 600;">Percentage</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">Total Testable Stories</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.totalTestableStoriesCount}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">100.0%</td>
                                        </tr>
                                        <tr style="background-color: #f8f9fa;">
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">Agent Augmented = No</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.agentNoCount}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.agentNoPercentage}%</td>
                                        </tr>
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #dee2e6; font-weight: 500;">Delayed Delivery = Yes</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.delayedYesCount}</td>
                                            <td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">${data.delayedYesPercentage}%</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;
                }
            });
            
            html += '</div>';
            
            // Update the POC data container
            document.getElementById('poc-data-container').innerHTML = html;
        }
        
        function resetPOCFilters() {
            // Reset all filters to 'all'
            document.getElementById('poc-ad-filter').value = 'all';
            document.getElementById('poc-sm-filter').value = 'all';
            document.getElementById('poc-sm-filter').disabled = true;
            document.getElementById('poc-sm-filter').innerHTML = '<option value="all">All SM POCs</option>';
            document.getElementById('poc-m-filter').value = 'all';
            document.getElementById('poc-m-filter').disabled = true;
            document.getElementById('poc-m-filter').innerHTML = '<option value="all">All M POCs</option>';
            document.getElementById('poc-node-filter').value = 'all';
            document.getElementById('poc-node-filter').disabled = true;
            document.getElementById('poc-node-filter').innerHTML = '<option value="all">All Nodes</option>';
            
            // Re-render with all data
            renderPOCData();
        }
        
        // Initialize POC data when tab is first shown
        document.addEventListener('DOMContentLoaded', function() {
            renderPOCData();
        });
        
"""
html_content += """
        function showTab(tabId) {
            // Hide all tab contents
            const tabContents = document.querySelectorAll('.tab-content');
            tabContents.forEach(content => {
                content.classList.remove('active');
            });

            // Remove active class from all tabs
            const tabs = document.querySelectorAll('.tab');
            tabs.forEach(tab => {
                tab.classList.remove('active');
            });

            // Show selected tab content
            document.getElementById(tabId).classList.add('active');

            // Add active class to clicked tab
            event.target.classList.add('active');
            
            // Render POC data when POC metrics tab is shown
            if (tabId === 'poc-metrics') {
                renderPOCData();
            }
        }

        function updateADFilters() {
            const adFilter = document.getElementById('ad-ad-filter');
            const selectedAD = adFilter.value;

            // Apply filter to AD POC items
            const adPocItems = document.querySelectorAll('.ad-poc-item');
            adPocItems.forEach(item => {
                const itemAD = item.getAttribute('data-ad-poc');
                
                if (selectedAD === 'all' || itemAD === selectedAD) {
                    item.style.display = 'block';
                } else {
                    item.style.display = 'none';
                }
            });
        }

        function updateSMFilters(changedLevel) {
            const adFilter = document.getElementById('sm-ad-filter');
            const smFilter = document.getElementById('sm-sm-filter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;

            // Update SM POC filter based on AD selection
            if (changedLevel === 'ad') {
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                
                if (selectedAD !== 'all' && smPocHierarchy[selectedAD]) {
                    const smPocs = smPocHierarchy[selectedAD].sort();
                    smPocs.forEach(sm => {
                        smFilter.innerHTML += `<option value="${sm}">${sm}</option>`;
                    });
                } else if (selectedAD === 'all') {
                    // Show all SM POCs
                    const allSMPocs = new Set();
                    Object.values(smPocHierarchy).forEach(smList => {
                        smList.forEach(sm => allSMPocs.add(sm));
                    });
                    Array.from(allSMPocs).sort().forEach(sm => {
                        smFilter.innerHTML += `<option value="${sm}">${sm}</option>`;
                    });
                }
            }

            // Apply filters to SM POC items
            const smPocItems = document.querySelectorAll('.sm-poc-item');
            smPocItems.forEach(item => {
                const itemAD = item.getAttribute('data-ad-poc');
                const itemSM = item.getAttribute('data-sm-poc');
                
                let showItem = true;
                
                if (selectedAD !== 'all' && itemAD !== selectedAD) {
                    showItem = false;
                }
                if (selectedSM !== 'all' && itemSM !== selectedSM) {
                    showItem = false;
                }
                
                item.style.display = showItem ? 'block' : 'none';
            });
        }

        function updateMFilters(changedLevel) {
            const adFilter = document.getElementById('m-ad-filter');
            const smFilter = document.getElementById('m-sm-filter');
            const mFilter = document.getElementById('m-m-filter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;
            const selectedM = mFilter.value;

            // Update SM POC filter based on AD selection
            if (changedLevel === 'ad') {
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                if (selectedAD !== 'all' && mPocHierarchy[selectedAD]) {
                    const smPocs = Object.keys(mPocHierarchy[selectedAD]).sort();
                    smPocs.forEach(sm => {
                        smFilter.innerHTML += `<option value="${sm}">${sm}</option>`;
                    });
                } else if (selectedAD === 'all') {
                    // Show all SM POCs
                    const allSMPocs = new Set();
                    Object.values(mPocHierarchy).forEach(smObj => {
                        Object.keys(smObj).forEach(sm => allSMPocs.add(sm));
                    });
                    Array.from(allSMPocs).sort().forEach(sm => {
                        smFilter.innerHTML += `<option value="${sm}">${sm}</option>`;
                    });
                }
            }

            // Update M POC filter based on SM selection
            if (changedLevel === 'ad' || changedLevel === 'sm') {
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                if (selectedAD !== 'all' && selectedSM !== 'all' && 
                    mPocHierarchy[selectedAD] && mPocHierarchy[selectedAD][selectedSM]) {
                    const mPocs = mPocHierarchy[selectedAD][selectedSM].sort();
                    mPocs.forEach(m => {
                        mFilter.innerHTML += `<option value="${m}">${m}</option>`;
                    });
                } else if (selectedSM !== 'all') {
                    // Show all M POCs under selected SM across all ADs
                    const allMPocs = new Set();
                    Object.values(mPocHierarchy).forEach(smObj => {
                        if (smObj[selectedSM]) {
                            smObj[selectedSM].forEach(m => allMPocs.add(m));
                        }
                    });
                    Array.from(allMPocs).sort().forEach(m => {
                        mFilter.innerHTML += `<option value="${m}">${m}</option>`;
                    });
                } else if (selectedAD !== 'all') {
                    // Show all M POCs under selected AD
                    const allMPocs = new Set();
                    if (mPocHierarchy[selectedAD]) {
                        Object.values(mPocHierarchy[selectedAD]).forEach(mList => {
                            mList.forEach(m => allMPocs.add(m));
                        });
                    }
                    Array.from(allMPocs).sort().forEach(m => {
                        mFilter.innerHTML += `<option value="${m}">${m}</option>`;
                    });
                }
            }

            // Apply filters to M POC items
            const mPocItems = document.querySelectorAll('.m-poc-item');
            mPocItems.forEach(item => {
                const itemAD = item.getAttribute('data-ad-poc');
                const itemSM = item.getAttribute('data-sm-poc');
                const itemM = item.getAttribute('data-m-poc');
                
                let showItem = true;
                
                if (selectedAD !== 'all' && itemAD !== selectedAD) {
                    showItem = false;
                }
                if (selectedSM !== 'all' && itemSM !== selectedSM) {
                    showItem = false;
                }
                if (selectedM !== 'all' && itemM !== selectedM) {
                    showItem = false;
                }
                
                item.style.display = showItem ? 'block' : 'none';
            });
        }
    </script>
</body>
</html>
"""

# Write to file
output_file = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Dashboard_Complete_Quality_Metrics.html'
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\nDashboard generated successfully!")
print(f"Output file: {output_file}")
print(f"\nSummary:")
print(f"Total Scrums: {total_scrums}")
print(f"Total Stories: {total_stories}")
print(f"Total Story Points: {total_story_points}")
print(f"Total Bugs: {total_bugs}")
print(f"AD POCs: {len(top_ad_pocs)}")
print(f"SM POCs: {len(top_sm_pocs)}")
print(f"M POCs: {len(top_m_pocs)}")
