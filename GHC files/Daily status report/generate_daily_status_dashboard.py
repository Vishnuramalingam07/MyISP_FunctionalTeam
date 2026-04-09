import pandas as pd
import os
import subprocess
import sys
from datetime import datetime

# Define paths
base_dir = r"C:\Users\vishnu.ramalingam\MyISP_Tools\GHC files\Daily status report"
python_exe = sys.executable

# Scripts to run in order
script1 = os.path.join(base_dir, "process_uat_status.py")
script2 = os.path.join(base_dir, "generate_bug_summary.py")
script3 = os.path.join(base_dir, "generate_story_summary_detailed.py")
script4 = os.path.join(base_dir, "generate_overall_defect_summary.py")

print("=" * 80)
print("DAILY STATUS REPORT GENERATOR")
print("=" * 80)
print()

# Run process_uat_status.py
print("Step 1: Processing UAT Status...")
print("-" * 80)
result1 = subprocess.run([python_exe, script1], capture_output=False)
if result1.returncode != 0:
    print("Error running process_uat_status.py")
    sys.exit(1)
print()

# Run generate_bug_summary.py
print("Step 2: Generating Bug Summary...")
print("-" * 80)
result2 = subprocess.run([python_exe, script2], capture_output=False)
if result2.returncode != 0:
    print("Error running generate_bug_summary.py")
    sys.exit(1)
print()

# Run generate_story_summary_detailed.py
print("Step 3: Generating Story Summary...")
print("-" * 80)
result3 = subprocess.run([python_exe, script3], capture_output=False)
if result3.returncode != 0:
    print("Error running generate_story_summary_detailed.py")
    sys.exit(1)
print()

# Run generate_overall_defect_summary.py
print("Step 4: Generating Overall Defect Summary...")
print("-" * 80)
result4 = subprocess.run([python_exe, script4], capture_output=False)
if result4.returncode != 0:
    print("Error running generate_overall_defect_summary.py")
    sys.exit(1)
print()

# Read generated Excel files
print("Step 5: Reading generated Excel files...")
print("-" * 80)
story_summary_path = os.path.join(base_dir, "Story_summary.xlsx")
bug_summary_path = os.path.join(base_dir, "Bug_summary.xlsx")
milestone_dates_path = os.path.join(base_dir, "Milestone dates.xlsx")
overall_defect_summary_path = os.path.join(base_dir, "Over all Defect Summary.xlsx")

# Read Story Summary
story_summary_df = pd.read_excel(story_summary_path, sheet_name='Story Summary')
testable_stories_df = pd.read_excel(story_summary_path, sheet_name='Testable stories')

# Read Consolidated PT_UAT Status
consolidated_pt_uat_df = pd.read_excel(story_summary_path, sheet_name='Consolidated PT_UAT Status')

# Replace "UAT NOT Applicable" with "UAT NA" for better column width
consolidated_pt_uat_df = consolidated_pt_uat_df.replace('UAT NOT Applicable', 'UAT NA')

# Get unique Product Owners for filter
unique_product_owners = sorted(consolidated_pt_uat_df['Product Owner'].dropna().unique().tolist())

# Read Bug Summary
bug_summary_df = pd.read_excel(bug_summary_path)

# Sort bug_summary_df by Severity (Critical, High, Medium, Low)
severity_order = {'1 - Critical': 0, '2 - High': 1, '3 - Medium': 2, '4 - Low': 3}
bug_summary_df['Severity_Order'] = bug_summary_df['Severity'].map(severity_order)
bug_summary_df = bug_summary_df.sort_values('Severity_Order').drop('Severity_Order', axis=1)

# Read Milestone Dates
milestone_dates_df = pd.read_excel(milestone_dates_path)

# Read Overall Defect Summary
overall_defect_summary_df = pd.read_excel(overall_defect_summary_path)

print(f"  Story Summary: {len(story_summary_df)} total stories")
print(f"  Testable Stories: {len(testable_stories_df)} PT testable stories")
print(f"  Bug Summary: {len(bug_summary_df)} nodes with bugs")
print(f"  Milestone Dates: {len(milestone_dates_df)} milestones")
print(f"  Overall Defect Summary: {len(overall_defect_summary_df)} total defects")
print()

# Calculate metrics from Story Summary
print("Step 6: Calculating metrics...")
print("-" * 80)

total_stories = len(story_summary_df)
testing_na_stories = int(story_summary_df['Testing NA stories'].sum())
pt_testable_stories = int(story_summary_df['PT testable stories'].sum())
uat_testable_stories = int(story_summary_df['UAT Testable Stories'].sum())
pt_delivered = int(story_summary_df['PT delivered'].sum())
pt_not_delivered = int(story_summary_df['PT NOT delivered'].sum())
uat_delivered = int(story_summary_df['UAT delivered'].sum())
uat_not_delivered = int(story_summary_df['UAT NOT delivered'].sum())

# Calculate PT Execution % and Pass % from testable stories
# Convert to numeric, treating empty/non-numeric as 0
testable_stories_df['Passed'] = pd.to_numeric(testable_stories_df['Passed'], errors='coerce').fillna(0)
testable_stories_df['Failed'] = pd.to_numeric(testable_stories_df['Failed'], errors='coerce').fillna(0)
testable_stories_df['Blocked'] = pd.to_numeric(testable_stories_df['Blocked'], errors='coerce').fillna(0)
testable_stories_df['Not Run'] = pd.to_numeric(testable_stories_df['Not Run'], errors='coerce').fillna(0)

total_passed = int(testable_stories_df['Passed'].sum())
total_failed = int(testable_stories_df['Failed'].sum())
total_blocked = int(testable_stories_df['Blocked'].sum())
total_not_run = int(testable_stories_df['Not Run'].sum())
total_tests = total_passed + total_failed + total_blocked + total_not_run

# PT Execution % = (Passed + Failed) / Total
pt_execution_pct = ((total_passed + total_failed) / total_tests * 100) if total_tests > 0 else 0

# PT Pass % = Passed / (Passed + Failed + Blocked)
pt_pass_pct = (total_passed / (total_passed + total_failed + total_blocked) * 100) if (total_passed + total_failed + total_blocked) > 0 else 0

# Read UAT metrics from Status excel - UAT status sheet
status_excel_path = os.path.join(base_dir, "Status excel.xlsx")
uat_status_df = pd.read_excel(status_excel_path, sheet_name='UAT status')
if len(uat_status_df) > 0:
    uat_row = uat_status_df.iloc[0]
    # Use simple conversion with default values
    uat_total_tests = int(uat_row.get('Total UAT Test Cases', 0)) if pd.notna(uat_row.get('Total UAT Test Cases')) else 0
    uat_passed = int(uat_row.get('Passed', 0)) if pd.notna(uat_row.get('Passed')) else 0
    uat_failed = int(uat_row.get('Failed', 0)) if pd.notna(uat_row.get('Failed')) else 0
    uat_blocked = int(uat_row.get('Blocked', 0)) if pd.notna(uat_row.get('Blocked')) else 0
    uat_not_run = int(uat_row.get('Not Run', 0)) if pd.notna(uat_row.get('Not Run')) else 0
    uat_exec_val = uat_row.get('UAT Execution %', 0)
    uat_execution_pct = float(uat_exec_val * 100) if pd.notna(uat_exec_val) else 0  # Convert to percentage
    uat_pass_val = uat_row.get('UAT Pass %', 0)
    uat_pass_pct = float(uat_pass_val * 100) if pd.notna(uat_pass_val) else 0  # Convert to percentage
else:
    uat_total_tests = uat_passed = uat_failed = uat_blocked = uat_not_run = 0
    uat_execution_pct = uat_pass_pct = 0

# Calculate bug metrics from the new format
total_bugs = len(bug_summary_df)

# Map severity to counts
severity_map = {'1 - Critical': 0, '2 - High': 0, '3 - Medium': 0, '4 - Low': 0}
for severity in bug_summary_df['Severity'].value_counts().items():
    severity_map[severity[0]] = severity[1]

critical_bugs = severity_map.get('1 - Critical', 0)
high_bugs = severity_map.get('2 - High', 0)
medium_bugs = severity_map.get('3 - Medium', 0)
low_bugs = severity_map.get('4 - Low', 0)

# Calculate Active defects (State NOT in 'Resolved', 'Ready to Deploy')
active_bugs_df = bug_summary_df[~bug_summary_df['State'].isin(['Resolved', 'Ready to Deploy'])]
total_active_bugs = len(active_bugs_df)
active_critical = len(active_bugs_df[active_bugs_df['Severity'] == '1 - Critical'])
active_high = len(active_bugs_df[active_bugs_df['Severity'] == '2 - High'])
active_medium = len(active_bugs_df[active_bugs_df['Severity'] == '3 - Medium'])
active_low = len(active_bugs_df[active_bugs_df['Severity'] == '4 - Low'])

# Calculate RTD defects (State in 'Resolved', 'Ready to Deploy')
rtd_bugs_df = bug_summary_df[bug_summary_df['State'].isin(['Resolved', 'Ready to Deploy'])]
total_rtd_bugs = len(rtd_bugs_df)
rtd_critical = len(rtd_bugs_df[rtd_bugs_df['Severity'] == '1 - Critical'])
rtd_high = len(rtd_bugs_df[rtd_bugs_df['Severity'] == '2 - High'])
rtd_medium = len(rtd_bugs_df[rtd_bugs_df['Severity'] == '3 - Medium'])
rtd_low = len(rtd_bugs_df[rtd_bugs_df['Severity'] == '4 - Low'])

# Prepare Execution Summary data grouped by Parent and Parent Title

# Group by Parent, Parent Title, AD POC, SM POC, and M POC to show separate rows for each POC combination
# This ensures that if a Parent has stories from multiple M POCs, they appear as separate rows
execution_summary_df = testable_stories_df.groupby(['Parent', 'Parent Title', 'AD POC', 'SM POC', 'M POC']).agg({
    'ID': ['count', lambda x: ','.join(map(str, x.unique()))],  # Count and list of IDs
    'PT delivered': 'sum',
    'PT NOT delivered': 'sum',
    'UAT Testable Stories': 'sum',
    'UAT delivered': 'sum',
    'UAT NOT delivered': 'sum',
    'Passed': 'sum',
    'Failed': 'sum',
    'Blocked': 'sum',
    'Not Run': 'sum'
}).reset_index()

# Flatten column names
execution_summary_df.columns = ['Parent', 'Parent Title', 'AD POC', 'SM POC', 'M POC', 
                                 'Total Stories', 'Story IDs',
                                 'PT delivered', 'PT NOT delivered', 'UAT Testable Stories',
                                 'UAT delivered', 'UAT NOT delivered', 'Passed', 'Failed', 'Blocked', 'Not Run']

# Calculate Total = Passed + Failed + Blocked + Not Run
execution_summary_df['Total'] = (
    execution_summary_df['Passed'] + 
    execution_summary_df['Failed'] + 
    execution_summary_df['Blocked'] + 
    execution_summary_df['Not Run']
)

# Calculate Execution % and Pass % from numeric columns as decimal (0-1 range for proper % formatting)
# Execution % = (Passed + Failed) / Total (only if Total > 0)
execution_summary_df['Execution %'] = execution_summary_df.apply(
    lambda row: round((row['Passed'] + row['Failed']) / row['Total'], 4) if row['Total'] > 0 else 0,
    axis=1
)

# Pass % = Passed / (Passed + Failed + Blocked) (only if denominator > 0)
execution_summary_df['Pass %'] = execution_summary_df.apply(
    lambda row: round(row['Passed'] / (row['Passed'] + row['Failed'] + row['Blocked']), 4) 
    if (row['Passed'] + row['Failed'] + row['Blocked']) > 0 else 0,
    axis=1
)

# Create Module Level Execution Summary (grouped by TextVerification instead of Parent)
module_execution_summary_df = testable_stories_df.groupby(['TextVerification', 'AD POC', 'SM POC', 'M POC']).agg({
    'ID': ['count', lambda x: ','.join(map(str, x.unique()))],  # Count and list of IDs
    'PT delivered': 'sum',
    'PT NOT delivered': 'sum',
    'UAT Testable Stories': 'sum',
    'UAT delivered': 'sum',
    'UAT NOT delivered': 'sum',
    'Passed': 'sum',
    'Failed': 'sum',
    'Blocked': 'sum',
    'Not Run': 'sum'
}).reset_index()

# Flatten column names
module_execution_summary_df.columns = ['Module', 'AD POC', 'SM POC', 'M POC', 
                                       'Total Stories', 'Story IDs',
                                       'PT delivered', 'PT NOT delivered', 'UAT Testable Stories',
                                       'UAT delivered', 'UAT NOT delivered', 'Passed', 'Failed', 'Blocked', 'Not Run']

# Calculate Total = Passed + Failed + Blocked + Not Run
module_execution_summary_df['Total'] = (
    module_execution_summary_df['Passed'] + 
    module_execution_summary_df['Failed'] + 
    module_execution_summary_df['Blocked'] + 
    module_execution_summary_df['Not Run']
)

# Calculate Execution % and Pass % as decimal (0-1 range for proper % formatting)
module_execution_summary_df['Execution %'] = module_execution_summary_df.apply(
    lambda row: round((row['Passed'] + row['Failed']) / row['Total'], 4) if row['Total'] > 0 else 0,
    axis=1
)

module_execution_summary_df['Pass %'] = module_execution_summary_df.apply(
    lambda row: round(row['Passed'] / (row['Passed'] + row['Failed'] + row['Blocked']), 4) 
    if (row['Passed'] + row['Failed'] + row['Blocked']) > 0 else 0,
    axis=1
)

# Build POC hierarchy for filtering in Execution Summary
import json

# Create mappings from execution summary data
m_to_sm_mapping_exec = execution_summary_df.set_index('M POC')['SM POC'].to_dict()
m_to_ad_mapping_exec = execution_summary_df.set_index('M POC')['AD POC'].to_dict()
sm_to_ad_mapping_exec = execution_summary_df.set_index('SM POC')['AD POC'].to_dict()

# Build hierarchy: AD -> SM -> M
poc_hierarchy_exec = {}
for _, row in execution_summary_df.iterrows():
    ad_poc = row['AD POC']
    sm_poc = row['SM POC']
    m_poc = row['M POC']
    
    if ad_poc not in poc_hierarchy_exec:
        poc_hierarchy_exec[ad_poc] = {}
    if sm_poc not in poc_hierarchy_exec[ad_poc]:
        poc_hierarchy_exec[ad_poc][sm_poc] = []
    if m_poc not in poc_hierarchy_exec[ad_poc][sm_poc]:
        poc_hierarchy_exec[ad_poc][sm_poc].append(m_poc)

# Get unique lists for filters
unique_ad_pocs_exec = sorted(execution_summary_df['AD POC'].unique())
poc_hierarchy_json = json.dumps(poc_hierarchy_exec)

# Build POC hierarchy for Module Level Execution Summary
poc_hierarchy_module = {}
for _, row in module_execution_summary_df.iterrows():
    ad_poc = row['AD POC']
    sm_poc = row['SM POC']
    m_poc = row['M POC']
    
    if ad_poc not in poc_hierarchy_module:
        poc_hierarchy_module[ad_poc] = {}
    if sm_poc not in poc_hierarchy_module[ad_poc]:
        poc_hierarchy_module[ad_poc][sm_poc] = []
    if m_poc not in poc_hierarchy_module[ad_poc][sm_poc]:
        poc_hierarchy_module[ad_poc][sm_poc].append(m_poc)

unique_ad_pocs_module = sorted(module_execution_summary_df['AD POC'].unique())
poc_hierarchy_module_json = json.dumps(poc_hierarchy_module)

# Prepare Testing NA Stories summary from Story Summary sheet
# The Story Summary sheet already has aggregated Testing NA stories count by Node Name
# Filter for nodes that have Testing NA stories > 0
testing_na_df = story_summary_df[story_summary_df['Testing NA stories'] > 0].copy()

# Group by Node Name only and sum up Testing NA stories count
# Also get the first occurrence of POCs for each Node Name
testing_na_summary = testing_na_df.groupby('Node Name').agg({
    'AD POC': 'first',
    'SM POC': 'first', 
    'M POC': 'first',
    'Testing NA stories': 'sum'
}).reset_index()

# Rename column for display
testing_na_summary = testing_na_summary.rename(columns={'Testing NA stories': 'Testing NA Stories Count'})

# Calculate grand total (sum of all Testing NA stories counts)
testing_na_grand_total = int(testing_na_summary['Testing NA Stories Count'].sum())

# Calculate PT defects (StageFound NOT 'User Acceptance Test')
pt_bugs_df = bug_summary_df[bug_summary_df['StageFound'] != 'User Acceptance Test']
total_pt_bugs = len(pt_bugs_df)
pt_critical = len(pt_bugs_df[pt_bugs_df['Severity'] == '1 - Critical'])
pt_high = len(pt_bugs_df[pt_bugs_df['Severity'] == '2 - High'])
pt_medium = len(pt_bugs_df[pt_bugs_df['Severity'] == '3 - Medium'])
pt_low = len(pt_bugs_df[pt_bugs_df['Severity'] == '4 - Low'])

# Calculate PT Active defects
pt_active_bugs_df = pt_bugs_df[~pt_bugs_df['State'].isin(['Resolved', 'Ready to Deploy'])]
total_pt_active_bugs = len(pt_active_bugs_df)
pt_active_critical = len(pt_active_bugs_df[pt_active_bugs_df['Severity'] == '1 - Critical'])
pt_active_high = len(pt_active_bugs_df[pt_active_bugs_df['Severity'] == '2 - High'])
pt_active_medium = len(pt_active_bugs_df[pt_active_bugs_df['Severity'] == '3 - Medium'])
pt_active_low = len(pt_active_bugs_df[pt_active_bugs_df['Severity'] == '4 - Low'])

# Calculate PT RTD defects
pt_rtd_bugs_df = pt_bugs_df[pt_bugs_df['State'].isin(['Resolved', 'Ready to Deploy'])]
total_pt_rtd_bugs = len(pt_rtd_bugs_df)
pt_rtd_critical = len(pt_rtd_bugs_df[pt_rtd_bugs_df['Severity'] == '1 - Critical'])
pt_rtd_high = len(pt_rtd_bugs_df[pt_rtd_bugs_df['Severity'] == '2 - High'])
pt_rtd_medium = len(pt_rtd_bugs_df[pt_rtd_bugs_df['Severity'] == '3 - Medium'])
pt_rtd_low = len(pt_rtd_bugs_df[pt_rtd_bugs_df['Severity'] == '4 - Low'])

# Calculate UAT defects (StageFound = 'User Acceptance Test')
uat_bugs_df = bug_summary_df[bug_summary_df['StageFound'] == 'User Acceptance Test']
total_uat_bugs = len(uat_bugs_df)
uat_critical = len(uat_bugs_df[uat_bugs_df['Severity'] == '1 - Critical'])
uat_high = len(uat_bugs_df[uat_bugs_df['Severity'] == '2 - High'])
uat_medium = len(uat_bugs_df[uat_bugs_df['Severity'] == '3 - Medium'])
uat_low = len(uat_bugs_df[uat_bugs_df['Severity'] == '4 - Low'])

# Calculate UAT Active defects
uat_active_bugs_df = uat_bugs_df[~uat_bugs_df['State'].isin(['Resolved', 'Ready to Deploy'])]
total_uat_active_bugs = len(uat_active_bugs_df)
uat_active_critical = len(uat_active_bugs_df[uat_active_bugs_df['Severity'] == '1 - Critical'])
uat_active_high = len(uat_active_bugs_df[uat_active_bugs_df['Severity'] == '2 - High'])
uat_active_medium = len(uat_active_bugs_df[uat_active_bugs_df['Severity'] == '3 - Medium'])
uat_active_low = len(uat_active_bugs_df[uat_active_bugs_df['Severity'] == '4 - Low'])

# Calculate UAT RTD defects
uat_rtd_bugs_df = uat_bugs_df[uat_bugs_df['State'].isin(['Resolved', 'Ready to Deploy'])]
total_uat_rtd_bugs = len(uat_rtd_bugs_df)
uat_rtd_critical = len(uat_rtd_bugs_df[uat_rtd_bugs_df['Severity'] == '1 - Critical'])
uat_rtd_high = len(uat_rtd_bugs_df[uat_rtd_bugs_df['Severity'] == '2 - High'])
uat_rtd_medium = len(uat_rtd_bugs_df[uat_rtd_bugs_df['Severity'] == '3 - Medium'])
uat_rtd_low = len(uat_rtd_bugs_df[uat_rtd_bugs_df['Severity'] == '4 - Low'])

# Create POC-wise breakdown
def get_poc_breakdown(df, poc_column):
    breakdown = []
    for poc in df[poc_column].unique():
        if pd.isna(poc) or str(poc).strip() == '':
            poc = 'yet to assign'
        poc_df = df[df[poc_column] == poc] if poc != 'yet to assign' else df[df[poc_column].isna() | (df[poc_column] == '')]
        total = len(poc_df)
        critical = len(poc_df[poc_df['Severity'] == '1 - Critical'])
        high = len(poc_df[poc_df['Severity'] == '2 - High'])
        medium = len(poc_df[poc_df['Severity'] == '3 - Medium'])
        low = len(poc_df[poc_df['Severity'] == '4 - Low'])
        breakdown.append({
            'POC': poc,
            'Critical': critical,
            'High': high,
            'Medium': medium,
            'Low': low,
            'Total': total
        })
    df_result = pd.DataFrame(breakdown)
    # Separate 'yet to assign' rows and sort the rest by Total
    yet_to_assign = df_result[df_result['POC'] == 'yet to assign']
    others = df_result[df_result['POC'] != 'yet to assign'].sort_values('Total', ascending=False)
    # Concatenate with 'yet to assign' at the bottom
    return pd.concat([others, yet_to_assign], ignore_index=True)

ad_poc_breakdown = get_poc_breakdown(bug_summary_df, 'AD POC')
sm_poc_breakdown = get_poc_breakdown(bug_summary_df, 'SM POC')
m_poc_breakdown = get_poc_breakdown(bug_summary_df, 'M POC')

# Create detailed POC-wise defect breakdown for filtering
def get_detailed_poc_breakdown(df):
    """Get defect counts by POC with breakdown by Total/Active/RTD and PT/UAT"""
    # Create separate breakdowns for AD, SM, and M POCs
    ad_breakdown = {}
    sm_breakdown = {}
    m_breakdown = {}
    
    for _, row in df.iterrows():
        ad_poc = str(row.get('AD POC', '')).strip()
        sm_poc = str(row.get('SM POC', '')).strip()
        m_poc = str(row.get('M POC', '')).strip()
        severity = row.get('Severity', '')
        state = row.get('State', '')
        stage_found = row.get('StageFound', '')
        
        # Assign default value for empty POC values
        if not ad_poc or ad_poc == 'nan':
            ad_poc = 'Unassigned'
        if not sm_poc or sm_poc == 'nan':
            sm_poc = 'Unassigned'
        if not m_poc or m_poc == 'nan':
            m_poc = 'Unassigned'
            
        # Determine if PT or UAT
        is_uat = (stage_found == 'User Acceptance Test')
        is_pt = not is_uat
        
        # Determine if Active or RTD
        is_active = state not in ['Resolved', 'Ready to Deploy']
        is_rtd = not is_active
        
        # Map severity to key
        sev_map = {
            '1 - Critical': 'critical',
            '2 - High': 'high',
            '3 - Medium': 'medium',
            '4 - Low': 'low'
        }
        sev_key = sev_map.get(severity, 'low')
        
        # Helper function to update counts
        def update_poc_data(breakdown_dict, poc_key):
            if poc_key and poc_key not in breakdown_dict:
                breakdown_dict[poc_key] = {
                    'total': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'active': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'rtd': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'pt_total': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'pt_active': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'pt_rtd': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'uat_total': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'uat_active': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                    'uat_rtd': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0}
                }
            
            if poc_key:
                # Total counts
                breakdown_dict[poc_key]['total']['total'] += 1
                breakdown_dict[poc_key]['total'][sev_key] += 1
                
                # Active/RTD counts
                if is_active:
                    breakdown_dict[poc_key]['active']['total'] += 1
                    breakdown_dict[poc_key]['active'][sev_key] += 1
                else:
                    breakdown_dict[poc_key]['rtd']['total'] += 1
                    breakdown_dict[poc_key]['rtd'][sev_key] += 1
                
                # PT/UAT breakdown
                if is_pt:
                    breakdown_dict[poc_key]['pt_total']['total'] += 1
                    breakdown_dict[poc_key]['pt_total'][sev_key] += 1
                    if is_active:
                        breakdown_dict[poc_key]['pt_active']['total'] += 1
                        breakdown_dict[poc_key]['pt_active'][sev_key] += 1
                    else:
                        breakdown_dict[poc_key]['pt_rtd']['total'] += 1
                        breakdown_dict[poc_key]['pt_rtd'][sev_key] += 1
                else:  # UAT
                    breakdown_dict[poc_key]['uat_total']['total'] += 1
                    breakdown_dict[poc_key]['uat_total'][sev_key] += 1
                    if is_active:
                        breakdown_dict[poc_key]['uat_active']['total'] += 1
                        breakdown_dict[poc_key]['uat_active'][sev_key] += 1
                    else:
                        breakdown_dict[poc_key]['uat_rtd']['total'] += 1
                        breakdown_dict[poc_key]['uat_rtd'][sev_key] += 1
        
        # Update each POC type separately
        update_poc_data(ad_breakdown, ad_poc)
        update_poc_data(sm_breakdown, sm_poc)
        update_poc_data(m_breakdown, m_poc)
    
    # Merge all breakdowns into one structure with prefixes
    combined_breakdown = {}
    for poc, data in ad_breakdown.items():
        combined_breakdown[f"AD:{poc}"] = data
    for poc, data in sm_breakdown.items():
        combined_breakdown[f"SM:{poc}"] = data
    for poc, data in m_breakdown.items():
        combined_breakdown[f"M:{poc}"] = data
    
    return combined_breakdown

defect_summary_data = get_detailed_poc_breakdown(bug_summary_df)
defect_summary_json = json.dumps(defect_summary_data)

# Calculate initial display values by aggregating all M: entries (to avoid triple-counting)
initial_data = {
    'total': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'active': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'rtd': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'pt_total': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'pt_active': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'pt_rtd': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'uat_total': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'uat_active': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
    'uat_rtd': {'total': 0, 'critical': 0, 'high': 0, 'medium': 0, 'low': 0}
}

# Sum all M: entries
for key, poc_data in defect_summary_data.items():
    if key.startswith('M:'):
        for category in initial_data:
            for severity in initial_data[category]:
                initial_data[category][severity] += poc_data[category][severity]

# Update display variables for Total Defect breakdown table
total_bugs = initial_data['total']['total']
critical_bugs = initial_data['total']['critical']
high_bugs = initial_data['total']['high']
medium_bugs = initial_data['total']['medium']
low_bugs = initial_data['total']['low']

total_active_bugs = initial_data['active']['total']
active_critical = initial_data['active']['critical']
active_high = initial_data['active']['high']
active_medium = initial_data['active']['medium']
active_low = initial_data['active']['low']

total_rtd_bugs = initial_data['rtd']['total']
rtd_critical = initial_data['rtd']['critical']
rtd_high = initial_data['rtd']['high']
rtd_medium = initial_data['rtd']['medium']
rtd_low = initial_data['rtd']['low']

# Update display variables for PT Defect Breakdown table
total_pt_bugs = initial_data['pt_total']['total']
pt_critical = initial_data['pt_total']['critical']
pt_high = initial_data['pt_total']['high']
pt_medium = initial_data['pt_total']['medium']
pt_low = initial_data['pt_total']['low']

total_pt_active_bugs = initial_data['pt_active']['total']
pt_active_critical = initial_data['pt_active']['critical']
pt_active_high = initial_data['pt_active']['high']
pt_active_medium = initial_data['pt_active']['medium']
pt_active_low = initial_data['pt_active']['low']

total_pt_rtd_bugs = initial_data['pt_rtd']['total']
pt_rtd_critical = initial_data['pt_rtd']['critical']
pt_rtd_high = initial_data['pt_rtd']['high']
pt_rtd_medium = initial_data['pt_rtd']['medium']
pt_rtd_low = initial_data['pt_rtd']['low']

# Update display variables for UAT Defect Breakdown table
total_uat_bugs = initial_data['uat_total']['total']
uat_critical = initial_data['uat_total']['critical']
uat_high = initial_data['uat_total']['high']
uat_medium = initial_data['uat_total']['medium']
uat_low = initial_data['uat_total']['low']

total_uat_active_bugs = initial_data['uat_active']['total']
uat_active_critical = initial_data['uat_active']['critical']
uat_active_high = initial_data['uat_active']['high']
uat_active_medium = initial_data['uat_active']['medium']
uat_active_low = initial_data['uat_active']['low']

total_uat_rtd_bugs = initial_data['uat_rtd']['total']
uat_rtd_critical = initial_data['uat_rtd']['critical']
uat_rtd_high = initial_data['uat_rtd']['high']
uat_rtd_medium = initial_data['uat_rtd']['medium']
uat_rtd_low = initial_data['uat_rtd']['low']

# Get unique states from bug_summary_df for the State filter
unique_states = sorted(bug_summary_df['State'].unique())

# Function to categorize defects based on Defect Record
def categorize_defect(defect_record):
    if pd.isna(defect_record):
        return 'Feature related'
    defect_record_str = str(defect_record).strip()
    if defect_record_str.lower() == 'sanity' or defect_record_str.lower() == 'prod sanity':
        return 'Sanity'
    elif defect_record_str.lower() == 'regression':
        return 'Regression'
    elif 'existing prod issue' in defect_record_str.lower() or defect_record_str.lower() == 'prod user bugs':
        return 'Existing Prod Issue'
    else:
        return 'Feature related'

# Add Defect Category column to bug_summary_df
bug_summary_df['Defect Category'] = bug_summary_df['Defect Record'].apply(categorize_defect)

# Create Node breakdown
node_breakdown = []
for node in bug_summary_df['Node Name'].unique():
    node_df = bug_summary_df[bug_summary_df['Node Name'] == node]
    ad_poc = node_df['AD POC'].iloc[0] if not node_df['AD POC'].isna().all() and str(node_df['AD POC'].iloc[0]).strip() != '' else 'yet to assign'
    sm_poc = node_df['SM POC'].iloc[0] if not node_df['SM POC'].isna().all() and str(node_df['SM POC'].iloc[0]).strip() != '' else 'yet to assign'
    m_poc = node_df['M POC'].iloc[0] if not node_df['M POC'].isna().all() and str(node_df['M POC'].iloc[0]).strip() != '' else 'yet to assign'
    
    # Classify by StageFound: UAT if "User Acceptance Test", otherwise PT
    pt_df = node_df[node_df['StageFound'] != 'User Acceptance Test']
    uat_df = node_df[node_df['StageFound'] == 'User Acceptance Test']
    
    # Calculate counts for PT
    pt_total = len(pt_df)
    pt_critical = len(pt_df[pt_df['Severity'] == '1 - Critical'])
    pt_high = len(pt_df[pt_df['Severity'] == '2 - High'])
    pt_medium = len(pt_df[pt_df['Severity'] == '3 - Medium'])
    pt_low = len(pt_df[pt_df['Severity'] == '4 - Low'])
    
    # Calculate counts for UAT
    uat_total = len(uat_df)
    uat_critical = len(uat_df[uat_df['Severity'] == '1 - Critical'])
    uat_high = len(uat_df[uat_df['Severity'] == '2 - High'])
    uat_medium = len(uat_df[uat_df['Severity'] == '3 - Medium'])
    uat_low = len(uat_df[uat_df['Severity'] == '4 - Low'])
    
    total = len(node_df)
    critical = len(node_df[node_df['Severity'] == '1 - Critical'])
    high = len(node_df[node_df['Severity'] == '2 - High'])
    medium = len(node_df[node_df['Severity'] == '3 - Medium'])
    low = len(node_df[node_df['Severity'] == '4 - Low'])
    
    # Calculate counts for each state (severity breakdown)
    state_counts = {}
    for state in unique_states:
        state_df = node_df[node_df['State'] == state]
        state_counts[state] = {
            'Critical': len(state_df[state_df['Severity'] == '1 - Critical']),
            'High': len(state_df[state_df['Severity'] == '2 - High']),
            'Medium': len(state_df[state_df['Severity'] == '3 - Medium']),
            'Low': len(state_df[state_df['Severity'] == '4 - Low']),
            'Total': len(state_df)
        }
    
    # Calculate counts for each defect category (severity breakdown)
    defect_categories = ['Sanity', 'Regression', 'Existing Prod Issue', 'Feature related']
    category_counts = {}
    for category in defect_categories:
        category_df = node_df[node_df['Defect Category'] == category]
        category_counts[category] = {
            'Critical': len(category_df[category_df['Severity'] == '1 - Critical']),
            'High': len(category_df[category_df['Severity'] == '2 - High']),
            'Medium': len(category_df[category_df['Severity'] == '3 - Medium']),
            'Low': len(category_df[category_df['Severity'] == '4 - Low']),
            'Total': len(category_df)
        }
    
    node_data = {
        'Node Name': node,
        'AD POC': ad_poc,
        'SM POC': sm_poc,
        'M POC': m_poc,
        'Critical': critical,
        'High': high,
        'Medium': medium,
        'Low': low,
        'Total': total,
        'PT_Critical': pt_critical,
        'PT_High': pt_high,
        'PT_Medium': pt_medium,
        'PT_Low': pt_low,
        'PT_Total': pt_total,
        'UAT_Critical': uat_critical,
        'UAT_High': uat_high,
        'UAT_Medium': uat_medium,
        'UAT_Low': uat_low,
        'UAT_Total': uat_total
    }
    
    # Add state-specific counts to node_data
    for state in unique_states:
        state_key = state.replace(' ', '_').replace('-', '_')
        node_data[f'{state_key}_Critical'] = state_counts[state]['Critical']
        node_data[f'{state_key}_High'] = state_counts[state]['High']
        node_data[f'{state_key}_Medium'] = state_counts[state]['Medium']
        node_data[f'{state_key}_Low'] = state_counts[state]['Low']
        node_data[f'{state_key}_Total'] = state_counts[state]['Total']
    
    # Add category-specific counts to node_data
    for category in defect_categories:
        category_key = category.replace(' ', '_')
        node_data[f'{category_key}_Critical'] = category_counts[category]['Critical']
        node_data[f'{category_key}_High'] = category_counts[category]['High']
        node_data[f'{category_key}_Medium'] = category_counts[category]['Medium']
        node_data[f'{category_key}_Low'] = category_counts[category]['Low']
        node_data[f'{category_key}_Total'] = category_counts[category]['Total']
    
    node_breakdown.append(node_data)
node_breakdown_df = pd.DataFrame(node_breakdown)
# Separate 'yet to assign' rows and sort the rest by Total
yet_to_assign_nodes = node_breakdown_df[
    (node_breakdown_df['AD POC'] == 'yet to assign') | 
    (node_breakdown_df['SM POC'] == 'yet to assign') | 
    (node_breakdown_df['M POC'] == 'yet to assign')
]
others_nodes = node_breakdown_df[
    (node_breakdown_df['AD POC'] != 'yet to assign') & 
    (node_breakdown_df['SM POC'] != 'yet to assign') & 
    (node_breakdown_df['M POC'] != 'yet to assign')
].sort_values('Total', ascending=False)
# Concatenate with 'yet to assign' at the bottom
node_breakdown_df = pd.concat([others_nodes, yet_to_assign_nodes], ignore_index=True)

# Build POC hierarchy for Defect Node Breakdown filtering
poc_hierarchy_defect = {}
for _, row in node_breakdown_df.iterrows():
    ad_poc = row['AD POC']
    sm_poc = row['SM POC']
    m_poc = row['M POC']
    
    if ad_poc not in poc_hierarchy_defect:
        poc_hierarchy_defect[ad_poc] = {}
    if sm_poc not in poc_hierarchy_defect[ad_poc]:
        poc_hierarchy_defect[ad_poc][sm_poc] = []
    if m_poc not in poc_hierarchy_defect[ad_poc][sm_poc]:
        poc_hierarchy_defect[ad_poc][sm_poc].append(m_poc)

unique_ad_pocs_defect = sorted(node_breakdown_df['AD POC'].unique())
unique_sm_pocs_defect = sorted(node_breakdown_df['SM POC'].unique())
unique_m_pocs_defect = sorted(node_breakdown_df['M POC'].unique())
poc_hierarchy_defect_json = json.dumps(poc_hierarchy_defect)

# Prepare Key Highlights and Risks data
from datetime import datetime
import pytz
current_date_obj = pd.Timestamp.now(tz='UTC')

# Get PT cutoff date from milestone_dates_df - use "Cut off for PT hand over" column explicitly
pt_cutoff_date = None
if len(milestone_dates_df) > 0 and 'Cut off for PT hand over' in milestone_dates_df.columns:
    pt_cutoff_value = milestone_dates_df.iloc[0]['Cut off for PT hand over']
    if pd.notna(pt_cutoff_value) and isinstance(pt_cutoff_value, pd.Timestamp):
        pt_cutoff_date = pd.Timestamp(pt_cutoff_value).tz_localize('UTC') if pt_cutoff_value.tz is None else pt_cutoff_value

# Helper function to extract date only (before 'T') from datetime strings
def extract_date_only(date_series):
    """Extract date portion from datetime strings (before 'T')"""
    # Extract date string before 'T' and convert to datetime
    date_str = date_series.astype(str).str.split('T').str[0]
    # Convert to datetime without timezone
    return pd.to_datetime(date_str, errors='coerce')

# Get current date (date only, no time) - timezone naive
current_date_only = pd.Timestamp.now().normalize()

# Create lowercase State column for case-insensitive comparison
testable_stories_df['State_lower'] = testable_stories_df['State'].str.lower()

# 1. User stories - Planned PT date Not filled
# Exclude stories in these states (delivered/completed)
pt_date_exclude_states = ['Ready to Test', 'Ready for UAT', 'In Test', 'Closed', 'Ready for test', 
                          'Ready for E2E test', 'blocked', 'blocked in pt', 'blocked in uat', 
                          'PT in test', 'UAT in test', 'Test complete']
pt_date_exclude_states_lower = [s.lower() for s in pt_date_exclude_states]

planned_pt_not_filled_filtered = testable_stories_df[
    (testable_stories_df['PT testable stories'] == 1) &
    (testable_stories_df['Planned for PT Date'].isna() | (testable_stories_df['Planned for PT Date'] == '')) &
    (~testable_stories_df['State_lower'].isin(pt_date_exclude_states_lower))
].copy()
planned_pt_not_filled = planned_pt_not_filled_filtered.groupby(['AD POC', 'SM POC', 'M POC']).size().reset_index(name='Count').sort_values('Count', ascending=False)

# 2. User stories - Not Delivered to PT as per plan
pt_delivered_states = ['Ready to Test', 'Ready for UAT', 'In Test', 'Closed', 'Ready for test', 
                       'Test complete', 'Ready for E2E Test', 'Blocked', 'blocked in pt', 
                       'blocked in uat', 'uat in test', 'pt in test', 'awaiting uat deployment']
# Convert to lowercase for case-insensitive comparison
pt_delivered_states_lower = [s.lower() for s in pt_delivered_states]
not_delivered_pt_filtered = testable_stories_df[
    (testable_stories_df['PT testable stories'] == 1) &
    (~testable_stories_df['State_lower'].isin(pt_delivered_states_lower)) &
    (pd.notna(testable_stories_df['Planned for PT Date'])) &
    (extract_date_only(testable_stories_df['Planned for PT Date']) < current_date_only)
].copy()
not_delivered_pt = not_delivered_pt_filtered.groupby(['AD POC', 'SM POC', 'M POC']).size().reset_index(name='Count').sort_values('Count', ascending=False)

# 3. User stories - Not Delivered to UAT as per plan
uat_delivered_states = ['Removed', 'ready for uat', 'Test Complete', 'blocked in UAT', 'UAT in test']
# Convert to lowercase for case-insensitive comparison
uat_delivered_states_lower = [s.lower() for s in uat_delivered_states]
not_delivered_uat_filtered = testable_stories_df[
    (testable_stories_df['UAT Testable Stories'] == 1) &
    (~testable_stories_df['State_lower'].isin(uat_delivered_states_lower)) &
    (pd.notna(testable_stories_df['Planned UAT Date'])) &
    (extract_date_only(testable_stories_df['Planned UAT Date']) < current_date_only)
].copy()
not_delivered_uat = not_delivered_uat_filtered.groupby(['AD POC', 'SM POC', 'M POC']).size().reset_index(name='Count').sort_values('Count', ascending=False)

# 4. User stories - Beyond PT Cut off Date
beyond_pt_cutoff = pd.DataFrame()
if pt_cutoff_date:
    # Convert to timezone-naive if it's timezone-aware
    pt_cutoff_date_only = pt_cutoff_date.normalize()
    if pt_cutoff_date_only.tz is not None:
        pt_cutoff_date_only = pt_cutoff_date_only.tz_localize(None)
    pt_cutoff_states = ['Ready for UAT', 'Ready to Test', 'Ready for Test', 'In Test', 
                        'ready for E2E test', 'Test Complete', 'blocked', 'closed', 
                        'blocked in pt', 'blocked in uat', 'uat in test', 'pt in test']
    # Convert to lowercase for case-insensitive comparison
    pt_cutoff_states_lower = [s.lower() for s in pt_cutoff_states]
    beyond_pt_cutoff_filtered = testable_stories_df[
        (testable_stories_df['PT testable stories'] == 1) &
        (~testable_stories_df['State_lower'].isin(pt_cutoff_states_lower)) &
        (pd.notna(testable_stories_df['Planned for PT Date'])) &
        (extract_date_only(testable_stories_df['Planned for PT Date']) > pt_cutoff_date_only)
    ].copy()
    beyond_pt_cutoff = beyond_pt_cutoff_filtered.groupby(['AD POC', 'SM POC', 'M POC']).size().reset_index(name='Count').sort_values('Count', ascending=False)

# 5. User stories - Parallel UAT
parallel_uat_filtered = testable_stories_df[
    (testable_stories_df['UAT Testable Stories'] == 1) &
    (testable_stories_df['Tags'].str.contains('Parallel UAT', case=False, na=False))
].copy()
parallel_uat = parallel_uat_filtered.groupby(['AD POC', 'SM POC', 'M POC']).size().reset_index(name='Count').sort_values('Count', ascending=False)

# 6. User stories - Added Post User story Freeze date and CCB Approved
ccb_approved_filtered = testable_stories_df[
    (testable_stories_df['PT testable stories'] == 1) &
    (testable_stories_df['Tags'].str.contains('CCB Approved', case=False, na=False))
].copy()
ccb_approved = ccb_approved_filtered.groupby(['AD POC', 'SM POC', 'M POC']).size().reset_index(name='Count').sort_values('Count', ascending=False)

# Calculate totals for highlights section
planned_pt_not_filled_count = planned_pt_not_filled['Count'].sum() if len(planned_pt_not_filled) > 0 else 0
not_delivered_pt_count = not_delivered_pt['Count'].sum() if len(not_delivered_pt) > 0 else 0
not_delivered_uat_count = not_delivered_uat['Count'].sum() if len(not_delivered_uat) > 0 else 0
beyond_pt_cutoff_count = beyond_pt_cutoff['Count'].sum() if len(beyond_pt_cutoff) > 0 else 0
parallel_uat_count = parallel_uat['Count'].sum() if len(parallel_uat) > 0 else 0
ccb_approved_count = ccb_approved['Count'].sum() if len(ccb_approved) > 0 else 0

# Prepare Overall Defect Summary data for JavaScript
print("Preparing Overall Defect Summary data...")

# State mappings - mutually exclusive categories
active_states = ['Active', 'New', 'Blocked', 'Re-open', 'BA clarification', 
                'Blocked in PT', 'Blocked in UAT', 'Deferred']
fixed_ready_states = ['Ready to Deploy', 'Resolved']
under_testing_states = ['In Test', 'Ready to Test', 'Rejected', 'UAT In Test', 'PT In Test']
closed_states = ['Closed', 'Monitoring', 'Ready for Prod Deployment']

# Prepare defect data as list of dictionaries for JavaScript
overall_defects_data = []
for _, row in overall_defect_summary_df.iterrows():
    overall_defects_data.append({
        'id': int(row['ID']) if pd.notna(row['ID']) else 0,
        'state': str(row['State']) if pd.notna(row['State']) else '',
        'stageFound': str(row['StageFound']) if pd.notna(row['StageFound']) else '',
        'severity': str(row['Severity']) if pd.notna(row['Severity']) else '',
        'adPoc': str(row['AD POC']) if pd.notna(row['AD POC']) else '',
        'smPoc': str(row['SM POC']) if pd.notna(row['SM POC']) else '',
        'mPoc': str(row['M POC']) if pd.notna(row['M POC']) else ''
    })

import json
overall_defects_json = json.dumps(overall_defects_data)
active_states_json = json.dumps(active_states)
fixed_ready_states_json = json.dumps(fixed_ready_states)
under_testing_states_json = json.dumps(under_testing_states)
closed_states_json = json.dumps(closed_states)

print(f"  Prepared {len(overall_defects_data)} defects for dashboard")

print(f"  Total Stories: {total_stories}")
print(f"  PT Execution %: {pt_execution_pct:.2f}%")
print(f"  PT Pass %: {pt_pass_pct:.2f}%")
print(f"  Total Bugs: {total_bugs}")
print()

# Generate HTML Dashboard
print("Step 5: Generating HTML Dashboard...")
print("-" * 80)

current_date = datetime.now().strftime("%B %d, %Y")

# Format footer date with proper ordinal suffix (1st, 2nd, 3rd, 4th, etc.)
def get_ordinal_suffix(day):
    if 10 <= day % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return suffix

now = datetime.now()
day = now.day
ordinal_suffix = get_ordinal_suffix(day)
current_date_footer = now.strftime(f"%d{ordinal_suffix} %b %Y")

# Extract Final Go-Live date from milestone_dates_df
release_name = "myISP Release"
if len(milestone_dates_df) > 0 and 'Final Go-Live' in milestone_dates_df.columns:
    final_golive_value = milestone_dates_df.iloc[0]['Final Go-Live']
    if pd.notna(final_golive_value):
        if isinstance(final_golive_value, pd.Timestamp):
            release_name = final_golive_value.strftime("%d %b %Y").upper() + " myISP Release"
        else:
            try:
                golive_date = pd.to_datetime(final_golive_value)
                release_name = golive_date.strftime("%d %b %Y").upper() + " myISP Release"
            except:
                release_name = "myISP Release"

html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Status Report - {current_date}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}
        
        .header .subtitle {{
            font-size: 1.2em;
            opacity: 0.9;
        }}
        
        .tabs {{
            display: flex;
            background: #f5f5f5;
            border-bottom: 2px solid #ddd;
        }}
        
        .tab {{
            flex: 1;
            padding: 20px;
            text-align: center;
            cursor: pointer;
            background: #f5f5f5;
            border: none;
            font-size: 1.1em;
            font-weight: 600;
            color: #555;
            transition: all 0.3s;
        }}
        
        .tab:hover {{
            background: #e0e0e0;
        }}
        
        .tab.active {{
            background: white;
            color: #667eea;
            border-bottom: 3px solid #667eea;
        }}
        
        .tab-content {{
            display: none;
            padding: 40px;
            animation: fadeIn 0.5s;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        .sub-tabs {{
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            border-bottom: 2px solid #e0e0e0;
        }}
        
        .sub-tab {{
            padding: 12px 24px;
            cursor: pointer;
            background: #f9f9f9;
            border: none;
            font-size: 1em;
            font-weight: 500;
            color: #666;
            transition: all 0.3s;
            border-radius: 8px 8px 0 0;
        }}
        
        .sub-tab:hover {{
            background: #e8e8e8;
            color: #333;
        }}
        
        .sub-tab.active {{
            background: white;
            color: #667eea;
            border-bottom: 3px solid #667eea;
            font-weight: 600;
        }}
        
        .sub-tab-content {{
            display: none;
            animation: fadeIn 0.3s;
        }}
        
        .sub-tab-content.active {{
            display: block;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        
        .metrics-grid-compact {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 15px;
            margin-bottom: 40px;
        }}
        
        .metrics-grid-single-row {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 15px;
            margin-bottom: 30px;
        }}
        
        .metrics-grid-defects {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 15px;
            margin-bottom: 30px;
        }}
        
        .metric-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s, box-shadow 0.3s;
        }}
        
        .metric-card-compact {{
            color: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.1);
            text-align: center;
        }}
        
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.2);
        }}
        
        .metric-card.green {{
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        }}
        
        .metric-card.orange {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }}
        
        .metric-card.blue {{
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        }}
        
        .metric-card.red {{
            background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
        }}
        
        .metric-card.purple {{
            background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
            color: #333;
        }}
        
        .metric-label {{
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 10px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .metric-label-compact {{
            font-size: 0.75em;
            opacity: 0.9;
            margin-bottom: 5px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        
        .metric-value {{
            font-size: 2.5em;
            font-weight: bold;
            margin-bottom: 5px;
        }}
        
        .metric-value-compact {{
            font-size: 1.8em;
            font-weight: bold;
        }}
        
        .metric-subtitle {{
            font-size: 0.85em;
            opacity: 0.8;
        }}
        
        .section-title {{
            font-size: 1.8em;
            color: #333;
            margin: 30px 0 20px 0;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
        }}
        
        .bug-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 20px;
        }}
        
        .bug-card {{
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            color: white;
            font-weight: 600;
        }}
        
        .bug-card.critical {{
            background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%);
        }}
        
        .bug-card.high {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }}
        
        .bug-card.medium {{
            background: linear-gradient(135deg, #ffd89b 0%, #19547b 100%);
        }}
        
        .bug-card.low {{
            background: #FFD700;
            color: #333;
        }}
        
        .bug-card.dark-red {{
            background: #8B0000;
        }}
        
        .bug-card.amber {{
            background: #FFBF00;
            color: #333;
        }}
        
        .bug-card .count {{
            font-size: 2em;
            display: block;
            margin-bottom: 5px;
        }}
        
        .section-heading {{
            font-size: 1.4em;
            color: #667eea;
            margin: 30px 0 15px 0;
            padding: 10px;
            background: #f5f5f5;
            border-left: 5px solid #667eea;
            font-weight: 600;
        }}
        
        .table-wrapper {{
            max-height: 600px;
            overflow-y: auto;
            overflow-x: auto;
            border: 1px solid #ddd;
            border-radius: 8px;
            margin-top: 20px;
        }}
        
        .milestone-table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            font-size: 0.9em;
            border: 1px solid #ddd;
        }}
        
        .milestone-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 8px;
            text-align: left;
            font-weight: 600;
            font-size: 0.85em;
            position: sticky;
            top: 0;
            z-index: 10;
            border: 1px solid #555;
        }}
        
        .total-row {{
            background: #e8f4f8 !important;
            font-weight: bold !important;
            border-top: 3px solid #1e3a8a !important;
            border-bottom: 3px solid #1e3a8a !important;
            border-left: 3px solid #1e3a8a !important;
            border-right: 3px solid #1e3a8a !important;
            position: -webkit-sticky !important;
            position: sticky !important;
            top: 40px !important;
            z-index: 9 !important;
            box-shadow: 0 3px 6px rgba(0,0,0,0.15) !important;
        }}
        
        .total-row td {{
            border: 1px solid #1e3a8a !important;
            padding: 10px 8px !important;
        }}
        
        .milestone-table td {{
            padding: 10px 8px;
            border: 1px solid #ddd;
            font-size: 0.85em;
            word-wrap: break-word;
            max-width: 150px;
        }}
        
        .milestone-table tr:hover {{
            background: #f5f5f5;
        }}
        
        .milestone-table tr:last-child td {{
            border-bottom: none;
        }}
        
        .footer {{
            text-align: center;
            padding: 20px;
            background: #f5f5f5;
            color: #666;
            font-size: 0.9em;
        }}
        
        .percentage-bar {{
            width: 100%;
            height: 8px;
            background: rgba(255,255,255,0.3);
            border-radius: 4px;
            margin-top: 10px;
            overflow: hidden;
        }}
        
        .percentage-fill {{
            height: 100%;
            background: white;
            border-radius: 4px;
            transition: width 1s ease-in-out;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Daily Status Report</h1>
            <div class="subtitle">Release Name - {release_name}</div>
            <div class="subtitle">Daily Status Report as on - {current_date}</div>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="showTab('overview')">Overview</button>
            <button class="tab" onclick="showTab('execution')">Parent Level PT Execution Status</button>
            <button class="tab" onclick="showTab('consolidated')">Parent Level - PT and UAT Execution Status</button>
            <button class="tab" onclick="showTab('defects')">Open Defects Summary</button>
            <button class="tab" onclick="showTab('highlights')">Key Highlights and Risks</button>
            <button class="tab" onclick="showTab('overall-defects')">Over All Defect Summary</button>
        </div>
        
        <div id="overview" class="tab-content active">
            <h2 class="section-title">📅 Milestone Dates</h2>
            <table class="milestone-table">
                <thead>
                    <tr>
"""

# Add milestone table headers dynamically at the beginning of overview
for col in milestone_dates_df.columns:
    html_content += f"                        <th>{col}</th>\n"

html_content += """                    </tr>
                </thead>
                <tbody>
"""

# Add milestone table rows
for _, row in milestone_dates_df.iterrows():
    html_content += "                    <tr>\n"
    for col in milestone_dates_df.columns:
        value = row[col]
        # Format dates if they are datetime objects
        if isinstance(value, pd.Timestamp):
            value = value.strftime("%B %d, %Y")
        html_content += f"                        <td>{value}</td>\n"
    html_content += "                    </tr>\n"

html_content += f"""                </tbody>
            </table>
            
            <h2 class="section-title">📈 Story Summary</h2>
            <div class="metrics-grid">
                <div class="metric-card blue">
                    <div class="metric-label">Total Stories</div>
                    <div class="metric-value">{total_stories}</div>
                </div>
                
                <div class="metric-card orange">
                    <div class="metric-label">Testing NA Stories</div>
                    <div class="metric-value">{testing_na_stories}</div>
                </div>
                
                <div class="metric-card green">
                    <div class="metric-label">PT Testable Stories</div>
                    <div class="metric-value">{pt_testable_stories}</div>
                </div>
                
                <div class="metric-card purple">
                    <div class="metric-label">UAT Testable Stories</div>
                    <div class="metric-value">{uat_testable_stories}</div>
                </div>
            </div>
            
            <div class="metrics-grid">
                <div class="metric-card green">
                    <div class="metric-label">PT Delivered</div>
                    <div class="metric-value">{pt_delivered}</div>
                    <div class="metric-subtitle">Stories in delivery states</div>
                </div>
                
                <div class="metric-card red">
                    <div class="metric-label">PT NOT Delivered</div>
                    <div class="metric-value">{pt_not_delivered}</div>
                    <div class="metric-subtitle">Stories pending delivery</div>
                </div>
                
                <div class="metric-card green">
                    <div class="metric-label">UAT Delivered</div>
                    <div class="metric-value">{uat_delivered}</div>
                    <div class="metric-subtitle">Stories in UAT delivery</div>
                </div>
                
                <div class="metric-card red">
                    <div class="metric-label">UAT NOT Delivered</div>
                    <div class="metric-value">{uat_not_delivered}</div>
                    <div class="metric-subtitle">Stories pending UAT</div>
                </div>
            </div>
            
            <h2 class="section-title">📊 PT Execution Summary</h2>
            <div class="metrics-grid-single-row">
                <div class="metric-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                    <div class="metric-label">Total Test Cases</div>
                    <div class="metric-value">{total_tests}</div>
                    <div class="metric-subtitle">All test cases combined</div>
                </div>
                
                <div class="metric-card green">
                    <div class="metric-label">Passed</div>
                    <div class="metric-value">{total_passed}</div>
                    <div class="metric-subtitle">Successfully executed</div>
                </div>
                
                <div class="metric-card red">
                    <div class="metric-label">Failed</div>
                    <div class="metric-value">{total_failed}</div>
                    <div class="metric-subtitle">Execution failures</div>
                </div>
                
                <div class="metric-card orange">
                    <div class="metric-label">Blocked</div>
                    <div class="metric-value">{total_blocked}</div>
                    <div class="metric-subtitle">Blocked test cases</div>
                </div>
                
                <div class="metric-card" style="background: #9E9E9E;">
                    <div class="metric-label">Not Run</div>
                    <div class="metric-value">{total_not_run}</div>
                    <div class="metric-subtitle">Yet to be executed</div>
                </div>
                
                <div class="metric-card blue">
                    <div class="metric-label">PT Execution %</div>
                    <div class="metric-value">{pt_execution_pct:.2f}%</div>
                    <div class="metric-subtitle">Passed: {total_passed} | Failed: {total_failed}</div>
                    <div class="percentage-bar">
                        <div class="percentage-fill" style="width: {pt_execution_pct}%"></div>
                    </div>
                </div>
                
                <div class="metric-card green">
                    <div class="metric-label">PT Pass %</div>
                    <div class="metric-value">{pt_pass_pct:.2f}%</div>
                    <div class="metric-subtitle">Passed: {total_passed} | Total Executed: {total_passed + total_failed + total_blocked}</div>
                    <div class="percentage-bar">
                        <div class="percentage-fill" style="width: {pt_pass_pct}%"></div>
                    </div>
                </div>
            </div>
            
            <h2 class="section-title">📊 UAT Execution Summary</h2>
            <div class="metrics-grid-single-row">
                <div class="metric-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                    <div class="metric-label">Total UAT Test Cases</div>
                    <div class="metric-value">{uat_total_tests}</div>
                    <div class="metric-subtitle">All UAT test cases</div>
                </div>
                
                <div class="metric-card green">
                    <div class="metric-label">Passed</div>
                    <div class="metric-value">{uat_passed}</div>
                    <div class="metric-subtitle">Successfully executed</div>
                </div>
                
                <div class="metric-card red">
                    <div class="metric-label">Failed</div>
                    <div class="metric-value">{uat_failed}</div>
                    <div class="metric-subtitle">Execution failures</div>
                </div>
                
                <div class="metric-card orange">
                    <div class="metric-label">Blocked</div>
                    <div class="metric-value">{uat_blocked}</div>
                    <div class="metric-subtitle">Blocked test cases</div>
                </div>
                
                <div class="metric-card" style="background: #9E9E9E;">
                    <div class="metric-label">Not Run</div>
                    <div class="metric-value">{uat_not_run}</div>
                    <div class="metric-subtitle">Yet to be executed</div>
                </div>
                
                <div class="metric-card blue">
                    <div class="metric-label">UAT Execution %</div>
                    <div class="metric-value">{uat_execution_pct:.2f}%</div>
                    <div class="metric-subtitle">Passed: {uat_passed} | Failed: {uat_failed}</div>
                    <div class="percentage-bar">
                        <div class="percentage-fill" style="width: {uat_execution_pct}%"></div>
                    </div>
                </div>
                
                <div class="metric-card green">
                    <div class="metric-label">UAT Pass %</div>
                    <div class="metric-value">{uat_pass_pct:.2f}%</div>
                    <div class="metric-subtitle">Passed: {uat_passed} | Total Executed: {uat_passed + uat_failed + uat_blocked}</div>
                    <div class="percentage-bar">
                        <div class="percentage-fill" style="width: {uat_pass_pct}%"></div>
                    </div>
                </div>
            </div>
            
            <h2 class="section-title">🐛 Defect Summary</h2>
            <div class="metrics-grid-defects">
                <div class="metric-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                    <div class="metric-label">Total Defects</div>
                    <div class="metric-value">{total_bugs}</div>
                </div>
                <div class="bug-card" style="background: #8B0000;">
                    <span class="count">{critical_bugs}</span>
                    <div>Critical</div>
                </div>
                <div class="bug-card" style="background: #FF0000;">
                    <span class="count">{high_bugs}</span>
                    <div>High</div>
                </div>
                <div class="bug-card" style="background: #FFBF00; color: #333;">
                    <span class="count">{medium_bugs}</span>
                    <div>Medium</div>
                </div>
                <div class="bug-card" style="background: #FFFACD; color: #333;">
                    <span class="count">{low_bugs}</span>
                    <div>Low</div>
                </div>
            </div>
            
            <h3 style="margin-top: 25px; margin-bottom: 15px; color: #2c3e50;">Defects Breakdown by AD POC</h3>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>Critical</th>
                            <th>High</th>
                            <th>Medium</th>
                            <th>Low</th>
                            <th>Total</th>
                        </tr>
                    </thead>
                    <tbody>
"""

for _, row in ad_poc_breakdown.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['POC']}</td>
                            <td>{int(row['Critical'])}</td>
                            <td>{int(row['High'])}</td>
                            <td>{int(row['Medium'])}</td>
                            <td>{int(row['Low'])}</td>
                            <td>{int(row['Total'])}</td>
                        </tr>
"""

# Add total row for AD POC
ad_total_critical = int(ad_poc_breakdown['Critical'].sum())
ad_total_high = int(ad_poc_breakdown['High'].sum())
ad_total_medium = int(ad_poc_breakdown['Medium'].sum())
ad_total_low = int(ad_poc_breakdown['Low'].sum())
ad_total_all = int(ad_poc_breakdown['Total'].sum())

html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td>Total</td>
                            <td>{ad_total_critical}</td>
                            <td>{ad_total_high}</td>
                            <td>{ad_total_medium}</td>
                            <td>{ad_total_low}</td>
                            <td>{ad_total_all}</td>
                        </tr>
"""  

html_content += """                    </tbody>
                </table>
            </div>
            
            <h3 style="margin-top: 25px; margin-bottom: 15px; color: #2c3e50;">Defects Breakdown by SM POC</h3>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>SM POC</th>
                            <th>Critical</th>
                            <th>High</th>
                            <th>Medium</th>
                            <th>Low</th>
                            <th>Total</th>
                        </tr>
                    </thead>
                    <tbody>
"""

for _, row in sm_poc_breakdown.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['POC']}</td>
                            <td>{int(row['Critical'])}</td>
                            <td>{int(row['High'])}</td>
                            <td>{int(row['Medium'])}</td>
                            <td>{int(row['Low'])}</td>
                            <td>{int(row['Total'])}</td>
                        </tr>
"""

# Add total row for SM POC
sm_total_critical = int(sm_poc_breakdown['Critical'].sum())
sm_total_high = int(sm_poc_breakdown['High'].sum())
sm_total_medium = int(sm_poc_breakdown['Medium'].sum())
sm_total_low = int(sm_poc_breakdown['Low'].sum())
sm_total_all = int(sm_poc_breakdown['Total'].sum())

html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td>Total</td>
                            <td>{sm_total_critical}</td>
                            <td>{sm_total_high}</td>
                            <td>{sm_total_medium}</td>
                            <td>{sm_total_low}</td>
                            <td>{sm_total_all}</td>
                        </tr>
"""  

html_content += """                    </tbody>
                </table>
            </div>
        </div>
        
        <div id="execution" class="tab-content">
            <h2 class="section-title">📊 Parent Level PT Execution Status</h2>
            <div style="margin-bottom: 25px; padding: 20px; background: #f8f9fa; border-radius: 10px;">
                <div style="display: grid; grid-template-columns: repeat(3, 1fr) auto; gap: 20px; align-items: end;">
                    <div>
                        <label for="exec-ad-filter" style="font-weight: bold; font-size: 1em; margin-bottom: 8px; display: block; color: #2c3e50;">AD POC:</label>
                        <select id="exec-ad-filter" onchange="updateExecFilters('ad')" style="width: 100%; padding: 10px; font-size: 1em; border: 2px solid #667eea; border-radius: 8px; background: white; cursor: pointer;">
                            <option value="all">All AD POCs</option>
"""

# Add AD POC options
for ad_poc in unique_ad_pocs_exec:
    html_content += f"""                            <option value="{ad_poc}">{ad_poc}</option>
"""

html_content += """                        </select>
                    </div>
                    <div>
                        <label for="exec-sm-filter" style="font-weight: bold; font-size: 1em; margin-bottom: 8px; display: block; color: #2c3e50;">SM POC:</label>
                        <select id="exec-sm-filter" onchange="updateExecFilters('sm')" style="width: 100%; padding: 10px; font-size: 1em; border: 2px solid #667eea; border-radius: 8px; background: white; cursor: pointer;" disabled>
                            <option value="all">All SM POCs</option>
                        </select>
                    </div>
                    <div>
                        <label for="exec-m-filter" style="font-weight: bold; font-size: 1em; margin-bottom: 8px; display: block; color: #2c3e50;">M POC:</label>
                        <select id="exec-m-filter" onchange="updateExecFilters('m')" style="width: 100%; padding: 10px; font-size: 1em; border: 2px solid #667eea; border-radius: 8px; background: white; cursor: pointer;" disabled>
                            <option value="all">All M POCs</option>
                        </select>
                    </div>
                    <div>
                        <label style="font-weight: bold; font-size: 1em; margin-bottom: 8px; display: block; color: transparent;">Action:</label>
                        <button onclick="resetExecFilters()" style="width: 100%; padding: 10px; font-size: 1em; border: 2px solid #f5576c; border-radius: 8px; background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; cursor: pointer; font-weight: bold; transition: all 0.3s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
                            🔄 Reset Filters
                        </button>
                    </div>
                </div>
            </div>
"""

html_content += f"""            
            <h3 style="margin: 30px 0 15px 0; color: #2c3e50; font-size: 1.3em;">📊 Story Summary</h3>
            <div class="metrics-grid" style="margin-bottom: 30px;">
                <div class="metric-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                    <div class="metric-label">Total Stories</div>
                    <div class="metric-value" id="exec-summary-total-stories">{total_stories}</div>
                </div>
                <div class="metric-card" style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
                    <div class="metric-label">Testable Stories</div>
                    <div class="metric-value" id="exec-summary-testable-stories">{pt_testable_stories}</div>
                </div>
                <div class="metric-card" style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);">
                    <div class="metric-label">Testing NA Stories</div>
                    <div class="metric-value" id="exec-summary-testing-na-stories">{testing_na_stories}</div>
                </div>
            </div>
"""

html_content += """            
            <h3 style="margin: 20px 0 15px 0; color: #2c3e50; font-size: 1.3em;">Testable Stories</h3>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>Parent</th>
                            <th>Parent Title</th>
                            <th>Total PT</th>
                            <th>PT - D</th>
                            <th>PT - ND</th>
                            <th>Total UAT</th>
                            <th>UAT - D</th>
                            <th>UAT - ND</th>
                            <th>Passed</th>
                            <th>Failed</th>
                            <th>Blocked</th>
                            <th>Not Run</th>
                            <th>Total</th>
                            <th>PT Exec</th>
                            <th>PT Pass</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Calculate totals first for Parent level
total_stories = int(execution_summary_df['Total Stories'].sum())
total_pt_delivered = int(execution_summary_df['PT delivered'].sum())
total_pt_not_delivered = int(execution_summary_df['PT NOT delivered'].sum())
total_uat_testable = int(execution_summary_df['UAT Testable Stories'].sum())
total_uat_delivered = int(execution_summary_df['UAT delivered'].sum())
total_uat_not_delivered = int(execution_summary_df['UAT NOT delivered'].sum())
total_passed_sum = int(execution_summary_df['Passed'].sum())
total_failed_sum = int(execution_summary_df['Failed'].sum())
total_blocked_sum = int(execution_summary_df['Blocked'].sum())
total_not_run_sum = int(execution_summary_df['Not Run'].sum())
total_total = int(execution_summary_df['Total'].sum())

# Calculate PT Exec and PT Pass using formula (matching Tab 1 and Tab 3)
# PT Exec = (Passed + Failed) / Total
total_execution_pct = round((total_passed_sum + total_failed_sum) / total_total, 4) if total_total > 0 else 0
# PT Pass = Passed / (Passed + Failed + Blocked)
total_pass_pct = round(total_passed_sum / (total_passed_sum + total_failed_sum + total_blocked_sum), 4) if (total_passed_sum + total_failed_sum + total_blocked_sum) > 0 else 0

# Add total row at the top
html_content += f"""                        <tr id="exec-total-row" class="total-row">
                            <td colspan="2" style="text-align: left;">TOTAL</td>
                            <td>{total_stories}</td>
                            <td>{total_pt_delivered}</td>
                            <td>{total_pt_not_delivered}</td>
                            <td>{total_uat_testable}</td>
                            <td>{total_uat_delivered}</td>
                            <td>{total_uat_not_delivered}</td>
                            <td>{total_passed_sum}</td>
                            <td>{total_failed_sum}</td>
                            <td>{total_blocked_sum}</td>
                            <td>{total_not_run_sum}</td>
                            <td>{total_total}</td>
                            <td>{total_execution_pct:.2%}</td>
                            <td>{total_pass_pct:.2%}</td>
                        </tr>
"""

# Add execution summary rows
for _, row in execution_summary_df.iterrows():
    html_content += f"""                    <tr class="exec-row" data-ad-poc="{row['AD POC']}" data-sm-poc="{row['SM POC']}" data-m-poc="{row['M POC']}" data-story-ids="{row['Story IDs']}" 
                        data-stories="{int(row['Total Stories'])}" data-pt-delivered="{int(row['PT delivered'])}" data-pt-not-delivered="{int(row['PT NOT delivered'])}" 
                        data-uat-testable="{int(row['UAT Testable Stories'])}" data-uat-delivered="{int(row['UAT delivered'])}" data-uat-not-delivered="{int(row['UAT NOT delivered'])}" 
                        data-passed="{int(row['Passed'])}" data-failed="{int(row['Failed'])}" data-blocked="{int(row['Blocked'])}" 
                        data-not-run="{int(row['Not Run'])}" data-total="{int(row['Total'])}">
                        <td>{row['Parent']}</td>
                        <td>{row['Parent Title']}</td>
                        <td>{int(row['Total Stories'])}</td>
                        <td>{int(row['PT delivered'])}</td>
                        <td>{int(row['PT NOT delivered'])}</td>
                        <td>{int(row['UAT Testable Stories'])}</td>
                        <td>{int(row['UAT delivered'])}</td>
                        <td>{int(row['UAT NOT delivered'])}</td>
                        <td>{int(row['Passed'])}</td>
                        <td>{int(row['Failed'])}</td>
                        <td>{int(row['Blocked'])}</td>
                        <td>{int(row['Not Run'])}</td>
                        <td>{int(row['Total'])}</td>
                        <td>{row['Execution %']:.2%}</td>
                        <td>{row['Pass %']:.2%}</td>
                    </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            
            <h2 class="section-title" style="margin-top: 50px;">📋 Testing Not Applicable Stories</h2>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Node Name</th>
                            <th>Testing NA Stories Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add grand total row at the top with id for dynamic updates
html_content += f"""                        <tr id="testing-na-grand-total-row" style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.1em;">
                            <td colspan="4" style="text-align: center;">Grand Total</td>
                            <td id="testing-na-grand-total">{testing_na_grand_total}</td>
                        </tr>
"""

# Add Testing NA stories rows
for _, row in testing_na_summary.iterrows():
    count = int(row['Testing NA Stories Count'])
    html_content += f"""                        <tr class="exec-row testing-na-row" data-ad-poc="{row['AD POC']}" data-sm-poc="{row['SM POC']}" data-m-poc="{row['M POC']}" data-testing-na-count="{count}">
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Node Name']}</td>
                            <td>{count}</td>
                        </tr>
"""

html_content += """                    </tbody>
                </table>
            </div>
            
            <!-- Note Section -->
            <div style="margin-top: 30px; padding: 12px; background: #f8f9fa; border-radius: 8px; border-left: 4px solid #667eea; font-size: 0.75em;">
                <h4 style="color: #2c3e50; margin-bottom: 8px; font-size: 0.9em;">📝 Note:</h4>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                    <ul style="list-style-type: disc; margin-left: 20px; line-height: 1.4; font-size: 0.85em;">
                        <li><strong>Total PT:</strong> Total Stories Tested by PT</li>
                        <li><strong>PT - D:</strong> Stories delivered to PT Testing</li>
                        <li><strong>PT - ND:</strong> Stories Not delivered to PT Testing</li>
                        <li><strong>PT Exec:</strong> PT Execution %</li>
                        <li><strong>PT Pass:</strong> PT Pass %</li>
                    </ul>
                    <ul style="list-style-type: disc; margin-left: 20px; line-height: 1.4; font-size: 0.85em;">
                        <li><strong>Total UAT:</strong> Total Stories Tested by UAT</li>
                        <li><strong>UAT - D:</strong> Stories delivered to UAT Testing</li>
                        <li><strong>UAT - ND:</strong> Stories Not delivered to UAT Testing</li>
                        <li><strong>UAT Exec:</strong> UAT Execution %</li>
                        <li><strong>UAT Pass:</strong> UAT Pass %</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <div id="consolidated" class="tab-content">
            <h2 class="section-title">📊 Parent Level - PT and UAT Execution Status</h2>
            
            <!-- Product Owner Filter -->
            <div class="filter-section" style="margin-bottom: 20px; display: flex; align-items: center; gap: 15px;">
                <label for="poFilter" style="font-weight: bold;">Product Owner:</label>
                <select id="poFilter" onchange="filterByProductOwner()" style="padding: 8px; border-radius: 5px; border: 1px solid #ddd; font-size: 1em; min-width: 250px;">
                    <option value="All">All</option>
"""

# Add Product Owner filter options
for po in unique_product_owners:
    html_content += f"""                    <option value="{po}">{po}</option>
"""

html_content += """                </select>
                <button onclick="resetPOFilter()" style="padding: 8px 16px; border-radius: 5px; border: 2px solid #f5576c; background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; cursor: pointer; font-weight: bold; font-size: 1em; transition: all 0.3s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
                    🔄 Reset Filter
                </button>
            </div>
            
            <div class="table-wrapper">
                <table class="milestone-table" id="consolidatedTable">
                    <thead>
                        <tr>
                            <th>Parent ID</th>
                            <th>Parent Title</th>
                            <th>Total PT</th>
                            <th>PT - D</th>
                            <th>PT - ND</th>
                            <th>PT Exec</th>
                            <th>PT Pass</th>
                            <th>Total UAT</th>
                            <th>UAT - D</th>
                            <th>UAT - ND</th>
                            <th>UAT Exec</th>
                            <th>UAT Pass</th>
                            <th style="width: 40px;">Total</th>
                            <th style="width: 40px;">C</th>
                            <th style="width: 40px;">H</th>
                            <th style="width: 40px;">M</th>
                            <th style="width: 40px;">L</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Calculate totals for consolidated PT/UAT status
# Convert to numeric to handle any string values
cons_total_pt_stories = int(pd.to_numeric(consolidated_pt_uat_df['Total PT stories'], errors='coerce').fillna(0).sum())
cons_total_pt_delivered = int(pd.to_numeric(consolidated_pt_uat_df['PT Delivered'], errors='coerce').fillna(0).sum())
cons_total_pt_not_delivered = int(pd.to_numeric(consolidated_pt_uat_df['PT Not Delivered'], errors='coerce').fillna(0).sum())
cons_total_uat_stories = int(pd.to_numeric(consolidated_pt_uat_df['Total UAT stories'], errors='coerce').fillna(0).sum())
cons_total_uat_delivered = int(pd.to_numeric(consolidated_pt_uat_df['UAT Delivered'], errors='coerce').fillna(0).sum())
cons_total_uat_not_delivered = int(pd.to_numeric(consolidated_pt_uat_df['UAT Not Delivered'], errors='coerce').fillna(0).sum())
cons_total_bugs = int(pd.to_numeric(consolidated_pt_uat_df['Total Open Bugs for Parent ID'], errors='coerce').fillna(0).sum())
cons_total_critical = int(pd.to_numeric(consolidated_pt_uat_df['Critical'], errors='coerce').fillna(0).sum())
cons_total_high = int(pd.to_numeric(consolidated_pt_uat_df['High'], errors='coerce').fillna(0).sum())
cons_total_medium = int(pd.to_numeric(consolidated_pt_uat_df['Medium'], errors='coerce').fillna(0).sum())
cons_total_low = int(pd.to_numeric(consolidated_pt_uat_df['Low'], errors='coerce').fillna(0).sum())

# Calculate execution and pass percentages for totals
# For PT: Get sum of passed/failed/blocked/not run from the detailed story data
# We need to calculate based on actual test execution data
cons_total_pt_passed = 0
cons_total_pt_failed = 0
cons_total_pt_blocked = 0
cons_total_pt_not_run = 0
cons_total_pt_total = 0

# Get execution data from story_summary_df grouped by Parent
for parent_id in consolidated_pt_uat_df['Parent ID'].unique():
    parent_stories = story_summary_df[story_summary_df['Parent'] == parent_id]
    
    # PT execution data
    pt_testable = parent_stories[parent_stories['PT testable stories'] == 1]
    if not pt_testable.empty:
        pt_passed = int(pd.to_numeric(pt_testable['Passed'], errors='coerce').fillna(0).sum())
        pt_failed = int(pd.to_numeric(pt_testable['Failed'], errors='coerce').fillna(0).sum())
        pt_blocked = int(pd.to_numeric(pt_testable['Blocked'], errors='coerce').fillna(0).sum())
        pt_not_run = int(pd.to_numeric(pt_testable['Not Run'], errors='coerce').fillna(0).sum())
        
        cons_total_pt_passed += pt_passed
        cons_total_pt_failed += pt_failed
        cons_total_pt_blocked += pt_blocked
        cons_total_pt_not_run += pt_not_run
        cons_total_pt_total += (pt_passed + pt_failed + pt_blocked + pt_not_run)

# Calculate PT percentages
if cons_total_pt_total > 0:
    # PT Execution % = (Passed + Failed) / Total - matching Parent Level PT execution Status tab
    cons_pt_exec_pct = ((cons_total_pt_passed + cons_total_pt_failed) / cons_total_pt_total) * 100
    # PT Pass % = Passed / (Passed + Failed + Blocked) - same as in Parent Level PT execution Status
    cons_pt_pass_pct = (cons_total_pt_passed / (cons_total_pt_passed + cons_total_pt_failed + cons_total_pt_blocked)) * 100 if (cons_total_pt_passed + cons_total_pt_failed + cons_total_pt_blocked) > 0 else 0
else:
    cons_pt_exec_pct = 0
    cons_pt_pass_pct = 0

# Calculate UAT percentages using simple average from consolidated_pt_uat_df
# UAT percentages come from the Excel file, we calculate simple average (not weighted)
uat_exec_sum = 0
uat_pass_sum = 0
uat_count = 0

for _, row in consolidated_pt_uat_df.iterrows():
    # Get UAT story count
    uat_stories = pd.to_numeric(row['Total UAT stories'], errors='coerce')
    if pd.isna(uat_stories) or uat_stories == 0:
        continue
    
    # Get UAT Execution% and Pass%
    uat_exec = row['UAT Execution%']
    uat_pass = row['UAT Pass%']
    
    # Include only numeric values (not 'UAT NA' strings)
    if isinstance(uat_exec, (int, float)) and isinstance(uat_pass, (int, float)):
        uat_exec_sum += uat_exec
        uat_pass_sum += uat_pass
        uat_count += 1

# Calculate simple averages
if uat_count > 0:
    cons_uat_exec_pct = (uat_exec_sum / uat_count) * 100
    cons_uat_pass_pct = (uat_pass_sum / uat_count) * 100
else:
    cons_uat_exec_pct = 0
    cons_uat_pass_pct = 0

# Add total row
html_content += f"""                        <tr id="consolidated-total-row" class="total-row" style="position: -webkit-sticky !important; position: sticky !important; top: 40px !important; z-index: 9 !important; background: #e8f4f8 !important; box-shadow: 0 3px 6px rgba(0,0,0,0.15) !important;" data-hardcoded-pt-exec="{cons_pt_exec_pct:.2f}" data-hardcoded-pt-pass="{cons_pt_pass_pct:.2f}" data-hardcoded-uat-exec="{uat_execution_pct:.2f}" data-hardcoded-uat-pass="{uat_pass_pct:.2f}">
                            <td colspan="2" style="text-align: left;">TOTAL</td>
                            <td id="cons-total-pt-stories">{cons_total_pt_stories}</td>
                            <td id="cons-total-pt-delivered">{cons_total_pt_delivered}</td>
                            <td id="cons-total-pt-not-delivered">{cons_total_pt_not_delivered}</td>
                            <td id="cons-total-pt-exec">{cons_pt_exec_pct:.2f}%</td>
                            <td id="cons-total-pt-pass">{cons_pt_pass_pct:.2f}%</td>
                            <td id="cons-total-uat-stories">{cons_total_uat_stories}</td>
                            <td id="cons-total-uat-delivered">{cons_total_uat_delivered}</td>
                            <td id="cons-total-uat-not-delivered">{cons_total_uat_not_delivered}</td>
                            <td id="cons-total-uat-exec">{uat_execution_pct:.2f}%</td>
                            <td id="cons-total-uat-pass">{uat_pass_pct:.2f}%</td>
                            <td id="cons-total-bugs">{cons_total_bugs}</td>
                            <td id="cons-total-critical">{cons_total_critical}</td>
                            <td id="cons-total-high">{cons_total_high}</td>
                            <td id="cons-total-medium">{cons_total_medium}</td>
                            <td id="cons-total-low">{cons_total_low}</td>
                        </tr>
"""

# Add rows for Consolidated PT_UAT Status
for _, row in consolidated_pt_uat_df.iterrows():
    product_owner = row.get('Product Owner', 'N/A')
    parent_id = row['Parent ID']
    parent_title = row['Parent Title']
    total_pt_stories = row['Total PT stories']
    pt_delivered = row['PT Delivered']
    pt_not_delivered = row['PT Not Delivered']
    pt_exec = row['PT Execution%']
    pt_pass = row['PT Pass%']
    
    # Convert UAT values to numeric, handling 'UAT NA' strings
    total_uat_stories = pd.to_numeric(row['Total UAT stories'], errors='coerce')
    total_uat_stories = 0 if pd.isna(total_uat_stories) else int(total_uat_stories)
    
    uat_delivered = pd.to_numeric(row['UAT Delivered'], errors='coerce')
    uat_delivered = 0 if pd.isna(uat_delivered) else int(uat_delivered)
    
    uat_not_delivered = pd.to_numeric(row['UAT Not Delivered'], errors='coerce')
    uat_not_delivered = 0 if pd.isna(uat_not_delivered) else int(uat_not_delivered)
    
    uat_exec = row['UAT Execution%']
    uat_pass = row['UAT Pass%']
    
    row_total_bugs = row['Total Open Bugs for Parent ID']
    row_critical = row['Critical']
    row_high = row['High']
    row_medium = row['Medium']
    row_low = row['Low']
    
    # Get execution data for this parent
    parent_stories = story_summary_df[story_summary_df['Parent'] == parent_id]
    
    # PT execution counts
    pt_testable = parent_stories[parent_stories['PT testable stories'] == 1]
    pt_passed = int(pd.to_numeric(pt_testable['Passed'], errors='coerce').fillna(0).sum()) if not pt_testable.empty else 0
    pt_failed = int(pd.to_numeric(pt_testable['Failed'], errors='coerce').fillna(0).sum()) if not pt_testable.empty else 0
    pt_blocked = int(pd.to_numeric(pt_testable['Blocked'], errors='coerce').fillna(0).sum()) if not pt_testable.empty else 0
    pt_not_run = int(pd.to_numeric(pt_testable['Not Run'], errors='coerce').fillna(0).sum()) if not pt_testable.empty else 0
    pt_total = pt_passed + pt_failed + pt_blocked + pt_not_run
    
    # UAT execution counts - UAT data comes from UAT Status Excel, set to 0 for now
    uat_passed = 0
    uat_failed = 0
    uat_blocked = 0
    uat_not_run = 0
    uat_total = 0
    
    # Format percentages
    if isinstance(pt_exec, (int, float)):
        pt_exec_display = f"{pt_exec * 100:.2f}%"
    else:
        pt_exec_display = str(pt_exec)
    
    if isinstance(pt_pass, (int, float)):
        pt_pass_display = f"{pt_pass * 100:.2f}%"
    else:
        pt_pass_display = str(pt_pass)
    
    if isinstance(uat_exec, (int, float)):
        uat_exec_display = f"{uat_exec * 100:.2f}%"
    else:
        uat_exec_display = str(uat_exec)
    
    if isinstance(uat_pass, (int, float)):
        uat_pass_display = f"{uat_pass * 100:.2f}%"
    else:
        uat_pass_display = str(uat_pass)
    
    # Store UAT percentages as numeric values for simple average calculation
    # Use -1 as marker for 'UAT NA' to distinguish from actual 0% values
    uat_exec_num = uat_exec if isinstance(uat_exec, (int, float)) else -1
    uat_pass_num = uat_pass if isinstance(uat_pass, (int, float)) else -1
    
    # Store PT percentages as numeric values for simple average calculation
    pt_exec_num = pt_exec if isinstance(pt_exec, (int, float)) else 0
    pt_pass_num = pt_pass if isinstance(pt_pass, (int, float)) else 0
    
    html_content += f"""                        <tr data-po="{product_owner}" 
                            data-pt-stories="{int(total_pt_stories)}" data-pt-delivered="{int(pt_delivered)}" data-pt-not-delivered="{int(pt_not_delivered)}"
                            data-pt-passed="{int(pt_passed)}" data-pt-failed="{int(pt_failed)}" data-pt-blocked="{int(pt_blocked)}" data-pt-not-run="{int(pt_not_run)}" data-pt-total="{int(pt_total)}"
                            data-pt-exec-pct="{pt_exec_num}" data-pt-pass-pct="{pt_pass_num}"
                            data-uat-stories="{int(total_uat_stories)}" data-uat-delivered="{int(uat_delivered)}" data-uat-not-delivered="{int(uat_not_delivered)}"
                            data-uat-exec-pct="{uat_exec_num}" data-uat-pass-pct="{uat_pass_num}"
                            data-bugs="{int(row_total_bugs)}" data-critical="{int(row_critical)}" data-high="{int(row_high)}" data-medium="{int(row_medium)}" data-low="{int(row_low)}">
                            <td>{parent_id}</td>
                            <td>{parent_title}</td>
                            <td>{total_pt_stories}</td>
                            <td>{pt_delivered}</td>
                            <td>{pt_not_delivered}</td>
                            <td>{pt_exec_display}</td>
                            <td>{pt_pass_display}</td>
                            <td>{total_uat_stories}</td>
                            <td>{uat_delivered}</td>
                            <td>{uat_not_delivered}</td>
                            <td>{uat_exec_display}</td>
                            <td>{uat_pass_display}</td>
                            <td>{row_total_bugs}</td>
                            <td>{row_critical}</td>
                            <td>{row_high}</td>
                            <td>{row_medium}</td>
                            <td>{row_low}</td>
                        </tr>
"""

html_content += """                    </tbody>
                </table>
            </div>
            
            <!-- Note Section -->
            <div style="margin-top: 30px; padding: 12px; background: #f8f9fa; border-radius: 8px; border-left: 4px solid #667eea; font-size: 0.75em;">
                <h4 style="color: #2c3e50; margin-bottom: 8px; font-size: 0.9em;">📝 Note:</h4>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                    <ul style="list-style-type: disc; margin-left: 20px; line-height: 1.4; font-size: 0.85em;">
                        <li><strong>Total PT:</strong> Total Stories Tested by PT</li>
                        <li><strong>PT - D:</strong> Stories delivered to PT Testing</li>
                        <li><strong>PT - ND:</strong> Stories Not delivered to PT Testing</li>
                        <li><strong>PT Exec:</strong> PT Execution %</li>
                        <li><strong>PT Pass:</strong> PT Pass %</li>
                    </ul>
                    <ul style="list-style-type: disc; margin-left: 20px; line-height: 1.4; font-size: 0.85em;">
                        <li><strong>Total UAT:</strong> Total Stories Tested by UAT</li>
                        <li><strong>UAT - D:</strong> Stories delivered to UAT Testing</li>
                        <li><strong>UAT - ND:</strong> Stories Not delivered to UAT Testing</li>
                        <li><strong>UAT Exec:</strong> UAT Execution %</li>
                        <li><strong>UAT Pass:</strong> UAT Pass %</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <div id="defects" class="tab-content">
            <h2 class="section-title">🐛 Defect Summary</h2>
            <p style="text-align: center; margin-bottom: 30px; font-size: 0.95em;">
                <strong>Defect Query:</strong> 
                <a href="https://dev.azure.com/accenturecio08/AutomationProcess_29697/_queries/query/360e2468-0efc-4818-bf48-a126887c81e4/" 
                   target="_blank" 
                   style="color: #0066cc; text-decoration: none; font-weight: 500;">
                    Open Defects Query
                </a>
            </p>
            
            <!-- Filters Section -->
            <div class="filter-section" style="margin-bottom: 30px;">
                <div class="filter-group">
                    <label for="defectSummaryAdPocFilter">AD POC:</label>
                    <select id="defectSummaryAdPocFilter" onchange="filterDefectSummary()">
                        <option value="">All</option>
"""

# Add AD POC options for defect summary filter
for ad_poc in unique_ad_pocs_defect:
    html_content += f"""                        <option value=\"{ad_poc}\">{ad_poc}</option>
"""

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectSummarySmPocFilter">SM POC:</label>
                    <select id="defectSummarySmPocFilter" onchange="filterDefectSummary()" disabled>
                        <option value="">All</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectSummaryMPocFilter">M POC:</label>
                    <select id="defectSummaryMPocFilter" onchange="filterDefectSummary()" disabled>
                        <option value="">All</option>
"""

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label style="visibility: hidden;">Action:</label>
                    <button onclick="resetDefectSummaryFilters()" style="padding: 4px 8px; border-radius: 3px; border: 1px solid #f5576c; background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; cursor: pointer; font-weight: normal; font-size: 11px; transition: all 0.3s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
                        🔄 Reset Filters
                    </button>
                </div>
            </div>
"""

html_content += f"""            
            <div class="section-heading">Total Defect breakdown</div>
            <div class="table-wrapper">
                <table class="milestone-table" id="totalDefectTable" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr>
                            <th style="width: 30%;">Category</th>
                            <th style="width: 14%;">Total</th>
                            <th style="width: 14%; background: #8B0000;">Critical</th>
                            <th style="width: 14%; background: #FF0000;">High</th>
                            <th style="width: 14%; background: #FFBF00; color: #333;">Medium</th>
                            <th style="width: 14%; background: #FFFACD; color: #333;">Low</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="font-weight: bold;">Total Bugs</td>
                            <td style="font-weight: bold;">{total_bugs}</td>
                            <td>{critical_bugs}</td>
                            <td>{high_bugs}</td>
                            <td>{medium_bugs}</td>
                            <td>{low_bugs}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold;">Active</td>
                            <td style="font-weight: bold;">{total_active_bugs}</td>
                            <td>{active_critical}</td>
                            <td>{active_high}</td>
                            <td>{active_medium}</td>
                            <td>{active_low}</td>
                        </tr>
                        <tr style="background-color: #e8f5e9;">
                            <td style="font-weight: bold;">Resolved/RTD</td>
                            <td style="font-weight: bold;">{total_rtd_bugs}</td>
                            <td>{rtd_critical}</td>
                            <td>{rtd_high}</td>
                            <td>{rtd_medium}</td>
                            <td>{rtd_low}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div class="section-heading" style="margin-top: 40px;">PT Defect Breakdown</div>
            <div class="table-wrapper">
                <table class="milestone-table" id="ptDefectTable" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr>
                            <th style="width: 30%;">Category</th>
                            <th style="width: 14%;">Total</th>
                            <th style="width: 14%; background: #8B0000;">Critical</th>
                            <th style="width: 14%; background: #FF0000;">High</th>
                            <th style="width: 14%; background: #FFBF00; color: #333;">Medium</th>
                            <th style="width: 14%; background: #FFFACD; color: #333;">Low</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="font-weight: bold;">Total PT</td>
                            <td style="font-weight: bold;">{total_pt_bugs}</td>
                            <td>{pt_critical}</td>
                            <td>{pt_high}</td>
                            <td>{pt_medium}</td>
                            <td>{pt_low}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold;">Active PT</td>
                            <td style="font-weight: bold;">{total_pt_active_bugs}</td>
                            <td>{pt_active_critical}</td>
                            <td>{pt_active_high}</td>
                            <td>{pt_active_medium}</td>
                            <td>{pt_active_low}</td>
                        </tr>
                        <tr style="background-color: #e8f5e9;">
                            <td style="font-weight: bold;">Resolved/RTD PT</td>
                            <td style="font-weight: bold;">{total_pt_rtd_bugs}</td>
                            <td>{pt_rtd_critical}</td>
                            <td>{pt_rtd_high}</td>
                            <td>{pt_rtd_medium}</td>
                            <td>{pt_rtd_low}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div class="section-heading" style="margin-top: 40px;">UAT Defect Breakdown</div>
            <div class="table-wrapper">
                <table class="milestone-table" id="uatDefectTable" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr>
                            <th style="width: 30%;">Category</th>
                            <th style="width: 14%;">Total</th>
                            <th style="width: 14%; background: #8B0000;">Critical</th>
                            <th style="width: 14%; background: #FF0000;">High</th>
                            <th style="width: 14%; background: #FFBF00; color: #333;">Medium</th>
                            <th style="width: 14%; background: #FFFACD; color: #333;">Low</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="font-weight: bold;">Total UAT</td>
                            <td style="font-weight: bold;">{total_uat_bugs}</td>
                            <td>{uat_critical}</td>
                            <td>{uat_high}</td>
                            <td>{uat_medium}</td>
                            <td>{uat_low}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold;">Active UAT</td>
                            <td style="font-weight: bold;">{total_uat_active_bugs}</td>
                            <td>{uat_active_critical}</td>
                            <td>{uat_active_high}</td>
                            <td>{uat_active_medium}</td>
                            <td>{uat_active_low}</td>
                        </tr>
                        <tr style="background-color: #e8f5e9;">
                            <td style="font-weight: bold;">Resolved/RTD UAT</td>
                            <td style="font-weight: bold;">{total_uat_rtd_bugs}</td>
                            <td>{uat_rtd_critical}</td>
                            <td>{uat_rtd_high}</td>
                            <td>{uat_rtd_medium}</td>
                            <td>{uat_rtd_low}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
"""

html_content += """            
            <div class="section-heading">Detailed Defect Breakdown by Node</div>
            
            <div class="filter-container">
                <div class="filter-group">
                    <label for="defectAdPocFilter">AD POC:</label>
                    <select id="defectAdPocFilter" onchange="updateDefectFilters()">
                        <option value="all">All AD POCs</option>
"""

# Add AD POC options for defect filter
for ad_poc in unique_ad_pocs_defect:
    html_content += f"""                        <option value="{ad_poc}">{ad_poc}</option>
"""

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectSmPocFilter">SM POC:</label>
                    <select id="defectSmPocFilter" onchange="updateDefectFilters()" disabled>
                        <option value="all">All SM POCs</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectMPocFilter">M POC:</label>
                    <select id="defectMPocFilter" onchange="updateDefectFilters()" disabled>
                        <option value="all">All M POCs</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectStageFilter">Stage Found:</label>
                    <select id="defectStageFilter" onchange="updateDefectFilters()">
                        <option value="all">All</option>
                        <option value="PT">PT</option>
                        <option value="UAT">UAT</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectStateFilter">State:</label>
                    <select id="defectStateFilter" onchange="updateDefectFilters()">
                        <option value="all">All</option>
"""

# Add State filter options
for state in unique_states:
    html_content += f"""                        <option value="{state}">{state}</option>
"""

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="defectCategoryFilter">Defect Category:</label>
                    <select id="defectCategoryFilter" onchange="updateDefectFilters()">
                        <option value="all">All</option>
                        <option value="Sanity">Sanity</option>
                        <option value="Regression">Regression</option>
                        <option value="Existing Prod Issue">Existing Prod Issue</option>
                        <option value="Feature related">Feature related</option>
                    </select>
                </div>
            </div>
            
            <div class="sub-tabs" style="margin-top: 20px;">
                <button class="sub-tab active" onclick="showDefectSubTab('countByNode')">Defect Count by Node Name</button>
                <button class="sub-tab" onclick="showDefectSubTab('defectDetails')">Defect Details</button>
            </div>
            
            <div id="countByNode" class="sub-tab-content active">
            <div class="table-wrapper">
                <table class="milestone-table" id="defectNodeTable">
                    <thead>
                        <tr>
                            <th>Node Name</th>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Critical</th>
                            <th>High</th>
                            <th>Medium</th>
                            <th>Low</th>
                            <th>Total</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Calculate initial totals for all defects
total_critical = node_breakdown_df['Critical'].sum()
total_high = node_breakdown_df['High'].sum()
total_medium = node_breakdown_df['Medium'].sum()
total_low = node_breakdown_df['Low'].sum()
total_all = node_breakdown_df['Total'].sum()

# Add total row at the top with id for dynamic updates
html_content += f"""                        <tr id="defect-total-row" style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.1em;">
                            <td colspan="4" style="text-align: center;">TOTAL</td>
                            <td id="defect-total-critical">{total_critical}</td>
                            <td id="defect-total-high">{total_high}</td>
                            <td id="defect-total-medium">{total_medium}</td>
                            <td id="defect-total-low">{total_low}</td>
                            <td id="defect-total-all"><strong>{total_all}</strong></td>
                        </tr>
"""

# Add node breakdown rows with data attributes including stage-specific, state-specific, and category-specific counts
for _, row in node_breakdown_df.iterrows():
    # Build state data attributes dynamically
    state_attrs = ''
    for state in unique_states:
        state_key = state.replace(' ', '_').replace('-', '_')
        state_attrs += f'data-{state_key.lower()}-critical="{row[f"{state_key}_Critical"]}" '
        state_attrs += f'data-{state_key.lower()}-high="{row[f"{state_key}_High"]}" '
        state_attrs += f'data-{state_key.lower()}-medium="{row[f"{state_key}_Medium"]}" '
        state_attrs += f'data-{state_key.lower()}-low="{row[f"{state_key}_Low"]}" '
        state_attrs += f'data-{state_key.lower()}-total="{row[f"{state_key}_Total"]}" '
    
    # Build category data attributes dynamically
    category_attrs = ''
    defect_categories = ['Sanity', 'Regression', 'Existing_Prod_Issue', 'Feature_related']
    for category in defect_categories:
        category_attrs += f'data-{category.lower()}-critical="{row[f"{category}_Critical"]}" '
        category_attrs += f'data-{category.lower()}-high="{row[f"{category}_High"]}" '
        category_attrs += f'data-{category.lower()}-medium="{row[f"{category}_Medium"]}" '
        category_attrs += f'data-{category.lower()}-low="{row[f"{category}_Low"]}" '
        category_attrs += f'data-{category.lower()}-total="{row[f"{category}_Total"]}" '
    
    html_content += f"""                        <tr data-ad-poc="{row['AD POC']}" data-sm-poc="{row['SM POC']}" data-m-poc="{row['M POC']}" 
                            data-pt-critical="{row['PT_Critical']}" data-pt-high="{row['PT_High']}" data-pt-medium="{row['PT_Medium']}" data-pt-low="{row['PT_Low']}" data-pt-total="{row['PT_Total']}"
                            data-uat-critical="{row['UAT_Critical']}" data-uat-high="{row['UAT_High']}" data-uat-medium="{row['UAT_Medium']}" data-uat-low="{row['UAT_Low']}" data-uat-total="{row['UAT_Total']}"
                            {state_attrs}
                            {category_attrs}>
                            <td>{row['Node Name']}</td>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Critical']}</td>
                            <td>{row['High']}</td>
                            <td>{row['Medium']}</td>
                            <td>{row['Low']}</td>
                            <td><strong>{row['Total']}</strong></td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            </div>
            
            <div id="defectDetails" class="sub-tab-content">
            <div class="table-wrapper">
                <table class="milestone-table" id="defectDetailsTable">
                    <thead>
                        <tr>
                            <th>Node Name</th>
                            <th>Defect Record</th>
                            <th>TextVerification</th>
                            <th>ID</th>
                            <th>Title</th>
                            <th>Severity</th>
                            <th>State</th>
                            <th>StageFound</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add defect details rows from bug_summary_df
for _, row in bug_summary_df.iterrows():
    node_name = row.get('Node Name', '')
    defect_record = row.get('Defect Record', '')
    text_verification = row.get('TextVerification', '')
    bug_id = row.get('ID', '')
    title = row.get('Title', '')
    severity = row.get('Severity', '')
    state = row.get('State', '')
    stage_found = row.get('StageFound', '')
    
    # Get POC information for filtering
    ad_poc = row.get('AD POC', 'yet to assign')
    sm_poc = row.get('SM POC', 'yet to assign')
    m_poc = row.get('M POC', 'yet to assign')
    defect_category = row.get('Defect Category', 'Feature related')
    
    html_content += f"""                        <tr data-ad-poc="{ad_poc}" data-sm-poc="{sm_poc}" data-m-poc="{m_poc}" 
                            data-stage="{stage_found}" data-state="{state}" data-category="{defect_category}">
                            <td>{node_name}</td>
                            <td>{defect_record}</td>
                            <td>{text_verification}</td>
                            <td>{bug_id}</td>
                            <td>{title}</td>
                            <td>{severity}</td>
                            <td>{state}</td>
                            <td>{stage_found}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            </div>
        </div>
        
        <div id="highlights" class="tab-content">
            <h2 class="section-title">🔍 Key Highlights and Risks</h2>
            
            <div class="section-heading">User stories - Planned PT date Not filled ({planned_pt_not_filled_count} stories)</div>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add rows for Planned PT date Not filled
for _, row in planned_pt_not_filled.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Count']}</td>
                        </tr>
"""

# Add total row
html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td colspan="3">Total</td>
                            <td>{planned_pt_not_filled_count}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            
            <div class="section-heading">User stories - Not Delivered to PT as per plan ({not_delivered_pt_count} stories)</div>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add rows for Not Delivered to PT
for _, row in not_delivered_pt.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Count']}</td>
                        </tr>
"""

# Add total row
html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td colspan="3">Total</td>
                            <td>{not_delivered_pt_count}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            
            <div class="section-heading">User stories - Not Delivered to UAT as per plan ({not_delivered_uat_count} stories)</div>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add rows for Not Delivered to UAT
for _, row in not_delivered_uat.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Count']}</td>
                        </tr>
"""

# Add total row
html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td colspan="3">Total</td>
                            <td>{not_delivered_uat_count}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            
            <div class="section-heading">User stories - Beyond PT Cut off Date ({beyond_pt_cutoff_count} stories)</div>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add rows for Beyond PT Cut off Date
for _, row in beyond_pt_cutoff.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Count']}</td>
                        </tr>
"""

# Add total row
html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td colspan="3">Total</td>
                            <td>{beyond_pt_cutoff_count}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            
            <div class="section-heading">User stories - Parallel UAT ({parallel_uat_count} stories)</div>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add rows for Parallel UAT
for _, row in parallel_uat.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Count']}</td>
                        </tr>
"""

# Add total row
html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td colspan="3">Total</td>
                            <td>{parallel_uat_count}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
            
            <div class="section-heading">User stories - Added Post User story Freeze date and CCB Approved ({ccb_approved_count} stories)</div>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Add rows for CCB Approved
for _, row in ccb_approved.iterrows():
    html_content += f"""                        <tr>
                            <td>{row['AD POC']}</td>
                            <td>{row['SM POC']}</td>
                            <td>{row['M POC']}</td>
                            <td>{row['Count']}</td>
                        </tr>
"""

# Add total row
html_content += f"""                        <tr style="background-color: #667eea; color: white; font-weight: bold; font-size: 1.05em;">
                            <td colspan="3">Total</td>
                            <td>{ccb_approved_count}</td>
                        </tr>
"""

html_content += f"""                    </tbody>
                </table>
            </div>
        </div>
        
        <div id="overall-defects" class="tab-content">
            <h2 class="section-title">📋 Over All Defect Summary</h2>
            <p style="text-align: center; margin-bottom: 30px; font-size: 0.95em;">
                <span style="font-weight: 600; color: #333;">Defect Query:</span> 
                <a href="https://dev.azure.com/accenturecio08/AutomationProcess_29697/_queries/query/7e2da104-c0c5-4558-9082-abf2c349f015/" 
                   target="_blank" 
                   style="color: #0066cc; text-decoration: none; font-weight: 500;">
                    Complete Defect List
                </a>
            </p>
            
            <!-- Filters Section -->
            <div class="filter-section">
                <div class="filter-group">
                    <label for="overallDefectAdPocFilter">AD POC:</label>
                    <select id="overallDefectAdPocFilter" onchange="filterOverallDefects()">
                        <option value="">All</option>
"""

# Add unique AD POC options
ad_pocs = sorted(overall_defect_summary_df['AD POC'].dropna().unique())
for poc in ad_pocs:
    html_content += f"""                        <option value="{poc}">{poc}</option>\n"""

html_content += """                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="overallDefectSmPocFilter">SM POC:</label>
                    <select id="overallDefectSmPocFilter" onchange="filterOverallDefects()" disabled>
                        <option value="">All</option>
                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="overallDefectMPocFilter">M POC:</label>
                    <select id="overallDefectMPocFilter" onchange="filterOverallDefects()" disabled>
                        <option value="">All</option>
"""

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label style="visibility: hidden;">Action:</label>
                    <button onclick="resetOverallDefectFilters()" style="padding: 4px 8px; border-radius: 3px; border: 1px solid #f5576c; background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; cursor: pointer; font-weight: normal; font-size: 11px; transition: all 0.3s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
                        🔄 Reset Filters
                    </button>
                </div>
            </div>
            
            <!-- Over All Defect Count Section -->
            <h2 class="section-title" style="margin-top: 30px;">📊 Over All Defect Count</h2>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th style="width: 30%"></th>
                            <th style="background: #333; color: white;">Total</th>
                            <th style="background: #8B0000; color: white;">Critical</th>
                            <th style="background: #FF0000; color: white;">High</th>
                            <th style="background: #FFBF00; color: #333;">Medium</th>
                            <th style="background: #FFFACD; color: #333;">Low</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr style="color: black;">
                            <td style="font-weight: bold; background: #f8f9fa;">Total bugs raised</td>
                            <td id="overallTotalBugs">-</td>
                            <td id="overallCritical">-</td>
                            <td id="overallHigh">-</td>
                            <td id="overallMedium">-</td>
                            <td id="overallLow">-</td>
                        </tr>
                        <tr style="color: red;">
                            <td style="font-weight: bold; background: #ffebee;">Active</td>
                            <td id="overallActive">-</td>
                            <td id="overallActiveCritical">-</td>
                            <td id="overallActiveHigh">-</td>
                            <td id="overallActiveMedium">-</td>
                            <td id="overallActiveLow">-</td>
                        </tr>
                        <tr style="color: #FF8C00;">
                            <td style="font-weight: bold; background: #fff8e1;">Fixed and Ready to Deploy</td>
                            <td id="overallFixedReady">-</td>
                            <td id="overallFixedCritical">-</td>
                            <td id="overallFixedHigh">-</td>
                            <td id="overallFixedMedium">-</td>
                            <td id="overallFixedLow">-</td>
                        </tr>
                        <tr style="color: green;">
                            <td style="font-weight: bold; background: #e8f5e9;">Under Testing</td>
                            <td id="overallUnderTesting">-</td>
                            <td id="overallTestingCritical">-</td>
                            <td id="overallTestingHigh">-</td>
                            <td id="overallTestingMedium">-</td>
                            <td id="overallTestingLow">-</td>
                        </tr>
                        <tr style="color: blue;">
                            <td style="font-weight: bold; background: #e3f2fd;">Closed</td>
                            <td id="overallClosed">-</td>
                            <td id="overallClosedCritical">-</td>
                            <td id="overallClosedHigh">-</td>
                            <td id="overallClosedMedium">-</td>
                            <td id="overallClosedLow">-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!-- PT Defect Count Section -->
            <h2 class="section-title" style="margin-top: 50px;">📊 PT Defect Count</h2>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th style="width: 30%"></th>
                            <th style="background: #333; color: white;">Total</th>
                            <th style="background: #8B0000; color: white;">Critical</th>
                            <th style="background: #FF0000; color: white;">High</th>
                            <th style="background: #FFBF00; color: #333;">Medium</th>
                            <th style="background: #FFFACD; color: #333;">Low</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr style="color: black;">
                            <td style="font-weight: bold; background: #f8f9fa;">Total bugs raised</td>
                            <td id="ptTotalBugs">-</td>
                            <td id="ptCritical">-</td>
                            <td id="ptHigh">-</td>
                            <td id="ptMedium">-</td>
                            <td id="ptLow">-</td>
                        </tr>
                        <tr style="color: red;">
                            <td style="font-weight: bold; background: #ffebee;">Active</td>
                            <td id="ptActive">-</td>
                            <td id="ptActiveCritical">-</td>
                            <td id="ptActiveHigh">-</td>
                            <td id="ptActiveMedium">-</td>
                            <td id="ptActiveLow">-</td>
                        </tr>
                        <tr style="color: #FF8C00;">
                            <td style="font-weight: bold; background: #fff8e1;">Fixed and Ready to Deploy</td>
                            <td id="ptFixedReady">-</td>
                            <td id="ptFixedCritical">-</td>
                            <td id="ptFixedHigh">-</td>
                            <td id="ptFixedMedium">-</td>
                            <td id="ptFixedLow">-</td>
                        </tr>
                        <tr style="color: green;">
                            <td style="font-weight: bold; background: #e8f5e9;">Under Testing</td>
                            <td id="ptUnderTesting">-</td>
                            <td id="ptTestingCritical">-</td>
                            <td id="ptTestingHigh">-</td>
                            <td id="ptTestingMedium">-</td>
                            <td id="ptTestingLow">-</td>
                        </tr>
                        <tr style="color: blue;">
                            <td style="font-weight: bold; background: #e3f2fd;">Closed</td>
                            <td id="ptClosed">-</td>
                            <td id="ptClosedCritical">-</td>
                            <td id="ptClosedHigh">-</td>
                            <td id="ptClosedMedium">-</td>
                            <td id="ptClosedLow">-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!-- UAT Defect Count Section -->
            <h2 class="section-title" style="margin-top: 50px;">📊 UAT Defect Count</h2>
            <div class="table-wrapper">
                <table class="milestone-table">
                    <thead>
                        <tr>
                            <th style="width: 30%"></th>
                            <th style="background: #333; color: white;">Total</th>
                            <th style="background: #8B0000; color: white;">Critical</th>
                            <th style="background: #FF0000; color: white;">High</th>
                            <th style="background: #FFBF00; color: #333;">Medium</th>
                            <th style="background: #FFFACD; color: #333;">Low</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr style="color: black;">
                            <td style="font-weight: bold; background: #f8f9fa;">Total bugs raised</td>
                            <td id="uatTotalBugs">-</td>
                            <td id="uatCritical">-</td>
                            <td id="uatHigh">-</td>
                            <td id="uatMedium">-</td>
                            <td id="uatLow">-</td>
                        </tr>
                        <tr style="color: red;">
                            <td style="font-weight: bold; background: #ffebee;">Active</td>
                            <td id="uatActive">-</td>
                            <td id="uatActiveCritical">-</td>
                            <td id="uatActiveHigh">-</td>
                            <td id="uatActiveMedium">-</td>
                            <td id="uatActiveLow">-</td>
                        </tr>
                        <tr style="color: #FF8C00;">
                            <td style="font-weight: bold; background: #fff8e1;">Fixed and Ready to Deploy</td>
                            <td id="uatFixedReady">-</td>
                            <td id="uatFixedCritical">-</td>
                            <td id="uatFixedHigh">-</td>
                            <td id="uatFixedMedium">-</td>
                            <td id="uatFixedLow">-</td>
                        </tr>
                        <tr style="color: green;">
                            <td style="font-weight: bold; background: #e8f5e9;">Under Testing</td>
                            <td id="uatUnderTesting">-</td>
                            <td id="uatTestingCritical">-</td>
                            <td id="uatTestingHigh">-</td>
                            <td id="uatTestingMedium">-</td>
                            <td id="uatTestingLow">-</td>
                        </tr>
                        <tr style="color: blue;">
                            <td style="font-weight: bold; background: #e3f2fd;">Closed</td>
                            <td id="uatClosed">-</td>
                            <td id="uatClosedCritical">-</td>
                            <td id="uatClosedHigh">-</td>
                            <td id="uatClosedMedium">-</td>
                            <td id="uatClosedLow">-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
"""

html_content += f"""        
        <div class="footer">
            Generated on {current_date} | Daily Status Report Dashboard
        </div>
    </div>
    
    <script>
        // POC hierarchies for filtering
"""
html_content += f"        const pocHierarchyExec = {poc_hierarchy_json};\n"
html_content += f"        const pocHierarchyModule = {poc_hierarchy_module_json};\n"
html_content += f"        const pocHierarchyDefect = {poc_hierarchy_defect_json};\n"
html_content += f"        const defectSummaryData = {defect_summary_json};\n"
html_content += f"        const overallDefectsData = {overall_defects_json};\n"
html_content += f"        const activeStates = {active_states_json};\n"
html_content += f"        const fixedReadyStates = {fixed_ready_states_json};\n"
html_content += f"        const underTestingStates = {under_testing_states_json};\n"
html_content += f"        const closedStates = {closed_states_json};\n"
html_content += """
        function showTab(tabName) {
            // Hide all tab contents
            const contents = document.querySelectorAll('.tab-content');
            contents.forEach(content => content.classList.remove('active'));
            
            // Remove active class from all tabs
            const tabs = document.querySelectorAll('.tab');
            tabs.forEach(tab => tab.classList.remove('active'));
            
            // Show selected tab content
            document.getElementById(tabName).classList.add('active');
            
            // Add active class to clicked tab button
            const clickedTab = document.querySelector('.tab[onclick*="' + tabName + '"]');
            if (clickedTab) {
                clickedTab.classList.add('active');
            }
            
            // Initialize overall defects when that tab is shown
            if (tabName === 'overall-defects') {
                filterOverallDefects();
            }
        }
        
        function filterByProductOwner() {
            const selectedPO = document.getElementById('poFilter').value;
            const table = document.getElementById('consolidatedTable');
            const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            
            // Initialize totals
            let totalPTStories = 0;
            let totalPTDelivered = 0;
            let totalPTNotDelivered = 0;
            let totalPTPassed = 0;
            let totalPTFailed = 0;
            let totalPTBlocked = 0;
            let totalPTNotRun = 0;
            let totalPTTotal = 0;
            
            let totalUATStories = 0;
            let totalUATDelivered = 0;
            let totalUATNotDelivered = 0;
            let ptExecSum = 0;
            let ptPassSum = 0;
            let ptCount = 0;
            let uatExecSum = 0;
            let uatPassSum = 0;
            let uatCount = 0;
            
            let totalBugs = 0;
            let totalCritical = 0;
            let totalHigh = 0;
            let totalMedium = 0;
            let totalLow = 0;
            
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                
                // Skip the total row
                if (row.id === 'consolidated-total-row') {
                    continue;
                }
                
                const po = row.getAttribute('data-po');
                
                if (selectedPO === 'All' || po === selectedPO) {
                    row.style.display = '';
                    
                    // Add to totals
                    const rowPTStories = parseInt(row.getAttribute('data-pt-stories')) || 0;
                    totalPTStories += rowPTStories;
                    totalPTDelivered += parseInt(row.getAttribute('data-pt-delivered')) || 0;
                    totalPTNotDelivered += parseInt(row.getAttribute('data-pt-not-delivered')) || 0;
                    totalPTPassed += parseInt(row.getAttribute('data-pt-passed')) || 0;
                    totalPTFailed += parseInt(row.getAttribute('data-pt-failed')) || 0;
                    totalPTBlocked += parseInt(row.getAttribute('data-pt-blocked')) || 0;
                    totalPTNotRun += parseInt(row.getAttribute('data-pt-not-run')) || 0;
                    totalPTTotal += parseInt(row.getAttribute('data-pt-total')) || 0;
                    
                    // For PT percentages, calculate simple average (not weighted)
                    const ptExecPct = parseFloat(row.getAttribute('data-pt-exec-pct'));
                    const ptPassPct = parseFloat(row.getAttribute('data-pt-pass-pct'));
                    
                    // Include only rows with valid PT percentages (> 0)
                    if (rowPTStories > 0 && ptExecPct >= 0 && ptPassPct >= 0) {
                        ptExecSum += ptExecPct;
                        ptPassSum += ptPassPct;
                        ptCount++;
                    }
                    
                    const rowUATStories = parseInt(row.getAttribute('data-uat-stories')) || 0;
                    totalUATStories += rowUATStories;
                    totalUATDelivered += parseInt(row.getAttribute('data-uat-delivered')) || 0;
                    totalUATNotDelivered += parseInt(row.getAttribute('data-uat-not-delivered')) || 0;
                    
                    // For UAT percentages, calculate simple average (not weighted)
                    const uatExecPct = parseFloat(row.getAttribute('data-uat-exec-pct'));
                    const uatPassPct = parseFloat(row.getAttribute('data-uat-pass-pct'));
                    
                    // Include only rows with valid UAT percentages (>= 0), exclude UAT NA rows (marked with -1)
                    if (rowUATStories > 0 && uatExecPct >= 0 && uatPassPct >= 0) {
                        uatExecSum += uatExecPct;
                        uatPassSum += uatPassPct;
                        uatCount++;
                    }
                    
                    totalBugs += parseInt(row.getAttribute('data-bugs')) || 0;
                    totalCritical += parseInt(row.getAttribute('data-critical')) || 0;
                    totalHigh += parseInt(row.getAttribute('data-high')) || 0;
                    totalMedium += parseInt(row.getAttribute('data-medium')) || 0;
                    totalLow += parseInt(row.getAttribute('data-low')) || 0;
                } else {
                    row.style.display = 'none';
                }
            }
            
            // Calculate percentages
            // PT Execution % = simple average of row percentages
            const ptExecPct = ptCount > 0 ? ((ptExecSum / ptCount) * 100).toFixed(2) : '0.00';
            // PT Pass % = simple average of row percentages
            const ptPassPct = ptCount > 0 ? ((ptPassSum / ptCount) * 100).toFixed(2) : '0.00';
            // UAT Execution % = simple average of row percentages
            const uatExecPct = uatCount > 0 ? ((uatExecSum / uatCount) * 100).toFixed(2) : '0.00';
            // UAT Pass % = simple average of row percentages
            const uatPassPct = uatCount > 0 ? ((uatPassSum / uatCount) * 100).toFixed(2) : '0.00';
            
            // Update total row
            document.getElementById('cons-total-pt-stories').textContent = totalPTStories;
            document.getElementById('cons-total-pt-delivered').textContent = totalPTDelivered;
            document.getElementById('cons-total-pt-not-delivered').textContent = totalPTNotDelivered;
            
            // PT Exec and PT Pass: use hardcoded values when filter is "All", otherwise use simple average
            if (selectedPO === 'All') {
                // Use hardcoded values from initial load
                const totalRow = document.getElementById('consolidated-total-row');
                const hardcodedPtExec = totalRow.getAttribute('data-hardcoded-pt-exec');
                const hardcodedPtPass = totalRow.getAttribute('data-hardcoded-pt-pass');
                document.getElementById('cons-total-pt-exec').textContent = hardcodedPtExec + '%';
                document.getElementById('cons-total-pt-pass').textContent = hardcodedPtPass + '%';
            } else {
                // Use simple average from filtered rows
                document.getElementById('cons-total-pt-exec').textContent = ptExecPct + '%';
                document.getElementById('cons-total-pt-pass').textContent = ptPassPct + '%';
            }
            
            document.getElementById('cons-total-uat-stories').textContent = totalUATStories;
            document.getElementById('cons-total-uat-delivered').textContent = totalUATDelivered;
            document.getElementById('cons-total-uat-not-delivered').textContent = totalUATNotDelivered;
            
            // UAT Exec and UAT Pass: use hardcoded values when filter is "All", otherwise calculate simple average
            if (selectedPO === 'All') {
                // Use hardcoded values from Overview tab
                const totalRow = document.getElementById('consolidated-total-row');
                const hardcodedUatExec = totalRow.getAttribute('data-hardcoded-uat-exec');
                const hardcodedUatPass = totalRow.getAttribute('data-hardcoded-uat-pass');
                document.getElementById('cons-total-uat-exec').textContent = hardcodedUatExec + '%';
                document.getElementById('cons-total-uat-pass').textContent = hardcodedUatPass + '%';
            } else {
                // Calculate simple average from filtered rows
                document.getElementById('cons-total-uat-exec').textContent = uatExecPct + '%';
                document.getElementById('cons-total-uat-pass').textContent = uatPassPct + '%';
            }
            
            document.getElementById('cons-total-bugs').textContent = totalBugs;
            document.getElementById('cons-total-critical').textContent = totalCritical;
            document.getElementById('cons-total-high').textContent = totalHigh;
            document.getElementById('cons-total-medium').textContent = totalMedium;
            document.getElementById('cons-total-low').textContent = totalLow;
        }
        
        function resetPOFilter() {
            // Reset Product Owner filter to "All"
            document.getElementById('poFilter').value = 'All';
            
            // Trigger the filter update to refresh the display
            filterByProductOwner();
        }
        
        function showDefectSubTab(subTabName) {
            // Hide all sub-tab contents
            const subContents = document.querySelectorAll('.sub-tab-content');
            subContents.forEach(content => content.classList.remove('active'));
            
            // Remove active class from all sub-tabs
            const subTabs = document.querySelectorAll('.sub-tab');
            subTabs.forEach(tab => tab.classList.remove('active'));
            
            // Show selected sub-tab content
            document.getElementById(subTabName).classList.add('active');
            
            // Add active class to clicked sub-tab button
            const clickedSubTab = document.querySelector('.sub-tab[onclick*="' + subTabName + '"]');
            if (clickedSubTab) {
                clickedSubTab.classList.add('active');
            }
        }

        function updateExecFilters(changedLevel) {
            const adFilter = document.getElementById('exec-ad-filter');
            const smFilter = document.getElementById('exec-sm-filter');
            const mFilter = document.getElementById('exec-m-filter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;
            const selectedM = mFilter.value;

            // Update SM POC filter based on AD selection
            if (changedLevel === 'ad') {
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                if (selectedAD !== 'all' && pocHierarchyExec[selectedAD]) {
                    const smPocs = Object.keys(pocHierarchyExec[selectedAD]).sort();
                    smPocs.forEach(sm => {
                        smFilter.innerHTML += '<option value="' + sm + '">' + sm + '</option>';
                    });
                    smFilter.disabled = false; // Enable SM POC dropdown
                    mFilter.disabled = true;   // Keep M POC disabled until SM is selected
                } else if (selectedAD === 'all') {
                    smFilter.disabled = true;  // Disable SM POC when "All" is selected
                    mFilter.disabled = true;   // Disable M POC as well
                }
            }

            // Update M POC filter based on AD and SM selection
            if (changedLevel === 'ad' || changedLevel === 'sm') {
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                if (selectedAD !== 'all' && selectedSM !== 'all' && 
                    pocHierarchyExec[selectedAD] && pocHierarchyExec[selectedAD][selectedSM]) {
                    const mPocs = pocHierarchyExec[selectedAD][selectedSM].sort();
                    mPocs.forEach(m => {
                        mFilter.innerHTML += '<option value="' + m + '">' + m + '</option>';
                    });
                    mFilter.disabled = false; // Enable M POC dropdown
                } else if (selectedAD !== 'all' && selectedSM === 'all' && pocHierarchyExec[selectedAD]) {
                    // Keep M POC disabled when SM is "All"
                    mFilter.disabled = true;
                } else if (selectedAD === 'all') {
                    mFilter.disabled = true;  // Keep disabled when AD is "All"
                }
            }

            // Apply filter to execution summary rows and recalculate totals
            // Exclude Testing NA rows from the main execution summary processing
            const execRows = document.querySelectorAll('.exec-row:not(.testing-na-row)');
            let totals = {
                stories: 0, ptDelivered: 0, ptNotDelivered: 0, uatTestable: 0, uatDelivered: 0, uatNotDelivered: 0,
                passed: 0, failed: 0, blocked: 0, notRun: 0, total: 0
            };
            
            // Always consolidate by Parent ID to avoid duplicate counting when same parent has multiple POC combinations
            {
                // Group rows by Parent ID and consolidate (only for rows matching the filter)
                const parentMap = new Map();
                
                execRows.forEach(row => {
                    const rowAD = row.getAttribute('data-ad-poc');
                    const rowSM = row.getAttribute('data-sm-poc');
                    const rowM = row.getAttribute('data-m-poc');
                    
                    // Check if row matches filter criteria
                    let matchesFilter = true;
                    if (selectedAD !== 'all' && rowAD !== selectedAD) {
                        matchesFilter = false;
                    }
                    if (selectedSM !== 'all' && rowSM !== selectedSM) {
                        matchesFilter = false;
                    }
                    if (selectedM !== 'all' && rowM !== selectedM) {
                        matchesFilter = false;
                    }
                    
                    if (!matchesFilter) {
                        return; // Skip this row
                    }
                    
                    const cells = row.cells;
                    const parentId = cells[0].textContent;
                    const parentTitle = cells[1].textContent;
                    const storyIds = row.getAttribute('data-story-ids') || '';
                    
                    if (!parentMap.has(parentId)) {
                        parentMap.set(parentId, {
                            title: parentTitle,
                            storyIdSet: new Set(),
                            ptDelivered: 0, ptNotDelivered: 0, uatTestable: 0, 
                            uatDelivered: 0, uatNotDelivered: 0,
                            passed: 0, failed: 0, blocked: 0, notRun: 0, total: 0,
                            rows: []
                        });
                    }
                    
                    const parent = parentMap.get(parentId);
                    
                    // Add unique story IDs to set
                    if (storyIds) {
                        storyIds.split(',').forEach(id => parent.storyIdSet.add(id.trim()));
                    }
                    
                    // Read from data attributes (original values) instead of cell text to avoid accumulation
                    parent.ptDelivered += parseInt(row.getAttribute('data-pt-delivered')) || 0;
                    parent.ptNotDelivered += parseInt(row.getAttribute('data-pt-not-delivered')) || 0;
                    parent.uatTestable += parseInt(row.getAttribute('data-uat-testable')) || 0;
                    parent.uatDelivered += parseInt(row.getAttribute('data-uat-delivered')) || 0;
                    parent.uatNotDelivered += parseInt(row.getAttribute('data-uat-not-delivered')) || 0;
                    parent.passed += parseInt(row.getAttribute('data-passed')) || 0;
                    parent.failed += parseInt(row.getAttribute('data-failed')) || 0;
                    parent.blocked += parseInt(row.getAttribute('data-blocked')) || 0;
                    parent.notRun += parseInt(row.getAttribute('data-not-run')) || 0;
                    parent.total += parseInt(row.getAttribute('data-total')) || 0;
                    parent.rows.push(row);
                });
                
                // Hide all rows first
                execRows.forEach(row => row.style.display = 'none');
                
                // Show and update first row of each parent group with consolidated data
                parentMap.forEach((parent, parentId) => {
                    const firstRow = parent.rows[0];
                    firstRow.style.display = '';
                    
                    const cells = firstRow.cells;
                    const uniqueStoryCount = parent.storyIdSet.size;
                    cells[2].textContent = uniqueStoryCount;
                    cells[3].textContent = parent.ptDelivered;
                    cells[4].textContent = parent.ptNotDelivered;
                    cells[5].textContent = parent.uatTestable;
                    cells[6].textContent = parent.uatDelivered;
                    cells[7].textContent = parent.uatNotDelivered;
                    cells[8].textContent = parent.passed;
                    cells[9].textContent = parent.failed;
                    cells[10].textContent = parent.blocked;
                    cells[11].textContent = parent.notRun;
                    cells[12].textContent = parent.total;
                    
                    // Calculate consolidated percentages
                    const execPct = parent.total > 0 ? 
                        ((parent.passed + parent.failed) / parent.total * 100).toFixed(2) : '0.00';
                    const passPct = (parent.passed + parent.failed + parent.blocked) > 0 ? 
                        (parent.passed / (parent.passed + parent.failed + parent.blocked) * 100).toFixed(2) : '0.00';
                    
                    cells[13].textContent = execPct + '%';
                    cells[14].textContent = passPct + '%';
                    
                    // Add to totals (use unique story count)
                    totals.stories += uniqueStoryCount;
                    totals.ptDelivered += parent.ptDelivered;
                    totals.ptNotDelivered += parent.ptNotDelivered;
                    totals.uatTestable += parent.uatTestable;
                    totals.uatDelivered += parent.uatDelivered;
                    totals.uatNotDelivered += parent.uatNotDelivered;
                    totals.passed += parent.passed;
                    totals.failed += parent.failed;
                    totals.blocked += parent.blocked;
                    totals.notRun += parent.notRun;
                    totals.total += parent.total;
                });
            }
            
            // Update total row (cells start at index 1 because first cell has colspan=2)
            const totalRow = document.getElementById('exec-total-row');
            if (totalRow) {
                const cells = totalRow.cells;
                // Note: cells[0] is "TOTAL" with colspan=2, so data starts at cells[1]
                cells[1].textContent = totals.stories;
                cells[2].textContent = totals.ptDelivered;
                cells[3].textContent = totals.ptNotDelivered;
                cells[4].textContent = totals.uatTestable;
                cells[5].textContent = totals.uatDelivered;
                cells[6].textContent = totals.uatNotDelivered;
                cells[7].textContent = totals.passed;
                cells[8].textContent = totals.failed;
                cells[9].textContent = totals.blocked;
                cells[10].textContent = totals.notRun;
                cells[11].textContent = totals.total;
                
                // Calculate PT Exec and PT Pass using formula (matching Tab 1 and Tab 3)
                // PT Exec = (Passed + Failed) / Total
                const execPct = totals.total > 0 ? ((totals.passed + totals.failed) / totals.total * 100).toFixed(2) : '0.00';
                // PT Pass = Passed / (Passed + Failed + Blocked)
                const passPct = (totals.passed + totals.failed + totals.blocked) > 0 ? 
                    (totals.passed / (totals.passed + totals.failed + totals.blocked) * 100).toFixed(2) : '0.00';
                
                cells[12].textContent = execPct + '%';
                cells[13].textContent = passPct + '%';
            }
            
            // Update Testing NA Stories grand total based on visible Testing NA rows only
            // First, filter Testing NA rows based on POC selections
            let testingNATotal = 0;
            const testingNARows = document.querySelectorAll('.testing-na-row');
            testingNARows.forEach(row => {
                const rowAD = row.getAttribute('data-ad-poc');
                const rowSM = row.getAttribute('data-sm-poc');
                const rowM = row.getAttribute('data-m-poc');
                
                let showRow = true;
                
                if (selectedAD !== 'all' && rowAD !== selectedAD) {
                    showRow = false;
                }
                if (selectedSM !== 'all' && rowSM !== selectedSM) {
                    showRow = false;
                }
                if (selectedM !== 'all' && rowM !== selectedM) {
                    showRow = false;
                }
                
                row.style.display = showRow ? '' : 'none';
                
                if (showRow) {
                    // Get the count from data attribute
                    const count = parseInt(row.getAttribute('data-testing-na-count')) || 0;
                    testingNATotal += count;
                }
            });
            
            // Update the grand total display
            const grandTotalElement = document.getElementById('testing-na-grand-total');
            if (grandTotalElement) {
                grandTotalElement.textContent = testingNATotal;
            }
            
            // Update Story Summary tiles for Parent Level PT execution Status
            const execSummaryTotalStories = document.getElementById('exec-summary-total-stories');
            const execSummaryTestableStories = document.getElementById('exec-summary-testable-stories');
            const execSummaryTestingNAStories = document.getElementById('exec-summary-testing-na-stories');
            
            if (execSummaryTotalStories) {
                execSummaryTotalStories.textContent = totals.stories + testingNATotal;
            }
            if (execSummaryTestableStories) {
                execSummaryTestableStories.textContent = totals.stories;
            }
            if (execSummaryTestingNAStories) {
                execSummaryTestingNAStories.textContent = testingNATotal;
            }
        }
        
        function resetExecFilters() {
            // Reset all three filters to 'all'
            const adFilter = document.getElementById('exec-ad-filter');
            const smFilter = document.getElementById('exec-sm-filter');
            const mFilter = document.getElementById('exec-m-filter');
            
            // Set all to 'all' and disable SM and M POC dropdowns
            adFilter.value = 'all';
            smFilter.value = 'all';
            mFilter.value = 'all';
            smFilter.disabled = true;
            mFilter.disabled = true;
            
            // Trigger update which will show all rows and recalculate totals
            updateExecFilters('ad');
        }
        
        function updateModuleFilters(changedLevel) {
            const adFilter = document.getElementById('module-ad-filter');
            const smFilter = document.getElementById('module-sm-filter');
            const mFilter = document.getElementById('module-m-filter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;
            const selectedM = mFilter.value;

            // Update SM POC filter based on AD selection
            if (changedLevel === 'ad') {
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                if (selectedAD !== 'all' && pocHierarchyModule[selectedAD]) {
                    const smPocs = Object.keys(pocHierarchyModule[selectedAD]).sort();
                    smPocs.forEach(sm => {
                        smFilter.innerHTML += '<option value="' + sm + '">' + sm + '</option>';
                    });
                } else if (selectedAD === 'all') {
                    const allSMPocs = new Set();
                    Object.values(pocHierarchyModule).forEach(adData => {
                        Object.keys(adData).forEach(sm => allSMPocs.add(sm));
                    });
                    Array.from(allSMPocs).sort().forEach(sm => {
                        smFilter.innerHTML += '<option value="' + sm + '">' + sm + '</option>';
                    });
                }
            }

            // Update M POC filter based on AD and SM selection
            if (changedLevel === 'ad' || changedLevel === 'sm') {
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                if (selectedAD !== 'all' && selectedSM !== 'all' && 
                    pocHierarchyModule[selectedAD] && pocHierarchyModule[selectedAD][selectedSM]) {
                    const mPocs = pocHierarchyModule[selectedAD][selectedSM].sort();
                    mPocs.forEach(m => {
                        mFilter.innerHTML += '<option value="' + m + '">' + m + '</option>';
                    });
                } else if (selectedAD !== 'all' && selectedSM === 'all' && pocHierarchyModule[selectedAD]) {
                    const allMPocs = new Set();
                    Object.values(pocHierarchyModule[selectedAD]).forEach(mList => {
                        mList.forEach(m => allMPocs.add(m));
                    });
                    Array.from(allMPocs).sort().forEach(m => {
                        mFilter.innerHTML += '<option value="' + m + '">' + m + '</option>';
                    });
                } else if (selectedAD === 'all') {
                    const allMPocs = new Set();
                    Object.values(pocHierarchyModule).forEach(adData => {
                        Object.values(adData).forEach(mList => {
                            mList.forEach(m => allMPocs.add(m));
                        });
                    });
                    Array.from(allMPocs).sort().forEach(m => {
                        mFilter.innerHTML += '<option value="' + m + '">' + m + '</option>';
                    });
                }
            }

            // Apply filter to module rows and recalculate totals
            const moduleRows = document.querySelectorAll('.module-row');
            let totals = {
                stories: 0, ptDelivered: 0, ptNotDelivered: 0, uatTestable: 0, uatDelivered: 0, uatNotDelivered: 0,
                passed: 0, failed: 0, blocked: 0, notRun: 0, total: 0
            };
            
            // Check if all filters are set to "all" - need to consolidate by Module
            const consolidateByModule = (selectedAD === 'all' && selectedSM === 'all' && selectedM === 'all');
            
            if (consolidateByModule) {
                // Group rows by Module and consolidate
                const moduleMap = new Map();
                
                moduleRows.forEach(row => {
                    const cells = row.cells;
                    const moduleName = cells[0].textContent;
                    const storyIds = row.getAttribute('data-story-ids') || '';
                    
                    if (!moduleMap.has(moduleName)) {
                        moduleMap.set(moduleName, {
                            storyIdSet: new Set(),
                            ptDelivered: 0, ptNotDelivered: 0, uatTestable: 0, 
                            uatDelivered: 0, uatNotDelivered: 0,
                            passed: 0, failed: 0, blocked: 0, notRun: 0, total: 0,
                            rows: []
                        });
                    }
                    
                    const module = moduleMap.get(moduleName);
                    
                    // Add unique story IDs to set
                    if (storyIds) {
                        storyIds.split(',').forEach(id => module.storyIdSet.add(id.trim()));
                    }
                    
                    module.ptDelivered += parseInt(cells[2].textContent) || 0;
                    module.ptNotDelivered += parseInt(cells[3].textContent) || 0;
                    module.uatTestable += parseInt(cells[4].textContent) || 0;
                    module.uatDelivered += parseInt(cells[5].textContent) || 0;
                    module.uatNotDelivered += parseInt(cells[6].textContent) || 0;
                    module.passed += parseInt(cells[7].textContent) || 0;
                    module.failed += parseInt(cells[8].textContent) || 0;
                    module.blocked += parseInt(cells[9].textContent) || 0;
                    module.notRun += parseInt(cells[10].textContent) || 0;
                    module.total += parseInt(cells[11].textContent) || 0;
                    module.rows.push(row);
                });
                
                // Hide all rows first
                moduleRows.forEach(row => row.style.display = 'none');
                
                // Show and update first row of each module group with consolidated data
                moduleMap.forEach((module, moduleName) => {
                    const firstRow = module.rows[0];
                    firstRow.style.display = '';
                    
                    const cells = firstRow.cells;
                    const uniqueStoryCount = module.storyIdSet.size;
                    cells[1].textContent = uniqueStoryCount;
                    cells[2].textContent = module.ptDelivered;
                    cells[3].textContent = module.ptNotDelivered;
                    cells[4].textContent = module.uatTestable;
                    cells[5].textContent = module.uatDelivered;
                    cells[6].textContent = module.uatNotDelivered;
                    cells[7].textContent = module.passed;
                    cells[8].textContent = module.failed;
                    cells[9].textContent = module.blocked;
                    cells[10].textContent = module.notRun;
                    cells[11].textContent = module.total;
                    
                    // Calculate consolidated percentages
                    const execPct = module.total > 0 ? 
                        ((module.passed + module.failed) / module.total * 100).toFixed(2) : '0.00';
                    const passPct = (module.passed + module.failed + module.blocked) > 0 ? 
                        (module.passed / (module.passed + module.failed + module.blocked) * 100).toFixed(2) : '0.00';
                    
                    cells[12].textContent = execPct + '%';
                    cells[13].textContent = passPct + '%';
                    
                    // Add to totals (use unique story count)
                    totals.stories += uniqueStoryCount;
                    totals.ptDelivered += module.ptDelivered;
                    totals.ptNotDelivered += module.ptNotDelivered;
                    totals.uatTestable += module.uatTestable;
                    totals.uatDelivered += module.uatDelivered;
                    totals.uatNotDelivered += module.uatNotDelivered;
                    totals.passed += module.passed;
                    totals.failed += module.failed;
                    totals.blocked += module.blocked;
                    totals.notRun += module.notRun;
                    totals.total += module.total;
                });
            } else {
                // Normal filtering - show rows matching selected POCs
                moduleRows.forEach(row => {
                    const rowAD = row.getAttribute('data-ad-poc');
                    const rowSM = row.getAttribute('data-sm-poc');
                    const rowM = row.getAttribute('data-m-poc');
                    
                    let showRow = true;
                    
                    if (selectedAD !== 'all' && rowAD !== selectedAD) {
                        showRow = false;
                    }
                    if (selectedSM !== 'all' && rowSM !== selectedSM) {
                        showRow = false;
                    }
                    if (selectedM !== 'all' && rowM !== selectedM) {
                        showRow = false;
                    }
                    
                    row.style.display = showRow ? '' : 'none';
                    
                    // If row is visible, add to totals - read from data attributes (original values)
                    if (showRow) {
                        totals.stories += parseInt(row.getAttribute('data-stories')) || 0;
                        totals.ptDelivered += parseInt(row.getAttribute('data-pt-delivered')) || 0;
                        totals.ptNotDelivered += parseInt(row.getAttribute('data-pt-not-delivered')) || 0;
                        totals.uatTestable += parseInt(row.getAttribute('data-uat-testable')) || 0;
                        totals.uatDelivered += parseInt(row.getAttribute('data-uat-delivered')) || 0;
                        totals.uatNotDelivered += parseInt(row.getAttribute('data-uat-not-delivered')) || 0;
                        totals.passed += parseInt(row.getAttribute('data-passed')) || 0;
                        totals.failed += parseInt(row.getAttribute('data-failed')) || 0;
                        totals.blocked += parseInt(row.getAttribute('data-blocked')) || 0;
                        totals.notRun += parseInt(row.getAttribute('data-not-run')) || 0;
                        totals.total += parseInt(row.getAttribute('data-total')) || 0;
                    }
                });
            }
            
            // Update total row
            const totalRow = document.getElementById('module-total-row');
            if (totalRow) {
                const cells = totalRow.cells;
                // Note: cells[0] is "TOTAL", so data starts at cells[1]
                cells[1].textContent = totals.stories;
                cells[2].textContent = totals.ptDelivered;
                cells[3].textContent = totals.ptNotDelivered;
                cells[4].textContent = totals.uatTestable;
                cells[5].textContent = totals.uatDelivered;
                cells[6].textContent = totals.uatNotDelivered;
                cells[7].textContent = totals.passed;
                cells[8].textContent = totals.failed;
                cells[9].textContent = totals.blocked;
                cells[10].textContent = totals.notRun;
                cells[11].textContent = totals.total;
                
                // Calculate execution % and pass % from actual totals
                const execPct = totals.total > 0 ? ((totals.passed + totals.failed) / totals.total * 100).toFixed(2) : '0.00';
                const passPct = (totals.passed + totals.failed + totals.blocked) > 0 ? 
                    (totals.passed / (totals.passed + totals.failed + totals.blocked) * 100).toFixed(2) : '0.00';
                
                cells[12].textContent = execPct + '%';
                cells[13].textContent = passPct + '%';
            }
            
            // Update Testing NA Stories grand total for Module Level based on visible Testing NA rows only
            let moduleTestingNATotal = 0;
            const moduleTestingNARows = document.querySelectorAll('.module-testing-na-row');
            moduleTestingNARows.forEach(row => {
                if (row.style.display !== 'none') {
                    // Get the count from data attribute
                    const count = parseInt(row.getAttribute('data-testing-na-count')) || 0;
                    moduleTestingNATotal += count;
                }
            });
            
            // Update the grand total display for Module Level
            const moduleGrandTotalElement = document.getElementById('module-testing-na-grand-total');
            if (moduleGrandTotalElement) {
                moduleGrandTotalElement.textContent = moduleTestingNATotal;
            }
            
            // Update Story Summary tiles for Module Level PT execution Status
            const moduleSummaryTotalStories = document.getElementById('module-summary-total-stories');
            const moduleSummaryTestableStories = document.getElementById('module-summary-testable-stories');
            const moduleSummaryTestingNAStories = document.getElementById('module-summary-testing-na-stories');
            
            if (moduleSummaryTotalStories) {
                moduleSummaryTotalStories.textContent = totals.stories + moduleTestingNATotal;
            }
            if (moduleSummaryTestableStories) {
                moduleSummaryTestableStories.textContent = totals.stories;
            }
            if (moduleSummaryTestingNAStories) {
                moduleSummaryTestingNAStories.textContent = moduleTestingNATotal;
            }
        }
        
        function updateDefectFilters() {
            const adFilter = document.getElementById('defectAdPocFilter');
            const smFilter = document.getElementById('defectSmPocFilter');
            const mFilter = document.getElementById('defectMPocFilter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;
            const selectedM = mFilter.value;
            
            // Update SM POC dropdown based on selected AD POC
            if (selectedAD === 'all') {
                // If "All AD POCs" selected, disable SM and M filters
                smFilter.disabled = true;
                mFilter.disabled = true;
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
            } else {
                // Enable SM filter and populate based on selected AD POC
                smFilter.disabled = false;
                smFilter.innerHTML = '<option value="all">All SM POCs</option>';
                
                const smPocs = Object.keys(pocHierarchyDefect[selectedAD] || {});
                smPocs.forEach(sm => {
                    const option = document.createElement('option');
                    option.value = sm;
                    option.textContent = sm;
                    if (sm === selectedSM) {
                        option.selected = true;
                    }
                    smFilter.appendChild(option);
                });
                
                // If SM POC was previously selected and still valid, keep it
                if (selectedSM !== 'all' && !smPocs.includes(selectedSM)) {
                    smFilter.value = 'all';
                }
            }
            
            // Update M POC dropdown based on selected SM POC
            if (selectedSM === 'all' || selectedAD === 'all') {
                mFilter.disabled = true;
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
            } else {
                mFilter.disabled = false;
                mFilter.innerHTML = '<option value="all">All M POCs</option>';
                
                const mPocs = pocHierarchyDefect[selectedAD]?.[selectedSM] || [];
                mPocs.forEach(m => {
                    const option = document.createElement('option');
                    option.value = m;
                    option.textContent = m;
                    if (m === selectedM) {
                        option.selected = true;
                    }
                    mFilter.appendChild(option);
                });
                
                // If M POC was previously selected and still valid, keep it
                if (selectedM !== 'all' && !mPocs.includes(selectedM)) {
                    mFilter.value = 'all';
                }
            }
            
            // Filter table rows
            const table = document.getElementById('defectNodeTable');
            if (!table) return;
            
            const tbody = table.querySelector('tbody');
            const rows = tbody.querySelectorAll('tr');
            
            // Get StageFound, State, and Category filter values
            const stageFilter = document.getElementById('defectStageFilter');
            const selectedStage = stageFilter ? stageFilter.value : 'all';
            
            const stateFilter = document.getElementById('defectStateFilter');
            const selectedState = stateFilter ? stateFilter.value : 'all';
            
            const categoryFilter = document.getElementById('defectCategoryFilter');
            const selectedCategory = categoryFilter ? categoryFilter.value : 'all';
            
            // Initialize totals for visible rows
            let totalCritical = 0;
            let totalHigh = 0;
            let totalMedium = 0;
            let totalLow = 0;
            let totalAll = 0;
            
            rows.forEach(row => {
                // Skip the total row itself
                if (row.id === 'defect-total-row') return;
                const rowAD = row.getAttribute('data-ad-poc');
                const rowSM = row.getAttribute('data-sm-poc');
                const rowM = row.getAttribute('data-m-poc');
                
                let showRow = true;
                
                // Apply AD POC filter
                if (selectedAD !== 'all' && rowAD !== selectedAD) {
                    showRow = false;
                }
                
                // Apply SM POC filter
                if (selectedSM !== 'all' && rowSM !== selectedSM) {
                    showRow = false;
                }
                
                // Apply M POC filter
                if (selectedM !== 'all' && rowM !== selectedM) {
                    showRow = false;
                }
                
                // Update row display and cell values based on StageFound, State, and Category filters
                if (showRow) {
                    const cells = row.cells;
                    let critical, high, medium, low, total;
                    
                    // If Category filter is selected, use category-specific counts (highest priority)
                    if (selectedCategory !== 'all') {
                        const categoryKey = selectedCategory.toLowerCase().replace(/ /g, '_');
                        critical = parseInt(row.getAttribute(`data-${categoryKey}-critical`)) || 0;
                        high = parseInt(row.getAttribute(`data-${categoryKey}-high`)) || 0;
                        medium = parseInt(row.getAttribute(`data-${categoryKey}-medium`)) || 0;
                        low = parseInt(row.getAttribute(`data-${categoryKey}-low`)) || 0;
                        total = parseInt(row.getAttribute(`data-${categoryKey}-total`)) || 0;
                        
                        cells[4].textContent = critical;
                        cells[5].textContent = high;
                        cells[6].textContent = medium;
                        cells[7].textContent = low;
                        cells[8].innerHTML = '<strong>' + total + '</strong>';
                    } else if (selectedState !== 'all') {
                        // If State filter is selected, use state-specific counts
                        const stateKey = selectedState.toLowerCase().replace(/ /g, '_').replace(/-/g, '_');
                        critical = parseInt(row.getAttribute(`data-${stateKey}-critical`)) || 0;
                        high = parseInt(row.getAttribute(`data-${stateKey}-high`)) || 0;
                        medium = parseInt(row.getAttribute(`data-${stateKey}-medium`)) || 0;
                        low = parseInt(row.getAttribute(`data-${stateKey}-low`)) || 0;
                        total = parseInt(row.getAttribute(`data-${stateKey}-total`)) || 0;
                        
                        cells[4].textContent = critical;
                        cells[5].textContent = high;
                        cells[6].textContent = medium;
                        cells[7].textContent = low;
                        cells[8].innerHTML = '<strong>' + total + '</strong>';
                    } else if (selectedStage === 'PT') {
                        // Show PT counts
                        critical = parseInt(row.getAttribute('data-pt-critical')) || 0;
                        high = parseInt(row.getAttribute('data-pt-high')) || 0;
                        medium = parseInt(row.getAttribute('data-pt-medium')) || 0;
                        low = parseInt(row.getAttribute('data-pt-low')) || 0;
                        total = parseInt(row.getAttribute('data-pt-total')) || 0;
                        
                        cells[4].textContent = critical;
                        cells[5].textContent = high;
                        cells[6].textContent = medium;
                        cells[7].textContent = low;
                        cells[8].innerHTML = '<strong>' + total + '</strong>';
                    } else if (selectedStage === 'UAT') {
                        // Show UAT counts
                        critical = parseInt(row.getAttribute('data-uat-critical')) || 0;
                        high = parseInt(row.getAttribute('data-uat-high')) || 0;
                        medium = parseInt(row.getAttribute('data-uat-medium')) || 0;
                        low = parseInt(row.getAttribute('data-uat-low')) || 0;
                        total = parseInt(row.getAttribute('data-uat-total')) || 0;
                        
                        cells[4].textContent = critical;
                        cells[5].textContent = high;
                        cells[6].textContent = medium;
                        cells[7].textContent = low;
                        cells[8].innerHTML = '<strong>' + total + '</strong>';
                    } else {
                        // Show all counts (original values)
                        const ptCritical = parseInt(row.getAttribute('data-pt-critical')) || 0;
                        const ptHigh = parseInt(row.getAttribute('data-pt-high')) || 0;
                        const ptMedium = parseInt(row.getAttribute('data-pt-medium')) || 0;
                        const ptLow = parseInt(row.getAttribute('data-pt-low')) || 0;
                        const ptTotal = parseInt(row.getAttribute('data-pt-total')) || 0;
                        const uatCritical = parseInt(row.getAttribute('data-uat-critical')) || 0;
                        const uatHigh = parseInt(row.getAttribute('data-uat-high')) || 0;
                        const uatMedium = parseInt(row.getAttribute('data-uat-medium')) || 0;
                        const uatLow = parseInt(row.getAttribute('data-uat-low')) || 0;
                        const uatTotal = parseInt(row.getAttribute('data-uat-total')) || 0;
                        
                        critical = ptCritical + uatCritical;
                        high = ptHigh + uatHigh;
                        medium = ptMedium + uatMedium;
                        low = ptLow + uatLow;
                        total = ptTotal + uatTotal;
                        
                        cells[4].textContent = critical;
                        cells[5].textContent = high;
                        cells[6].textContent = medium;
                        cells[7].textContent = low;
                        cells[8].innerHTML = '<strong>' + total + '</strong>';
                    }
                    
                    // Add to running totals
                    totalCritical += critical;
                    totalHigh += high;
                    totalMedium += medium;
                    totalLow += low;
                    totalAll += total;
                }
                
                row.style.display = showRow ? '' : 'none';
            });
            
            // Update the total row with calculated totals
            const totalCriticalElement = document.getElementById('defect-total-critical');
            const totalHighElement = document.getElementById('defect-total-high');
            const totalMediumElement = document.getElementById('defect-total-medium');
            const totalLowElement = document.getElementById('defect-total-low');
            const totalAllElement = document.getElementById('defect-total-all');
            
            if (totalCriticalElement) totalCriticalElement.textContent = totalCritical;
            if (totalHighElement) totalHighElement.textContent = totalHigh;
            if (totalMediumElement) totalMediumElement.textContent = totalMedium;
            if (totalLowElement) totalLowElement.textContent = totalLow;
            if (totalAllElement) totalAllElement.innerHTML = '<strong>' + totalAll + '</strong>';
            
            // Also filter the Defect Details table
            const detailsTable = document.getElementById('defectDetailsTable');
            if (detailsTable) {
                const detailsRows = detailsTable.querySelectorAll('tbody tr');
                detailsRows.forEach(row => {
                    const rowAD = row.getAttribute('data-ad-poc');
                    const rowSM = row.getAttribute('data-sm-poc');
                    const rowM = row.getAttribute('data-m-poc');
                    const rowStage = row.getAttribute('data-stage');
                    const rowState = row.getAttribute('data-state');
                    const rowCategory = row.getAttribute('data-category');
                    
                    let showDetailRow = true;
                    
                    // Apply POC filters
                    if (selectedAD !== 'all' && rowAD !== selectedAD) showDetailRow = false;
                    if (selectedSM !== 'all' && rowSM !== selectedSM) showDetailRow = false;
                    if (selectedM !== 'all' && rowM !== selectedM) showDetailRow = false;
                    
                    // Apply Category filter
                    if (selectedCategory !== 'all' && rowCategory !== selectedCategory) showDetailRow = false;
                    
                    // Apply State filter
                    if (selectedState !== 'all' && rowState !== selectedState) showDetailRow = false;
                    
                    // Apply Stage filter
                    if (selectedStage === 'PT' && rowStage === 'User Acceptance Test') showDetailRow = false;
                    if (selectedStage === 'UAT' && rowStage !== 'User Acceptance Test') showDetailRow = false;
                    
                    row.style.display = showDetailRow ? '' : 'none';
                });
            }
        }
        
        // Function to filter defect summary tables by POC
        function filterDefectSummary() {
            const adFilter = document.getElementById('defectSummaryAdPocFilter');
            const smFilter = document.getElementById('defectSummarySmPocFilter');
            const mFilter = document.getElementById('defectSummaryMPocFilter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;
            const selectedM = mFilter.value;
            
            // Update SM POC dropdown based on selected AD POC
            if (!selectedAD) {
                // If "All" selected, disable SM and M filters
                smFilter.disabled = true;
                mFilter.disabled = true;
                smFilter.innerHTML = '<option value="">All</option>';
                mFilter.innerHTML = '<option value="">All</option>';
            } else {
                // Enable SM filter and populate based on selected AD POC
                smFilter.disabled = false;
                smFilter.innerHTML = '<option value="">All</option>';
                
                const smPocs = Object.keys(pocHierarchyDefect[selectedAD] || {});
                smPocs.forEach(sm => {
                    const option = document.createElement('option');
                    option.value = sm;
                    option.textContent = sm;
                    if (sm === selectedSM) {
                        option.selected = true;
                    }
                    smFilter.appendChild(option);
                });
                
                // If SM POC was previously selected and still valid, keep it
                if (selectedSM && !smPocs.includes(selectedSM)) {
                    smFilter.value = '';
                }
            }
            
            // Update M POC dropdown based on selected SM POC
            if (!selectedSM || !selectedAD) {
                mFilter.disabled = true;
                mFilter.innerHTML = '<option value="">All</option>';
            } else {
                mFilter.disabled = false;
                mFilter.innerHTML = '<option value="">All</option>';
                
                const mPocs = pocHierarchyDefect[selectedAD]?.[selectedSM] || [];
                mPocs.forEach(m => {
                    const option = document.createElement('option');
                    option.value = m;
                    option.textContent = m;
                    if (m === selectedM) {
                        option.selected = true;
                    }
                    mFilter.appendChild(option);
                });
                
                // If M POC was previously selected and still valid, keep it
                if (selectedM && !mPocs.includes(selectedM)) {
                    mFilter.value = '';
                }
            }
            
            // Get current filter values after updates
            const adPocFilter = adFilter.value;
            const smPocFilter = smFilter.value;
            const mPocFilter = mFilter.value;
            
            // Initialize aggregated data structure
            let data = {
                total: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                active: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                rtd: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                pt_total: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                pt_active: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                pt_rtd: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                uat_total: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                uat_active: {total: 0, critical: 0, high: 0, medium: 0, low: 0},
                uat_rtd: {total: 0, critical: 0, high: 0, medium: 0, low: 0}
            };
            
            // Determine which filter to use (M POC takes precedence, then SM, then AD)
            let filterKey = '';
            let filterValue = '';
            
            if (mPocFilter) {
                filterKey = 'M:' + mPocFilter;
                filterValue = mPocFilter;
            } else if (smPocFilter) {
                filterKey = 'SM:' + smPocFilter;
                filterValue = smPocFilter;
            } else if (adPocFilter) {
                filterKey = 'AD:' + adPocFilter;
                filterValue = adPocFilter;
            }
            
            // Aggregate data based on filter
            if (filterKey && defectSummaryData[filterKey]) {
                // Use the specific POC data
                data = defectSummaryData[filterKey];
            } else {
                // Aggregate all matching POCs
                Object.keys(defectSummaryData).forEach(key => {
                    // If no filter, include all
                    // If AD filter, include all AD: keys that match
                    // If SM filter, include all SM: keys that match
                    // If M filter, include all M: keys that match
                    let includeKey = false;
                    
                    if (!filterValue) {
                        // No filter - include all but avoid duplication
                        // Only count M: entries to avoid triple-counting
                        includeKey = key.startsWith('M:');
                    } else if (adPocFilter && !smPocFilter && !mPocFilter) {
                        // AD filter only - count all M: entries for this AD
                        // Since we can't directly map, we'll use AD: key
                        includeKey = key === 'AD:' + adPocFilter;
                    } else if (smPocFilter && !mPocFilter) {
                        // SM filter only
                        includeKey = key === 'SM:' + smPocFilter;
                    } else if (mPocFilter) {
                        // M filter
                        includeKey = key === 'M:' + mPocFilter;
                    }
                    
                    if (includeKey) {
                        const pocData = defectSummaryData[key];
                        Object.keys(data).forEach(category => {
                            Object.keys(data[category]).forEach(severity => {
                                data[category][severity] += pocData[category][severity];
                            });
                        });
                    }
                });
            }
            
            // Update Total Defect breakdown table
            const totalTable = document.getElementById('totalDefectTable');
            if (totalTable) {
                const rows = totalTable.querySelectorAll('tbody tr');
                if (rows[0]) {
                    rows[0].cells[1].textContent = data.total.total;
                    rows[0].cells[2].textContent = data.total.critical;
                    rows[0].cells[3].textContent = data.total.high;
                    rows[0].cells[4].textContent = data.total.medium;
                    rows[0].cells[5].textContent = data.total.low;
                }
                if (rows[1]) {
                    rows[1].cells[1].textContent = data.active.total;
                    rows[1].cells[2].textContent = data.active.critical;
                    rows[1].cells[3].textContent = data.active.high;
                    rows[1].cells[4].textContent = data.active.medium;
                    rows[1].cells[5].textContent = data.active.low;
                }
                if (rows[2]) {
                    rows[2].cells[1].textContent = data.rtd.total;
                    rows[2].cells[2].textContent = data.rtd.critical;
                    rows[2].cells[3].textContent = data.rtd.high;
                    rows[2].cells[4].textContent = data.rtd.medium;
                    rows[2].cells[5].textContent = data.rtd.low;
                }
            }
            
            // Update PT Defect Breakdown table
            const ptTable = document.getElementById('ptDefectTable');
            if (ptTable) {
                const rows = ptTable.querySelectorAll('tbody tr');
                if (rows[0]) {
                    rows[0].cells[1].textContent = data.pt_total.total;
                    rows[0].cells[2].textContent = data.pt_total.critical;
                    rows[0].cells[3].textContent = data.pt_total.high;
                    rows[0].cells[4].textContent = data.pt_total.medium;
                    rows[0].cells[5].textContent = data.pt_total.low;
                }
                if (rows[1]) {
                    rows[1].cells[1].textContent = data.pt_active.total;
                    rows[1].cells[2].textContent = data.pt_active.critical;
                    rows[1].cells[3].textContent = data.pt_active.high;
                    rows[1].cells[4].textContent = data.pt_active.medium;
                    rows[1].cells[5].textContent = data.pt_active.low;
                }
                if (rows[2]) {
                    rows[2].cells[1].textContent = data.pt_rtd.total;
                    rows[2].cells[2].textContent = data.pt_rtd.critical;
                    rows[2].cells[3].textContent = data.pt_rtd.high;
                    rows[2].cells[4].textContent = data.pt_rtd.medium;
                    rows[2].cells[5].textContent = data.pt_rtd.low;
                }
            }
            
            // Update UAT Defect Breakdown table
            const uatTable = document.getElementById('uatDefectTable');
            if (uatTable) {
                const rows = uatTable.querySelectorAll('tbody tr');
                if (rows[0]) {
                    rows[0].cells[1].textContent = data.uat_total.total;
                    rows[0].cells[2].textContent = data.uat_total.critical;
                    rows[0].cells[3].textContent = data.uat_total.high;
                    rows[0].cells[4].textContent = data.uat_total.medium;
                    rows[0].cells[5].textContent = data.uat_total.low;
                }
                if (rows[1]) {
                    rows[1].cells[1].textContent = data.uat_active.total;
                    rows[1].cells[2].textContent = data.uat_active.critical;
                    rows[1].cells[3].textContent = data.uat_active.high;
                    rows[1].cells[4].textContent = data.uat_active.medium;
                    rows[1].cells[5].textContent = data.uat_active.low;
                }
                if (rows[2]) {
                    rows[2].cells[1].textContent = data.uat_rtd.total;
                    rows[2].cells[2].textContent = data.uat_rtd.critical;
                    rows[2].cells[3].textContent = data.uat_rtd.high;
                    rows[2].cells[4].textContent = data.uat_rtd.medium;
                    rows[2].cells[5].textContent = data.uat_rtd.low;
                }
            }
        }
        
        // Function to filter overall defects
        function filterOverallDefects() {
            const adFilter = document.getElementById('overallDefectAdPocFilter');
            const smFilter = document.getElementById('overallDefectSmPocFilter');
            const mFilter = document.getElementById('overallDefectMPocFilter');
            
            const selectedAD = adFilter.value;
            const selectedSM = smFilter.value;
            const selectedM = mFilter.value;
            
            // Update SM POC dropdown based on selected AD POC
            if (!selectedAD) {
                // If "All" selected, disable SM and M filters
                smFilter.disabled = true;
                mFilter.disabled = true;
                smFilter.innerHTML = '<option value="">All</option>';
                mFilter.innerHTML = '<option value="">All</option>';
            } else {
                // Enable SM filter and populate based on selected AD POC
                smFilter.disabled = false;
                smFilter.innerHTML = '<option value="">All</option>';
                
                const smPocs = Object.keys(pocHierarchyDefect[selectedAD] || {});
                smPocs.forEach(sm => {
                    const option = document.createElement('option');
                    option.value = sm;
                    option.textContent = sm;
                    if (sm === selectedSM) {
                        option.selected = true;
                    }
                    smFilter.appendChild(option);
                });
                
                // If SM POC was previously selected and still valid, keep it
                if (selectedSM && !smPocs.includes(selectedSM)) {
                    smFilter.value = '';
                }
            }
            
            // Update M POC dropdown based on selected SM POC
            if (!selectedSM || !selectedAD) {
                mFilter.disabled = true;
                mFilter.innerHTML = '<option value="">All</option>';
            } else {
                mFilter.disabled = false;
                mFilter.innerHTML = '<option value="">All</option>';
                
                const mPocs = pocHierarchyDefect[selectedAD]?.[selectedSM] || [];
                mPocs.forEach(m => {
                    const option = document.createElement('option');
                    option.value = m;
                    option.textContent = m;
                    if (m === selectedM) {
                        option.selected = true;
                    }
                    mFilter.appendChild(option);
                });
                
                // If M POC was previously selected and still valid, keep it
                if (selectedM && !mPocs.includes(selectedM)) {
                    mFilter.value = '';
                }
            }
            
            // Get current filter values after updates
            const adPocFilter = adFilter.value;
            const smPocFilter = smFilter.value;
            const mPocFilter = mFilter.value;
            
            // Filter defects based on selected POCs
            const filteredDefects = overallDefectsData.filter(defect => {
                if (adPocFilter && defect.adPoc !== adPocFilter) return false;
                if (smPocFilter && defect.smPoc !== smPocFilter) return false;
                if (mPocFilter && defect.mPoc !== mPocFilter) return false;
                return true;
            });
            
            // Helper function to count by severity
            function countBySeverity(defects) {
                return {
                    total: defects.length,
                    critical: defects.filter(d => d.severity === '1 - Critical').length,
                    high: defects.filter(d => d.severity === '2 - High').length,
                    medium: defects.filter(d => d.severity === '3 - Medium').length,
                    low: defects.filter(d => d.severity === '4 - Low').length
                };
            }
            
            // ===== OVERALL DEFECT COUNTS =====
            const overallCounts = countBySeverity(filteredDefects);
            const overallActiveDefects = filteredDefects.filter(d => activeStates.includes(d.state));
            const overallActiveCounts = countBySeverity(overallActiveDefects);
            const overallFixedDefects = filteredDefects.filter(d => fixedReadyStates.includes(d.state));
            const overallFixedCounts = countBySeverity(overallFixedDefects);
            const overallTestingDefects = filteredDefects.filter(d => underTestingStates.includes(d.state));
            const overallTestingCounts = countBySeverity(overallTestingDefects);
            const overallClosedDefects = filteredDefects.filter(d => closedStates.includes(d.state));
            const overallClosedCounts = countBySeverity(overallClosedDefects);
            
            // Update Overall section - Total bugs raised
            document.getElementById('overallTotalBugs').textContent = overallCounts.total;
            document.getElementById('overallCritical').textContent = overallCounts.critical;
            document.getElementById('overallHigh').textContent = overallCounts.high;
            document.getElementById('overallMedium').textContent = overallCounts.medium;
            document.getElementById('overallLow').textContent = overallCounts.low;
            
            // Update Overall section - Active
            document.getElementById('overallActive').textContent = overallActiveCounts.total;
            document.getElementById('overallActiveCritical').textContent = overallActiveCounts.critical;
            document.getElementById('overallActiveHigh').textContent = overallActiveCounts.high;
            document.getElementById('overallActiveMedium').textContent = overallActiveCounts.medium;
            document.getElementById('overallActiveLow').textContent = overallActiveCounts.low;
            
            // Update Overall section - Fixed and Ready to Deploy
            document.getElementById('overallFixedReady').textContent = overallFixedCounts.total;
            document.getElementById('overallFixedCritical').textContent = overallFixedCounts.critical;
            document.getElementById('overallFixedHigh').textContent = overallFixedCounts.high;
            document.getElementById('overallFixedMedium').textContent = overallFixedCounts.medium;
            document.getElementById('overallFixedLow').textContent = overallFixedCounts.low;
            
            // Update Overall section - Under Testing
            document.getElementById('overallUnderTesting').textContent = overallTestingCounts.total;
            document.getElementById('overallTestingCritical').textContent = overallTestingCounts.critical;
            document.getElementById('overallTestingHigh').textContent = overallTestingCounts.high;
            document.getElementById('overallTestingMedium').textContent = overallTestingCounts.medium;
            document.getElementById('overallTestingLow').textContent = overallTestingCounts.low;
            
            // Update Overall section - Closed
            document.getElementById('overallClosed').textContent = overallClosedCounts.total;
            document.getElementById('overallClosedCritical').textContent = overallClosedCounts.critical;
            document.getElementById('overallClosedHigh').textContent = overallClosedCounts.high;
            document.getElementById('overallClosedMedium').textContent = overallClosedCounts.medium;
            document.getElementById('overallClosedLow').textContent = overallClosedCounts.low;
            
            // ===== PT DEFECT COUNTS =====
            const ptDefects = filteredDefects.filter(d => d.stageFound !== 'User Acceptance Test');
            const ptCounts = countBySeverity(ptDefects);
            const ptActiveDefects = ptDefects.filter(d => activeStates.includes(d.state));
            const ptActiveCounts = countBySeverity(ptActiveDefects);
            const ptFixedDefects = ptDefects.filter(d => fixedReadyStates.includes(d.state));
            const ptFixedCounts = countBySeverity(ptFixedDefects);
            const ptTestingDefects = ptDefects.filter(d => underTestingStates.includes(d.state));
            const ptTestingCounts = countBySeverity(ptTestingDefects);
            const ptClosedDefects = ptDefects.filter(d => closedStates.includes(d.state));
            const ptClosedCounts = countBySeverity(ptClosedDefects);
            
            // Update PT section - Total bugs raised
            document.getElementById('ptTotalBugs').textContent = ptCounts.total;
            document.getElementById('ptCritical').textContent = ptCounts.critical;
            document.getElementById('ptHigh').textContent = ptCounts.high;
            document.getElementById('ptMedium').textContent = ptCounts.medium;
            document.getElementById('ptLow').textContent = ptCounts.low;
            
            // Update PT section - Active
            document.getElementById('ptActive').textContent = ptActiveCounts.total;
            document.getElementById('ptActiveCritical').textContent = ptActiveCounts.critical;
            document.getElementById('ptActiveHigh').textContent = ptActiveCounts.high;
            document.getElementById('ptActiveMedium').textContent = ptActiveCounts.medium;
            document.getElementById('ptActiveLow').textContent = ptActiveCounts.low;
            
            // Update PT section - Fixed and Ready to Deploy
            document.getElementById('ptFixedReady').textContent = ptFixedCounts.total;
            document.getElementById('ptFixedCritical').textContent = ptFixedCounts.critical;
            document.getElementById('ptFixedHigh').textContent = ptFixedCounts.high;
            document.getElementById('ptFixedMedium').textContent = ptFixedCounts.medium;
            document.getElementById('ptFixedLow').textContent = ptFixedCounts.low;
            
            // Update PT section - Under Testing
            document.getElementById('ptUnderTesting').textContent = ptTestingCounts.total;
            document.getElementById('ptTestingCritical').textContent = ptTestingCounts.critical;
            document.getElementById('ptTestingHigh').textContent = ptTestingCounts.high;
            document.getElementById('ptTestingMedium').textContent = ptTestingCounts.medium;
            document.getElementById('ptTestingLow').textContent = ptTestingCounts.low;
            
            // Update PT section - Closed
            document.getElementById('ptClosed').textContent = ptClosedCounts.total;
            document.getElementById('ptClosedCritical').textContent = ptClosedCounts.critical;
            document.getElementById('ptClosedHigh').textContent = ptClosedCounts.high;
            document.getElementById('ptClosedMedium').textContent = ptClosedCounts.medium;
            document.getElementById('ptClosedLow').textContent = ptClosedCounts.low;
            
            // ===== UAT DEFECT COUNTS =====
            const uatDefects = filteredDefects.filter(d => d.stageFound === 'User Acceptance Test');
            const uatCounts = countBySeverity(uatDefects);
            const uatActiveDefects = uatDefects.filter(d => activeStates.includes(d.state));
            const uatActiveCounts = countBySeverity(uatActiveDefects);
            const uatFixedDefects = uatDefects.filter(d => fixedReadyStates.includes(d.state));
            const uatFixedCounts = countBySeverity(uatFixedDefects);
            const uatTestingDefects = uatDefects.filter(d => underTestingStates.includes(d.state));
            const uatTestingCounts = countBySeverity(uatTestingDefects);
            const uatClosedDefects = uatDefects.filter(d => closedStates.includes(d.state));
            const uatClosedCounts = countBySeverity(uatClosedDefects);
            
            // Update UAT section - Total bugs raised
            document.getElementById('uatTotalBugs').textContent = uatCounts.total;
            document.getElementById('uatCritical').textContent = uatCounts.critical;
            document.getElementById('uatHigh').textContent = uatCounts.high;
            document.getElementById('uatMedium').textContent = uatCounts.medium;
            document.getElementById('uatLow').textContent = uatCounts.low;
            
            // Update UAT section - Active
            document.getElementById('uatActive').textContent = uatActiveCounts.total;
            document.getElementById('uatActiveCritical').textContent = uatActiveCounts.critical;
            document.getElementById('uatActiveHigh').textContent = uatActiveCounts.high;
            document.getElementById('uatActiveMedium').textContent = uatActiveCounts.medium;
            document.getElementById('uatActiveLow').textContent = uatActiveCounts.low;
            
            // Update UAT section - Fixed and Ready to Deploy
            document.getElementById('uatFixedReady').textContent = uatFixedCounts.total;
            document.getElementById('uatFixedCritical').textContent = uatFixedCounts.critical;
            document.getElementById('uatFixedHigh').textContent = uatFixedCounts.high;
            document.getElementById('uatFixedMedium').textContent = uatFixedCounts.medium;
            document.getElementById('uatFixedLow').textContent = uatFixedCounts.low;
            
            // Update UAT section - Under Testing
            document.getElementById('uatUnderTesting').textContent = uatTestingCounts.total;
            document.getElementById('uatTestingCritical').textContent = uatTestingCounts.critical;
            document.getElementById('uatTestingHigh').textContent = uatTestingCounts.high;
            document.getElementById('uatTestingMedium').textContent = uatTestingCounts.medium;
            document.getElementById('uatTestingLow').textContent = uatTestingCounts.low;
            
            // Update UAT section - Closed
            document.getElementById('uatClosed').textContent = uatClosedCounts.total;
            document.getElementById('uatClosedCritical').textContent = uatClosedCounts.critical;
            document.getElementById('uatClosedHigh').textContent = uatClosedCounts.high;
            document.getElementById('uatClosedMedium').textContent = uatClosedCounts.medium;
            document.getElementById('uatClosedLow').textContent = uatClosedCounts.low;
        }
        
        function resetDefectSummaryFilters() {
            // Reset all Defect Summary filter dropdowns
            document.getElementById('defectSummaryAdPocFilter').value = '';
            document.getElementById('defectSummarySmPocFilter').value = '';
            document.getElementById('defectSummarySmPocFilter').disabled = true;
            document.getElementById('defectSummaryMPocFilter').value = '';
            document.getElementById('defectSummaryMPocFilter').disabled = true;
            
            // Trigger the filter update to refresh the display
            filterDefectSummary();
        }
        
        function resetOverallDefectFilters() {
            // Reset all Overall Defect filter dropdowns
            document.getElementById('overallDefectAdPocFilter').value = '';
            document.getElementById('overallDefectSmPocFilter').value = '';
            document.getElementById('overallDefectSmPocFilter').disabled = true;
            document.getElementById('overallDefectMPocFilter').value = '';
            document.getElementById('overallDefectMPocFilter').disabled = true;
            
            // Trigger the filter update to refresh the display
            filterOverallDefects();
        }
        
        // Initialize filters on page load
        window.addEventListener('DOMContentLoaded', function() {
            updateExecFilters('ad');
            updateModuleFilters('ad');
            filterOverallDefects(); // Initialize overall defects display
        });
    </script>
</body>
</html>
"""

# Save HTML dashboard with minification
dashboard_path = os.path.join(base_dir, "Daily_Status_Dashboard.html")

# Minify HTML to reduce file size
import re

# Step 1: Minify CSS - remove comments and excess whitespace
def minify_css(css_text):
    # Remove CSS comments
    css_text = re.sub(r'/\*.*?\*/', '', css_text, flags=re.DOTALL)
    # Remove whitespace around CSS punctuation
    css_text = re.sub(r'\s*([{}:;,])\s*', r'\1', css_text)
    # Remove multiple spaces
    css_text = re.sub(r'\s+', ' ', css_text)
    # Remove spaces around braces
    css_text = re.sub(r'\s*{\s*', '{', css_text)
    css_text = re.sub(r'\s*}\s*', '}', css_text)
    return css_text.strip()

minified_content = html_content

# Minify CSS within <style> tags
style_pattern = r'<style>(.*?)</style>'
matches = re.findall(style_pattern, minified_content, re.DOTALL)
for css_block in matches:
    minified_css = minify_css(css_block)
    minified_content = minified_content.replace(f'<style>{css_block}</style>', f'<style>{minified_css}</style>', 1)

# Step 2: Minify HTML - remove excess whitespace safely (preserve inline attributes)
# Remove whitespace between closing and opening tags only (not within tags)
minified_content = re.sub(r'>\s+<', '><', minified_content)
# Remove leading/trailing whitespace on lines but preserve single spaces
minified_content = re.sub(r'\n\s+', '\n', minified_content)
# Remove empty lines
minified_content = re.sub(r'\n\n+', '\n', minified_content)

with open(dashboard_path, 'w', encoding='utf-8') as f:
    f.write(minified_content)

print(f"  Dashboard created: {dashboard_path}")
print()

print("=" * 80)
print("DAILY STATUS REPORT GENERATION COMPLETE!")
print("=" * 80)
print()
print(f"Output Files:")
print(f"  1. {story_summary_path}")
print(f"  2. {bug_summary_path}")
print(f"  3. {dashboard_path}")
print()
print("Open the HTML dashboard in your browser to view the report.")
