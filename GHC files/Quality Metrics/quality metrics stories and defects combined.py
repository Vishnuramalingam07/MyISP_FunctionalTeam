import pandas as pd
import os

# File paths
base_path = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics'
bug_file = os.path.join(base_path, 'Bug_summary_Final.xlsx')
story_file = os.path.join(base_path, 'Story_summary_final.xlsx')
output_file = os.path.join(base_path, 'Quality Metrics Complete Input file for stories and bugs.xlsx')

print("="*80)
print("COMBINING STORY AND BUG DATA FOR QUALITY METRICS")
print("="*80)

# Read Excel files to get sheet names
bug_excel = pd.ExcelFile(bug_file)
story_excel = pd.ExcelFile(story_file)

print(f"\nBug file sheets: {bug_excel.sheet_names}")
print(f"Story file sheets: {story_excel.sheet_names}")

# Get common sheet names
common_sheets = list(set(bug_excel.sheet_names) & set(story_excel.sheet_names))
print(f"\nCommon sheets to process: {common_sheets}")

# Define the exact column order for output
output_columns = [
    'AD POC',
    'SM POC',
    'M POC',
    'Node Name',
    'Testing Not Applicable Stories',
    'Testable stories',
    'Total story Points',
    'Total Bugs',
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
    'Valid bugs',
    'Invalid bugs',
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

# Story file column mapping
story_columns_map = {
    'Area Name': 'Node Name',
    'M POC': 'M POC',
    'SM POC': 'SM POC',
    'AD POC': 'AD POC',
    'Testing Not Applicable Stories': 'Testing Not Applicable Stories',
    'Testable stories': 'Testable stories',
    'Story Points': 'Total story Points'
}

# Bug file column mapping
bug_columns_map = {
    'Node Name': 'Node Name',
    'M POC': 'M POC',
    'SM POC': 'SM POC',
    'AD POC': 'AD POC',
    'Total Bugs': 'Total Bugs',
    'Total Critical': 'Total Critical',
    'Total High': 'Total High',
    'Total Medium': 'Total Medium',
    'Total Low': 'Total Low',
    'Total PT Bugs': 'Total PT Bugs',
    'PT Critical': 'PT Critical',
    'PT High': 'PT High',
    'PT Medium': 'PT Medium',
    'PT Low': 'PT Low',
    'Total UAT Bugs': 'Total UAT Bugs',
    'UAT Critical': 'UAT Critical',
    'UAT High': 'UAT High',
    'UAT Medium': 'UAT Medium',
    'UAT Low': 'UAT Low',
    'Valid bugs': 'Valid bugs',
    'Invalid bugs': 'Invalid bugs',
    'Valid Critical': 'Valid Critical',
    'Valid High': 'Valid High',
    'Valid Medium': 'Valid Medium',
    'Valid Low': 'Valid Low',
    'Invalid Critical': 'Invalid Critical',
    'Invalid High': 'Invalid High',
    'Invalid Medium': 'Invalid Medium',
    'Invalid Low': 'Invalid Low',
    'PT Valid Critical': 'PT Valid Critical',
    'PT Valid High': 'PT Valid High',
    'PT Valid Medium': 'PT Valid Medium',
    'PT Valid Low': 'PT Valid Low',
    'PT Invalid Critical': 'PT Invalid Critical',
    'PT Invalid High': 'PT Invalid High',
    'PT Invalid Medium': 'PT Invalid Medium',
    'PT Invalid Low': 'PT Invalid Low',
    'UAT Valid Critical': 'UAT Valid Critical',
    'UAT Valid High': 'UAT Valid High',
    'UAT Valid Medium': 'UAT Valid Medium',
    'UAT Valid Low': 'UAT Valid Low',
    'UAT Invalid Critical': 'UAT Invalid Critical',
    'UAT Invalid High': 'UAT Invalid High',
    'UAT Invalid Medium': 'UAT Invalid Medium',
    'UAT Invalid Low': 'UAT Invalid Low'
}

# Process each common sheet
combined_data = {}

for sheet_name in common_sheets:
    print(f"\n{'='*80}")
    print(f"Processing Sheet: {sheet_name}")
    print(f"{'='*80}")
    
    # Read story data
    story_df = pd.read_excel(story_file, sheet_name=sheet_name)
    print(f"Story data - rows: {len(story_df)}")
    print(f"Story columns: {list(story_df.columns)}")
    
    # Read bug data
    bug_df = pd.read_excel(bug_file, sheet_name=sheet_name)
    print(f"Bug data - rows: {len(bug_df)}")
    print(f"Bug columns: {list(bug_df.columns)}")
    
    # Exclude Grand Total rows - handle different column names
    if 'Area Name' in story_df.columns:
        story_df = story_df[story_df['Area Name'] != 'Grand Total'].copy()
    
    # Bug file might have 'Node Name' or 'Area Name'
    node_col_in_bug = None
    if 'Node Name' in bug_df.columns:
        node_col_in_bug = 'Node Name'
        bug_df = bug_df[bug_df['Node Name'] != 'Grand Total'].copy()
    elif 'Area Name' in bug_df.columns:
        node_col_in_bug = 'Area Name'
        bug_df = bug_df[bug_df['Area Name'] != 'Grand Total'].copy()
        # Rename to Node Name for consistency
        bug_df = bug_df.rename(columns={'Area Name': 'Node Name'})
    
    print(f"After removing Grand Total - Story: {len(story_df)}, Bug: {len(bug_df)}")
    
    # Rename columns in story data
    story_df_renamed = story_df.rename(columns=story_columns_map)
    
    # Select only relevant columns from story data
    story_cols_to_keep = ['Node Name', 'M POC', 'SM POC', 'AD POC', 
                          'Testing Not Applicable Stories', 'Testable stories', 
                          'Total story Points']
    story_df_selected = story_df_renamed[story_cols_to_keep].copy()
    
    # Select only relevant columns from bug data
    bug_cols_to_keep = list(bug_columns_map.values())
    # Keep only columns that exist in bug_df
    bug_cols_available = [col for col in bug_cols_to_keep if col in bug_df.columns]
    bug_df_selected = bug_df[bug_cols_available].copy()
    
    # Merge story and bug data on Node Name
    merged_df = pd.merge(
        story_df_selected,
        bug_df_selected,
        on='Node Name',
        how='outer',
        suffixes=('_story', '_bug')
    )
    
    # Handle POC columns - prefer bug data if both exist, otherwise use story data
    for poc_col in ['M POC', 'SM POC', 'AD POC']:
        if f'{poc_col}_bug' in merged_df.columns:
            merged_df[poc_col] = merged_df[f'{poc_col}_bug'].fillna(merged_df.get(f'{poc_col}_story', ''))
            merged_df = merged_df.drop(columns=[f'{poc_col}_bug', f'{poc_col}_story'], errors='ignore')
        elif f'{poc_col}_story' in merged_df.columns:
            merged_df = merged_df.rename(columns={f'{poc_col}_story': poc_col})
    
    # Fill NaN values with 0 for numeric columns
    numeric_columns = [col for col in output_columns if col not in ['AD POC', 'SM POC', 'M POC', 'Node Name']]
    for col in numeric_columns:
        if col not in merged_df.columns:
            merged_df[col] = 0
        else:
            merged_df[col] = merged_df[col].fillna(0)
    
    # Fill NaN values with empty string for text columns
    text_columns = ['AD POC', 'SM POC', 'M POC', 'Node Name']
    for col in text_columns:
        if col not in merged_df.columns:
            merged_df[col] = ''
        else:
            merged_df[col] = merged_df[col].fillna('')
    
    # Ensure all output columns exist
    for col in output_columns:
        if col not in merged_df.columns:
            if col in ['AD POC', 'SM POC', 'M POC', 'Node Name']:
                merged_df[col] = ''
            else:
                merged_df[col] = 0
    
    # Select and reorder columns
    merged_df = merged_df[output_columns]
    
    # Sort by Node Name
    merged_df = merged_df.sort_values('Node Name').reset_index(drop=True)
    
    combined_data[sheet_name] = merged_df
    
    print(f"✓ Combined data - rows: {len(merged_df)}")
    print(f"  Stories nodes: {len(story_df)}, Bug nodes: {len(bug_df)}, Combined nodes: {len(merged_df)}")

# Save to Excel with multiple sheets
print(f"\n{'='*80}")
print("Saving combined data to Excel...")
print(f"{'='*80}")

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df in combined_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"✓ Sheet '{sheet_name}' written - {len(df)} rows, {len(df.columns)} columns")

print(f"\n{'='*80}")
print(f"✅ SUCCESS! Combined quality metrics saved to:")
print(f"   {output_file}")
print(f"\n📊 Summary:")
print(f"   Total sheets: {len(combined_data)}")
print(f"   Sheet names: {', '.join(combined_data.keys())}")
print(f"   Columns per sheet: {len(output_columns)}")
print(f"{'='*80}")
