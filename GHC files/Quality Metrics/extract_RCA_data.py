import pandas as pd
import os
import requests

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"

# Define releases with their query IDs
releases = {
    "Nov 8": "4d468701-65dd-431c-916a-c0c69a14788d",
    "Dec 13": "dbe5fecb-3b47-40f4-96d9-d4e6947e750a",
    "Jan 10": "7e2da104-c0c5-4558-9082-abf2c349f015"
}

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

# Output folder
output_folder = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics'

# Read mapping file and add POC columns
mapping_path = r'C:\Users\d.sampathkumar\GHC files\POC mapping\POD Mapping sheet_Updated.csv'
if os.path.exists(mapping_path):
    mapping_df = pd.read_csv(mapping_path)
    # Create mapping dictionaries (case-insensitive)
    m_poc_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['M POC'])}
    sm_poc_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['SM Name'])}
    ad_poc_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['AD POC'])}
    mapping_available = True
else:
    print(f"Warning: Mapping file not found at {mapping_path}. POC columns will be empty.")
    mapping_available = False

# Define DEV Countable RCA values
dev_countable_rca_values = [
    'Code Issue',
    'UX Design / UI Issue',
    'RCA Not filled - Dev Missed',
    'Integration Issue',
    'Design Issue',
    'New Feature / Enhancement',
    'Performance Issue',
    'Lack of environment knowledge / Inadequate version control on environment/ Lack of availability of the proper environment',
    'Incorrect Environment Setup / Inadequate Documentation',
    'Documentation Error/ Inadequate Documentation',
    'Environment Issue',
    'Existing Prod Issue',
    'Requirement Clarity Issue'
]

# Function to fetch and process RCA data for a release
def fetch_release_rca_data(release_name, query_id):
    print(f"\n{'='*70}")
    print(f"Processing Release: {release_name}")
    print(f"{'='*70}")
    
    api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"
    
    # Fetch work item IDs from query
    print("Fetching work items from ADO query...")
    response = requests.get(api_url, auth=("", pat), headers=headers)
    response.raise_for_status()
    work_items = response.json()["workItems"]
    ids = [str(item["id"]) for item in work_items]

    if not ids:
        print(f"No work items found for {release_name}.")
        return None, None

    print(f"Found {len(ids)} work items to process.")

    # Fetch work item details with required fields
    def fetch_details(ids):
        url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
        payload = {
            "ids": ids,
            "fields": [
                "System.Id",
                "System.Title",
                "System.AreaPath",
                "Custom.mySPRCA"  # mySP RCA field
            ]
        }
        response = requests.post(url, auth=("", pat), headers=headers, json=payload)
        response.raise_for_status()
        return response.json()["value"]

    # Fetch all work item details in batches
    batch_size = 200
    all_details = []
    print("Fetching work item details...")
    for i in range(0, len(ids), batch_size):
        batch_ids = ids[i:i + batch_size]
        all_details.extend(fetch_details(batch_ids))
        print(f"  Processed {min(i + batch_size, len(ids))}/{len(ids)} work items")

    # Prepare data list
    data = []
    for item in all_details:
        fields = item["fields"]
        area_path = fields.get("System.AreaPath", "")
        
        # Extract Node Name (last part of Area Path)
        node_name = area_path.split('\\')[-1] if area_path else ''
        
        data.append({
            "ID": fields.get("System.Id"),
            "Title": fields.get("System.Title"),
            "Node Name": node_name,
            "mySP RCA": fields.get("Custom.mySPRCA", "")
        })

    df = pd.DataFrame(data)

    # Add POC columns using Node Name lookup
    if mapping_available:
        df.insert(3, 'M POC', df['Node Name'].apply(lambda x: m_poc_dict.get(str(x).lower(), '')))
        df.insert(4, 'SM POC', df['Node Name'].apply(lambda x: sm_poc_dict.get(str(x).lower(), '')))
        df.insert(5, 'AD POC', df['Node Name'].apply(lambda x: ad_poc_dict.get(str(x).lower(), '')))
    else:
        df.insert(3, 'M POC', '')
        df.insert(4, 'SM POC', '')
        df.insert(5, 'AD POC', '')

    # Add DEV Countable RCA and DEV NOT countable RCA columns
    def classify_rca(rca_value):
        # Check if empty/blank or in the countable list
        if pd.isna(rca_value) or str(rca_value).strip() == '':
            return 'DEV Countable RCA', ''
        elif rca_value in dev_countable_rca_values:
            return 'DEV Countable RCA', ''
        else:
            return '', 'DEV NOT countable RCA'

    df[['DEV Countable RCA', 'DEV NOT countable RCA']] = df['mySP RCA'].apply(
        lambda x: pd.Series(classify_rca(x))
    )

    # Create RCA Summary by Node Name
    # Group by Node Name and mySP RCA to get counts
    summary_data = []
    for node_name in df['Node Name'].unique():
        node_df = df[df['Node Name'] == node_name]
        
        # Get POC values for this node
        m_poc = node_df['M POC'].iloc[0] if len(node_df) > 0 else ''
        sm_poc = node_df['SM POC'].iloc[0] if len(node_df) > 0 else ''
        ad_poc = node_df['AD POC'].iloc[0] if len(node_df) > 0 else ''
        
        # Get count for each unique RCA value
        rca_counts = node_df['mySP RCA'].value_counts()
        
        for rca_value, count in rca_counts.items():
            # Handle NaN/empty values
            rca_display = rca_value if pd.notna(rca_value) and str(rca_value).strip() != '' else 'Dev Not filled mySP RCA'
            
            # Determine RCA Type based on the classification logic
            if pd.isna(rca_value) or str(rca_value).strip() == '':
                rca_type = 'DEV Countable RCA'
            elif rca_value in dev_countable_rca_values:
                rca_type = 'DEV Countable RCA'
            else:
                rca_type = 'DEV NOT countable RCA'
            
            summary_data.append({
                'Node Name': node_name,
                'AD POC': ad_poc,
                'SM POC': sm_poc,
                'M POC': m_poc,
                'RCA Type': rca_type,
                'mySP RCA': rca_display,
                'Count': int(count)
            })

    summary_df = pd.DataFrame(summary_data)

    # Sort by Node Name and Count (descending)
    summary_df = summary_df.sort_values(['Node Name', 'Count'], ascending=[True, False])
    
    print(f"\n✓ {release_name} - Processed {len(df)} work items")
    print(f"  - Unique Node Names: {df['Node Name'].nunique()}")
    print(f"  - Work items with mySP RCA: {df['mySP RCA'].notna().sum()}")
    print(f"  - Unique RCA values: {summary_df['mySP RCA'].nunique()}")
    
    return df, summary_df

# Process all releases
print("="*70)
print("RCA DATA EXTRACTION FOR MULTIPLE RELEASES")
print("="*70)

all_details_dfs = {}
all_summary_dfs = {}

for release_name, query_id in releases.items():
    details_df, summary_df = fetch_release_rca_data(release_name, query_id)
    if details_df is not None:
        all_details_dfs[release_name] = details_df
        all_summary_dfs[release_name] = summary_df

# Save RCA Details to separate Excel file with multiple sheets
if all_details_dfs:
    details_output_file = os.path.join(output_folder, "RCA_Details_All_Releases.xlsx")
    with pd.ExcelWriter(details_output_file, engine='openpyxl') as writer:
        for release_name, df in all_details_dfs.items():
            df.to_excel(writer, sheet_name=release_name, index=False)
    
    print(f"\n{'='*70}")
    print(f"✓ RCA Details saved to: {details_output_file}")
    print(f"  Sheets: {', '.join(all_details_dfs.keys())}")

# Save RCA Summary to separate Excel file with multiple sheets
if all_summary_dfs:
    summary_output_file = os.path.join(output_folder, "RCA_Summary_Final.xlsx")
    with pd.ExcelWriter(summary_output_file, engine='openpyxl') as writer:
        for release_name, df in all_summary_dfs.items():
            df.to_excel(writer, sheet_name=release_name, index=False)
    
    print(f"✓ RCA Summary saved to: {summary_output_file}")
    print(f"  Sheets: {', '.join(all_summary_dfs.keys())}")

print(f"\n{'='*70}")
print("✅ ALL RELEASES PROCESSED SUCCESSFULLY!")
print(f"{'='*70}")
