import pandas as pd
import os
import requests

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"

# Define queries for each release
queries = {
    "Nov 8": "4d468701-65dd-431c-916a-c0c69a14788d",
    "Dec 13": "dbe5fecb-3b47-40f4-96d9-d4e6947e750a",
    "Jan 10": "7e2da104-c0c5-4558-9082-abf2c349f015"
}

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

# Fetch work item IDs from query
def fetch_work_items(query_id):
    api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"
    response = requests.get(api_url, auth=("", pat), headers=headers)
    response.raise_for_status()
    work_items = response.json()["workItems"]
    ids = [str(item["id"]) for item in work_items]
    return ids

# Fetch work item details with StageFound field
def fetch_details(ids):
    url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
    payload = {
        "ids": ids,
        "fields": [
            "System.Id",
            "System.AreaPath",
            "System.State",
            "System.Title",
            "Microsoft.VSTS.Common.Severity",
            "Custom.Application",
            "Custom.StageFound",  # Correct field name for StageFound
            "Custom.mySPRCA"  # Add mySP RCA field
        ]
    }
    response = requests.post(url, auth=("", pat), headers=headers, json=payload)
    response.raise_for_status()
    return response.json()["value"]

# Process data for a release
def process_release_data(release_name, query_id):
    print(f"\nProcessing {release_name}...")
    
    # Fetch work item IDs
    ids = fetch_work_items(query_id)
    
    if not ids:
        print(f"No work items found for {release_name}.")
        return None, None
    
    print(f"Found {len(ids)} work items for {release_name}")
    
    # Fetch all work item details in batches
    batch_size = 200
    all_details = []
    for i in range(0, len(ids), batch_size):
        batch_ids = ids[i:i + batch_size]
        all_details.extend(fetch_details(batch_ids))
    
    # Prepare bug list DataFrame with StageFound
    bug_data = []
    for item in all_details:
        fields = item["fields"]
        bug_data.append({
            "System.Id": fields.get("System.Id"),
            "System.AreaPath": fields.get("System.AreaPath"),
            "System.State": fields.get("System.State"),
            "Microsoft.VSTS.Common.Severity": fields.get("Microsoft.VSTS.Common.Severity"),
            "Custom.Application": fields.get("Custom.Application"),
            "System.Title": fields.get("System.Title"),
            "StageFound": fields.get("Custom.StageFound"),  # Fetch StageFound
            "mySP RCA": fields.get("Custom.mySPRCA")  # Fetch mySP RCA
        })
    df = pd.DataFrame(bug_data)
    
    # Add the "Team Raised" column based on "StageFound"
    df["Team Raised"] = df["StageFound"].apply(
        lambda x: "UAT" if str(x).strip().lower() == "user acceptance test" else "PT"
    )
    
    # Add the "Defect Category" column based on "mySP RCA"
    invalid_rca_values = [
        "Not reproduceable / Invalid data",
        "Insufficient application knowledge / Invalid data",
        "Duplicate Bug",
        "Inadequacy of available test data / Incorrect Data",
        "Access Rights Issue"
    ]
    
    df["Defect Category"] = df["mySP RCA"].apply(
        lambda x: "InValid" if pd.notnull(x) and str(x).strip() in invalid_rca_values else "Valid"
    )
    
    # Extract only the last part of Area Path and normalize to title case for consistency
    df["Area Name"] = df["System.AreaPath"].apply(lambda x: str(x).split('\\')[-1].strip().title() if pd.notnull(x) else '')
    
    return df, all_details

# Calculate bug summary
def calculate_bug_summary(group):
    total_bugs = len(group)
    
    # Total by severity
    total_critical = len(group[group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    total_high = len(group[group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    total_medium = len(group[group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    total_low = len(group[group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # PT bugs
    pt_group = group[group["Team Raised"] == "PT"]
    total_pt_bugs = len(pt_group)
    pt_critical = len(pt_group[pt_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    pt_high = len(pt_group[pt_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    pt_medium = len(pt_group[pt_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    pt_low = len(pt_group[pt_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # UAT bugs
    uat_group = group[group["Team Raised"] == "UAT"]
    total_uat_bugs = len(uat_group)
    uat_critical = len(uat_group[uat_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    uat_high = len(uat_group[uat_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    uat_medium = len(uat_group[uat_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    uat_low = len(uat_group[uat_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # Valid and Invalid bugs
    valid_bugs = len(group[group["Defect Category"] == "Valid"])
    invalid_bugs = len(group[group["Defect Category"] == "InValid"])
    
    # Valid bugs by severity
    valid_group = group[group["Defect Category"] == "Valid"]
    valid_critical = len(valid_group[valid_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    valid_high = len(valid_group[valid_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    valid_medium = len(valid_group[valid_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    valid_low = len(valid_group[valid_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # Invalid bugs by severity
    invalid_group = group[group["Defect Category"] == "InValid"]
    invalid_critical = len(invalid_group[invalid_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    invalid_high = len(invalid_group[invalid_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    invalid_medium = len(invalid_group[invalid_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    invalid_low = len(invalid_group[invalid_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # PT Valid bugs by severity
    pt_valid_group = group[(group["Team Raised"] == "PT") & (group["Defect Category"] == "Valid")]
    pt_valid_critical = len(pt_valid_group[pt_valid_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    pt_valid_high = len(pt_valid_group[pt_valid_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    pt_valid_medium = len(pt_valid_group[pt_valid_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    pt_valid_low = len(pt_valid_group[pt_valid_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # PT Invalid bugs by severity
    pt_invalid_group = group[(group["Team Raised"] == "PT") & (group["Defect Category"] == "InValid")]
    pt_invalid_critical = len(pt_invalid_group[pt_invalid_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    pt_invalid_high = len(pt_invalid_group[pt_invalid_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    pt_invalid_medium = len(pt_invalid_group[pt_invalid_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    pt_invalid_low = len(pt_invalid_group[pt_invalid_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # UAT Valid bugs by severity
    uat_valid_group = group[(group["Team Raised"] == "UAT") & (group["Defect Category"] == "Valid")]
    uat_valid_critical = len(uat_valid_group[uat_valid_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    uat_valid_high = len(uat_valid_group[uat_valid_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    uat_valid_medium = len(uat_valid_group[uat_valid_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    uat_valid_low = len(uat_valid_group[uat_valid_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    # UAT Invalid bugs by severity
    uat_invalid_group = group[(group["Team Raised"] == "UAT") & (group["Defect Category"] == "InValid")]
    uat_invalid_critical = len(uat_invalid_group[uat_invalid_group["Microsoft.VSTS.Common.Severity"] == "1 - Critical"])
    uat_invalid_high = len(uat_invalid_group[uat_invalid_group["Microsoft.VSTS.Common.Severity"] == "2 - High"])
    uat_invalid_medium = len(uat_invalid_group[uat_invalid_group["Microsoft.VSTS.Common.Severity"] == "3 - Medium"])
    uat_invalid_low = len(uat_invalid_group[uat_invalid_group["Microsoft.VSTS.Common.Severity"] == "4 - Low"])
    
    return pd.Series({
        'Total Bugs': total_bugs,
        'Total Critical': total_critical,
        'Total High': total_high,
        'Total Medium': total_medium,
        'Total Low': total_low,
        'Total PT Bugs': total_pt_bugs,
        'PT Critical': pt_critical,
        'PT High': pt_high,
        'PT Medium': pt_medium,
        'PT Low': pt_low,
        'Total UAT Bugs': total_uat_bugs,
        'UAT Critical': uat_critical,
        'UAT High': uat_high,
        'UAT Medium': uat_medium,
        'UAT Low': uat_low,
        'Valid bugs': valid_bugs,
        'Invalid bugs': invalid_bugs,
        'Valid Critical': valid_critical,
        'Valid High': valid_high,
        'Valid Medium': valid_medium,
        'Valid Low': valid_low,
        'Invalid Critical': invalid_critical,
        'Invalid High': invalid_high,
        'Invalid Medium': invalid_medium,
        'Invalid Low': invalid_low,
        'PT Valid Critical': pt_valid_critical,
        'PT Valid High': pt_valid_high,
        'PT Valid Medium': pt_valid_medium,
        'PT Valid Low': pt_valid_low,
        'PT Invalid Critical': pt_invalid_critical,
        'PT Invalid High': pt_invalid_high,
        'PT Invalid Medium': pt_invalid_medium,
        'PT Invalid Low': pt_invalid_low,
        'UAT Valid Critical': uat_valid_critical,
        'UAT Valid High': uat_valid_high,
        'UAT Valid Medium': uat_valid_medium,
        'UAT Valid Low': uat_valid_low,
        'UAT Invalid Critical': uat_invalid_critical,
        'UAT Invalid High': uat_invalid_high,
        'UAT Invalid Medium': uat_invalid_medium,
        'UAT Invalid Low': uat_invalid_low
    })

# Read mapping file
mapping_path = r'C:\Users\d.sampathkumar\GHC files\POC mapping\POD Mapping sheet_Updated.csv'
mapping_df = pd.read_csv(mapping_path)

# Create mapping dictionaries (case-insensitive)
m_poc_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['M POC'])}
sm_poc_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['SM Name'])}
ad_poc_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['AD POC'])}

# Process all releases
output_folder = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics'
output_excel = os.path.join(output_folder, "Bug_summary_Final.xlsx")

# Create Excel writer
with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    all_csv_data = {}
    
    for release_name, query_id in queries.items():
        df, all_details = process_release_data(release_name, query_id)
        
        if df is None:
            continue
        
        # Group by Area Name and calculate summary
        summary = df.groupby('Area Name').apply(calculate_bug_summary).reset_index()
        
        # Add POC columns
        summary.insert(1, 'M POC', summary['Area Name'].apply(lambda x: m_poc_dict.get(str(x).lower(), '')))
        summary.insert(2, 'SM POC', summary['Area Name'].apply(lambda x: sm_poc_dict.get(str(x).lower(), '')))
        summary.insert(3, 'AD POC', summary['Area Name'].apply(lambda x: ad_poc_dict.get(str(x).lower(), '')))
        
        # Sort by Total Bugs descending
        summary = summary.sort_values(by='Total Bugs', ascending=False)
        
        # Add Grand Total row
        grand_total = pd.DataFrame([{
            'Area Name': 'Grand Total',
            'M POC': '',
            'SM POC': '',
            'AD POC': '',
            'Total Bugs': summary['Total Bugs'].sum(),
            'Total Critical': summary['Total Critical'].sum(),
            'Total High': summary['Total High'].sum(),
            'Total Medium': summary['Total Medium'].sum(),
            'Total Low': summary['Total Low'].sum(),
            'Total PT Bugs': summary['Total PT Bugs'].sum(),
            'PT Critical': summary['PT Critical'].sum(),
            'PT High': summary['PT High'].sum(),
            'PT Medium': summary['PT Medium'].sum(),
            'PT Low': summary['PT Low'].sum(),
            'Total UAT Bugs': summary['Total UAT Bugs'].sum(),
            'UAT Critical': summary['UAT Critical'].sum(),
            'UAT High': summary['UAT High'].sum(),
            'UAT Medium': summary['UAT Medium'].sum(),
            'UAT Low': summary['UAT Low'].sum(),
            'Valid bugs': summary['Valid bugs'].sum(),
            'Invalid bugs': summary['Invalid bugs'].sum(),
            'Valid Critical': summary['Valid Critical'].sum(),
            'Valid High': summary['Valid High'].sum(),
            'Valid Medium': summary['Valid Medium'].sum(),
            'Valid Low': summary['Valid Low'].sum(),
            'Invalid Critical': summary['Invalid Critical'].sum(),
            'Invalid High': summary['Invalid High'].sum(),
            'Invalid Medium': summary['Invalid Medium'].sum(),
            'Invalid Low': summary['Invalid Low'].sum(),
            'PT Valid Critical': summary['PT Valid Critical'].sum(),
            'PT Valid High': summary['PT Valid High'].sum(),
            'PT Valid Medium': summary['PT Valid Medium'].sum(),
            'PT Valid Low': summary['PT Valid Low'].sum(),
            'PT Invalid Critical': summary['PT Invalid Critical'].sum(),
            'PT Invalid High': summary['PT Invalid High'].sum(),
            'PT Invalid Medium': summary['PT Invalid Medium'].sum(),
            'PT Invalid Low': summary['PT Invalid Low'].sum(),
            'UAT Valid Critical': summary['UAT Valid Critical'].sum(),
            'UAT Valid High': summary['UAT Valid High'].sum(),
            'UAT Valid Medium': summary['UAT Valid Medium'].sum(),
            'UAT Valid Low': summary['UAT Valid Low'].sum(),
            'UAT Invalid Critical': summary['UAT Invalid Critical'].sum(),
            'UAT Invalid High': summary['UAT Invalid High'].sum(),
            'UAT Invalid Medium': summary['UAT Invalid Medium'].sum(),
            'UAT Invalid Low': summary['UAT Invalid Low'].sum()
        }])
        
        summary = pd.concat([summary, grand_total], ignore_index=True)
        
        # Write to Excel sheet
        summary.to_excel(writer, sheet_name=release_name, index=False)
        print(f"  Added sheet '{release_name}' to Excel")
        
        # Save CSV extract for this release
        all_csv_data[release_name] = df

print(f"\nBug summary saved to: {output_excel}")

# Save separate CSV files for each release
for release_name, df in all_csv_data.items():
    output_csv = os.path.join(output_folder, f"Bug_TFS_Extract_{release_name.replace(' ', '_')}.csv")
    df.to_csv(output_csv, index=False)
    print(f"Bug extract for {release_name} saved to: {output_csv}")