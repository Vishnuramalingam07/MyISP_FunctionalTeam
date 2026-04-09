import pandas as pd
import os
import requests
from datetime import datetime

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"
query_id = "d853f2f3-992e-423e-b9b8-860ab841108c"
api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

print("Fetching defects from ADO query...")

# Fetch work item IDs from query
response = requests.get(api_url, auth=("", pat), headers=headers)
response.raise_for_status()
work_items = response.json()["workItems"]
ids = [str(item["id"]) for item in work_items]

if not ids:
    print("No defects found for this query.")
    exit()

print(f"Found {len(ids)} defects")

# Fetch work item details with specified fields
def fetch_details(ids):
    url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
    payload = {
        "ids": ids,
        "fields": [
            "System.Id",
            "System.Title",
            "System.State",
            "System.AreaPath",
            "System.WorkItemType",
            "System.CreatedDate",
            "Microsoft.VSTS.Common.ClosedDate",
            "Custom.DefectRecord",
            "Custom.TextVerification",
            "Custom.StageFound",
            "Microsoft.VSTS.Common.Severity",
            "Custom.Category",
            "Custom.mySPRCA",
            "Custom.23bfcf97-0a58-4d60-9787-e54ef96208a0"  # Re open count field
        ]
    }
    response = requests.post(url, auth=("", pat), headers=headers, json=payload)
    if response.status_code != 200:
        print(f"Error response: {response.text}")
    response.raise_for_status()
    return response.json()["value"]

# Fetch all work item details in batches
batch_size = 200
all_details = []
for i in range(0, len(ids), batch_size):
    batch_ids = [int(id) for id in ids[i:i + batch_size]]
    all_details.extend(fetch_details(batch_ids))
    print(f"  Processed {min(i + batch_size, len(ids))}/{len(ids)} defects")

print("Processing defects...")

# Read mapping file
mapping_path = r"C:\Users\vishnu.ramalingam\MyISP_Tools\GHC files\POC mapping\POD Mapping sheet_Updated.csv"
print(f"Reading mapping file: {mapping_path}")
mapping_df = pd.read_csv(mapping_path)

# Create mapping dictionaries (case-insensitive)
node_to_m_poc = {str(k).strip().lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['M POC'])}
node_to_sm_poc = {str(k).strip().lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['SM Name'])}
node_to_ad_poc = {str(k).strip().lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['AD POC'])}

# Prepare data
data = []
for item in all_details:
    fields = item["fields"]
    area_path = fields.get("System.AreaPath", "")
    node_name = area_path.split('\\')[-1] if area_path else ""
    
    # Get POC mapping based on Node Name
    node_name_lower = node_name.strip().lower()
    m_poc = node_to_m_poc.get(node_name_lower, "")
    sm_poc = node_to_sm_poc.get(node_name_lower, "")
    ad_poc = node_to_ad_poc.get(node_name_lower, "")
    
    # Get Created Date and Closed Date
    created_date = fields.get("System.CreatedDate", "")
    if created_date:
        try:
            created_date = pd.to_datetime(created_date).strftime("%Y-%m-%d")
        except:
            created_date = str(created_date)
    
    closed_date = fields.get("Microsoft.VSTS.Common.ClosedDate", "")
    if closed_date:
        try:
            closed_date = pd.to_datetime(closed_date).strftime("%Y-%m-%d")
        except:
            closed_date = str(closed_date)
    
    data.append({
        "ID": fields.get("System.Id", ""),
        "Title": fields.get("System.Title", ""),
        "Defect Record": fields.get("Custom.DefectRecord", ""),
        "TextVerification": fields.get("Custom.TextVerification", ""),
        "Node Name": node_name,
        "StageFound": fields.get("Custom.StageFound", ""),
        "Severity": fields.get("Microsoft.VSTS.Common.Severity", ""),
        "State": fields.get("System.State", ""),
        "Category": fields.get("Custom.Category", ""),
        "mySP RCA": fields.get("Custom.mySPRCA", ""),
        "Re open Count": fields.get("Custom.23bfcf97-0a58-4d60-9787-e54ef96208a0", ""),
        "Created Date": created_date,
        "Closed Date": closed_date,
        "Work Item Type": fields.get("System.WorkItemType", ""),
        "AD POC": ad_poc,
        "SM POC": sm_poc,
        "M POC": m_poc
    })

df = pd.DataFrame(data)

# Define output directory
output_dir = r"C:\Users\vishnu.ramalingam\MyISP_Tools\GHC files\Daily status report"

# Save to Excel
output_path = os.path.join(output_dir, "Over all Defect Summary.xlsx")
df.to_excel(output_path, index=False)

print(f"\nData exported successfully!")
print(f"Output file: {output_path}")
print(f"\nSummary:")
print(f"Total Defects: {len(df)}")
if 'Severity' in df.columns:
    print("\nBy Severity:")
    for severity, count in df['Severity'].value_counts().items():
        print(f"  {severity}: {count}")
if 'State' in df.columns:
    print("\nBy State:")
    for state, count in df['State'].value_counts().items():
        print(f"  {state}: {count}")
