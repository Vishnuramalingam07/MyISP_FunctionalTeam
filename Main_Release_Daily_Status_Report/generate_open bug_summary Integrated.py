import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore", category=Warning, module="requests")
import requests
from datetime import datetime
from config import OUTPUT_DIR, POD_MAPPING_PATH, OPEN_BUG_SUMMARY_FILE

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"
query_id = "82f9cead-6354-49bc-ba99-b5aaf885c525"
api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

print("Fetching work items from ADO query...")

# Fetch work item IDs from query
response = requests.get(api_url, auth=("", pat), headers=headers)
response.raise_for_status()
work_items = response.json()["workItems"]
ids = [str(item["id"]) for item in work_items]

if not ids:
    print("No work items found for this query.")
    exit()

print(f"Found {len(ids)} work items")

# Fetch work item details with specified fields
def fetch_details(ids):
    url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
    payload = {
        "ids": ids,
        "fields": [
            "System.Id",
            "System.Title",
            "System.AreaPath",
            "System.State",
            "Microsoft.VSTS.Common.Priority",
            "Microsoft.VSTS.Common.Severity",
            "Custom.DefectRecord",
            "Custom.TextVerification",
            "Custom.StageFound"
        ]
    }
    response = requests.post(url, auth=("", pat), headers=headers, json=payload)
    response.raise_for_status()
    return response.json()["value"]

# Fetch all work item details in batches
batch_size = 200
all_details = []
for i in range(0, len(ids), batch_size):
    batch_ids = ids[i:i + batch_size]
    all_details.extend(fetch_details(batch_ids))
    print(f"  Processed {min(i + batch_size, len(ids))}/{len(ids)} work items")

print("Processing work items...")

# Read mapping file
mapping_path = POD_MAPPING_PATH
print(f"Reading mapping file: {mapping_path}")
mapping_df = pd.read_csv(mapping_path)

# Create mapping dictionaries (case-insensitive)
node_to_m_poc = {str(k).strip().lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['M POC'])}
node_to_sm_poc = {str(k).strip().lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['SM Name'])}
node_to_ad_poc = {str(k).strip().lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['AD POC'])}

# Create output data with detailed bug information
data = []
for item in all_details:
    fields = item["fields"]
    
    # Extract basic fields
    bug_id = fields.get("System.Id", "")
    title = fields.get("System.Title", "")
    state = fields.get("System.State", "")
    
    # Extract area path and node name
    area_path = fields.get("System.AreaPath", "")
    node_name = area_path.split('\\')[-1] if area_path else ""
    
    # Get severity
    severity = fields.get("Microsoft.VSTS.Common.Severity", "")
    
    # Get custom fields
    defect_record = fields.get("Custom.DefectRecord", "")
    text_verification = fields.get("Custom.TextVerification", "")
    stage_found = fields.get("Custom.StageFound", "")
    
    # Get POC mappings
    node_name_lower = node_name.strip().lower()
    m_poc = node_to_m_poc.get(node_name_lower, "")
    sm_poc = node_to_sm_poc.get(node_name_lower, "")
    ad_poc = node_to_ad_poc.get(node_name_lower, "")
    
    data.append({
        "ID": bug_id,
        "Title": title,
        "Severity": severity,
        "Node Name": node_name,
        "AD POC": ad_poc,
        "SM POC": sm_poc,
        "M POC": m_poc,
        "Defect Record": defect_record,
        "TextVerification": text_verification,
        "State": state,
        "StageFound": stage_found
    })

# Create DataFrame
df = pd.DataFrame(data)

# Define output directory
output_dir = OUTPUT_DIR

# Save to Excel
output_path = OPEN_BUG_SUMMARY_FILE
df.to_excel(output_path, index=False, sheet_name='Bug Summary')

print(f"\nData exported successfully!")
print(f"Output file: {output_path}")
print(f"\nSummary:")
print(f"Total Bugs: {len(all_details)}")
print(f"Total Unique Nodes: {df['Node Name'].nunique()}")

# Count by severity
severity_counts = df['Severity'].value_counts()
for severity, count in severity_counts.items():
    print(f"{severity}: {count}")
