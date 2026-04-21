import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore", category=Warning, module="requests")
import requests
from datetime import datetime
import time
from config import OUTPUT_DIR, POD_MAPPING_PATH, OVERALL_DEFECT_SUMMARY_FILE

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"
query_id = "948ba984-b405-457b-ad15-4ca29387a9fb"
api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

print("Fetching defects from ADO query...")

# Fetch work item IDs from query with timeout
try:
    response = requests.get(api_url, auth=("", pat), headers=headers, timeout=30)
    response.raise_for_status()
    work_items = response.json()["workItems"]
    ids = [str(item["id"]) for item in work_items]
except requests.exceptions.Timeout:
    print("Error: Request timeout while fetching defects. Check your network connection.")
    exit(1)
except requests.exceptions.RequestException as e:
    print(f"Error fetching defects: {e}")
    exit(1)

if not ids:
    print("No defects found for this query.")
    exit()

print(f"Found {len(ids)} defects")

# Fetch work item details with specified fields and retry logic
def fetch_details(ids, max_retries=3):
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
    
    for attempt in range(max_retries):
        try:
            response = requests.post(url, auth=("", pat), headers=headers, json=payload, timeout=60)
            if response.status_code in [429]:  # Rate limit
                wait_time = 2 ** attempt
                print(f"    Rate limited. Waiting {wait_time} seconds before retry...")
                time.sleep(wait_time)
                continue
            if response.status_code != 200:
                print(f"    Error response: {response.status_code}")
                response.raise_for_status()
            return response.json()["value"]
        except requests.exceptions.Timeout:
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"    Request timeout. Waiting {wait_time} seconds before retry (attempt {attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
            else:
                print(f"    Error: Request timeout after {max_retries} attempts")
                raise
        except requests.exceptions.RequestException as e:
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"    Connection error: {e}. Waiting {wait_time} seconds before retry...")
                time.sleep(wait_time)
            else:
                print(f"    Error fetching details: {e}")
                raise
    
    return []

# Fetch all work item details in batches
batch_size = 200
all_details = []
for i in range(0, len(ids), batch_size):
    batch_ids = [int(id) for id in ids[i:i + batch_size]]
    try:
        all_details.extend(fetch_details(batch_ids))
        print(f"  Processed {min(i + batch_size, len(ids))}/{len(ids)} defects")
    except Exception as e:
        print(f"  Warning: Failed to fetch batch starting at {i}: {e}")
        print(f"  Continuing with next batch...")
        continue

print("Processing defects...")

# Read mapping file
mapping_path = POD_MAPPING_PATH
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
output_dir = OUTPUT_DIR

# Save to Excel
output_path = OVERALL_DEFECT_SUMMARY_FILE
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
