import pandas as pd
import os
import requests
from datetime import datetime

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"
query_id = "432f1a3b-09ba-43a1-8f86-a93480ccf557"
api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

print("=" * 80)
print("PRODUCT OWNER DETAILS GENERATOR")
print("=" * 80)
print()

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
            "System.Parent",
            "Custom.RequirementRequestor"
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

print("\nProcessing work items...")

# Debug: Print sample work item to check field names
if all_details:
    print("\n" + "=" * 80)
    print("DEBUG: Sample work item fields:")
    print("=" * 80)
    sample_item = all_details[0]
    print(f"Work Item ID: {sample_item.get('id', 'N/A')}")
    print("\nAvailable fields:")
    for field_name, field_value in sample_item.get('fields', {}).items():
        print(f"  {field_name}: {field_value}")
    print("=" * 80)
    print()

# Function to extract enterprise ID from email format
def extract_enterprise_id(requirement_requestor):
    """
    Extract enterprise ID from format like 'Garg, Puneet <puneet.d.garg@accenture.com>'
    Returns 'puneet.d.garg'
    """
    if not requirement_requestor:
        return ""
    
    req_str = str(requirement_requestor).strip()
    
    # Check if it's in the format with < and @
    if '<' in req_str and '@' in req_str:
        # Extract text between < and @
        start = req_str.index('<') + 1
        end = req_str.index('@')
        return req_str[start:end].strip()
    
    return ""

# Group data by Parent
parent_data = {}
parent_ids_set = set()

for item in all_details:
    fields = item["fields"]
    item_id = fields.get("System.Id", "")
    
    # Extract Parent ID
    parent = fields.get("System.Parent")
    parent_id = ""
    if isinstance(parent, dict):
        parent_id = str(parent.get("id", ""))
    elif parent:
        parent_id = str(parent)
    
    # Skip if no parent
    if not parent_id:
        continue
    
    parent_ids_set.add(parent_id)
    
    # Get Requirement Requestor - try different possible field names
    req_requestor = fields.get("Custom.RequirementRequestor", "")
    if not req_requestor:
        req_requestor = fields.get("Custom.Requirement Requestor", "")
    if not req_requestor:
        req_requestor = fields.get("RequirementRequestor", "")
    
    # Handle dict format (contains uniqueName with email)
    if isinstance(req_requestor, dict):
        # Get uniqueName which contains the email like "puneet.d.garg@accenture.com"
        unique_name = req_requestor.get("uniqueName", "")
        if unique_name and '@' in unique_name:
            # Extract enterprise ID (part before @)
            enterprise_id = unique_name.split('@')[0].strip()
        else:
            enterprise_id = ""
    elif req_requestor:
        # If it's a string, try to extract enterprise ID
        enterprise_id = extract_enterprise_id(req_requestor)
    else:
        enterprise_id = ""
    
    # Initialize parent entry if not exists
    if parent_id not in parent_data:
        parent_data[parent_id] = {
            'story_ids': [],
            'requirement_requestors': set()
        }
    
    # Add story ID and enterprise ID to the parent's data
    parent_data[parent_id]['story_ids'].append(item_id)
    if enterprise_id:  # Only add non-empty enterprise IDs
        parent_data[parent_id]['requirement_requestors'].add(enterprise_id)

# Fetch parent titles from ADO
print("\nFetching parent titles from ADO...")
parent_titles = {}
if parent_ids_set:
    parent_ids_list = list(parent_ids_set)
    # Fetch parent details in batches
    for i in range(0, len(parent_ids_list), batch_size):
        batch_ids = [int(pid) for pid in parent_ids_list[i:i + batch_size]]
        url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
        payload = {
            "ids": batch_ids,
            "fields": ["System.Id", "System.Title"]
        }
        response = requests.post(url, auth=("", pat), headers=headers, json=payload)
        response.raise_for_status()
        parent_details = response.json()["value"]
        for parent in parent_details:
            parent_id = str(parent["id"])
            parent_title = parent["fields"].get("System.Title", "")
            parent_titles[parent_id] = parent_title
        print(f"  Fetched {min(i + batch_size, len(parent_ids_list))}/{len(parent_ids_list)} parent titles")

# Prepare final data
final_data = []
for parent_id, parent_info in parent_data.items():
    # Get parent title
    parent_title = parent_titles.get(parent_id, "")
    
    # Combine requirement requestors with " and "
    requestors = sorted(list(parent_info['requirement_requestors']))
    product_owner = " and ".join(requestors) if requestors else ""
    
    final_data.append({
        "Parent ID": parent_id,
        "Parent Title": parent_title,
        "Product Owner": product_owner
    })

# Create DataFrame
df = pd.DataFrame(final_data)

# Sort by Parent ID
df = df.sort_values('Parent ID').reset_index(drop=True)

print(f"\nProcessed {len(df)} unique parents")
print(f"Parents with Product Owner: {len(df[df['Product Owner'] != ''])}")
print(f"Parents without Product Owner: {len(df[df['Product Owner'] == ''])}")

# Save to Excel
output_path = r"C:\Users\vishnu.ramalingam\MyISP_Tools\GHC files\Daily status report\PO Details.xlsx"
df.to_excel(output_path, index=False, sheet_name='Product Owner Details')

print(f"\n✓ Excel file generated: {output_path}")
print("=" * 80)
