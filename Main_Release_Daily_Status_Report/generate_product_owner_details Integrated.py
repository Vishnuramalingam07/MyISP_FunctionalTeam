import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore", category=Warning, module="requests")
import requests
from datetime import datetime
from config import PO_DETAILS_FILE

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"
query_id = "730dc08e-2b34-4d8f-a74e-6b7c74a05071"
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
            "Custom.RequirementRequestor",
            "System.AreaPath"
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
    
    # Extract Node Name from Area Path (last part after backslash)
    area_path = fields.get("System.AreaPath", "")
    node_name = area_path.split('\\')[-1] if area_path else ""
    
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
            'requirement_requestors': set(),
            'node_names': []
        }
    
    # Add story ID and enterprise ID to the parent's data
    parent_data[parent_id]['story_ids'].append(item_id)
    parent_data[parent_id]['node_names'].append(node_name)
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
empty_product_owner_parents = []

for parent_id, parent_info in parent_data.items():
    # Get parent title
    parent_title = parent_titles.get(parent_id, "")
    
    # Combine requirement requestors with " and "
    requestors = sorted(list(parent_info['requirement_requestors']))
    product_owner = " and ".join(requestors) if requestors else ""
    
    # Track parents with empty product owner for secondary check
    if not product_owner:
        empty_product_owner_parents.append({
            'parent_id': parent_id,
            'parent_title': parent_title,
            'node_names': parent_info['node_names']
        })
    
    final_data.append({
        "Parent ID": parent_id,
        "Parent Title": parent_title,
        "Product Owner": product_owner,
        "_node_names": parent_info['node_names']  # Temporary field for processing
    })

# Secondary check: For empty product owners, check if all stories have myISP_IMS as Node Name
print("\nPerforming secondary check for empty Product Owners...")
secondary_filled = 0

for parent_info in empty_product_owner_parents:
    parent_id = parent_info['parent_id']
    node_names = parent_info['node_names']
    
    # Check if all node names are "myISP_IMS"
    if node_names and all(node_name.strip().lower() == "myisp_ims" for node_name in node_names if node_name):
        # Fill Product Owner as "IMS BAU"
        for item in final_data:
            if item['Parent ID'] == parent_id:
                item['Product Owner'] = "IMS BAU"
                secondary_filled += 1
                print(f"  [OK] Parent {parent_id}: Filled 'IMS BAU' (all {len(node_names)} stories have myISP_IMS)")
                break

print(f"Secondary check complete: Filled {secondary_filled} Product Owners based on Node Name match")

# Remove temporary field
for item in final_data:
    del item['_node_names']

# Tertiary check: For still-empty product owners, check if Parent Title contains "Angular"
print("\nPerforming tertiary check for remaining empty Product Owners...")
tertiary_filled = 0

for item in final_data:
    if not item['Product Owner']:  # Only process if still empty
        parent_title = item['Parent Title']
        if "angular" in parent_title.lower():
            item['Product Owner'] = "Angular upgrade stories"
            tertiary_filled += 1
            print(f"  [OK] Parent {item['Parent ID']}: Filled 'Angular upgrade stories' (Title: {parent_title})")

print(f"Tertiary check complete: Filled {tertiary_filled} Product Owners based on Angular keyword match")

# Create DataFrame
df = pd.DataFrame(final_data)

# Sort by Parent ID
df = df.sort_values('Parent ID').reset_index(drop=True)

print(f"\nProcessed {len(df)} unique parents")
print(f"Parents with Product Owner: {len(df[df['Product Owner'] != ''])}")
print(f"Parents without Product Owner: {len(df[df['Product Owner'] == ''])}")

# Save to Excel
output_path = PO_DETAILS_FILE
df.to_excel(output_path, index=False, sheet_name='Product Owner Details')

print(f"\n[OK] Excel file generated: {output_path}")
print("=" * 80)
