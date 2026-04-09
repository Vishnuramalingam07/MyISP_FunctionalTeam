import pandas as pd
import os

import requests

# Azure DevOps query info
org = "accenturecio08"
project = "AutomationProcess_29697"
query_id = "88002d3a-c9ca-46f9-aaa3-95a8f54dbb53"
api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"


# Get PAT from environment variable (e.g., AZURE_DEVOPS_PAT)
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
	raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}


# Fetch work item IDs from query with error handling
response = requests.get(api_url, auth=("", pat), headers=headers)
response.raise_for_status()
try:
	work_items = response.json()["workItems"]
except Exception as e:
	print("Failed to parse JSON response from Azure DevOps API.")
	print(f"Status code: {response.status_code}")
	print(f"Response text:\n{response.text}")
	raise
ids = [str(item["id"]) for item in work_items]

if not ids:
	print("No work items found for this query.")
	exit()

# Fetch work item details in batches
def fetch_details(ids):
	url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
	body = {
		"ids": ids,
		"fields": ["System.Id", "System.Title", "Node Name", "SM Name"]
	}
	r = requests.post(url, json=body, auth=("", pat), headers=headers)
	r.raise_for_status()
	return r.json()["value"]



# Fetch all work item details using correct field name
def fetch_details(ids):
	url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
	body = {
		"ids": ids,
		"fields": [
			"System.Id",
			"System.Title",
			"System.State",
			"Custom.Application",
			"Microsoft.VSTS.Common.Severity",
			"System.AreaPath"
		]
	}
	r = requests.post(url, json=body, auth=("", pat), headers=headers)
	r.raise_for_status()
	return r.json()["value"]

batch_size = 200
all_details = []
for i in range(0, len(ids), batch_size):
	batch_ids = ids[i:i+batch_size]
	all_details.extend(fetch_details(batch_ids))

# Prepare bug list DataFrame
bug_data = []
for item in all_details:
	fields = item["fields"]
	bug_data.append(fields)
bug_df = pd.DataFrame(bug_data)

# Read mapping file
mapping_path = r'C:\Users\d.sampathkumar\GHC files\POC mapping\POD Mapping sheet_Updated.csv'
mapping_df = pd.read_csv(mapping_path)

print("Mapping sheet columns:", mapping_df.columns.tolist())
print("Bug list columns:", bug_df.columns.tolist())

# Build mapping dictionaries (case-insensitive)
mapping_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['SM Name'])}
ad_mapping_dict = {str(k).lower(): v for k, v in zip(mapping_df['Node Name'], mapping_df['AD POC'])}

# Derive Node Name from System.AreaPath
if 'System.AreaPath' in bug_df.columns:
    bug_df['Node Name'] = bug_df['System.AreaPath'].apply(lambda x: str(x).split('\\')[-1] if pd.notnull(x) else '')
    bug_df['SM POC'] = bug_df['Node Name'].apply(lambda x: mapping_dict.get(str(x).lower(), None))
    bug_df['AD POC'] = bug_df['Node Name'].apply(lambda x: ad_mapping_dict.get(str(x).lower(), None))
    # Insert columns after 'System.AreaPath'
    area_idx = bug_df.columns.get_loc('System.AreaPath')
    cols = bug_df.columns.tolist()
    cols.insert(area_idx + 1, cols.pop(cols.index('Node Name')))
    cols.insert(area_idx + 2, cols.pop(cols.index('SM POC')))
    cols.insert(area_idx + 3, cols.pop(cols.index('AD POC')))
    bug_df = bug_df[cols]
else:
    print("Column 'System.AreaPath' not found in bug list. Please check the column names above and update the script with the correct name.")
    exit(1)

# Save updated file
output_path = os.path.join(os.path.dirname(__file__), "Final Bug list with POC added.csv")
bug_df.to_csv(output_path, index=False)
print(f"Updated file saved as: {output_path}")