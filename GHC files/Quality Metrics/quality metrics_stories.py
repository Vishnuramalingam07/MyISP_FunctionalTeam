import pandas as pd
import os
import requests

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"

# Define releases with their query IDs
releases = {
    "Nov 8": "cb331587-df88-42b3-8c90-8a9ec2a65b05",
    "Dec 13": "91f2ce95-f4d5-42d3-9157-cf594ef01241",
    "Jan 10": "ef7cd3f5-1854-4472-bc86-8dae882b7a75"
}

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")
headers = {"Content-Type": "application/json"}

# Output folder
output_folder = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics'

# Read mapping file
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

# Function to fetch and process stories for a release
def fetch_release_data(release_name, query_id):
    print(f"\n{'='*60}")
    print(f"Processing Release: {release_name}")
    print(f"{'='*60}")
    
    story_api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"
    
    # Fetch story work item IDs from query
    print("Fetching story work items from ADO query...")
    response = requests.get(story_api_url, auth=("", pat), headers=headers)
    response.raise_for_status()
    story_work_items = response.json()["workItems"]
    story_ids = [str(item["id"]) for item in story_work_items]
    
    if not story_ids:
        print(f"No story work items found for {release_name}.")
        return None, None
    
    print(f"Found {len(story_ids)} stories to process.")
    
    # Fetch story work item details with Tags and Story Points fields
    def fetch_story_details(ids):
        url = f"https://dev.azure.com/{org}/_apis/wit/workitemsbatch?api-version=7.0"
        payload = {
            "ids": ids,
            "fields": [
                "System.Id",
                "System.AreaPath",
                "System.State",
                "System.Title",
                "System.Tags",
                "System.WorkItemType",
                "Microsoft.VSTS.Scheduling.StoryPoints"
            ]
        }
        response = requests.post(url, auth=("", pat), headers=headers, json=payload)
        response.raise_for_status()
        return response.json()["value"]
    
    # Fetch all story work item details in batches
    batch_size = 200
    all_story_details = []
    print("Fetching story details...")
    for i in range(0, len(story_ids), batch_size):
        batch_ids = story_ids[i:i + batch_size]
        all_story_details.extend(fetch_story_details(batch_ids))
        print(f"  Processed {min(i + batch_size, len(story_ids))}/{len(story_ids)} stories")
    
    # Prepare story data with Testing NA, Testable stories, and Story Points columns
    story_data = []
    for item in all_story_details:
        fields = item["fields"]
        tags = str(fields.get("System.Tags", "")).lower()
        story_points = fields.get("Microsoft.VSTS.Scheduling.StoryPoints", 0)
        
        # Convert story points to float, default to 0 if None
        try:
            story_points = float(story_points) if story_points is not None else 0
        except (ValueError, TypeError):
            story_points = 0
        
        # Check if "Testing NA" is in tags, but exclude "Testing NA - Only For UAT"
        is_testing_na = False
        if "testing na" in tags:
            if "testing na - only for uat" not in tags:
                is_testing_na = True
        
        story_data.append({
            "System.Id": fields.get("System.Id"),
            "System.AreaPath": fields.get("System.AreaPath"),
            "System.State": fields.get("System.State"),
            "System.Title": fields.get("System.Title"),
            "System.WorkItemType": fields.get("System.WorkItemType"),
            "System.Tags": fields.get("System.Tags"),
            "Story Points": story_points,
            "Testing Not Applicable Stories": 1 if is_testing_na else 0,
            "Testable stories": 0 if is_testing_na else 1
        })
    
    story_df = pd.DataFrame(story_data)
    
    # Extract only the last part of Area Path for stories and normalize to title case
    story_df["Area Name"] = story_df["System.AreaPath"].apply(lambda x: str(x).split('\\')[-1].strip().title() if pd.notnull(x) else '')
    
    # Add POC columns to story data
    if mapping_available:
        story_df.insert(2, 'M POC', story_df['Area Name'].apply(lambda x: m_poc_dict.get(str(x).lower(), '')))
        story_df.insert(3, 'SM POC', story_df['Area Name'].apply(lambda x: sm_poc_dict.get(str(x).lower(), '')))
        story_df.insert(4, 'AD POC', story_df['Area Name'].apply(lambda x: ad_poc_dict.get(str(x).lower(), '')))
    
    # Save detailed story DataFrame to CSV
    output_story_csv = os.path.join(output_folder, f"Story_TFS_Extract_{release_name.replace(' ', '_')}.csv")
    story_df.to_csv(output_story_csv, index=False)
    print(f"✓ Story list saved to: {output_story_csv}")
    
    # Create summary by Area Name
    story_summary = story_df.groupby('Area Name').agg({
        'Testing Not Applicable Stories': 'sum',
        'Testable stories': 'sum',
        'System.Id': 'count',
        'Story Points': 'sum'
    }).reset_index()
    story_summary.rename(columns={'System.Id': 'Total Stories'}, inplace=True)
    
    # Add POC columns to summary
    if mapping_available:
        story_summary.insert(1, 'M POC', story_summary['Area Name'].apply(lambda x: m_poc_dict.get(str(x).lower(), '')))
        story_summary.insert(2, 'SM POC', story_summary['Area Name'].apply(lambda x: sm_poc_dict.get(str(x).lower(), '')))
        story_summary.insert(3, 'AD POC', story_summary['Area Name'].apply(lambda x: ad_poc_dict.get(str(x).lower(), '')))
    
    # Add Grand Total row
    grand_total = pd.DataFrame([{
        'Area Name': 'Grand Total',
        'M POC': '',
        'SM POC': '',
        'AD POC': '',
        'Total Stories': story_summary['Total Stories'].sum(),
        'Testing Not Applicable Stories': story_summary['Testing Not Applicable Stories'].sum(),
        'Testable stories': story_summary['Testable stories'].sum(),
        'Story Points': story_summary['Story Points'].sum()
    }])
    
    story_summary = pd.concat([story_summary, grand_total], ignore_index=True)
    
    # Print summary statistics
    print(f"\n=== {release_name} Summary Statistics ===")
    print(f"Total Stories: {len(story_df)}")
    print(f"Total Story Points: {story_df['Story Points'].sum():.0f}")
    print(f"Testing Not Applicable Stories: {story_df['Testing Not Applicable Stories'].sum()}")
    print(f"Testable Stories: {story_df['Testable stories'].sum()}")
    print(f"Unique Areas: {story_df['Area Name'].nunique()}")
    
    return story_summary, story_df

# Process all releases
all_summaries = {}
print("\n" + "="*60)
print("STARTING DATA EXTRACTION FOR ALL RELEASES")
print("="*60)

for release_name, query_id in releases.items():
    summary_df, detail_df = fetch_release_data(release_name, query_id)
    if summary_df is not None:
        all_summaries[release_name] = summary_df

# Save all summaries to Excel file with separate sheets
if all_summaries:
    output_excel = os.path.join(output_folder, "Story_summary_final.xlsx")
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for release_name, summary_df in all_summaries.items():
            summary_df.to_excel(writer, sheet_name=release_name, index=False)
    
    print(f"\n{'='*60}")
    print(f"✓ All summaries saved to Excel: {output_excel}")
    print(f"  Sheets created: {', '.join(all_summaries.keys())}")
    print(f"{'='*60}")
else:
    print("\nNo data to save.")
