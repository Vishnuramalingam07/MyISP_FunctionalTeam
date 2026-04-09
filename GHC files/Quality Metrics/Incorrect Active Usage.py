import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
from datetime import datetime
import os
import asyncio
import aiohttp
from aiohttp import BasicAuth
from dotenv import load_dotenv

# Load environment variables
load_dotenv('ADO_SECRETS.env')

# Configuration
PAT = os.getenv('ADO_PAT_MAIN', '')
if not PAT:
    raise ValueError("ADO_PAT_MAIN not found in ADO_SECRETS.env file")
ORG = 'accenturecio08'
PROJECT = 'AutomationProcess_29697'

# Define queries for each release
QUERIES = {
    "Nov 8": "4d468701-65dd-431c-916a-c0c69a14788d",
    "Dec 13": "dbe5fecb-3b47-40f4-96d9-d4e6947e750a",
    "Jan 10": "7e2da104-c0c5-4558-9082-abf2c349f015"
}

# Paths
MAPPING_FILE = r'C:\Users\d.sampathkumar\GHC files\POC mapping\POD Mapping sheet_Updated.csv'
OUTPUT_FILE = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Defects_with_Incorrect_Active_Usage.xlsx'

# Test Bucket States - states where defect should be in tester's hands
TEST_BUCKET_STATES = ['Ready to Test', 'Rejected', 'For Rejection', 'In Test', 'PT In Test', 'UAT In Test']

def fetch_ado_query_results(query_id):
    """Fetch work items from ADO query"""
    print(f"🔍 Fetching data from ADO query: {query_id}")
    
    # Set up authentication
    auth = HTTPBasicAuth('', PAT)
    headers = {'Content-Type': 'application/json'}
    
    # Execute query to get work item IDs
    query_url = f'https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/wiql/{query_id}?api-version=7.0'
    
    try:
        response = requests.get(query_url, auth=auth, headers=headers)
        response.raise_for_status()
        query_result = response.json()
        
        work_item_ids = [item['id'] for item in query_result.get('workItems', [])]
        
        if not work_item_ids:
            print("❌ No work items found in query")
            return []
        
        print(f"✅ Found {len(work_item_ids)} work items")
        
        # Fetch work item details in batches of 200
        batch_size = 200
        all_work_items = []
        
        for i in range(0, len(work_item_ids), batch_size):
            batch = work_item_ids[i:i + batch_size]
            ids_str = ','.join(map(str, batch))
            
            # Request specific fields
            fields = [
                'System.Id',
                'System.AreaPath',
                'System.Title',
                'System.State',
                'System.CreatedDate',
                'Microsoft.VSTS.Common.ClosedDate',
                'Microsoft.VSTS.Common.Severity'
            ]
            
            details_url = f'https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems?ids={ids_str}&fields={",".join(fields)}&api-version=7.0'
            
            details_response = requests.get(details_url, auth=auth, headers=headers)
            details_response.raise_for_status()
            details_result = details_response.json()
            
            all_work_items.extend(details_result.get('value', []))
            print(f"📥 Fetched {len(all_work_items)}/{len(work_item_ids)} work items...")
        
        return all_work_items
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Error fetching data from ADO: {e}")
        return []

def extract_node_name(area_path):
    """Extract Node Name from Area Path"""
    if not area_path:
        return ''
    
    # Area Path format: AutomationProcess_29697\Node Name or similar
    parts = area_path.split('\\')
    if len(parts) > 1:
        return parts[-1]  # Return the last part
    return area_path

async def check_incorrect_active_usage_async(session, work_item_id, semaphore):
    """Check if a defect was incorrectly moved to Active instead of Re-open after Test Bucket states"""
    async with semaphore:
        revisions_url = f'https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems/{work_item_id}/revisions?api-version=7.0'
        
        try:
            auth = BasicAuth('', PAT)
            async with session.get(revisions_url, auth=auth) as response:
                if response.status == 200:
                    data = await response.json()
                    revisions = data.get('value', [])
                    
                    # Track state transitions
                    previous_state = None
                    incorrect_active_transitions = []
                    was_in_test_bucket = False
                    
                    for revision in revisions:
                        current_state = revision.get('fields', {}).get('System.State')
                        changed_date = revision.get('fields', {}).get('System.ChangedDate')
                        changed_by = revision.get('fields', {}).get('System.ChangedBy', {}).get('displayName', 'Unknown')
                        
                        # Check if defect is in Test Bucket states
                        if current_state in TEST_BUCKET_STATES:
                            was_in_test_bucket = True
                        
                        # Detect incorrect workflow: Moving to 'Active' after being in Test Bucket
                        # This should have been 'Re-open' instead
                        if current_state == 'Active' and was_in_test_bucket and previous_state in TEST_BUCKET_STATES:
                            incorrect_active_transitions.append({
                                'from_state': previous_state,
                                'to_state': current_state,
                                'date': changed_date,
                                'changed_by': changed_by
                            })
                        
                        # Reset flag if moved to Re-open (correct workflow)
                        if current_state == 'Re-open':
                            was_in_test_bucket = False
                        
                        previous_state = current_state
                    
                    return work_item_id, incorrect_active_transitions
                else:
                    print(f"⚠️ Error fetching revisions for work item {work_item_id}: HTTP {response.status}")
                    return work_item_id, []
                    
        except Exception as e:
            print(f"⚠️ Error fetching revisions for work item {work_item_id}: {e}")
            return work_item_id, []

async def check_all_work_items_async(work_item_ids, max_concurrent=50):
    """Check all work items for incorrect Active usage concurrently"""
    print(f"🚀 Checking {len(work_item_ids)} work items for incorrect Active usage (max {max_concurrent} concurrent)...")
    
    semaphore = asyncio.Semaphore(max_concurrent)
    
    async with aiohttp.ClientSession() as session:
        tasks = [check_incorrect_active_usage_async(session, wid, semaphore) for wid in work_item_ids]
        
        # Process with progress updates
        results = {}
        completed = 0
        
        for coro in asyncio.as_completed(tasks):
            work_item_id, transitions = await coro
            results[work_item_id] = transitions
            completed += 1
            
            if completed % 100 == 0 or completed == len(work_item_ids):
                print(f"📊 Progress: {completed}/{len(work_item_ids)} work items checked...")
        
        return results

def process_work_items(work_items, incorrect_active_data):
    """Process work items and extract those with incorrect Active usage"""
    print(f"📊 Processing {len(work_items)} work items...")
    
    data = []
    for item in work_items:
        work_item_id = item.get('id')
        
        # Check if this work item has incorrect Active usage
        transitions = incorrect_active_data.get(work_item_id, [])
        
        if transitions:  # Only include defects with incorrect Active usage
            fields = item.get('fields', {})
            
            # Extract area path and node name
            area_path = fields.get('System.AreaPath', '')
            node_name = extract_node_name(area_path)
            
            # Get fields
            title = fields.get('System.Title', '')
            current_state = fields.get('System.State', '')
            created_date = fields.get('System.CreatedDate', '')
            closed_date = fields.get('Microsoft.VSTS.Common.ClosedDate', '')
            severity = fields.get('Microsoft.VSTS.Common.Severity', '')
            
            # Format transition details
            transition_details = []
            changed_by_list = []
            for trans in transitions:
                transition_details.append(f"{trans['from_state']}→Active on {trans['date']}")
                if trans['changed_by'] not in changed_by_list:
                    changed_by_list.append(trans['changed_by'])
            
            incorrect_active_details = '; '.join(transition_details)
            changed_by = ', '.join(changed_by_list)
            
            data.append({
                'ID': work_item_id,
                'Title': title,
                'Node Name': node_name,
                'Current State': current_state,
                'Incorrect Active Count': len(transitions),
                'Incorrect Active Details': incorrect_active_details,
                'Changed By': changed_by,
                'Created Date': created_date,
                'Closed Date': closed_date,
                'Severity': severity
            })
    
    return pd.DataFrame(data)

def load_mapping():
    """Load POC mapping from CSV"""
    print(f"📂 Loading POC mapping from: {MAPPING_FILE}")
    
    if not os.path.exists(MAPPING_FILE):
        print(f"❌ Mapping file not found: {MAPPING_FILE}")
        return pd.DataFrame()
    
    try:
        mapping_df = pd.read_csv(MAPPING_FILE)
        print(f"✅ Loaded {len(mapping_df)} mapping records")
        return mapping_df
    except Exception as e:
        print(f"❌ Error loading mapping file: {e}")
        return pd.DataFrame()

def add_poc_mapping(data_df, mapping_df):
    """Add POC columns to data based on Node Name mapping"""
    print("🔗 Adding POC mappings...")
    
    if mapping_df.empty:
        print("⚠️ Mapping data is empty, adding empty POC columns")
        data_df['AD POC'] = ''
        data_df['SM POC'] = ''
        data_df['M POC'] = ''
        return data_df
    
    # Clean and prepare data for case-insensitive matching
    data_df['Node Name'] = data_df['Node Name'].str.strip()
    data_df['Node Name_lower'] = data_df['Node Name'].str.replace(' ', '').str.lower()
    
    # Merge on Node Name
    mapping_df.columns = mapping_df.columns.str.strip()
    
    # Select only required columns from mapping
    if 'Node Name' in mapping_df.columns:
        mapping_subset = mapping_df[['Node Name', 'M POC', 'SM Name', 'AD POC']].copy()
        mapping_subset.columns = ['Node Name', 'M POC', 'SM POC', 'AD POC']
        
        # Clean mapping data
        mapping_subset['Node Name'] = mapping_subset['Node Name'].str.strip()
        mapping_subset['Node Name_lower'] = mapping_subset['Node Name'].str.replace(' ', '').str.lower()
        
        # Remove duplicates
        mapping_subset = mapping_subset.drop_duplicates(subset='Node Name_lower', keep='first')
        
        # Merge using lowercase for comparison
        result_df = data_df.merge(
            mapping_subset[['Node Name_lower', 'M POC', 'SM POC', 'AD POC']],
            on='Node Name_lower',
            how='left'
        )
        
        # Drop the temporary lowercase column
        result_df = result_df.drop('Node Name_lower', axis=1)
        
        # Fill NaN values with empty string
        result_df['AD POC'] = result_df['AD POC'].fillna('')
        result_df['SM POC'] = result_df['SM POC'].fillna('')
        result_df['M POC'] = result_df['M POC'].fillna('')
        
        matched_count = (result_df['AD POC'] != '').sum()
        print(f"✅ Matched {matched_count} records with POC data")
        
        return result_df
    else:
        print("❌ 'Node Name' column not found in mapping file")
        data_df['AD POC'] = ''
        data_df['SM POC'] = ''
        data_df['M POC'] = ''
        return data_df

def save_to_excel(all_releases_data):
    """Save DataFrame to Excel file with all releases combined"""
    print(f"\n💾 Saving Incorrect Active Usage Report to: {OUTPUT_FILE}")
    
    try:
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
        
        # Combine all releases
        all_defects = []
        for release_name, df in all_releases_data.items():
            if not df.empty:
                df_copy = df.copy()
                df_copy.insert(0, 'Release', release_name)
                all_defects.append(df_copy)
        
        if all_defects:
            combined_df = pd.concat(all_defects, ignore_index=True)
            
            # Reorder columns for better readability
            column_order = ['Release', 'ID', 'Title', 'Node Name', 'Current State', 
                          'Incorrect Active Count', 'Incorrect Active Details', 'Changed By',
                          'AD POC', 'SM POC', 'M POC', 
                          'Created Date', 'Closed Date', 'Severity']
            combined_df = combined_df[column_order]
            
            # Sort by Release and Incorrect Active Count (descending)
            combined_df = combined_df.sort_values(['Release', 'Incorrect Active Count'], 
                                                  ascending=[True, False])
            
            # Save to Excel
            with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Incorrect Active Usage', index=False)
            
            print(f"✅ Successfully saved report!")
            print(f"📊 File location: {OUTPUT_FILE}")
            print(f"📈 Total defects with incorrect Active usage: {len(combined_df)}")
            print(f"📈 Total incorrect Active transitions: {combined_df['Incorrect Active Count'].sum()}")
            
            # Print breakdown by release
            print("\n📋 Breakdown by Release:")
            for release in combined_df['Release'].unique():
                release_count = len(combined_df[combined_df['Release'] == release])
                print(f"  {release}: {release_count} defects")
        else:
            print(f"✅ No defects found with incorrect Active usage across all releases")
        
    except Exception as e:
        print(f"❌ Error saving to Excel: {e}")

def main():
    """Main execution function"""
    print("=" * 80)
    print("🔍 ADO Defect Analysis - Incorrect Active Usage Detection")
    print("=" * 80)
    print("\nThis script identifies defects that were moved to 'Active' state")
    print("instead of 'Re-open' after being in Test Bucket states.")
    print("=" * 80)
    
    # Load POC mapping once
    mapping_df = load_mapping()
    
    # Process each release
    all_releases_data = {}
    
    for release_name, query_id in QUERIES.items():
        print(f"\n{'='*80}")
        print(f"📋 Processing {release_name}")
        print(f"{'='*80}")
        
        # Step 1: Fetch data from ADO
        work_items = fetch_ado_query_results(query_id)
        if not work_items:
            print(f"❌ No data fetched for {release_name}. Skipping.")
            continue
        
        # Step 2: Extract work item IDs and check for incorrect Active usage
        work_item_ids = [item.get('id') for item in work_items]
        print(f"\n🔄 Analyzing revision history for incorrect Active usage...")
        incorrect_active_data = asyncio.run(check_all_work_items_async(work_item_ids, max_concurrent=50))
        
        # Count defects with issues
        defects_with_issues = sum(1 for transitions in incorrect_active_data.values() if transitions)
        total_transitions = sum(len(transitions) for transitions in incorrect_active_data.values())
        
        print(f"⚠️  Found {defects_with_issues} defects with incorrect Active usage")
        print(f"⚠️  Total incorrect Active transitions: {total_transitions}")
        
        # Step 3: Process work items (only those with incorrect Active usage)
        data_df = process_work_items(work_items, incorrect_active_data)
        
        # Step 4: Add POC mapping
        if not data_df.empty:
            final_df = add_poc_mapping(data_df, mapping_df)
        else:
            final_df = data_df
        
        # Store in dictionary
        all_releases_data[release_name] = final_df
        
        # Print summary for this release
        print(f"\n📈 {release_name} Summary:")
        print(f"  Total Defects Checked: {len(work_items)}")
        print(f"  Defects with Incorrect Active Usage: {len(final_df)}")
        if not final_df.empty:
            print(f"  Total Incorrect Active Transitions: {final_df['Incorrect Active Count'].sum()}")
    
    # Step 5: Save combined report to Excel
    if all_releases_data:
        print(f"\n{'='*80}")
        save_to_excel(all_releases_data)
        
        print("\n" + "=" * 80)
        print("✅ Analysis completed successfully!")
        print("=" * 80)
    else:
        print("\n❌ No data processed for any release.")

if __name__ == "__main__":
    main()

