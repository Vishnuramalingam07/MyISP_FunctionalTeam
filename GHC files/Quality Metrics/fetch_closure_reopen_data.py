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
OUTPUT_FILE = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Closure_reopen_with_POC.xlsx'

# Test Bucket States
# Priority order: First check for "Ready to Test", then check other states
PRIMARY_TEST_STATE = 'Ready to Test'
FALLBACK_TEST_STATES = ['Rejected', 'For Rejection', 'In Test', 'PT In Test', 'UAT In Test']
TEST_BUCKET_STATES = [PRIMARY_TEST_STATE] + FALLBACK_TEST_STATES

# Additional fallback states for defects that never went through testing
ALTERNATE_COMPLETION_STATES = ['Monitoring', 'Ready for Prod Deployment']

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
            
            # Request specific fields including Severity
            fields = [
                'System.Id',
                'System.AreaPath',
                'System.CreatedDate',
                'Microsoft.VSTS.Common.ClosedDate',
                'Microsoft.VSTS.Common.Severity',
                'Custom.ReopenCount'
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

async def get_reopen_count_and_test_bucket_time_async(session, work_item_id, semaphore):
    """Get the reopen count and appropriate Test Bucket state timestamp for a specific work item using revisions - async version
    
    For reopened defects OR defects with incorrect Active usage: Returns the LAST 'Ready to Test' time after the last reopen/incorrect Active
    For non-reopened defects: Returns the FIRST 'Ready to Test' time
    """
    async with semaphore:
        # Azure DevOps API returns max 200 revisions by default, need to fetch all pages
        all_revisions = []
        skip = 0
        page_size = 200
        
        while True:
            revisions_url = f'https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems/{work_item_id}/revisions?$skip={skip}&api-version=7.0'
            
            try:
                auth = BasicAuth('', PAT)
                async with session.get(revisions_url, auth=auth) as response:
                    if response.status == 200:
                        data = await response.json()
                        revisions = data.get('value', [])
                        
                        if not revisions:
                            break
                        
                        all_revisions.extend(revisions)
                        
                        # If we got less than page_size, we've reached the end
                        if len(revisions) < page_size:
                            break
                        
                        skip += page_size
                    else:
                        print(f"⚠️ Error fetching revisions for work item {work_item_id}: HTTP {response.status}")
                        return work_item_id, 0, None
            except Exception as e:
                print(f"⚠️ Error fetching revisions for work item {work_item_id}: {e}")
                return work_item_id, 0, None
        
        # Now process all revisions
        try:
                    
                    # Track reopen events and their timestamps
                    reopen_count = 0
                    previous_state = None
                    last_reopen_or_incorrect_active_date = None
                    was_in_test_bucket = False
                    
                    # Track all "Ready to Test" and fallback test state transitions
                    all_ready_to_test_times = []
                    all_fallback_test_times = []
                    
                    # Track alternate completion states (Monitoring, Ready for Prod deployment)
                    all_alternate_completion_times = []
                    
                    for revision in all_revisions:
                        current_state = revision.get('fields', {}).get('System.State')
                        changed_date = revision.get('fields', {}).get('System.ChangedDate')
                        
                        # Check if defect is/was in Test Bucket states
                        if current_state in TEST_BUCKET_STATES:
                            was_in_test_bucket = True
                        
                        # Detect INCORRECT workflow: Moving to 'Active' after being in Test Bucket
                        # This should have been 'Re-open' - treat it as a pseudo-reopen
                        if current_state == 'Active' and was_in_test_bucket and previous_state in TEST_BUCKET_STATES:
                            reopen_count += 1  # Count as pseudo-reopen
                            last_reopen_or_incorrect_active_date = changed_date
                            was_in_test_bucket = False  # Reset flag
                        
                        # Count whenever state changes to 'Re-open' from any other state (CORRECT workflow)
                        if current_state == 'Re-open' and previous_state != 'Re-open' and previous_state is not None:
                            reopen_count += 1
                            last_reopen_or_incorrect_active_date = changed_date
                            was_in_test_bucket = False  # Reset flag
                        
                        # Collect ALL "Ready to Test" state transitions with timestamps
                        if current_state == PRIMARY_TEST_STATE and previous_state != PRIMARY_TEST_STATE:
                            all_ready_to_test_times.append(changed_date)
                        
                        # Collect ALL fallback test state transitions with timestamps
                        elif current_state in FALLBACK_TEST_STATES and previous_state not in FALLBACK_TEST_STATES:
                            all_fallback_test_times.append(changed_date)
                        
                        # Collect alternate completion states (for defects that never went through testing)
                        elif current_state in ALTERNATE_COMPLETION_STATES and previous_state not in ALTERNATE_COMPLETION_STATES:
                            all_alternate_completion_times.append(changed_date)
                        
                        previous_state = current_state
                    
                    # Debug output for specific defect
                    if work_item_id == 4133807:
                        print(f"\n🔍 DEBUG Defect {work_item_id}:")
                        print(f"   Reopen count: {reopen_count}")
                        print(f"   Last reopen/incorrect active: {last_reopen_or_incorrect_active_date}")
                        print(f"   All Ready to Test times: {len(all_ready_to_test_times)} - {all_ready_to_test_times[:3] if all_ready_to_test_times else 'None'}")
                        print(f"   All fallback test times: {len(all_fallback_test_times)}")
                        print(f"   All alternate completion times: {len(all_alternate_completion_times)} - {all_alternate_completion_times[:3] if all_alternate_completion_times else 'None'}")
                    
                    # Determine which Test Bucket time to use based on reopen/incorrect Active status
                    appropriate_test_bucket_time = None
                    
                    if reopen_count > 0 and last_reopen_or_incorrect_active_date:
                        # For REOPENED defects OR defects with incorrect Active usage:
                        # Find the LAST "Ready to Test" after the last reopen/incorrect Active
                        ready_to_test_after_reopen = [
                            dt for dt in all_ready_to_test_times 
                            if dt > last_reopen_or_incorrect_active_date
                        ]
                        
                        if ready_to_test_after_reopen:
                            # Use the LAST occurrence after the last reopen/incorrect Active
                            appropriate_test_bucket_time = ready_to_test_after_reopen[-1]
                        else:
                            # Fallback: Check for fallback states after last reopen/incorrect Active
                            fallback_after_reopen = [
                                dt for dt in all_fallback_test_times 
                                if dt > last_reopen_or_incorrect_active_date
                            ]
                            if fallback_after_reopen:
                                appropriate_test_bucket_time = fallback_after_reopen[-1]
                            else:
                                # Fallback: Check for alternate completion states after last reopen
                                alternate_after_reopen = [
                                    dt for dt in all_alternate_completion_times 
                                    if dt > last_reopen_or_incorrect_active_date
                                ]
                                if alternate_after_reopen:
                                    appropriate_test_bucket_time = alternate_after_reopen[-1]
                                else:
                                    # Final fallback: Use first Ready to Test even if before reopen
                                    if all_ready_to_test_times:
                                        appropriate_test_bucket_time = all_ready_to_test_times[0]
                                    elif all_fallback_test_times:
                                        appropriate_test_bucket_time = all_fallback_test_times[0]
                                    elif all_alternate_completion_times:
                                        appropriate_test_bucket_time = all_alternate_completion_times[0]
                    else:
                        # For NON-REOPENED defects: Use the FIRST "Ready to Test" (existing logic)
                        if all_ready_to_test_times:
                            appropriate_test_bucket_time = all_ready_to_test_times[0]
                        elif all_fallback_test_times:
                            appropriate_test_bucket_time = all_fallback_test_times[0]
                        elif all_alternate_completion_times:
                            # If never went through test bucket, use alternate completion states
                            appropriate_test_bucket_time = all_alternate_completion_times[0]
                    
                    return work_item_id, reopen_count, appropriate_test_bucket_time
                    
        except Exception as e:
            print(f"⚠️ Error fetching revisions for work item {work_item_id}: {e}")
            return work_item_id, 0, None

async def fetch_all_reopen_counts_and_test_bucket_times(work_item_ids, max_concurrent=50):
    """Fetch reopen counts and appropriate Test Bucket timestamps for all work items concurrently
    
    For reopened defects OR defects with incorrect Active usage: Returns the LAST 'Ready to Test' time after the last reopen/incorrect Active
    For non-reopened defects: Returns the FIRST 'Ready to Test' time
    
    Note: Incorrect Active usage (Test Bucket → Active) is treated as a pseudo-reopen for metrics accuracy
    """
    print(f"🚀 Fetching reopen counts and appropriate Test Bucket state times for {len(work_item_ids)} work items (max {max_concurrent} concurrent)...")
    
    semaphore = asyncio.Semaphore(max_concurrent)
    
    async with aiohttp.ClientSession() as session:
        tasks = [get_reopen_count_and_test_bucket_time_async(session, wid, semaphore) for wid in work_item_ids]
        
        # Process with progress updates
        reopen_results = {}
        test_bucket_results = {}
        completed = 0
        
        for coro in asyncio.as_completed(tasks):
            work_item_id, reopen_count, first_test_bucket_time = await coro
            reopen_results[work_item_id] = reopen_count
            test_bucket_results[work_item_id] = first_test_bucket_time
            completed += 1
            
            if completed % 100 == 0 or completed == len(work_item_ids):
                print(f"📊 Progress: {completed}/{len(work_item_ids)} work items processed...")
        
        return reopen_results, test_bucket_results
    """Extract Node Name from Area Path"""
    if not area_path:
        return ''
    
    # Area Path format: AutomationProcess_29697\Node Name or similar
    parts = area_path.split('\\')
    if len(parts) > 1:
        return parts[-1]  # Return the last part
    return area_path

def process_work_items(work_items, reopen_counts, test_bucket_times):
    """Process work items and extract required fields"""
    print(f"📊 Processing {len(work_items)} work items...")
    
    data = []
    for item in work_items:
        fields = item.get('fields', {})
        work_item_id = item.get('id')
        
        # Extract area path and node name
        area_path = fields.get('System.AreaPath', '')
        node_name = extract_node_name(area_path)
        
        # Get reopen count from the async results
        reopen_count = reopen_counts.get(work_item_id, 0)
        
        # Get dates
        created_date = fields.get('System.CreatedDate', '')
        closed_date = fields.get('Microsoft.VSTS.Common.ClosedDate', '')
        
        # Get first Test Bucket state timestamp
        first_test_bucket_time = test_bucket_times.get(work_item_id)
        
        # If defect never went through Test Bucket states, use Closed Date as fallback
        if not first_test_bucket_time and closed_date:
            first_test_bucket_time = closed_date
        
        # Get Severity
        severity = fields.get('Microsoft.VSTS.Common.Severity', '')
        
        # Calculate Closure trend (hours between created and first Test Bucket state)
        closure_trend = ''
        closure_trend_hours = ''
        if created_date and first_test_bucket_time:
            try:
                created_dt = pd.to_datetime(created_date)
                test_bucket_dt = pd.to_datetime(first_test_bucket_time)
                hours_diff = (test_bucket_dt - created_dt).total_seconds() / 3600  # Convert to hours
                closure_trend_hours = round(hours_diff, 2)  # Round to 2 decimal places
                
                # Also calculate in days for SLA comparison
                days_diff = hours_diff / 24
                closure_trend = round(days_diff, 2)
            except Exception as e:
                print(f"⚠️ Error calculating closure trend for work item {work_item_id}: {e}")
                closure_trend = ''
                closure_trend_hours = ''
        
        # Calculate SLA based on Severity and Closure trend (in days)
        sla = ''
        if closure_trend != '' and severity:
            # Critical (1 - Critical) - must reach Test Bucket within 1 day (0-24 hours)
            if severity == '1 - Critical':
                sla = 'Met SLA' if closure_trend <= 1 else 'Not Met SLA'
            # High (2 - High) - must reach Test Bucket within 2 days (0-48 hours)
            elif severity == '2 - High':
                sla = 'Met SLA' if closure_trend <= 2 else 'Not Met SLA'
            # Medium (3 - Medium) - must reach Test Bucket within 3 days (0-72 hours)
            elif severity == '3 - Medium':
                sla = 'Met SLA' if closure_trend <= 3 else 'Not Met SLA'
            # Low (4 - Low) - must reach Test Bucket within 4 days (0-96 hours)
            elif severity == '4 - Low':
                sla = 'Met SLA' if closure_trend <= 4 else 'Not Met SLA'
        
        data.append({
            'ID': work_item_id,
            'Node Name': node_name,
            'Created Date': created_date,
            'Test Bucket State Time': first_test_bucket_time if first_test_bucket_time else '',
            'Closed Date': closed_date,
            'Closure Trend (Hours)': closure_trend_hours,
            'Closure Trend (Days)': closure_trend,
            'Re open Count': reopen_count,
            'Severity': severity,
            'SLA': sla
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
    # Strip whitespace and create lowercase version for matching
    data_df['Node Name'] = data_df['Node Name'].str.strip()
    # Remove spaces and convert to lowercase for robust matching
    data_df['Node Name_lower'] = data_df['Node Name'].str.replace(' ', '').str.lower()
    
    # Merge on Node Name
    # First, ensure column names are clean
    mapping_df.columns = mapping_df.columns.str.strip()
    
    # Select only required columns from mapping
    if 'Node Name' in mapping_df.columns:
        mapping_subset = mapping_df[['Node Name', 'M POC', 'SM Name', 'AD POC']].copy()
        mapping_subset.columns = ['Node Name', 'M POC', 'SM POC', 'AD POC']
        
        # Clean mapping data
        mapping_subset['Node Name'] = mapping_subset['Node Name'].str.strip()
        # Remove spaces and convert to lowercase for robust matching
        mapping_subset['Node Name_lower'] = mapping_subset['Node Name'].str.replace(' ', '').str.lower()
        
        # Remove duplicates from mapping if any exist after lowercase transformation
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
        unmatched_count = len(result_df) - matched_count
        print(f"✅ Matched {matched_count} records, {unmatched_count} unmatched")
        
        # Print some unmatched node names for debugging
        if unmatched_count > 0:
            unmatched_nodes = result_df[result_df['AD POC'] == '']['Node Name'].unique()[:10]
            print(f"⚠️ Sample unmatched Node Names: {', '.join(unmatched_nodes)}")
        
        return result_df
    else:
        print("❌ 'Node Name' column not found in mapping file")
        data_df['AD POC'] = ''
        data_df['SM POC'] = ''
        data_df['M POC'] = ''
        return data_df

def save_to_excel(all_releases_data):
    """Save DataFrame to Excel file with multiple sheets"""
    print(f"💾 Saving data to: {OUTPUT_FILE}")
    
    try:
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
        
        # Create Excel writer
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            for release_name, df in all_releases_data.items():
                # Reorder columns to show: ID, Node Name, Created Date, Test Bucket State Time, Closed Date, Closure Trend (Hours), Closure Trend (Days), Re open Count, AD POC, SM POC, M POC, Severity, SLA
                column_order = ['ID', 'Node Name', 'Created Date', 'Test Bucket State Time', 'Closed Date', 'Closure Trend (Hours)', 'Closure Trend (Days)', 'Re open Count', 'AD POC', 'SM POC', 'M POC', 'Severity', 'SLA']
                df = df[column_order]
                
                # Save to sheet
                df.to_excel(writer, sheet_name=release_name, index=False)
                print(f"  ✅ Added sheet '{release_name}' with {len(df)} records")
        
        print(f"✅ Successfully saved all releases to Excel")
        print(f"📊 File location: {OUTPUT_FILE}")
        
    except Exception as e:
        print(f"❌ Error saving to Excel: {e}")

def main():
    """Main execution function"""
    print("=" * 80)
    print("🚀 ADO Closure and Reopen Data Extraction (Multi-Release)")
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
        
        # Step 2: Extract work item IDs and fetch reopen counts and first Test Bucket times asynchronously
        work_item_ids = [item.get('id') for item in work_items]
        print(f"\n🔄 Fetching reopen counts and first Test Bucket state times from revision history...")
        reopen_counts, test_bucket_times = asyncio.run(fetch_all_reopen_counts_and_test_bucket_times(work_item_ids, max_concurrent=50))
        
        # Step 3: Process work items
        data_df = process_work_items(work_items, reopen_counts, test_bucket_times)
        
        # Step 4: Add POC mapping
        final_df = add_poc_mapping(data_df, mapping_df)
        
        # Store in dictionary
        all_releases_data[release_name] = final_df
        
        # Print summary for this release
        print(f"\n📈 {release_name} Summary:")
        print(f"  Total Records: {len(final_df)}")
        print(f"  Records with Test Bucket State: {final_df['Test Bucket State Time'].notna().sum()}")
        print(f"  Records with Closed Date: {final_df['Closed Date'].notna().sum()}")
        print(f"  Records with Reopens: {(final_df['Re open Count'] > 0).sum()}")
        print(f"  Total Reopens: {final_df['Re open Count'].sum()}")
    
    # Step 5: Save all releases to Excel with multiple sheets
    if all_releases_data:
        print(f"\n{'='*80}")
        save_to_excel(all_releases_data)
        
        print("\n" + "=" * 80)
        print("✅ Process completed successfully!")
        print("=" * 80)
        
        # Print overall summary
        print("\n📊 Overall Summary:")
        total_records = sum(len(df) for df in all_releases_data.values())
        print(f"  Total Releases Processed: {len(all_releases_data)}")
        print(f"  Total Records Across All Releases: {total_records}")
    else:
        print("\n❌ No data processed for any release.")

if __name__ == "__main__":
    main()
