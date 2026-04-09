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
    "Nov 8": "dd3f3053-ba0c-4c09-ae43-b84ee16b35bb",
    "Dec 13": "00243f6a-15ab-43cb-91f1-6df8e4ff688c",
    "Jan 10": "6afab8e6-52a6-4c35-ab77-2094dac31514"
}

# Paths
MAPPING_FILE = r'C:\Users\d.sampathkumar\GHC files\POC mapping\POD Mapping sheet_Updated.csv'
OUTPUT_FILE = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\Stories score summary.xlsx'

# PT States to look for in revision history (in priority order)
PT_STATES = ['Ready for Test', 'Ready for E2E Test', 'In Test', 'PT In Test']

def load_poc_mapping():
    """Load POC mapping from CSV file"""
    try:
        mapping_df = pd.read_csv(MAPPING_FILE)
        print(f"✅ Loaded POC mapping: {len(mapping_df)} nodes")
        return mapping_df
    except Exception as e:
        print(f"⚠️ Error loading POC mapping: {e}")
        return pd.DataFrame(columns=['Node Name', 'M POC', 'SM Name', 'AD POC'])

def extract_node_name(area_path):
    """Extract Node Name from Area Path"""
    if not area_path:
        return ''
    
    # Area Path format: AutomationProcess_29697\Node Name or similar
    parts = area_path.split('\\')
    if len(parts) > 1:
        return parts[-1]  # Return the last part
    return area_path

def get_poc_for_node(node_name, mapping_df):
    """Get POCs for a given node name from mapping"""
    if not node_name or mapping_df.empty:
        return '', '', ''
    
    # Find matching node in mapping (case-insensitive)
    match = mapping_df[mapping_df['Node Name'].str.lower() == node_name.lower()]
    
    if not match.empty:
        row = match.iloc[0]
        return (
            row.get('AD POC', ''),
            row.get('SM Name', ''),  # SM POC is in 'SM Name' column
            row.get('M POC', '')
        )
    
    return '', '', ''

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
        
        # Extract work item IDs
        work_items = query_result.get('workItems', [])
        work_item_ids = [item['id'] for item in work_items]
        
        if not work_item_ids:
            print(f"⚠️ No work items found in query")
            return []
        
        print(f"✅ Found {len(work_item_ids)} work items")
        
        # Fetch detailed work item data
        # Azure DevOps API allows fetching up to 200 work items at once
        all_work_items = []
        batch_size = 200
        
        for i in range(0, len(work_item_ids), batch_size):
            batch_ids = work_item_ids[i:i + batch_size]
            ids_param = ','.join(map(str, batch_ids))
            
            # Fetch all fields - don't specify fields parameter
            work_items_url = f'https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems?ids={ids_param}&api-version=7.0'
            
            response = requests.get(work_items_url, auth=auth, headers=headers)
            response.raise_for_status()
            batch_result = response.json()
            
            all_work_items.extend(batch_result.get('value', []))
            print(f"  Fetched {len(all_work_items)}/{len(work_item_ids)} work items...")
        
        return all_work_items
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Error fetching query results: {e}")
        return []

async def get_actual_pt_date_async(session, work_item_id, semaphore):
    """Get the first occurrence of 'Ready for E2E Test' or 'Ready for Test' state from revision history - async version"""
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
                        return work_item_id, None
            except Exception as e:
                print(f"⚠️ Error fetching revisions for work item {work_item_id}: {e}")
                return work_item_id, None
        
        # Now process all revisions to find first PT state
        try:
            previous_state = None
            
            # Track first occurrence of each PT state
            first_pt_state_dates = {}
            
            for revision in all_revisions:
                current_state = revision.get('fields', {}).get('System.State')
                changed_date = revision.get('fields', {}).get('System.ChangedDate')
                
                if current_state and changed_date:
                    # Check if state matches any PT state (case-insensitive)
                    current_state_lower = current_state.lower()
                    
                    for pt_state in PT_STATES:
                        pt_state_lower = pt_state.lower()
                        
                        # If current state matches a PT state and it's the first time we see it
                        if current_state_lower == pt_state_lower and pt_state not in first_pt_state_dates:
                            # Check if it's a transition TO this state (not already in it)
                            previous_state_lower = previous_state.lower() if previous_state else None
                            if previous_state_lower != pt_state_lower:
                                first_pt_state_dates[pt_state] = changed_date
                
                previous_state = current_state
            
            # Return the first PT state found based on priority order
            for pt_state in PT_STATES:
                if pt_state in first_pt_state_dates:
                    return work_item_id, first_pt_state_dates[pt_state]
            
            # If no PT state found, return None
            return work_item_id, None
            
        except Exception as e:
            print(f"⚠️ Error processing revisions for work item {work_item_id}: {e}")
            return work_item_id, None

async def fetch_all_actual_pt_dates(work_item_ids, max_concurrent=50):
    """Fetch actual PT dates for all work items concurrently"""
    print(f"🚀 Fetching Actual PT dates for {len(work_item_ids)} work items (max {max_concurrent} concurrent)...")
    
    semaphore = asyncio.Semaphore(max_concurrent)
    
    async with aiohttp.ClientSession() as session:
        tasks = [get_actual_pt_date_async(session, wid, semaphore) for wid in work_item_ids]
        
        # Process with progress updates
        actual_pt_dates = {}
        completed = 0
        
        for coro in asyncio.as_completed(tasks):
            work_item_id, actual_pt_date = await coro
            actual_pt_dates[work_item_id] = actual_pt_date
            completed += 1
            
            if completed % 50 == 0 or completed == len(work_item_ids):
                print(f"📊 Progress: {completed}/{len(work_item_ids)} work items processed...")
        
        return actual_pt_dates

def process_work_items(work_items, actual_pt_dates, mapping_df):
    """Process work items and extract required fields"""
    print(f"📊 Processing {len(work_items)} work items...")
    
    processed_data = []
    
    for item in work_items:
        item_id = item.get('id')
        fields = item.get('fields', {})
        
        # Extract fields - using correct field names with underscores
        agent_augmented = fields.get('Custom.AgentAugmentedDelivery_Development', '')
        planned_pt_date = fields.get('Custom.PlannedforPTDate', '')
        
        # Get actual PT date from the fetched data
        actual_pt_date = actual_pt_dates.get(item_id, '')
        
        # Calculate delayed delivery and delay time
        delayed_delivery = ''
        delay_time = ''
        
        if planned_pt_date and actual_pt_date:
            try:
                # Parse dates and extract only date part (ignore time)
                planned_date_obj = datetime.fromisoformat(planned_pt_date.replace('Z', '+00:00')).date()
                actual_date_obj = datetime.fromisoformat(actual_pt_date.replace('Z', '+00:00')).date()
                
                # Calculate difference
                date_diff = (actual_date_obj - planned_date_obj).days
                
                # Determine if delayed
                if actual_date_obj > planned_date_obj:
                    delayed_delivery = 'Yes'
                    delay_time = date_diff
                else:
                    delayed_delivery = 'No'
                    delay_time = 0
            except Exception as e:
                # If date parsing fails, leave as empty
                pass
        
        # Extract additional fields
        area_path = fields.get('System.AreaPath', '')
        node_name = extract_node_name(area_path)
        ut_results = fields.get('Custom.UTResultsAttached', '')
        code_review = fields.get('Custom.CodeReviewCompleted', '')
        
        # Get POCs from mapping
        ad_poc, sm_poc, m_poc = get_poc_for_node(node_name, mapping_df)
        
        processed_data.append({
            'ID': item_id,
            'Agent Augmented delivery_Development': agent_augmented,
            'Planned for PT Date': planned_pt_date,
            'Actual PT date': actual_pt_date,
            'Delayed Story Delivery': delayed_delivery,
            'Delay Time': delay_time,
            'Node Name': node_name,
            'UT Results Attached': ut_results,
            'Code Review Completed': code_review,
            'AD POC': ad_poc,
            'SM POC': sm_poc,
            'M POC': m_poc
        })
    
    return pd.DataFrame(processed_data)

def main():
    """Main execution function"""
    print("="*80)
    print("Stories Score Summary Data Fetch")
    print("="*80)
    
    # Load POC mapping
    print("\n📂 Loading POC mapping...")
    mapping_df = load_poc_mapping()
    
    # Dictionary to store DataFrames for each release
    all_sheets = {}
    
    # Process each release query
    for release_name, query_id in QUERIES.items():
        print(f"\n{'='*80}")
        print(f"Processing Release: {release_name}")
        print(f"{'='*80}")
        
        # Fetch work items from query
        work_items = fetch_ado_query_results(query_id)
        
        if not work_items:
            print(f"⚠️ No work items found for {release_name}, skipping...")
            all_sheets[release_name] = pd.DataFrame(columns=[
                'ID', 
                'Agent Augmented delivery_Development',
                'Planned for PT Date',
                'Actual PT date',
                'Delayed Story Delivery',
                'Delay Time',
                'Node Name',
                'UT Results Attached',
                'Code Review Completed',
                'AD POC',
                'SM POC',
                'M POC'
            ])
            continue
        
        # Extract work item IDs
        work_item_ids = [item['id'] for item in work_items]
        
        # Fetch actual PT dates asynchronously
        actual_pt_dates = asyncio.run(fetch_all_actual_pt_dates(work_item_ids))
        
        # Process work items into DataFrame
        df = process_work_items(work_items, actual_pt_dates, mapping_df)
        
        # Store in dictionary
        all_sheets[release_name] = df
        
        print(f"✅ Processed {len(df)} stories for {release_name}")
        print(f"   - Stories with Actual PT date: {df['Actual PT date'].notna().sum()}")
        print(f"   - Stories without Actual PT date: {df['Actual PT date'].isna().sum()}")
    
    # Write all sheets to Excel
    print(f"\n{'='*80}")
    print(f"Writing data to Excel file...")
    print(f"{'='*80}")
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  ✅ Sheet '{sheet_name}': {len(df)} records")
    
    print(f"\n{'='*80}")
    print(f"✅ Excel file created successfully!")
    print(f"📁 Location: {OUTPUT_FILE}")
    print(f"{'='*80}")
    
    # Summary
    total_stories = sum(len(df) for df in all_sheets.values())
    total_with_pt = sum(df['Actual PT date'].notna().sum() for df in all_sheets.values())
    
    print(f"\n📊 Summary:")
    print(f"  Total Stories: {total_stories}")
    print(f"  Stories with Actual PT date: {total_with_pt}")
    print(f"  Stories without Actual PT date: {total_stories - total_with_pt}")

if __name__ == "__main__":
    main()
