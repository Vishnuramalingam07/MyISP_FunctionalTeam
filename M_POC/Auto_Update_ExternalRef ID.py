import pandas as pd
import os
import requests
import json
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import sys
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Azure DevOps configuration
org = "accenturecio08"
project = "AutomationProcess_29697"

# Get PAT from environment variable
pat = os.environ.get("AZURE_DEVOPS_PAT")
if not pat:
    raise RuntimeError("Azure DevOps PAT not found in environment variable 'AZURE_DEVOPS_PAT'. Please set it before running the script.")

headers = {"Content-Type": "application/json"}

# Create session with connection pooling and retry strategy
session = requests.Session()
retry_strategy = Retry(
    total=3,
    backoff_factor=0.5,
    status_forcelist=[429, 500, 502, 503, 504]
)
adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=50, pool_maxsize=50)
session.mount("https://", adapter)
session.mount("http://", adapter)

# Rate limiting variables
rate_limit_lock = threading.Lock()
rate_limit_delay = 0

# ADO Query ID (configured via Settings in the web app)
query_id = "6badbdf6-4b17-4053-a23b-383c21ab4d39"  # ADO Query ID

# Mapping file path (works in both Windows and Docker)
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
mapping_path = os.path.join(project_root, "Common_Files", "M POC Mapping.csv")

# Read mapping file
print(f"Reading mapping file: {mapping_path}")
mapping_df = pd.read_csv(mapping_path, dtype=str)

# Create mapping dictionary (case-insensitive)
mapping_df['Node Name_clean'] = mapping_df['Node Name'].fillna('').str.strip().str.lower()
mapping_df['ExternalRef ID'] = mapping_df['ExternalRef ID'].fillna('').str.strip()
node_to_extref = dict(zip(mapping_df['Node Name_clean'], mapping_df['ExternalRef ID']))

print(f"Loaded {len(node_to_extref)} mapping entries")

def get_query_results(query_id):
    """Fetch work items from the query"""
    api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"
    print(f"\nFetching work items from query: {query_id}")
    
    response = session.get(api_url, auth=("", pat), headers=headers, timeout=10)
    response.raise_for_status()
    
    work_items = response.json()["workItems"]
    ids = [item["id"] for item in work_items]
    print(f"Found {len(ids)} work items")
    return ids

def get_work_items_batch(work_item_ids, batch_size=100):
    """Get details for multiple work items in batches - MUCH more efficient"""
    all_items = {}
    
    for i in range(0, len(work_item_ids), batch_size):
        batch = work_item_ids[i:i+batch_size]
        id_string = ",".join(str(id) for id in batch)
        
        url = f"https://dev.azure.com/{org}/{project}/_apis/wit/workItems?ids={id_string}&fields=System.AreaPath,Custom.ExternalRefID&api-version=7.0"
        
        try:
            # Apply rate limit delay before each batch
            global rate_limit_delay
            if rate_limit_delay > 0:
                time.sleep(rate_limit_delay)
            
            response = session.get(url, auth=("", pat), headers=headers, timeout=15)
            
            # Handle rate limiting
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 5))
                with rate_limit_lock:
                    rate_limit_delay = retry_after
                print(f"⚠️  Rate limited! Waiting {retry_after} seconds...")
                time.sleep(retry_after)
                # Retry the batch
                response = session.get(url, auth=("", pat), headers=headers, timeout=15)
            
            response.raise_for_status()
            items = response.json().get("value", [])
            for item in items:
                all_items[item["id"]] = item
        except Exception as e:
            print(f"❌ Error fetching batch {i//batch_size + 1}: {e}")
            continue
    
    return all_items

def update_work_items_batch(updates_list, batch_size=50):
    """Update multiple work items in batches - more efficient than one-by-one"""
    results = {}
    
    for i in range(0, len(updates_list), batch_size):
        batch = updates_list[i:i+batch_size]
        
        # Apply rate limit delay
        global rate_limit_delay
        if rate_limit_delay > 0:
            time.sleep(rate_limit_delay)
        
        # Prepare batch update request
        for wi_id, ext_ref_id in batch:
            url = f"https://dev.azure.com/{org}/{project}/_apis/wit/workItems/{wi_id}?api-version=7.0"
            
            patch_document = [
                {
                    "op": "add",
                    "path": "/fields/Custom.ExternalRefID",
                    "value": ext_ref_id
                }
            ]
            
            try:
                response = session.patch(
                    url,
                    auth=("", pat),
                    headers={"Content-Type": "application/json-patch+json"},
                    json=patch_document,
                    timeout=10
                )
                
                # Handle rate limiting
                if response.status_code == 429:
                    retry_after = int(response.headers.get('Retry-After', 5))
                    with rate_limit_lock:
                        rate_limit_delay = retry_after
                    time.sleep(retry_after)
                    # Retry
                    response = session.patch(
                        url,
                        auth=("", pat),
                        headers={"Content-Type": "application/json-patch+json"},
                        json=patch_document,
                        timeout=10
                    )
                
                if response.status_code == 200:
                    results[wi_id] = (True, "Updated")
                else:
                    results[wi_id] = (False, response.text[:200])
                    
            except Exception as e:
                results[wi_id] = (False, str(e)[:200])
    
    return results

def process_and_update_work_items(query_id):
    """Main process: Fetch, match, and update work items - OPTIMIZED"""
    
    # Step 1: Get work item IDs from query
    api_url = f"https://dev.azure.com/{org}/{project}/_apis/wit/wiql/{query_id}?api-version=7.0"
    print(f"\nFetching work items from query: {query_id}")
    
    try:
        response = session.get(api_url, auth=("", pat), headers=headers, timeout=10)
        response.raise_for_status()
        
        work_items = response.json()["workItems"]
        work_item_ids = [item["id"] for item in work_items]
        print(f"✅ Found {len(work_item_ids)} work items")
    except Exception as e:
        print(f"❌ Failed to fetch work items: {e}")
        return 0, 0, 0
    
    if not work_item_ids:
        print("No work items found in query")
        return 0, 0, 0
    
    # Step 2: Fetch ALL work item details using batch API - MUCH more efficient!
    print(f"\n📥 Fetching details for all {len(work_item_ids)} work items (batch method)...")
    work_items_data = get_work_items_batch(work_item_ids, batch_size=100)
    print(f"✅ Fetched details for {len(work_items_data)} work items")
    
    # Step 3: Match with mapping and prepare updates
    print(f"\n🔍 Matching work items with POC mapping...")
    updates_needed = []
    results = []
    
    for wi_id in work_item_ids:
        wi_details = work_items_data.get(wi_id)
        
        if not wi_details:
            results.append({
                "ID": wi_id,
                "Node Name": "N/A",
                "ExternalRef ID": "N/A",
                "Status": "Failed to fetch details"
            })
            continue
        
        fields = wi_details.get("fields", {})
        
        # Get Node Name from Area Path (last part)
        area_path = fields.get("System.AreaPath", "")
        node_name = area_path.split('\\')[-1] if area_path else ""
        node_name_clean = node_name.strip().lower() if node_name else ""
        
        # Get ExternalRef ID from mapping
        ext_ref_id = node_to_extref.get(node_name_clean, "")
        
        if not ext_ref_id:
            results.append({
                "ID": wi_id,
                "Node Name": node_name,
                "ExternalRef ID": "N/A",
                "Status": "No mapping found"
            })
            continue
        
        # Check if already has this value
        current_extref = fields.get("Custom.ExternalRefID", "")
        if current_extref == ext_ref_id:
            results.append({
                "ID": wi_id,
                "Node Name": node_name,
                "ExternalRef ID": ext_ref_id,
                "Status": "Already updated"
            })
            continue
        
        # Need to update
        updates_needed.append((wi_id, ext_ref_id, node_name))
    
    print(f"✅ Matching complete:")
    print(f"   - Need update: {len(updates_needed)}")
    print(f"   - Already updated: {len([r for r in results if r['Status'] == 'Already updated'])}")
    print(f"   - No mapping found: {len([r for r in results if r['Status'] == 'No mapping found'])}")
    print(f"   - Failed to fetch: {len([r for r in results if r['Status'] == 'Failed to fetch details'])}")
    
    # Step 4: Update work items with controlled concurrency and rate limiting
    if updates_needed:
        print(f"\n📤 Updating {len(updates_needed)} work items...")
        
        # Process updates in smaller batches with light threading (10 workers max)
        update_batches = []
        for i in range(0, len(updates_needed), 5):
            batch = updates_needed[i:i+5]
            update_batches.append(batch)
        
        max_workers = min(10, len(update_batches))  # Max 10 concurrent batches
        update_results = {}
        processed = [0]
        processed_lock = threading.Lock()
        
        def process_update_batch(batch):
            batch_results = {}
            for wi_id, ext_ref_id, node_name in batch:
                url = f"https://dev.azure.com/{org}/{project}/_apis/wit/workItems/{wi_id}?api-version=7.0"
                
                patch_document = [
                    {
                        "op": "add",
                        "path": "/fields/Custom.ExternalRefID",
                        "value": ext_ref_id
                    }
                ]
                
                try:
                    global rate_limit_delay
                    if rate_limit_delay > 0:
                        time.sleep(rate_limit_delay)
                    
                    response = session.patch(
                        url,
                        auth=("", pat),
                        headers={"Content-Type": "application/json-patch+json"},
                        json=patch_document,
                        timeout=10
                    )
                    
                    if response.status_code == 429:
                        retry_after = int(response.headers.get('Retry-After', 5))
                        with rate_limit_lock:
                            rate_limit_delay = retry_after
                        time.sleep(retry_after)
                        response = session.patch(
                            url,
                            auth=("", pat),
                            headers={"Content-Type": "application/json-patch+json"},
                            json=patch_document,
                            timeout=10
                        )
                    
                    if response.status_code == 200:
                        batch_results[wi_id] = (True, ext_ref_id, node_name)
                    else:
                        batch_results[wi_id] = (False, ext_ref_id, node_name, response.text[:100])
                        
                except Exception as e:
                    batch_results[wi_id] = (False, ext_ref_id, node_name, str(e)[:100])
                
                with processed_lock:
                    processed[0] += 1
                    if processed[0] % 50 == 0:
                        print(f"   Progress: {processed[0]}/{len(updates_needed)} items updated...")
            
            return batch_results
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(process_update_batch, batch): idx for idx, batch in enumerate(update_batches)}
            
            for future in as_completed(futures):
                batch_results = future.result()
                update_results.update(batch_results)
        
        # Process update results
        for wi_id in [u[0] for u in updates_needed]:
            if wi_id in update_results:
                result = update_results[wi_id]
                if result[0]:  # Success
                    results.append({
                        "ID": wi_id,
                        "Node Name": result[2],
                        "ExternalRef ID": result[1],
                        "Status": "Updated"
                    })
                else:  # Failed
                    error_msg = result[3] if len(result) > 3 else "Unknown error"
                    results.append({
                        "ID": wi_id,
                        "Node Name": result[2],
                        "ExternalRef ID": result[1],
                        "Status": f"Error: {error_msg}"
                    })
    
    # Step 5: Print summary
    print("\n" + "="*70)
    print("FINAL SUMMARY")
    print("="*70)
    print(f"Total work items processed: {len(work_item_ids)}")
    print(f"✅ Updated: {len([r for r in results if r['Status'] == 'Updated'])}")
    print(f"⏭️  Already updated: {len([r for r in results if r['Status'] == 'Already updated'])}")
    print(f"⏭️  Skipped (no mapping): {len([r for r in results if r['Status'] == 'No mapping found'])}")
    print(f"❌ Failed: {len([r for r in results if 'Error' in r['Status'] or r['Status'] == 'Failed to fetch details'])}")
    print("="*70)
    
    # Save results to CSV
    results_df = pd.DataFrame(results)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = r"C:\Users\vishnu.ramalingam\MyISP_Tools\M_POC"
    output_file = f"{output_dir}\ADO_Update_Results_{timestamp}.csv"
    results_df.to_csv(output_file, index=False)
    print(f"\n💾 Results saved to: {output_file}")
    
    return len([r for r in results if r['Status'] == 'Updated']), len([r for r in results if r['Status'] == 'No mapping found']), len([r for r in results if 'Error' in r['Status'] or r['Status'] == 'Failed to fetch details'])

# Main execution
if __name__ == "__main__":
    print("="*70)
    print("ADO WORK ITEM BATCH UPDATE - EXTERNALREF ID")
    print("="*70)
    
    # Get query ID from command line argument, configured default, or ask user
    if len(sys.argv) > 1:
        # Query ID passed as command line argument (from web app)
        query_id = sys.argv[1]
        print(f"\nUsing query ID from command line: {query_id}")
    elif query_id:
        # Use the configured default query ID
        print(f"\nUsing configured query ID: {query_id}")
    else:
        # Ask user for query ID
        query_id = input("\nEnter ADO Query ID: ").strip()
        
        if not query_id:
            print("❌ Query ID is required!")
            exit(1)
        
        print(f"Using query ID: {query_id}")
    
    try:
        updated, skipped, failed = process_and_update_work_items(query_id)
        print("\n✅ Process completed successfully!")
    except Exception as e:
        print(f"\n❌ Error: {e}")
