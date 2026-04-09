"""
Azure DevOps Test Plan Execution Report Generator - Custom Format (OPTIMIZED)
"""

import requests
import json
from datetime import datetime
import base64
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import shutil
import os
import sys
import io

# Fix console encoding for Windows to support emojis
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ============================================================================
# CONFIGURATION
# ============================================================================

from dotenv import load_dotenv
load_dotenv('ADO_SECRETS.env')

ADO_CONFIG = {
    'organization': 'accenturecio08',
    'project': 'AutomationProcess_29697',
    'plan_id': '4319862',
    'suite_id': '4319865',  # Regression Execution suite ID
    #'insprint_suite_id': '4358470',  # Insprint Execution suite ID
    'target_suite_name': 'PT Execution',
    #'insprint_suite_name': 'Insprint Execution',
    'pat_token': os.getenv('ADO_PAT_MAIN', ''),
    'max_workers': 10,  # Parallel API calls
}

if not ADO_CONFIG['pat_token']:
    raise ValueError("ADO_PAT_MAIN not found in ADO_SECRETS.env file")

# ============================================================================
# AZURE DEVOPS API CLIENT (OPTIMIZED)
# ============================================================================

class AzureDevOpsClient:
    def __init__(self, config):
        self.org = config['organization']
        self.project = config['project']
        self.plan_id = config['plan_id']
        self.suite_id = config['suite_id']
        self.target_suite_name = config.get('target_suite_name', 'Regression Execution')
        self.pat = config['pat_token']
        self.max_workers = config.get('max_workers', 10)
        
        # Encode PAT for Basic Auth
        auth_bytes = f":{self.pat}".encode('ascii')
        encoded_pat = base64.b64encode(auth_bytes).decode('ascii')
        
        self.headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Basic {encoded_pat}',
            'Accept': 'application/json'
        }
        
        self.session = requests.Session()
        self.session.headers.update(self.headers)
        
        # Cache for work items to avoid duplicate fetches
        self.work_item_cache = {}
        self.cache_lock = threading.Lock()
    
    def test_connection(self):
        """Test basic connection"""
        try:
            url = f"https://dev.azure.com/{self.org}/_apis/projects/{self.project}?api-version=7.0"
            print(f"\n🔍 Testing Connection...")
            
            response = self.session.get(url, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                print(f"   ✅ Connected to: {data.get('name', 'Unknown')}")
                return True
            else:
                print(f"   ❌ Failed: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"   ❌ Error: {e}")
            return False
    
    def get_test_plan(self):
        """Fetch test plan details"""
        try:
            url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/testplan/plans/{self.plan_id}?api-version=7.0"
            
            print(f"\n📋 Fetching Test Plan {self.plan_id}...")
            
            response = self.session.get(url, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                print(f"   ✅ Plan: {data.get('name', 'N/A')}")
                return data
            else:
                url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/test/plans/{self.plan_id}?api-version=7.0"
                response = self.session.get(url, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    print(f"   ✅ Plan: {data.get('name', 'N/A')}")
                    return data
                    
                print(f"   ❌ Failed: {response.status_code}")
                return None
                
        except Exception as e:
            print(f"   ❌ Error: {e}")
            return None
    
    def get_all_suites_in_plan(self):
        """Get all suites in the plan with pagination support"""
        try:
            url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/testplan/plans/{self.plan_id}/suites?api-version=7.0"
            
            print(f"\n📦 Fetching All Suites...")
            
            all_suites = []
            continuation_token = None
            
            while True:
                request_url = url
                if continuation_token:
                    request_url = f"{url}&continuationToken={continuation_token}"
                
                response = self.session.get(request_url, timeout=30)
                
                if response.status_code != 200:
                    url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/test/plans/{self.plan_id}/suites?api-version=7.0"
                    response = self.session.get(url, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    suites = data.get('value', [])
                    all_suites.extend(suites)
                    
                    continuation_token = response.headers.get('x-ms-continuationtoken')
                    if not continuation_token:
                        break
                else:
                    break
            
            print(f"   ✅ Found {len(all_suites)} suites")
            return all_suites
                
        except Exception as e:
            print(f"   ❌ Error: {e}")
            return []
    
    def verify_suite_exists(self, suite_id):
        """Verify if a specific suite exists"""
        try:
            url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/test/plans/{self.plan_id}/suites/{suite_id}?api-version=7.0"
            
            print(f"\n🔍 Verifying Suite {suite_id}...")
            
            response = self.session.get(url, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                suite_name = data.get('name', 'Unknown')
                print(f"   ✅ Suite exists: {suite_name}")
                return data
            else:
                print(f"   ❌ Suite not accessible (Status: {response.status_code})")
                return None
                
        except Exception as e:
            print(f"   ❌ Error: {e}")
            return None
    
    def get_test_points_from_suite(self, suite_id):
        """Fetch test points from a suite"""
        try:
            url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/test/plans/{self.plan_id}/suites/{suite_id}/points?api-version=7.0"
            
            response = self.session.get(url, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                return data.get('value', [])
            
            return []
                
        except Exception as e:
            return []
    
    def get_work_items_batch(self, work_item_ids):
        """Fetch multiple work items in a single API call (OPTIMIZED)"""
        if not work_item_ids:
            return {}
        
        try:
            # Remove duplicates and None values
            unique_ids = list(set([id for id in work_item_ids if id]))
            
            if not unique_ids:
                return {}
            
            # Azure DevOps supports batch requests with up to 200 IDs
            batch_size = 200
            all_work_items = {}
            
            for i in range(0, len(unique_ids), batch_size):
                batch_ids = unique_ids[i:i+batch_size]
                ids_param = ','.join(map(str, batch_ids))
                
                url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/wit/workitems?ids={ids_param}&api-version=7.0"
                
                response = self.session.get(url, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    work_items = data.get('value', [])
                    
                    for wi in work_items:
                        wi_id = wi.get('id')
                        all_work_items[wi_id] = wi
            
            return all_work_items
                
        except Exception as e:
            print(f"      ⚠️  Batch fetch error: {e}")
            return {}
    
    def get_child_suites_from_cache(self, parent_suite_id, all_suites):
        """Get child suites from cached all_suites list (OPTIMIZED)"""
        try:
            child_suites = []
            parent_suite_id_int = int(parent_suite_id)
            
            for suite in all_suites:
                parent = suite.get('parent', {}) or suite.get('parentSuite', {})
                parent_id = None
                
                if isinstance(parent, dict):
                    parent_id = parent.get('id')
                elif isinstance(parent, (int, str)):
                    try:
                        parent_id = int(parent)
                    except:
                        pass
                
                try:
                    parent_id_int = int(parent_id) if parent_id else None
                except:
                    parent_id_int = parent_id
                
                if parent_id_int == parent_suite_id_int:
                    child_suites.append(suite)
            
            return child_suites
                
        except Exception as e:
            return []
    
    def _build_suite_tree(self, root_suite_id, root_suite_name, all_suites):
        """Build complete suite tree with metadata"""
        suite_tree = []
        
        def traverse_suite(suite_id, suite_name, parent_lead=None, parent_module=None, test_type=None, depth=0):
            """
            Traverse suite hierarchy:
            - Depth 0: Root (Regression Execution)
            - Depth 1: Lead folders (Kavi, Pirtheebaa, etc.)
            - Depth 2: Module folders (SI OCP, SI DCTA, etc.)
            - Depth 3: Test Type folders (Automation, Manual)
            - Depth 4+: Deeper nesting
            """
            
            # Determine current level type
            current_lead = parent_lead
            current_module = parent_module
            current_test_type = test_type
            
            if depth == 1:
                # This is a Lead folder
                current_lead = suite_name
                current_module = None
                current_test_type = None
            elif depth == 2:
                # This is a Module folder
                current_module = suite_name
                current_test_type = None
            elif depth == 3:
                # This is Test Type folder (Automation or Manual)
                suite_name_lower = suite_name.lower()
                if 'automation' in suite_name_lower:
                    current_test_type = 'Automation'
                elif 'manual' in suite_name_lower:
                    current_test_type = 'Manual'    
                else:
                    # If not explicitly named, keep parent test type
                    current_test_type = test_type
            # depth >= 4: keep parent values
            
            suite_tree.append({
                'id': suite_id,
                'name': suite_name,
                'parent_lead': current_lead,
                'parent_module': current_module,
                'test_type': current_test_type,
                'depth': depth
            })
            
            # Get children
            children = self.get_child_suites_from_cache(suite_id, all_suites)
            
            for child in children:
                child_id = child.get('id')
                child_name = child.get('name', 'Unknown')
                
                traverse_suite(child_id, child_name, 
                             parent_lead=current_lead, 
                             parent_module=current_module,
                             test_type=current_test_type,
                             depth=depth+1)
        
        traverse_suite(root_suite_id, root_suite_name)
        return suite_tree
    
    def _collect_test_points_parallel(self, suite_tree):
        """Collect test points from all suites in parallel"""
        all_test_items = []
        
        def fetch_suite_tests(suite_info):
            suite_id = suite_info['id']
            suite_name = suite_info['name']
            parent_lead = suite_info['parent_lead']
            parent_module = suite_info['parent_module']
            test_type = suite_info['test_type']
            
            test_items = self.get_test_points_from_suite(suite_id)
            
            result = []
            for item in test_items:
                # Pass lead, module, and test type information
                result.append((item, suite_name, parent_lead, parent_module, test_type))
            
            return result
        
        # Use thread pool for parallel fetching
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(fetch_suite_tests, suite): suite for suite in suite_tree}
            
            completed = 0
            total = len(suite_tree)
            
            for future in as_completed(futures):
                completed += 1
                if completed % 10 == 0 or completed == total:
                    print(f"      Progress: {completed}/{total} suites processed", end='\r')
                
                try:
                    items = future.result()
                    all_test_items.extend(items)
                except Exception as e:
                    pass
        
        print()  # New line after progress
        return all_test_items

    def get_all_test_data_from_suite(self, suite_id, suite_name):
        """Get all test data from suite and children (OPTIMIZED)"""
        print(f"\n📊 Collecting Test Data from: {suite_name}...")
        print(f"   ⚡ Using optimized batch processing...")
        
        # Fetch all suites once
        print(f"\n   Step 1: Fetching suite hierarchy...")
        all_suites = self.get_all_suites_in_plan()
        
        # Build suite hierarchy
        print(f"   Step 2: Building suite tree...")
        suite_tree = self._build_suite_tree(suite_id, suite_name, all_suites)
        
        # Collect all test points in parallel
        print(f"   Step 3: Collecting test points from {len(suite_tree)} suites...")
        all_test_items = self._collect_test_points_parallel(suite_tree)
        
        if not all_test_items:
            print(f"   ⚠️  No test points found")
            return []
        
        print(f"   ✓ Found {len(all_test_items)} test items")
        
        # Extract unique work item IDs
        print(f"   Step 4: Extracting work item IDs...")
        work_item_ids = []
        for item_data in all_test_items:
            # Unpack the tuple with 5 elements now
            item, item_suite_name, parent_lead, parent_module, test_type = item_data
            
            test_case = item.get('testCase', {})
            work_item_id = test_case.get('id') or item.get('testCaseReference', {}).get('id')
            if work_item_id:
                work_item_ids.append(work_item_id)
        
        # Fetch all work items in batch
        print(f"   Step 5: Fetching {len(set(work_item_ids))} unique work items in batch...")
        work_items_dict = self.get_work_items_batch(work_item_ids)
        print(f"   ✓ Retrieved {len(work_items_dict)} work items")
        
        # Process test data
        print(f"   Step 6: Processing test data...")
        all_test_data = []
        
        for item_data in all_test_items:
            # Unpack the tuple with 5 elements now
            item, item_suite_name, parent_lead, parent_module, suite_test_type = item_data
            
            test_case = item.get('testCase', {})
            work_item_id = test_case.get('id') or item.get('testCaseReference', {}).get('id')
            
            # Determine test type based on suite folder structure
            if suite_test_type:
                # Use test type from suite hierarchy (Automation or Manual folder)
                test_type = suite_test_type
            else:
                # Fallback: Check work item tags/automation status
                work_item = work_items_dict.get(work_item_id) if work_item_id else None
                test_type = 'Manual'  # Default
                
                if work_item:
                    fields = work_item.get('fields', {})
                    tags = fields.get('System.Tags', '')
                    automation_status = fields.get('Microsoft.VSTS.TCM.AutomationStatus', '')
                    
                    if 'automation' in str(tags).lower() or 'automated' in str(automation_status).lower():
                        test_type = 'Automation'
            
            # Get outcome
            outcome = 'Not Run'
            if 'results' in item:
                outcome = item['results'].get('outcome', 'Not Run')
            elif 'lastResultOutcome' in item:
                outcome = item.get('lastResultOutcome', 'Not Run')
            elif 'outcome' in item:
                outcome = item.get('outcome', 'Not Run')
            
            # Get assigned to
            assigned_to = 'Unassigned'
            if 'assignedTo' in item and item['assignedTo']:
                assigned_to = item['assignedTo'].get('displayName', 'Unassigned')
            
            # Determine final lead and module names
            final_lead = parent_lead if parent_lead else 'Unassigned'
            final_module = parent_module if parent_module else item_suite_name
            
            # Override with assigned user's first name if available (only if no parent_lead)
            if assigned_to and assigned_to != 'Unassigned' and not parent_lead:
                try:
                    final_lead = assigned_to.split()[0]
                except:
                    final_lead = assigned_to
            
            all_test_data.append({
                'id': work_item_id or 'N/A',
                'name': test_case.get('name', 'N/A'),
                'suite': item_suite_name,
                'module': final_module,  # Use parent_module from hierarchy
                'lead': final_lead,      # Use parent_lead from hierarchy
                'type': test_type,       # Use test type from suite folder structure
                'outcome': outcome,
                'priority': item.get('priority', 2),
                'state': item.get('state', 'Active')
            })
        
        print(f"\n   ✅ Total test items collected: {len(all_test_data)}")
        return all_test_data
    
    def get_bugs_from_query(self, query_id):
        """Fetch bugs from ADO query"""
        try:
            url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/wit/wiql/{query_id}?api-version=7.0"
            
            print(f"\n🐛 Fetching Bugs from Query {query_id}...")
            
            response = self.session.get(url, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                work_items = data.get('workItems', [])
                
                if not work_items:
                    print(f"   ⚠️  No bugs found in query")
                    return []
                
                # Extract work item IDs
                bug_ids = [wi['id'] for wi in work_items]
                print(f"   ✓ Found {len(bug_ids)} bugs in query")
                
                # Fetch bug details in batch
                print(f"   Step 2: Fetching bug details...")
                bugs_dict = self.get_work_items_batch(bug_ids)
                
                # Process bug data
                print(f"   Step 3: Processing bug data...")
                bug_data = []
                
                for bug_id, bug_wi in bugs_dict.items():
                    fields = bug_wi.get('fields', {})
                    
                    # Get bug details
                    # Try different possible field names for ExternalRef ID
                    external_ref = (
                        fields.get('Custom.ExternalRefID') or 
                        fields.get('Custom.ExternalRegID') or
                        fields.get('Custom.ExternalRef') or
                        fields.get('ExternalRefID') or
                        fields.get('ExternalRegID') or
                        'Unassigned'
                    )
                    
                    # Extract Node Name from Area Path (last segment after \)
                    area_path = fields.get('System.AreaPath', 'N/A')
                    if area_path != 'N/A' and '\\' in area_path:
                        node_name = area_path.split('\\')[-1]
                    else:
                        node_name = area_path
                    
                    bug_info = {
                        'id': bug_id,
                        'title': fields.get('System.Title', 'N/A'),
                        'state': fields.get('System.State', 'N/A'),
                        'severity': fields.get('Microsoft.VSTS.Common.Severity', 'N/A'),
                        'priority': fields.get('Microsoft.VSTS.Common.Priority', 'N/A'),
                        'assigned_to': fields.get('System.AssignedTo', {}).get('displayName', 'Unassigned') if isinstance(fields.get('System.AssignedTo'), dict) else 'Unassigned',
                        'created_date': fields.get('System.CreatedDate', 'N/A'),
                        'mpoc': external_ref,  # ExternalRef ID as MPOC
                        'area_path': area_path,
                        'tags': fields.get('System.Tags', ''),
                        'text_verification': fields.get('Custom.TextVerification', 'N/A'),
                        'defect_record': fields.get('Custom.DefectRecord', 'N/A'),
                        'node_name': node_name,
                        'eta': fields.get('Custom.ETA', 'N/A'),
                        'stage_found': fields.get('Custom.StageFound', 'N/A'),
                        'text_verification1': fields.get('Custom.TextVerification1', 'N/A')
                    }
                    
                    bug_data.append(bug_info)
                
                print(f"   ✅ Processed {len(bug_data)} bugs")
                return bug_data
            else:
                print(f"   ❌ Failed: {response.status_code}")
                return []
                
        except Exception as e:
            print(f"   ❌ Error fetching bugs: {e}")
            return []
    
    def get_defects_by_tag_and_date(self, tags, created_after_date):
        """
        Fetch defects (bugs) based on tags and creation date
        
        Args:
            tags: List of tags to filter by (e.g., ['Insprint_Regression'])
            created_after_date: Date string in format 'YYYY-MM-DD' (e.g., '2026-02-12')
        
        Returns:
            List of defect/bug information dictionaries
        """
        try:
            # Convert tags to a list if it's a string
            if isinstance(tags, str):
                tags = [tags]
            
            # Build WIQL query
            # Tags in Azure DevOps are stored as a semicolon-separated string
            # We need to check if any of the tags exists in the System.Tags field
            tag_conditions = " OR ".join([f"[System.Tags] CONTAINS '{tag}'" for tag in tags])
            
            wiql_query = f"""
            SELECT [System.Id], [System.Title], [System.State], [System.CreatedDate]
            FROM WorkItems
            WHERE [System.TeamProject] = @project
                AND [System.WorkItemType] = 'Bug'
                AND ({tag_conditions})
                AND [System.CreatedDate] >= '{created_after_date}'
            ORDER BY [System.CreatedDate] DESC
            """
            
            url = f"https://dev.azure.com/{self.org}/{self.project}/_apis/wit/wiql?api-version=7.0"
            
            print(f"\n🐛 Fetching Defects with tags {tags} created after {created_after_date}...")
            
            response = self.session.post(
                url,
                json={"query": wiql_query},
                timeout=30
            )
            
            if response.status_code == 200:
                data = response.json()
                work_items = data.get('workItems', [])
                
                if not work_items:
                    print(f"   ⚠️  No defects found matching criteria")
                    return []
                
                # Extract work item IDs
                defect_ids = [wi['id'] for wi in work_items]
                print(f"   ✓ Found {len(defect_ids)} defects matching criteria")
                
                # Fetch defect details in batch
                print(f"   Step 2: Fetching defect details...")
                defects_dict = self.get_work_items_batch(defect_ids)
                
                # Process defect data
                print(f"   Step 3: Processing defect data...")
                defect_data = []
                
                for defect_id, defect_wi in defects_dict.items():
                    fields = defect_wi.get('fields', {})
                    
                    # Get defect details
                    # Try different possible field names for ExternalRef ID
                    external_ref = (
                        fields.get('Custom.ExternalRefID') or 
                        fields.get('Custom.ExternalRegID') or
                        fields.get('Custom.ExternalRef') or
                        fields.get('ExternalRefID') or
                        fields.get('ExternalRegID') or
                        'Unassigned'
                    )
                    
                    # Extract Node Name from Area Path (last segment after \)
                    area_path = fields.get('System.AreaPath', 'N/A')
                    if area_path != 'N/A' and '\\' in area_path:
                        node_name = area_path.split('\\')[-1]
                    else:
                        node_name = area_path
                    
                    defect_info = {
                        'id': defect_id,
                        'title': fields.get('System.Title', 'N/A'),
                        'state': fields.get('System.State', 'N/A'),
                        'severity': fields.get('Microsoft.VSTS.Common.Severity', 'N/A'),
                        'priority': fields.get('Microsoft.VSTS.Common.Priority', 'N/A'),
                        'assigned_to': fields.get('System.AssignedTo', {}).get('displayName', 'Unassigned') if isinstance(fields.get('System.AssignedTo'), dict) else 'Unassigned',
                        'created_date': fields.get('System.CreatedDate', 'N/A'),
                        'mpoc': external_ref,  # ExternalRef ID as MPOC
                        'area_path': area_path,
                        'tags': fields.get('System.Tags', ''),
                        'text_verification': fields.get('Custom.TextVerification', 'N/A'),
                        'defect_record': fields.get('Custom.DefectRecord', 'N/A'),
                        'node_name': node_name,
                        'eta': fields.get('Custom.ETA', 'N/A'),
                        'stage_found': fields.get('Custom.StageFound', 'N/A'),
                        'text_verification1': fields.get('Custom.TextVerification1', 'N/A')
                    }
                    
                    defect_data.append(defect_info)
                
                print(f"   ✅ Processed {len(defect_data)} defects")
                return defect_data
            else:
                print(f"   ❌ Failed: {response.status_code}")
                try:
                    error_detail = response.json()
                    print(f"   ❌ Error details: {error_detail}")
                except:
                    print(f"   ❌ Response: {response.text[:200]}")
                return []
                
        except Exception as e:
            print(f"   ❌ Error fetching defects: {e}")
            import traceback
            traceback.print_exc()
            return []

# ============================================================================
# HTML REPORT GENERATOR - CUSTOM FORMAT
# ============================================================================

class CustomHTMLReportGenerator:
    def __init__(self, test_data, plan_info=None, suite_name=None, bug_data=None, insprint_data=None, insprint_defects=None):
        self.test_data = test_data
        self.insprint_data = insprint_data or []
        self.plan_info = plan_info or {}
        self.suite_name = suite_name or 'Test Suite'
        self.timestamp = datetime.now().strftime("%B %d, %Y at %H:%M:%S")
        self.bug_data = bug_data or []
        self.insprint_defects = insprint_defects or []
    
    def organize_data_by_lead_module(self):
        """Organize test data by Lead and Module"""
        organized = defaultdict(lambda: defaultdict(lambda: {
            'manual': {'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0},
            'automation': {'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0}
        }))
        
        for test in self.test_data:
            lead = test['lead']
            module = test['module']
            test_type = test['type'].lower()
            outcome = test['outcome']
            
            # Increment total count
            organized[lead][module][test_type]['total'] += 1
            
            # Map outcome to correct category (case-insensitive matching)
            outcome_lower = outcome.lower()
            
            if outcome_lower in ['passed', 'pass']:
                organized[lead][module][test_type]['passed'] += 1
            elif outcome_lower in ['failed', 'fail']:
                organized[lead][module][test_type]['failed'] += 1
            elif outcome_lower in ['blocked', 'block']:
                organized[lead][module][test_type]['blocked'] += 1
            elif outcome_lower in ['not applicable', 'na', 'n/a', 'notapplicable']:
                organized[lead][module][test_type]['na'] += 1
            elif outcome_lower in ['not run', 'notrun', 'active', 'none', '']:
                organized[lead][module][test_type]['not_run'] += 1
            else:
                # Any unrecognized outcome goes to 'not_run'
                organized[lead][module][test_type]['not_run'] += 1
        
        return organized
    
    def calculate_percentages(self, data):
        """Calculate Pass% and Execution% for individual Lead/Module rows"""
        passed = data['passed']
        failed = data['failed']
        blocked = data['blocked']
        total = data['total']
        na = data['na']
        
        # Pass % = (Pass / (Pass + Fail + Blocked)) * 100
        denominator_pass = passed + failed + blocked
        pass_percentage = (passed / denominator_pass * 100) if denominator_pass > 0 else 0
        
        # Execution % = (Pass + Fail) / (Total - NA) * 100
        denominator_exec = total - na
        execution_percentage = ((passed + failed) / denominator_exec * 100) if denominator_exec > 0 else 0
        
        return pass_percentage, execution_percentage
    
    def calculate_grand_total_percentages(self, data):
        """Calculate Pass% and Execution% for Grand Total row (different formula)"""
        passed = data['passed']
        failed = data['failed']
        blocked = data['blocked']
        total = data['total']
        na = data['na']
        
        # Pass % = (Pass / (Pass + Fail + Blocked)) * 100
        denominator_pass = passed + failed + blocked
        pass_percentage = (passed / denominator_pass * 100) if denominator_pass > 0 else 0
        
        # Execution % = (Pass + Fail + Blocked) / (Total - NA) * 100
        numerator_exec = passed + failed + blocked
        denominator_exec = total - na
        execution_percentage = (numerator_exec / denominator_exec * 100) if denominator_exec > 0 else 0
        
        return pass_percentage, execution_percentage
    
    def calculate_grand_totals(self, organized_data):
        """Calculate grand totals"""
        grand_totals = {
            'manual': {'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0},
            'automation': {'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0}
        }
        
        for lead_data in organized_data.values():
            for module_data in lead_data.values():
                for test_type in ['manual', 'automation']:
                    for key in grand_totals[test_type]:
                        grand_totals[test_type][key] += module_data[test_type][key]
        
        return grand_totals
    
    def process_bug_data_by_mpoc(self):
        """Process bug data and group by MPOC with severity breakdown - includes multiple bug states"""
        bug_summary = defaultdict(lambda: {
            '1 - Critical': 0,
            '2 - High': 0,
            '3 - Medium': 0,
            '4 - Low': 0
        })
        
        # Define allowed bug states (case-insensitive)
        allowed_states = {
            'new', 'active', 'blocked', 'ready to deploy', 'resolved', 
            'ba clarification', 're-open', 'blocked in pt', 'blocked in uat', 'deferred'
        }
        
        # Filter bugs by allowed states
        filtered_bugs = [bug for bug in self.bug_data if bug['state'].lower() in allowed_states]
        
        # Track unique MPOC names for case-insensitive grouping
        mpoc_mapping = {}  # Maps lowercase MPOC to original casing
        
        for bug in filtered_bugs:
            raw_mpoc = bug['mpoc'] if bug['mpoc'] and bug['mpoc'] != 'Unassigned' else 'Unassigned'
            
            # Normalize MPOC name to avoid case duplicates
            mpoc_lower = raw_mpoc.lower()
            
            # Use the first occurrence's casing as the canonical form
            if mpoc_lower not in mpoc_mapping:
                mpoc_mapping[mpoc_lower] = raw_mpoc
            
            mpoc = mpoc_mapping[mpoc_lower]
            severity = bug['severity']
            
            # Map severity to standardized format
            if severity in bug_summary[mpoc]:
                bug_summary[mpoc][severity] += 1
            else:
                # Handle any non-standard severity values
                bug_summary[mpoc]['3 - Medium'] += 1
        
        return dict(bug_summary)
    
    def organize_data_by_lead_module_insprint(self):
        """Organize insprint test data by Lead and Module"""
        organized = defaultdict(lambda: defaultdict(lambda: {
            'manual': {'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0},
            'automation': {'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0}
        }))
        
        for test in self.insprint_data:
            lead = test['lead']
            module = test['module']
            test_type = test['type'].lower()
            outcome = test['outcome']
            
            # Increment total count
            organized[lead][module][test_type]['total'] += 1
            
            # Map outcome to correct category (case-insensitive matching)
            outcome_lower = outcome.lower()
            
            if outcome_lower in ['passed', 'pass']:
                organized[lead][module][test_type]['passed'] += 1
            elif outcome_lower in ['failed', 'fail']:
                organized[lead][module][test_type]['failed'] += 1
            elif outcome_lower in ['blocked', 'block']:
                organized[lead][module][test_type]['blocked'] += 1
            elif outcome_lower in ['not applicable', 'na', 'n/a', 'notapplicable']:
                organized[lead][module][test_type]['na'] += 1
            elif outcome_lower in ['not run', 'notrun', 'active', 'none', '']:
                organized[lead][module][test_type]['not_run'] += 1
            else:
                # Any unrecognized outcome goes to 'not_run'
                organized[lead][module][test_type]['not_run'] += 1
        
        return organized
    
    def generate_html(self):
        """Generate HTML report - Compact Design with Filters and Tabs"""
        organized_data = self.organize_data_by_lead_module()
        grand_totals = self.calculate_grand_totals(organized_data)
        manual_gt = grand_totals['manual']
        auto_gt = grand_totals['automation']
        gt_pass_pct, gt_exec_pct = self.calculate_grand_total_percentages(manual_gt)
        auto_gt_pass_pct, auto_gt_exec_pct = self.calculate_grand_total_percentages(auto_gt)
        # Calculate leads-wise manual summary
        leads_summary = self.calculate_leads_summary(organized_data)
        # Calculate leads-wise automation summary
        automation_leads_summary = self.calculate_automation_leads_summary(organized_data)
        # Calculate insprint data for summary
        insprint_organized = self.organize_data_by_lead_module_insprint()
        insprint_grand_totals = self.calculate_grand_totals(insprint_organized)
        insprint_manual_gt = insprint_grand_totals['manual']
        insprint_gt_pass_pct, insprint_gt_exec_pct = self.calculate_grand_total_percentages(insprint_manual_gt)
        plan_name = self.plan_info.get('name', f"Test Plan {ADO_CONFIG['plan_id']}")
        
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Test Execution Report - {self.suite_name}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #f0f2f5;
            padding: 10px;
        }}
        .container {{
            max-width: 100%;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 15px;
            text-align: center;
        }}
        .header h1 {{ font-size: 18px; margin-bottom: 4px; }}
        .header p {{ font-size: 11px; opacity: 0.9; margin: 2px 0; }}
        
        .tabs {{
            display: flex;
            background: #e9ecef;
            border-bottom: 2px solid #dee2e6;
        }}
        
        .tab {{
            padding: 10px 20px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            border: none;
            background: transparent;
            color: #495057;
            transition: all 0.3s;
            border-bottom: 3px solid transparent;
        }}
        
        .tab:hover {{
            background: #dee2e6;
        }}
        
        .tab.active {{
            background: white;
            color: #667eea;
            border-bottom: 3px solid #667eea;
        }}
        
        .tab-content {{
            display: none;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        .filter-section {{
            padding: 10px 15px;
            background: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
            position: relative;
            overflow: visible;
        }}
        
        .filter-group {{
            display: flex;
            align-items: center;
            gap: 8px;
            position: relative;
        }}
        
        .filter-group label {{
            font-size: 11px;
            font-weight: 600;
            color: #495057;
        }}
        
        .filter-group select {{
            padding: 5px 10px;
            font-size: 11px;
            border: 1px solid #ced4da;
            border-radius: 4px;
            background: white;
            cursor: pointer;
            min-width: 150px;
        }}
        
        .filter-group select:focus {{
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 2px rgba(102, 126, 234, 0.1);
        }}
        
        .reset-btn {{
            padding: 5px 15px;
            font-size: 11px;
            background: #667eea;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 600;
        }}
        
        .reset-btn:hover {{
            background: #5568d3;
        }}
        
        .filter-info {{
            margin-left: auto;
            font-size: 11px;
            color: #6c757d;
            font-weight: 600;
        }}
        
        /* Custom Dropdown Styles */
        .custom-dropdown {{
            position: relative;
            display: inline-block;
            z-index: 100;
        }}
        
        .dropdown-toggle {{
            padding: 5px 10px;
            font-size: 11px;
            border: 1px solid #ced4da;
            border-radius: 4px;
            background: white;
            cursor: pointer;
            min-width: 150px;
            text-align: left;
            font-weight: 600;
        }}
        
        .dropdown-toggle:hover {{
            background: #f8f9fa;
        }}
        
        .dropdown-menu {{
            display: none;
            position: absolute;
            top: 100%;
            left: 0;
            min-width: 200px;
            max-height: 350px;
            overflow-y: auto;
            overflow-x: hidden;
            background: white;
            border: 1px solid #ced4da;
            border-radius: 4px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            z-index: 9999;
            margin-top: 2px;
        }}
        
        .dropdown-menu.show {{
            display: block;
        }}
        
        .dropdown-item {{
            padding: 6px 12px;
        }}
        
        .dropdown-item label {{
            display: flex;
            align-items: center;
            cursor: pointer;
            font-size: 11px;
            margin: 0;
            font-weight: normal;
        }}
        
        .dropdown-item label:hover {{
            background: #f8f9fa;
        }}
        
        .dropdown-item input[type="checkbox"] {{
            margin-right: 8px;
            cursor: pointer;
        }}
        
        .dropdown-divider {{
            height: 1px;
            background: #e9ecef;
            margin: 4px 0;
        }}
        
        .report-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 0;
            font-size: 10px;
        }}
        
        .report-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 6px 4px;
            text-align: center;
            font-size: 9px;
            font-weight: 600;
            border: 2px solid #000 !important;
            line-height: 1.2;
        }}
        
        .report-table th.main-header {{
            background: linear-gradient(135deg, #4299e1 0%, #667eea 100%);
            font-size: 10px;
            padding: 5px 3px;
        }}
        
        .report-table td {{
            padding: 4px 3px;
            text-align: center;
            border: 2px solid #000 !important;
            font-size: 10px;
            line-height: 1.3;
        }}
        
        .report-table tbody tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        
        .report-table tbody tr:hover {{
            background: #e9ecef;
        }}
        
        .report-table tbody tr.hidden {{
            display: none;
        }}
        
        .sno-col {{ width: 30px; font-weight: 600; }}
        .lead-col {{ width: 70px; font-weight: 600; font-size: 9px; }}
        .module-col {{ width: 100px; font-size: 9px; }}
        .total-p1p2-col {{ width: 50px; font-weight: 700; background: #fd7e14 !important; color: white !important; }}
        .total-col {{ width: 40px; font-weight: 600; background: #e7f3ff !important; }}
        .pass-col {{ width: 35px; background: #d4edda !important; }}
        .fail-col {{ width: 35px; background: #f8d7da !important; }}
        .blocked-col {{ width: 35px; background: #fff3cd !important; }}
        .na-col {{ width: 35px; background: #e2e3e5 !important; }}
        .notrun-col {{ width: 40px; background: #cce5ff !important; }}
        .percentage-col {{ width: 45px; font-weight: 600; background: #fff9e6 !important; font-size: 9px; }}
        
        .grand-total-row {{
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
            color: white !important;
            font-weight: 700 !important;
            font-size: 10px !important;
        }}

        /* ========================================================================
           LEADS WISE EXECUTION STATUS TABLE STYLES (TAB 3 & TAB 4)
           Compact Professional Design
           ======================================================================== */
        
        .leads-table {{
            width: auto;
            max-width: 700px;
            margin: 10px auto;
            border-collapse: collapse;
            font-size: 10px;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        
        /* Header Row - Compact Purple Gradient */
        .leads-table thead {{
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        .leads-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
            padding: 5px 6px;
            text-align: center;
            font-size: 9px;
            font-weight: 700;
            border: 2px solid #000 !important;
            line-height: 1.1;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            white-space: nowrap;
        }}
        
        /* Data Cell Base Styles - Compact */
        .leads-table td {{
            padding: 4px 6px;
            text-align: center;
            border: 2px solid #000 !important;
            font-size: 10px;
            font-weight: 400;
            line-height: 1.2;
            color: #333 !important;
        }}
        
        /* Alternating Row Colors - Like Tab 1 */
        .leads-table tbody tr:nth-child(odd) td {{
            background: white !important;
        }}
        
        .leads-table tbody tr:nth-child(even) td {{
            background: #f8f9fa !important;
        }}
        
        /* Hover Effect - Matching Tab 1 */
        .leads-table tbody tr:not(.grand-total-row):hover td {{
            background: #e9ecef !important;
            cursor: pointer;
        }}
        
        /* Lead Name Column - Compact */
        .leads-table .lead-name-col {{
            width: 90px;
            font-weight: 700;
            text-align: left;
            padding-left: 8px;
            font-size: 10px;
            color: #1a1a1a !important;
            background: #e3f2fd !important;
        }}
        
        /* Pass Column - Compact */
        .leads-table .pass-col {{
            width: 40px;
            background: #d4edda !important;
            color: #155724 !important;
            font-weight: 600;
        }}
        
        /* Fail Column - Compact */
        .leads-table .fail-col {{
            width: 40px;
            background: #f8d7da !important;
            color: #721c24 !important;
            font-weight: 600;
        }}
        
        /* Blocked Column - Compact */
        .leads-table .blocked-col {{
            width: 40px;
            background: #fff3cd !important;
            color: #856404 !important;
            font-weight: 600;
        }}
        
        /* NA Column - Compact */
        .leads-table .na-col {{
            width: 40px;
            background: #e2e3e5 !important;
            color: #383d41 !important;
            font-weight: 600;
        }}
        
        /* Not Run Column - Compact */
        .leads-table .notrun-col {{
            width: 45px;
            background: #cce5ff !important;
            color: #004085 !important;
            font-weight: 600;
        }}
        
        /* Total Column - Compact */
        .leads-table .total-col {{
            width: 45px;
            font-weight: 700;
            background: #e7f3ff !important;
            color: #1a1a1a !important;
        }}
        
        /* Percentage Columns - Compact */
        .leads-table .percentage-col {{
            width: 50px;
            font-weight: 700;
            background: #fff9e6 !important;
            color: #1a1a1a !important;
            font-size: 10px;
        }}
        
        /* Grand Total Row - Compact Green Gradient */
        .leads-table .grand-total-row {{
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
            font-weight: 700 !important;
            font-size: 10px !important;
        }}
        
        .leads-table .grand-total-row td {{
            border: 2px solid #000 !important;
            padding: 5px 6px;
            color: white !important;
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
            font-weight: 700 !important;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }}
        
        .leads-table .grand-total-row .lead-name-col {{
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
            color: white !important;
            font-weight: 700 !important;
            text-align: left;
            padding-left: 8px;
            text-transform: uppercase;
        }}
        
        /* Override all status column colors for grand total */
        .leads-table .grand-total-row .pass-col,
        .leads-table .grand-total-row .fail-col,
        .leads-table .grand-total-row .blocked-col,
        .leads-table .grand-total-row .na-col,
        .leads-table .grand-total-row .notrun-col,
        .leads-table .grand-total-row .total-col,
        .leads-table .grand-total-row .percentage-col {{
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
            color: white !important;
            font-weight: 700 !important;
        }}
        
        /* Ensure grand total ignores alternating colors */
        .leads-table tbody tr.grand-total-row:nth-child(even) td,
        .leads-table tbody tr.grand-total-row:nth-child(odd) td {{
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
        }}
        
        /* Grand total hover effect */
        .leads-table tbody tr.grand-total-row:hover td {{
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%) !important;
            cursor: default;
        }}
        
        /* ========================================================================
           SUMMARY TABLE STYLES (TAB 1)
           Compact Professional Design
           ======================================================================== */
        
        .summary-table {{
            width: 100%;
            max-width: 900px;
            margin: 15px auto;
            border-collapse: collapse;
            font-size: 10px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.12);
            border-radius: 6px;
            overflow: hidden;
        }}
        
        /* Summary Table Header */
        .summary-table thead {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }}
        
        .summary-table th {{
            padding: 8px 10px;
            text-align: center;
            font-size: 9px;
            font-weight: 700;
            border: 2px solid #1e293b !important;
            color: white !important;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            line-height: 1.2;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }}
        
        /* Summary Table Data Cells */
        .summary-table td {{
            padding: 10px 8px;
            text-align: center;
            font-size: 11px;
            font-weight: 600;
            border: 2px solid #1e293b !important;
            color: white !important;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
            transition: all 0.3s ease;
        }}
        
        .summary-table tbody tr:hover td {{
            transform: scale(1.01);
            box-shadow: 0 3px 6px rgba(0,0,0,0.15);
            cursor: pointer;
        }}
        
        /* Pass Column - Enhanced Green */
        .summary-table .summary-pass {{
            background: linear-gradient(135deg, #10b981 0%, #34d399 100%) !important;
            box-shadow: inset 0 2px 4px rgba(16, 185, 129, 0.3);
        }}
        
        .summary-table .summary-pass:hover {{
            background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        }}
        
        /* Fail Column - Enhanced Red */
        .summary-table .summary-fail {{
            background: linear-gradient(135deg, #ef4444 0%, #f87171 100%) !important;
            box-shadow: inset 0 2px 4px rgba(239, 68, 68, 0.3);
        }}
        
        .summary-table .summary-fail:hover {{
            background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%) !important;
        }}
        
        /* Blocked Column - Enhanced Yellow/Orange */
        .summary-table .summary-blocked {{
            background: linear-gradient(135deg, #f59e0b 0%, #fbbf24 100%) !important;
            box-shadow: inset 0 2px 4px rgba(245, 158, 11, 0.3);
        }}
        
        .summary-table .summary-blocked:hover {{
            background: linear-gradient(135deg, #d97706 0%, #f59e0b 100%) !important;
        }}
        
        /* NA Column - Enhanced Gray */
        .summary-table .summary-na {{
            background: linear-gradient(135deg, #6b7280 0%, #9ca3af 100%) !important;
            box-shadow: inset 0 2px 4px rgba(107, 114, 128, 0.3);
        }}
        
        .summary-table .summary-na:hover {{
            background: linear-gradient(135deg, #4b5563 0%, #6b7280 100%) !important;
        }}
        
        /* Not Run Column - Enhanced Blue */
        .summary-table .summary-notrun {{
            background: linear-gradient(135deg, #3b82f6 0%, #60a5fa 100%) !important;
            box-shadow: inset 0 2px 4px rgba(59, 130, 246, 0.3);
        }}
        
        .summary-table .summary-notrun:hover {{
            background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%) !important;
        }}
        
        /* Total Column - Enhanced Cyan */
        .summary-table .summary-total {{
            background: linear-gradient(135deg, #06b6d4 0%, #22d3ee 100%) !important;
            box-shadow: inset 0 2px 4px rgba(6, 182, 212, 0.3);
            font-size: 12px;
            font-weight: 700;
        }}
        
        .summary-table .summary-total:hover {{
            background: linear-gradient(135deg, #0891b2 0%, #06b6d4 100%) !important;
        }}
        
        /* Pass% Column - Enhanced Lime Green */
        .summary-table .summary-pass-pct {{
            background: linear-gradient(135deg, #84cc16 0%, #a3e635 100%) !important;
            box-shadow: inset 0 2px 4px rgba(132, 204, 22, 0.3);
            font-size: 12px;
            font-weight: 700;
        }}
        
        .summary-table .summary-pass-pct:hover {{
            background: linear-gradient(135deg, #65a30d 0%, #84cc16 100%) !important;
        }}
        
        /* Exec% Column - Enhanced Emerald Green */
        .summary-table .summary-exec-pct {{
            background: linear-gradient(135deg, #10b981 0%, #34d399 100%) !important;
            box-shadow: inset 0 2px 4px rgba(16, 185, 129, 0.3);
            font-size: 12px;
            font-weight: 700;
        }}
        
        .summary-table .summary-exec-pct:hover {{
            background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        }}
        
        /* Responsive Design for Summary Table */
        @media (max-width: 1400px) {{
            .summary-table {{
                font-size: 9px;
                margin: 12px auto;
                max-width: 800px;
            }}
            
            .summary-table th {{
                font-size: 8px;
                padding: 7px 8px;
            }}
            
            .summary-table td {{
                font-size: 10px;
                padding: 8px 6px;
            }}
            
            .summary-table .summary-total,
            .summary-table .summary-pass-pct,
            .summary-table .summary-exec-pct {{
                font-size: 11px;
            }}
        }}
        
        @media (max-width: 1024px) {{
            .summary-table {{
                font-size: 8px;
                margin: 10px auto;
                max-width: 700px;
            }}
            
            .summary-table th {{
                font-size: 7px;
                padding: 6px 7px;
            }}
            
            .summary-table td {{
                font-size: 9px;
                padding: 7px 5px;
            }}
            
            .summary-table .summary-total,
            .summary-table .summary-pass-pct,
            .summary-table .summary-exec-pct {{
                font-size: 10px;
            }}
        }}
        
        @media (max-width: 768px) {{
            .summary-table {{
                font-size: 7px;
                margin: 8px;
                max-width: 100%;
            }}
            
            .summary-table th {{
                font-size: 6px;
                padding: 5px 6px;
            }}
            
            .summary-table td {{
                font-size: 8px;
                padding: 6px 4px;
            }}
            
            .summary-table .summary-total,
            .summary-table .summary-pass-pct,
            .summary-table .summary-exec-pct {{
                font-size: 9px;
            }}
        }}
        
        /* Print Styles for Summary Table */
        @media print {{
            .summary-table {{
                font-size: 9px;
                box-shadow: none;
                border-radius: 0;
            }}
            
            .summary-table th {{
                font-size: 8px;
                padding: 6px 8px;
                background: #667eea !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            
            .summary-table td {{
                padding: 7px 8px;
                font-size: 10px;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            
            .summary-table tbody tr:hover td {{
                transform: none;
                box-shadow: none;
            }}
        }}
        
        /* Bug Summary Table Styles - Compact Layout */
        .bug-summary-table {{
            width: auto;
            max-width: 750px;
            margin: 10px auto;
            border-collapse: collapse;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
            background: white;
            border: 2px solid #000;
            font-size: 10px;
        }}
        
        .bug-summary-table thead {{
            background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%);
            color: white;
        }}
        
        .bug-summary-table th {{
            padding: 6px 10px;
            text-align: center;
            font-size: 10px;
            font-weight: 700;
            border-right: 2px solid #000;
            border-bottom: 2px solid #000;
            line-height: 1.2;
            white-space: nowrap;
        }}
        
        .bug-summary-table th:last-child {{
            border-right: none;
        }}
        
        .bug-summary-table td {{
            padding: 5px 8px;
            text-align: center;
            font-size: 10px;
            border-bottom: 2px solid #000;
            border-right: 2px solid #000;
            line-height: 1.3;
        }}
        
        .bug-summary-table td:last-child {{
            border-right: none;
        }}
        
        .bug-summary-table tbody tr:hover {{
            background-color: #fef2f2;
        }}
        
        .bug-summary-table tbody tr:last-child td {{
            border-bottom: none;
        }}
        
        .bug-summary-table .bug-critical {{
            background: #fee2e2;
            color: #991b1b;
            font-weight: 600;
        }}
        
        .bug-summary-table .bug-high {{
            background: #fed7aa;
            color: #9a3412;
            font-weight: 600;
        }}
        
        .bug-summary-table .bug-medium {{
            background: #fef3c7;
            color: #92400e;
            font-weight: 600;
        }}
        
        .bug-summary-table .bug-low {{
            background: #dbeafe;
            color: #1e40af;
            font-weight: 600;
        }}
        
        .bug-summary-table .bug-total {{
            background: #f3f4f6;
            font-weight: 700;
            color: #1f2937;
        }}
        
        .bug-summary-table .grand-total-row {{
            background: #dc2626;
            color: white;
            font-weight: 700;
        }}
        
        .bug-summary-table .grand-total-row td {{
            border-bottom: none;
        }}
        
        /* Bug List Table Styles - Compact */
        .bug-list-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            font-size: 10px;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        
        .bug-list-table thead {{
            background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%);
            color: white;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        .bug-list-table th {{
            padding: 6px 8px;
            text-align: center;
            font-size: 9px;
            font-weight: 700;
            border: 2px solid #000 !important;
            line-height: 1.2;
            white-space: nowrap;
        }}
        
        .bug-list-table td {{
            padding: 5px 8px;
            text-align: left;
            font-size: 10px;
            border: 2px solid #000 !important;
            line-height: 1.3;
        }}
        
        .bug-list-table tbody tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        
        .bug-list-table tbody tr:hover {{
            background: #fef2f2;
        }}
        
        .bug-list-table tbody tr.hidden {{
            display: none;
        }}
        
        .bug-list-table .bug-mpoc-col {{
            width: 110px;
            font-weight: 600;
            text-align: left;
        }}
        
        .bug-list-table .bug-id-col {{
            width: 70px;
            text-align: center;
            font-weight: 600;
        }}
        
        .bug-list-table .bug-title-col {{
            min-width: 200px;
            max-width: 350px;
        }}
        
        .bug-list-table .bug-state-col {{
            width: 100px;
            text-align: center;
            font-weight: 600;
        }}
        
        .bug-list-table .bug-defect-col {{
            width: 110px;
            text-align: center;
        }}
        
        .bug-list-table .bug-severity-col {{
            width: 90px;
            text-align: center;
            font-weight: 600;
        }}
        
        .bug-list-table .bug-node-col {{
            width: 130px;
        }}
        
        .bug-list-table .bug-stage-col {{
            width: 110px;
            text-align: center;
        }}
        
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Regression Execution Report</h1>
            <p><strong>{plan_name}</strong> | <strong>Suite: {self.suite_name}</strong> | Generated: {self.timestamp}</p>
        </div>
        
        <!-- Tabs -->
        <div class="tabs">
            <button class="tab active" onclick="switchTab('summary')">📊 Execution Summary</button>
            <button class="tab" onclick="switchTab('detailed')">📋 Detailed Report - P1&P2</button>
            <button class="tab" onclick="switchTab('detailedInsprint')">📋 Detailed Report - Insprint</button>
            <button class="tab" onclick="switchTab('leadsManual')">👥 Leads Wise Status - Manual (P1)</button>
            <button class="tab" onclick="switchTab('leadsAutomation')">🤖 Leads Wise Status - Automation (P2)</button>
            <button class="tab" onclick="switchTab('leadsInsprint')">📦 Leads Wise Status - Insprint</button>
            <button class="tab" onclick="switchTab('bugList')">🐛 Overall Regression/Sanity Bug List</button>
            <button class="tab" onclick="switchTab('insprintDefects')">🏷️ Regression Defects (Insprint/Automation)</button>
        </div>

        <!-- Tab 1: Execution Summary -->
<div id="summaryTab" class="tab-content active">
    <div style="padding: 15px;">
        <h2 style="text-align: center; color: #667eea; margin-bottom: 20px; font-size: 16px;">📊 Execution Summary</h2>
        
        <div style="display: flex; gap: 20px; justify-content: center; align-items: flex-start; flex-wrap: wrap;">
            <!-- Manual (P1) Summary Table -->
            <div style="flex: 1; min-width: 450px; max-width: 550px;">
                <h3 style="text-align: center; color: #495057; margin-bottom: 10px; font-size: 13px; font-weight: 700;">Manual (P1) Summary</h3>
                <table class="summary-table" style="margin: 0 auto;">
                    <thead>
                        <tr>
                            <th class="summary-pass">Passed</th>
                            <th class="summary-fail">Failed</th>
                            <th class="summary-blocked">Blocked</th>
                            <th class="summary-na">NA</th>
                            <th class="summary-notrun">Not Run</th>
                            <th class="summary-total">Total Scenarios-P1</th>
                            <th class="summary-exec-pct">Execution %</th>
                            <th class="summary-pass-pct">Pass %</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="summary-pass">{manual_gt['passed']}</td>
                            <td class="summary-fail">{manual_gt['failed']}</td>
                            <td class="summary-blocked">{manual_gt['blocked']}</td>
                            <td class="summary-na">{manual_gt['na']}</td>
                            <td class="summary-notrun">{manual_gt['not_run']}</td>
                            <td class="summary-total">{manual_gt['total']}</td>
                            <td class="summary-exec-pct">{gt_exec_pct:.2f}</td>
                            <td class="summary-pass-pct">{gt_pass_pct:.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!-- Automation (P2) Summary Table -->
            <div style="flex: 1; min-width: 450px; max-width: 550px;">
                <h3 style="text-align: center; color: #495057; margin-bottom: 10px; font-size: 13px; font-weight: 700;">Automation (P2) Summary</h3>
                <table class="summary-table" style="margin: 0 auto;">
                    <thead>
                        <tr>
                            <th class="summary-pass">Passed</th>
                            <th class="summary-fail">Failed</th>
                            <th class="summary-blocked">Blocked</th>
                            <th class="summary-na">NA</th>
                            <th class="summary-notrun">Not Run</th>
                            <th class="summary-total">Total Scenarios-P2</th>
                            <th class="summary-exec-pct">Execution %</th>
                            <th class="summary-pass-pct">Pass %</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="summary-pass">{auto_gt['passed']}</td>
                            <td class="summary-fail">{auto_gt['failed']}</td>
                            <td class="summary-blocked">{auto_gt['blocked']}</td>
                            <td class="summary-na">{auto_gt['na']}</td>
                            <td class="summary-notrun">{auto_gt['not_run']}</td>
                            <td class="summary-total">{auto_gt['total']}</td>
                            <td class="summary-exec-pct">{auto_gt_exec_pct:.2f}</td>
                            <td class="summary-pass-pct">{auto_gt_pass_pct:.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- Insprint Summary Table (Center aligned) -->
        <div style="display: flex; justify-content: center; margin-top: 20px;">
            <div style="min-width: 450px; max-width: 550px;">
                <h3 style="text-align: center; color: #495057; margin-bottom: 10px; font-size: 13px; font-weight: 700;">Insprint Summary</h3>
                <table class="summary-table" style="margin: 0 auto;">
                    <thead>
                        <tr>
                            <th class="summary-pass">Passed</th>
                            <th class="summary-fail">Failed</th>
                            <th class="summary-blocked">Blocked</th>
                            <th class="summary-na">NA</th>
                            <th class="summary-notrun">Not Run</th>
                            <th class="summary-total">Total Scenarios</th>
                            <th class="summary-exec-pct">Execution %</th>
                            <th class="summary-pass-pct">Pass %</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="summary-pass">{insprint_manual_gt['passed']}</td>
                            <td class="summary-fail">{insprint_manual_gt['failed']}</td>
                            <td class="summary-blocked">{insprint_manual_gt['blocked']}</td>
                            <td class="summary-na">{insprint_manual_gt['na']}</td>
                            <td class="summary-notrun">{insprint_manual_gt['not_run']}</td>
                            <td class="summary-total">{insprint_manual_gt['total']}</td>
                            <td class="summary-exec-pct">{insprint_gt_exec_pct:.2f}</td>
                            <td class="summary-pass-pct">{insprint_gt_pass_pct:.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
"""
        
        # Add Bug Summary Table
        bug_summary = self.process_bug_data_by_mpoc()
        
        if bug_summary:
            html += """
        <!-- Bug Summary Section -->
        <div style="padding: 15px; margin-top: 20px;">
            <h2 style="text-align: center; color: #dc2626; margin-bottom: 20px; font-size: 16px;">🐛 Open Bug Summary (Regression/Sanity)</h2>
            
            <table class="bug-summary-table">
                <thead>
                    <tr>
                        <th style="background: #7c3aed;">MPOC</th>
                        <th class="bug-critical">1 - Critical</th>
                        <th class="bug-high">2 - High</th>
                        <th class="bug-medium">3 - Medium</th>
                        <th class="bug-low">4 - Low</th>
                        <th class="bug-total">Grand Total</th>
                    </tr>
                </thead>
                <tbody>
"""
            
            # Calculate grand totals for each severity
            total_critical = 0
            total_high = 0
            total_medium = 0
            total_low = 0
            
            # Add rows for each MPOC
            for mpoc in sorted(bug_summary.keys()):
                counts = bug_summary[mpoc]
                critical = counts['1 - Critical']
                high = counts['2 - High']
                medium = counts['3 - Medium']
                low = counts['4 - Low']
                mpoc_total = critical + high + medium + low
                
                total_critical += critical
                total_high += high
                total_medium += medium
                total_low += low
                
                html += f"""
                    <tr>
                        <td style="font-weight: 600; text-align: left; padding-left: 20px;">{mpoc}</td>
                        <td class="bug-critical">{critical}</td>
                        <td class="bug-high">{high}</td>
                        <td class="bug-medium">{medium}</td>
                        <td class="bug-low">{low}</td>
                        <td class="bug-total">{mpoc_total}</td>
                    </tr>
"""
            
            # Add Grand Total row
            grand_total_bugs = total_critical + total_high + total_medium + total_low
            
            html += f"""
                    <tr class="grand-total-row">
                        <td style="text-align: left; padding-left: 20px;">Grand Total</td>
                        <td>{total_critical}</td>
                        <td>{total_high}</td>
                        <td>{total_medium}</td>
                        <td>{total_low}</td>
                        <td>{grand_total_bugs}</td>
                    </tr>
                </tbody>
            </table>
        </div>
"""
        
        html += """
    </div>
</div>
        
        <!-- Tab 2: Detailed Report -->
        <div id="detailedTab" class="tab-content">
            <div class="filter-section">
                <div class="filter-group">
                    <label for="leadFilter">🔍 Filter by Lead:</label>
                    <select id="leadFilter" onchange="updateModuleOptions()">
                        <option value="all">-- All Leads --</option>
"""
        
        # Add unique lead options
        unique_leads = sorted(set(organized_data.keys()))
        for lead in unique_leads:
            html += f'                        <option value="{lead}">{lead}</option>\n'
        
        html += """                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="moduleFilter">🔍 Filter by Module:</label>
                    <select id="moduleFilter" onchange="applyFilters()">
                        <option value="all">-- All Modules --</option>
"""
        
        # Add unique module options
        unique_modules = set()
        for lead_data in organized_data.values():
            unique_modules.update(lead_data.keys())
        
        for module in sorted(unique_modules):
            html += f'                        <option value="{module}">{module}</option>\n'
        
        html += """                    </select>
                </div>
                
                <button class="reset-btn" onclick="resetFilters()">↻ Reset Filters</button>
                
                <div class="filter-info">
                    <span id="visibleCount">Showing all rows</span>
                </div>
            </div>
            
            <div class="table-wrapper">
                <table class="report-table">
                    <thead>
                        <tr>
                            <th rowspan="2" class="sno-col">S.No</th>
                            <th rowspan="2" class="lead-col">Lead</th>
                            <th rowspan="2" class="module-col">Module</th>
                            <th rowspan="2" class="total-p1p2-col">Total<br/>(P1+P2)</th>
                            <th colspan="8" class="main-header manual-header">Manual (P1)</th>
                            <th colspan="8" class="main-header automation-header">Automation (P2)</th>
                        </tr>
                        <tr>
                            <th class="manual-header">Total</th>
                            <th class="manual-header">Pass</th>
                            <th class="manual-header">Fail</th>
                            <th class="manual-header">Block</th>
                            <th class="manual-header">NA</th>
                            <th class="manual-header">Not Run</th>
                            <th class="manual-header">Exec%</th>
                            <th class="manual-header">Pass%</th>  
                            <th class="automation-header">Total</th>
                            <th class="automation-header">Pass</th>
                            <th class="automation-header">Fail</th>
                            <th class="automation-header">Block</th>
                            <th class="automation-header">NA</th>
                            <th class="automation-header">Not Run</th>
                             <th class="automation-header">Exec%</th>
                            <th class="automation-header">Pass%</th>           
                        </tr>
                    </thead>
                    <tbody id="reportTableBody">
"""
        
        # Add data rows (remove data attributes for percentage values since filters are removed)
        sno = 1
        for lead in sorted(organized_data.keys()):
            for module in sorted(organized_data[lead].keys()):
                manual = organized_data[lead][module]['manual']
                auto = organized_data[lead][module]['automation']
                
                total_p1p2 = manual['total'] + auto['total']
                
                manual_pass_pct, manual_exec_pct = self.calculate_percentages(manual)
                auto_pass_pct, auto_exec_pct = self.calculate_percentages(auto)
                
                html += f"""
                        <tr data-lead="{lead}" data-module="{module}">
                            <td class="sno-col">{sno}</td>
                            <td class="lead-col">{lead}</td>
                            <td class="module-col">{module}</td>
                            <td class="total-p1p2-col">{total_p1p2}</td>
                            <td class="total-col">{manual['total']}</td>
                            <td class="pass-col">{manual['passed']}</td>
                            <td class="fail-col">{manual['failed']}</td>
                            <td class="blocked-col">{manual['blocked']}</td>
                            <td class="na-col">{manual['na']}</td>
                            <td class="notrun-col">{manual['not_run']}</td>
                            <td class="percentage-col">{manual_exec_pct:.2f}%</td>
                            <td class="percentage-col">{manual_pass_pct:.2f}%</td>
                            <td class="total-col">{auto['total']}</td>
                            <td class="pass-col">{auto['passed']}</td>
                            <td class="fail-col">{auto['failed']}</td>
                            <td class="blocked-col">{auto['blocked']}</td>
                            <td class="na-col">{auto['na']}</td>
                            <td class="notrun-col">{auto['not_run']}</td>
                            <td class="percentage-col">{auto_exec_pct:.2f}%</td>
                            <td class="percentage-col">{auto_pass_pct:.2f}%</td>
                           
                        </tr>
"""
                sno += 1
        
        # Grand Total Row
        manual_gt = grand_totals['manual']
        auto_gt = grand_totals['automation']
        grand_total_p1p2 = manual_gt['total'] + auto_gt['total']
        
        manual_gt_pass_pct, manual_gt_exec_pct = self.calculate_grand_total_percentages(manual_gt)
        auto_gt_pass_pct, auto_gt_exec_pct = self.calculate_grand_total_percentages(auto_gt)
        
        html += f"""
                        <tr class="grand-total-row" id="grandTotalRow">
                            <td colspan="3">Grand Total</td>
                            <td>{grand_total_p1p2}</td>
                            <td>{manual_gt['total']}</td>
                            <td>{manual_gt['passed']}</td>
                            <td>{manual_gt['failed']}</td>
                            <td>{manual_gt['blocked']}</td>
                            <td>{manual_gt['na']}</td>
                            <td>{manual_gt['not_run']}</td>
                            <td>{manual_gt_exec_pct:.2f}%</td>
                            <td>{manual_gt_pass_pct:.2f}%</td>                            
                            <td>{auto_gt['total']}</td>
                            <td>{auto_gt['passed']}</td>
                            <td>{auto_gt['failed']}</td>
                            <td>{auto_gt['blocked']}</td>
                            <td>{auto_gt['na']}</td>
                            <td>{auto_gt['not_run']}</td>
                            <td>{auto_gt_exec_pct:.2f}%</td>
                            <td>{auto_gt_pass_pct:.2f}%</td>
                           
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- Tab 3: Detailed Report - Insprint -->
        <div id="detailedInsprintTab" class="tab-content">
            <div class="filter-section">
                <div class="filter-group">
                    <label for="leadFilterInsprint">🔍 Filter by Lead:</label>
                    <select id="leadFilterInsprint" onchange="updateModuleOptionsInsprint()">
                        <option value="all">-- All Leads --</option>
"""
        
        # Organize insprint data
        insprint_organized = self.organize_data_by_lead_module_insprint()
        
        # Add unique lead options for insprint
        insprint_unique_leads = sorted(set(insprint_organized.keys()))
        for lead in insprint_unique_leads:
            html += f'                        <option value="{lead}">{lead}</option>\n'
        
        html += """                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="moduleFilterInsprint">🔍 Filter by Module:</label>
                    <select id="moduleFilterInsprint" onchange="applyFiltersInsprint()">
                        <option value="all">-- All Modules --</option>
"""
        
        # Add unique module options for insprint
        insprint_unique_modules = set()
        for lead_data in insprint_organized.values():
            insprint_unique_modules.update(lead_data.keys())
        
        for module in sorted(insprint_unique_modules):
            html += f'                        <option value="{module}">{module}</option>\n'
        
        html += """                    </select>
                </div>
                
                <button class="reset-btn" onclick="resetFiltersInsprint()">↻ Reset Filters</button>
                
                <div class="filter-info">
                    <span id="visibleCountInsprint">Showing all rows</span>
                </div>
            </div>
            
            <div class="table-wrapper">
                <table class="report-table">
                    <thead>
                        <tr>
                            <th rowspan="2" class="sno-col">S.No</th>
                            <th rowspan="2" class="lead-col">Lead</th>
                            <th rowspan="2" class="module-col">Module</th>
                            <th rowspan="2" class="total-p1p2-col">Total</th>
                            <th colspan="8" class="main-header manual-header">Insprint</th>
                        </tr>
                        <tr>
                            <th class="manual-header">Total</th>
                            <th class="manual-header">Pass</th>
                            <th class="manual-header">Fail</th>
                            <th class="manual-header">Block</th>
                            <th class="manual-header">NA</th>
                            <th class="manual-header">Not Run</th>
                            <th class="manual-header">Exec%</th>
                            <th class="manual-header">Pass%</th>
                        </tr>
                    </thead>
                    <tbody id="reportTableBodyInsprint">
"""
        
        # Add insprint data rows
        insprint_grand_totals = self.calculate_grand_totals(insprint_organized)
        sno = 1
        for lead in sorted(insprint_organized.keys()):
            for module in sorted(insprint_organized[lead].keys()):
                manual = insprint_organized[lead][module]['manual']
                
                total = manual['total']
                
                manual_pass_pct, manual_exec_pct = self.calculate_percentages(manual)
                
                html += f"""
                        <tr data-lead="{lead}" data-module="{module}">
                            <td class="sno-col">{sno}</td>
                            <td class="lead-col">{lead}</td>
                            <td class="module-col">{module}</td>
                            <td class="total-p1p2-col">{total}</td>
                            <td class="total-col">{manual['total']}</td>
                            <td class="pass-col">{manual['passed']}</td>
                            <td class="fail-col">{manual['failed']}</td>
                            <td class="blocked-col">{manual['blocked']}</td>
                            <td class="na-col">{manual['na']}</td>
                            <td class="notrun-col">{manual['not_run']}</td>
                            <td class="percentage-col">{manual_exec_pct:.2f}%</td>
                            <td class="percentage-col">{manual_pass_pct:.2f}%</td>
                        </tr>
"""
                sno += 1
        
        # Insprint Grand Total Row
        insprint_manual_gt = insprint_grand_totals['manual']
        insprint_grand_total = insprint_manual_gt['total']
        
        insprint_manual_gt_pass_pct, insprint_manual_gt_exec_pct = self.calculate_grand_total_percentages(insprint_manual_gt)
        
        html += f"""
                        <tr class="grand-total-row" id="grandTotalRowInsprint">
                            <td colspan="3">Grand Total</td>
                            <td>{insprint_grand_total}</td>
                            <td>{insprint_manual_gt['total']}</td>
                            <td>{insprint_manual_gt['passed']}</td>
                            <td>{insprint_manual_gt['failed']}</td>
                            <td>{insprint_manual_gt['blocked']}</td>
                            <td>{insprint_manual_gt['na']}</td>
                            <td>{insprint_manual_gt['not_run']}</td>
                            <td>{insprint_manual_gt_exec_pct:.2f}%</td>
                            <td>{insprint_manual_gt_pass_pct:.2f}%</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        
                <!-- Tab 4: Leads Wise Status - Manual (P1) -->
        <div id="leadsManualTab" class="tab-content">
            <div style="padding: 20px;">
                <h2 style="text-align: center; color: #667eea; margin-bottom: 20px;">👥 Leads Wise Execution Status - Manual (P1)</h2>
                
                <div class="table-wrapper">
                    <table class="leads-table">
                        <thead>
                            <tr>
                                <th class="lead-name-col">Lead Name</th>
                                <th class="pass-col">Passed</th>
                                <th class="fail-col">Failed</th>
                                <th class="blocked-col">Blocked</th>
                                <th class="na-col">NA</th>
                                <th class="notrun-col">Not Run</th>
                                <th class="total-col">Total</th>
                                <th class="percentage-col">Execution %</th>
                                <th class="percentage-col">Pass %</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        
        # Add leads summary rows for Manual
        for lead in sorted(leads_summary.keys()):
            data = leads_summary[lead]
            pass_pct, exec_pct = self.calculate_grand_total_percentages(data)
            
            html += f"""
                        <tr>
                            <td class="lead-name-col">{lead}</td>
                            <td class="pass-col">{data['passed']}</td>
                            <td class="fail-col">{data['failed']}</td>
                            <td class="blocked-col">{data['blocked']}</td>
                            <td class="na-col">{data['na']}</td>
                            <td class="notrun-col">{data['not_run']}</td>
                            <td class="total-col">{data['total']}</td>
                            <td class="percentage-col">{exec_pct:.2f}</td>
                            <td class="percentage-col">{pass_pct:.2f}</td>
                        </tr>
"""
        
        # Add Grand Total row for manual leads summary
        html += f"""
                        <tr class="grand-total-row">
                            <td class="lead-name-col">Grand Total</td>
                            <td>{manual_gt['passed']}</td>
                            <td>{manual_gt['failed']}</td>
                            <td>{manual_gt['blocked']}</td>
                            <td>{manual_gt['na']}</td>
                            <td>{manual_gt['not_run']}</td>
                            <td>{manual_gt['total']}</td>
                            <td>{gt_exec_pct:.2f}</td>
                            <td>{gt_pass_pct:.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>    
        </div>
        
               <!-- Tab 4: Leads Wise Status - Automation (P2) -->
        <div id="leadsAutomationTab" class="tab-content">
            <div style="padding: 20px;">
                <h2 style="text-align: center; color: #667eea; margin-bottom: 20px;">🤖 Leads Wise Execution Status - Automation (P2)</h2>
                
                <div class="table-wrapper">
                    <table class="leads-table">
                        <thead>
                            <tr>
                                <th class="lead-name-col">Lead Name</th>
                                <th class="pass-col">Passed</th>
                                <th class="fail-col">Failed</th>
                                <th class="blocked-col">Blocked</th>
                                <th class="na-col">NA</th>
                                <th class="notrun-col">Not Run</th>
                                <th class="total-col">Total</th>
                                <th class="percentage-col">Execution %</th>
                                <th class="percentage-col">Pass %</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        
       
        
        # Add leads summary rows for Automation
        for lead in sorted(automation_leads_summary.keys()):
            data = automation_leads_summary[lead]
            pass_pct, exec_pct = self.calculate_grand_total_percentages(data)
            
            html += f"""
                        <tr>
                            <td class="lead-name-col">{lead}</td>
                            <td class="pass-col">{data['passed']}</td>
                            <td class="fail-col">{data['failed']}</td>
                            <td class="blocked-col">{data['blocked']}</td>
                            <td class="na-col">{data['na']}</td>
                            <td class="notrun-col">{data['not_run']}</td>
                            <td class="total-col">{data['total']}</td>
                            <td class="percentage-col">{exec_pct:.2f}</td>
                            <td class="percentage-col">{pass_pct:.2f}</td>
                        </tr>
"""
        
        # Add Grand Total row for automation leads summary
        html += f"""
                        <tr class="grand-total-row">
                            <td class="lead-name-col">Grand Total</td>
                            <td>{auto_gt['passed']}</td>
                            <td>{auto_gt['failed']}</td>
                            <td>{auto_gt['blocked']}</td>
                            <td>{auto_gt['na']}</td>
                            <td>{auto_gt['not_run']}</td>
                            <td>{auto_gt['total']}</td>
                            <td>{auto_gt_exec_pct:.2f}</td>
                            <td>{auto_gt_pass_pct:.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>    
        </div>
        
        <!-- Tab 6: Leads Wise Status - Insprint -->
        <div id="leadsInsprintTab" class="tab-content">
            <div style="padding: 20px;">
                <h2 style="text-align: center; color: #667eea; margin-bottom: 20px;">📦 Leads Wise Execution Status - Insprint</h2>
                
                <div class="table-wrapper">
                    <table class="leads-table">
                        <thead>
                            <tr>
                                <th class="lead-name-col">Lead Name</th>
                                <th class="pass-col">Passed</th>
                                <th class="fail-col">Failed</th>
                                <th class="blocked-col">Blocked</th>
                                <th class="na-col">NA</th>
                                <th class="notrun-col">Not Run</th>
                                <th class="total-col">Total</th>
                                <th class="percentage-col">Execution %</th>
                                <th class="percentage-col">Pass %</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        
        # Calculate insprint leads summary
        insprint_organized = self.organize_data_by_lead_module_insprint()
        insprint_leads_summary = self.calculate_insprint_leads_summary(insprint_organized)
        insprint_grand_totals = self.calculate_grand_totals(insprint_organized)
        insprint_manual_gt = insprint_grand_totals['manual']
        insprint_gt_pass_pct, insprint_gt_exec_pct = self.calculate_grand_total_percentages(insprint_manual_gt)
        
        # Add leads summary rows for Insprint
        for lead in sorted(insprint_leads_summary.keys()):
            data = insprint_leads_summary[lead]
            pass_pct, exec_pct = self.calculate_grand_total_percentages(data)
            
            html += f"""
                        <tr>
                            <td class="lead-name-col">{lead}</td>
                            <td class="pass-col">{data['passed']}</td>
                            <td class="fail-col">{data['failed']}</td>
                            <td class="blocked-col">{data['blocked']}</td>
                            <td class="na-col">{data['na']}</td>
                            <td class="notrun-col">{data['not_run']}</td>
                            <td class="total-col">{data['total']}</td>
                            <td class="percentage-col">{exec_pct:.2f}</td>
                            <td class="percentage-col">{pass_pct:.2f}</td>
                        </tr>
"""
        
        # Add Grand Total row for insprint leads summary
        html += f"""
                        <tr class="grand-total-row">
                            <td class="lead-name-col">Grand Total</td>
                            <td>{insprint_manual_gt['passed']}</td>
                            <td>{insprint_manual_gt['failed']}</td>
                            <td>{insprint_manual_gt['blocked']}</td>
                            <td>{insprint_manual_gt['na']}</td>
                            <td>{insprint_manual_gt['not_run']}</td>
                            <td>{insprint_manual_gt['total']}</td>
                            <td>{insprint_gt_exec_pct:.2f}</td>
                            <td>{insprint_gt_pass_pct:.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>    
        </div>
        
        <!-- Tab 7: Overall Regression/Sanity Bug List -->
        <div id="bugListTab" class="tab-content">
            <div style="padding: 20px;">
                <h2 style="text-align: center; color: #dc2626; margin-bottom: 15px; font-size: 16px;">🐛 Overall Regression/Sanity Bug List</h2>
                
                <!-- Bug List Filters -->
                <div class="filter-section">
                    <div class="filter-group">
                        <label for="bugMpocFilter">🔍 Filter by MPOC:</label>
                        <select id="bugMpocFilter" onchange="updateStateOptions(); applyBugFilters();">
                            <option value="all">-- All MPOCs --</option>
"""
        
        # Add unique MPOC options with case-insensitive deduplication and include Unassigned
        mpoc_dict = {}
        has_unassigned = False
        
        for bug in self.bug_data:
            mpoc = bug.get('mpoc', '')
            if mpoc in ['', 'N/A']:
                has_unassigned = True
            elif mpoc == 'Unassigned':
                has_unassigned = True
            else:
                # Use lowercase as key to prevent case-sensitive duplicates
                mpoc_lower = mpoc.lower()
                if mpoc_lower not in mpoc_dict:
                    mpoc_dict[mpoc_lower] = mpoc
        
        # Get unique MPOCs sorted
        unique_mpocs = sorted(mpoc_dict.values())
        
        # Add Unassigned first if exists
        if has_unassigned:
            html += f"""                            <option value="Unassigned">Unassigned</option>\n"""
        
        # Add other MPOCs
        for mpoc in unique_mpocs:
            html += f"""                            <option value="{mpoc}">{mpoc}</option>\n"""
        
        html += """                        </select>
                    </div>
                    
                    <div class="filter-group">
                        <label>🔍 Filter by State:</label>
                        <div class="custom-dropdown">
                            <button type="button" class="dropdown-toggle" id="stateDropdownToggle" onclick="toggleStateDropdown()">
                                All States ▼
                            </button>
                            <div class="dropdown-menu" id="stateDropdownMenu">
                                <div class="dropdown-item">
                                    <label>
                                        <input type="checkbox" value="all" checked onchange="toggleAllStates(this)"> All States
                                    </label>
                                </div>
                                <div class="dropdown-divider"></div>
"""
        
        # Add unique state options with checkboxes
        unique_states = sorted(set([bug['state'] for bug in self.bug_data if bug.get('state')]))
        for state in unique_states:
            html += f"""                                <div class="dropdown-item">
                                    <label>
                                        <input type="checkbox" class="state-checkbox" value="{state}" checked onchange="updateStateFilter()"> {state}
                                    </label>
                                </div>\n"""
        
        html += """                            </div>
                        </div>
                    </div>
                    
                    <button class="reset-btn" onclick="resetBugFilters()">↻ Reset Filters</button>
                    
                    <div class="filter-info">
                        <span id="bugVisibleCount">Showing all bugs</span>
                    </div>
                </div>
                
                <!-- Bug List Table -->
                <div class="table-wrapper">
                    <table class="bug-list-table">
                        <thead>
                            <tr>
                                <th class="bug-mpoc-col">ExternalRef ID</th>
                                <th class="bug-id-col">ID</th>
                                <th class="bug-title-col">Title</th>
                                <th class="bug-state-col">State</th>
                                <th class="bug-defect-col">Defect Record</th>
                                <th class="bug-severity-col">Severity</th>
                                <th class="bug-node-col">Node Name</th>
                                <th class="bug-stage-col">StageFound</th>
                            </tr>
                        </thead>
                        <tbody id="bugListTableBody">
"""
        
        # Add bug rows
        for bug in self.bug_data:
            bug_id = bug.get('id', 'N/A')
            mpoc = bug.get('mpoc', 'Unassigned')
            title = bug.get('title', 'N/A')
            state = bug.get('state', 'N/A')
            defect_record = bug.get('defect_record', 'N/A')
            severity = bug.get('severity', 'N/A')
            node_name = bug.get('node_name', 'N/A')
            stage_found = bug.get('stage_found', 'N/A')
            
            html += f"""
                            <tr data-mpoc="{mpoc}" data-state="{state}">
                                <td class="bug-mpoc-col">{mpoc}</td>
                                <td class="bug-id-col">{bug_id}</td>
                                <td class="bug-title-col">{title}</td>
                                <td class="bug-state-col">{state}</td>
                                <td class="bug-defect-col">{defect_record}</td>
                                <td class="bug-severity-col">{severity}</td>
                                <td class="bug-node-col">{node_name}</td>
                                <td class="bug-stage-col">{stage_found}</td>
                            </tr>
"""
        
        html += """                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- Tab 8: Regression Defects (Insprint_Regression and Automation Regression) -->
        <div id="insprintDefectsTab" class="tab-content">
            <div style="padding: 20px;">
                <h2 style="text-align: center; color: #7c3aed; margin-bottom: 15px; font-size: 16px;">🏷️ Regression Defects - Insprint_Regression & Automation Regression (Created After Feb 12, 2026)</h2>
                
                <!-- Insprint Defects Filters -->
                <div class="filter-section">
                    <div class="filter-group">
                        <label for="insprintMpocFilter">🔍 Filter by MPOC:</label>
                        <select id="insprintMpocFilter" onchange="updateInsprintStateOptions(); applyInsprintDefectFilters();">
                            <option value="all">-- All MPOCs --</option>
"""
        
        # Add unique MPOC options for insprint defects
        insprint_mpoc_dict = {}
        insprint_has_unassigned = False
        
        for defect in self.insprint_defects:
            mpoc = defect.get('mpoc', '')
            if mpoc in ['', 'N/A']:
                insprint_has_unassigned = True
            elif mpoc == 'Unassigned':
                insprint_has_unassigned = True
            else:
                mpoc_lower = mpoc.lower()
                if mpoc_lower not in insprint_mpoc_dict:
                    insprint_mpoc_dict[mpoc_lower] = mpoc
        
        unique_insprint_mpocs = sorted(insprint_mpoc_dict.values())
        
        if insprint_has_unassigned:
            html += f"""                            <option value="Unassigned">Unassigned</option>\n"""
        
        for mpoc in unique_insprint_mpocs:
            html += f"""                            <option value="{mpoc}">{mpoc}</option>\n"""
        
        html += """                        </select>
                    </div>
                    
                    <div class="filter-group">
                        <label>🔍 Filter by State:</label>
                        <div class="custom-dropdown">
                            <button type="button" class="dropdown-toggle" id="insprintStateDropdownToggle" onclick="toggleInsprintStateDropdown()">
                                All States ▼
                            </button>
                            <div class="dropdown-menu" id="insprintStateDropdownMenu">
                                <div class="dropdown-item">
                                    <label>
                                        <input type="checkbox" value="all" checked onchange="toggleAllInsprintStates(this)"> All States
                                    </label>
                                </div>
                                <div class="dropdown-divider"></div>
"""
        
        # Add unique state options for insprint defects
        unique_insprint_states = sorted(set([defect['state'] for defect in self.insprint_defects if defect.get('state')]))
        for state in unique_insprint_states:
            html += f"""                                <div class="dropdown-item">
                                    <label>
                                        <input type="checkbox" class="insprint-state-checkbox" value="{state}" checked onchange="updateInsprintStateFilter()"> {state}
                                    </label>
                                </div>\n"""
        
        html += """                            </div>
                        </div>
                    </div>
                    
                    <button class="reset-btn" onclick="resetInsprintDefectFilters()">↻ Reset Filters</button>
                    
                    <div class="filter-info">
                        <span id="insprintDefectVisibleCount">Showing all defects</span>
                    </div>
                </div>
                
                <!-- Insprint Defects Table -->
                <div class="table-wrapper">
                    <table class="bug-list-table">
                        <thead>
                            <tr>
                                <th class="bug-mpoc-col">ExternalRef ID</th>
                                <th class="bug-id-col">ID</th>
                                <th class="bug-title-col">Title</th>
                                <th class="bug-state-col">State</th>
                                <th class="bug-defect-col">Defect Record</th>
                                <th class="bug-severity-col">Severity</th>
                                <th class="bug-node-col">Node Name</th>
                                <th class="bug-stage-col">StageFound</th>
                                <th class="bug-stage-col">Created Date</th>
                            </tr>
                        </thead>
                        <tbody id="insprintDefectTableBody">
"""
        
        # Add insprint defect rows
        for defect in self.insprint_defects:
            defect_id = defect.get('id', 'N/A')
            mpoc = defect.get('mpoc', 'Unassigned')
            title = defect.get('title', 'N/A')
            state = defect.get('state', 'N/A')
            defect_record = defect.get('defect_record', 'N/A')
            severity = defect.get('severity', 'N/A')
            node_name = defect.get('node_name', 'N/A')
            stage_found = defect.get('stage_found', 'N/A')
            created_date = defect.get('created_date', 'N/A')
            
            # Format created_date if it's a datetime string
            if created_date != 'N/A' and 'T' in str(created_date):
                try:
                    from datetime import datetime
                    dt = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                    created_date = dt.strftime('%Y-%m-%d')
                except:
                    pass
            
            html += f"""
                            <tr data-mpoc="{mpoc}" data-state="{state}">
                                <td class="bug-mpoc-col">{mpoc}</td>
                                <td class="bug-id-col">{defect_id}</td>
                                <td class="bug-title-col">{title}</td>
                                <td class="bug-state-col">{state}</td>
                                <td class="bug-defect-col">{defect_record}</td>
                                <td class="bug-severity-col">{severity}</td>
                                <td class="bug-node-col">{node_name}</td>
                                <td class="bug-stage-col">{stage_found}</td>
                                <td class="bug-stage-col">{created_date}</td>
                            </tr>
"""
        
        html += """                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>© 2025 Test Execution Report</strong> | """ + f"{ADO_CONFIG['organization']} / {ADO_CONFIG['project']}" + """</p>
        </div>
    </div>
    
    <script>
        // Store bug data for dynamic filtering
        const bugData = [];
        
        // Store original module options for each lead
        const leadModuleMap = {};
        
        // Initialize bug data from table
        function initializeBugData() {
            const rows = document.querySelectorAll('#bugListTableBody tr');
            rows.forEach(row => {
                bugData.push({
                    mpoc: row.getAttribute('data-mpoc'),
                    state: row.getAttribute('data-state')
                });
            });
        }
        
        // Update state checkboxes based on selected MPOC
        function updateStateOptions() {
            const mpocFilter = document.getElementById('bugMpocFilter').value;
            
            // Get unique states for selected MPOC
            let availableStates = new Set();
            
            if (mpocFilter === 'all') {
                // Show all states
                bugData.forEach(bug => availableStates.add(bug.state));
            } else {
                // Show only states for selected MPOC
                bugData.forEach(bug => {
                    if (bug.mpoc === mpocFilter) {
                        availableStates.add(bug.state);
                    }
                });
            }
            
            // Update state checkboxes
            const stateCheckboxes = document.querySelectorAll('.state-checkbox');
            stateCheckboxes.forEach(cb => {
                const stateValue = cb.value;
                const isAvailable = availableStates.has(stateValue);
                
                // Disable/enable checkbox based on availability
                cb.disabled = !isAvailable;
                
                // Uncheck disabled checkboxes
                if (!isAvailable) {
                    cb.checked = false;
                } else {
                    // Check available checkboxes
                    cb.checked = true;
                }
            });
            
            // Update "All States" checkbox
            const allCheckbox = document.querySelector('input[value="all"]');
            const enabledCheckboxes = document.querySelectorAll('.state-checkbox:not(:disabled)');
            allCheckbox.checked = enabledCheckboxes.length > 0;
            
            updateStateFilter();
        }
        
        // Initialize lead-module mapping from table data
        function initializeLeadModuleMap() {
            const rows = document.querySelectorAll('#reportTableBody tr:not(.grand-total-row)');
            rows.forEach(row => {
                const lead = row.getAttribute('data-lead');
                const module = row.getAttribute('data-module');
                
                if (!leadModuleMap[lead]) {
                    leadModuleMap[lead] = new Set();
                }
                leadModuleMap[lead].add(module);
            });
        }
        
        // Update module dropdown based on selected lead
        function updateModuleOptions() {
            const leadFilter = document.getElementById('leadFilter').value;
            const moduleFilter = document.getElementById('moduleFilter');
            const currentModule = moduleFilter.value;
            
            // Clear existing options except "All Modules"
            moduleFilter.innerHTML = '<option value="all">-- All Modules --</option>';
            
            if (leadFilter === 'all') {
                // Show all modules if "All Leads" is selected
                const allModules = new Set();
                Object.values(leadModuleMap).forEach(modules => {
                    modules.forEach(module => allModules.add(module));
                });
                
                Array.from(allModules).sort().forEach(module => {
                    const option = document.createElement('option');
                    option.value = module;
                    option.textContent = module;
                    moduleFilter.appendChild(option);
                });
            } else {
                // Show only modules for selected lead
                const modulesForLead = leadModuleMap[leadFilter] || new Set();
                Array.from(modulesForLead).sort().forEach(module => {
                    const option = document.createElement('option');
                    option.value = module;
                    option.textContent = module;
                    moduleFilter.appendChild(option);
                });
            }
            
            // Try to restore previous module selection if it's still available
            const availableOptions = Array.from(moduleFilter.options).map(opt => opt.value);
            if (availableOptions.includes(currentModule)) {
                moduleFilter.value = currentModule;
            } else {
                moduleFilter.value = 'all';
            }
            
            // Apply filters after updating options
            applyFilters();
        }
        
        function switchTab(tabName) {
        // Hide all tabs
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
        });
    
        // Remove active class from all tab buttons
        document.querySelectorAll('.tab').forEach(btn => {
            btn.classList.remove('active');
        });
    
        // Show selected tab
        if (tabName === 'summary') {
            document.getElementById('summaryTab').classList.add('active');
            document.querySelectorAll('.tab')[0].classList.add('active');
        } else if (tabName === 'detailed') {
            document.getElementById('detailedTab').classList.add('active');
            document.querySelectorAll('.tab')[1].classList.add('active');
        } else if (tabName === 'detailedInsprint') {
            document.getElementById('detailedInsprintTab').classList.add('active');
            document.querySelectorAll('.tab')[2].classList.add('active');
        } else if (tabName === 'leadsManual') {
            document.getElementById('leadsManualTab').classList.add('active');
            document.querySelectorAll('.tab')[3].classList.add('active');
        } else if (tabName === 'leadsAutomation') {
            document.getElementById('leadsAutomationTab').classList.add('active');
            document.querySelectorAll('.tab')[4].classList.add('active');
        } else if (tabName === 'leadsInsprint') {
            document.getElementById('leadsInsprintTab').classList.add('active');
            document.querySelectorAll('.tab')[5].classList.add('active');
        } else if (tabName === 'bugList') {
            document.getElementById('bugListTab').classList.add('active');
            document.querySelectorAll('.tab')[6].classList.add('active');
        } else if (tabName === 'insprintDefects') {
            document.getElementById('insprintDefectsTab').classList.add('active');
            document.querySelectorAll('.tab')[7].classList.add('active');
        }
    }       
        function updateGrandTotal() {
            // Get all visible rows (not hidden and not grand total)
            const visibleRows = document.querySelectorAll('#reportTableBody tr:not(.grand-total-row):not(.hidden)');
            
            // Initialize totals
            let totalP1P2 = 0;
            let manualTotal = 0, manualPass = 0, manualFail = 0, manualBlock = 0, manualNA = 0, manualNotRun = 0;
            let autoTotal = 0, autoPass = 0, autoFail = 0, autoBlock = 0, autoNA = 0, autoNotRun = 0;
            
            // Sum up values from visible rows
            visibleRows.forEach(row => {
                const cells = row.querySelectorAll('td');
                // Column indices: 3=TotalP1P2, 4=ManTotal, 5=ManPass, 6=ManFail, 7=ManBlock, 8=ManNA, 9=ManNotRun
                // 12=AutoTotal, 13=AutoPass, 14=AutoFail, 15=AutoBlock, 16=AutoNA, 17=AutoNotRun
                totalP1P2 += parseInt(cells[3].textContent) || 0;
                manualTotal += parseInt(cells[4].textContent) || 0;
                manualPass += parseInt(cells[5].textContent) || 0;
                manualFail += parseInt(cells[6].textContent) || 0;
                manualBlock += parseInt(cells[7].textContent) || 0;
                manualNA += parseInt(cells[8].textContent) || 0;
                manualNotRun += parseInt(cells[9].textContent) || 0;
                autoTotal += parseInt(cells[12].textContent) || 0;
                autoPass += parseInt(cells[13].textContent) || 0;
                autoFail += parseInt(cells[14].textContent) || 0;
                autoBlock += parseInt(cells[15].textContent) || 0;
                autoNA += parseInt(cells[16].textContent) || 0;
                autoNotRun += parseInt(cells[17].textContent) || 0;
            });
            
            // Calculate percentages for Manual
            const manualDenomPass = manualPass + manualFail + manualBlock;
            const manualPassPct = manualDenomPass > 0 ? (manualPass / manualDenomPass * 100) : 0;
            const manualDenomExec = manualTotal - manualNA;
            const manualExecPct = manualDenomExec > 0 ? ((manualPass + manualFail + manualBlock) / manualDenomExec * 100) : 0;
            
            // Calculate percentages for Automation
            const autoDenomPass = autoPass + autoFail + autoBlock;
            const autoPassPct = autoDenomPass > 0 ? (autoPass / autoDenomPass * 100) : 0;
            const autoDenomExec = autoTotal - autoNA;
            const autoExecPct = autoDenomExec > 0 ? ((autoPass + autoFail + autoBlock) / autoDenomExec * 100) : 0;
            
            // Update grand total row
            const grandTotalRow = document.getElementById('grandTotalRow');
            if (grandTotalRow) {
                const cells = grandTotalRow.querySelectorAll('td');
                cells[1].textContent = totalP1P2;
                cells[2].textContent = manualTotal;
                cells[3].textContent = manualPass;
                cells[4].textContent = manualFail;
                cells[5].textContent = manualBlock;
                cells[6].textContent = manualNA;
                cells[7].textContent = manualNotRun;
                cells[8].textContent = manualExecPct.toFixed(2) + '%';
                cells[9].textContent = manualPassPct.toFixed(2) + '%';
                cells[10].textContent = autoTotal;
                cells[11].textContent = autoPass;
                cells[12].textContent = autoFail;
                cells[13].textContent = autoBlock;
                cells[14].textContent = autoNA;
                cells[15].textContent = autoNotRun;
                cells[16].textContent = autoExecPct.toFixed(2) + '%';
                cells[17].textContent = autoPassPct.toFixed(2) + '%';
            }
        }
        
        function applyFilters() {
            const leadFilter = document.getElementById('leadFilter').value;
            const moduleFilter = document.getElementById('moduleFilter').value;
            const rows = document.querySelectorAll('#reportTableBody tr:not(.grand-total-row)');
            
            let visibleCount = 0;
            
            rows.forEach(row => {
                const lead = row.getAttribute('data-lead');
                const module = row.getAttribute('data-module');
                
                const leadMatch = leadFilter === 'all' || lead === leadFilter;
                const moduleMatch = moduleFilter === 'all' || module === moduleFilter;
                
                if (leadMatch && moduleMatch) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });
            
            // Update visible count
            const totalRows = rows.length;
            document.getElementById('visibleCount').textContent = 
                `Showing ${visibleCount} of ${totalRows} rows`;
            
            // Update grand total based on visible rows
            updateGrandTotal();
        }
        
        function resetFilters() {
            document.getElementById('leadFilter').value = 'all';
            document.getElementById('moduleFilter').value = 'all';
            updateModuleOptions();
        }
        
        // Insprint Tab Filter Functions
        const leadModuleMapInsprint = {};
        
        // Initialize leadModuleMapInsprint
        if (document.getElementById('reportTableBodyInsprint')) {
            const insprintRows = document.querySelectorAll('#reportTableBodyInsprint tr:not(.grand-total-row)');
            insprintRows.forEach(row => {
                const lead = row.getAttribute('data-lead');
                const module = row.getAttribute('data-module');
                
                if (!leadModuleMapInsprint[lead]) {
                    leadModuleMapInsprint[lead] = new Set();
                }
                leadModuleMapInsprint[lead].add(module);
            });
        }
        
        function updateModuleOptionsInsprint() {
            const leadFilter = document.getElementById('leadFilterInsprint').value;
            const moduleFilter = document.getElementById('moduleFilterInsprint');
            const currentModule = moduleFilter.value;
            
            // Clear existing options except "All Modules"
            moduleFilter.innerHTML = '<option value="all">-- All Modules --</option>';
            
            if (leadFilter === 'all') {
                // Show all modules if "All Leads" is selected
                const allModules = new Set();
                Object.values(leadModuleMapInsprint).forEach(modules => {
                    modules.forEach(module => allModules.add(module));
                });
                
                Array.from(allModules).sort().forEach(module => {
                    const option = document.createElement('option');
                    option.value = module;
                    option.textContent = module;
                    moduleFilter.appendChild(option);
                });
            } else {
                // Show only modules for selected lead
                const modulesForLead = leadModuleMapInsprint[leadFilter] || new Set();
                Array.from(modulesForLead).sort().forEach(module => {
                    const option = document.createElement('option');
                    option.value = module;
                    option.textContent = module;
                    moduleFilter.appendChild(option);
                });
            }
            
            // Try to restore previous module selection if it's still available
            const availableOptions = Array.from(moduleFilter.options).map(opt => opt.value);
            if (availableOptions.includes(currentModule)) {
                moduleFilter.value = currentModule;
            } else {
                moduleFilter.value = 'all';
            }
            
            // Apply filters after updating options
            applyFiltersInsprint();
        }
        
        function updateGrandTotalInsprint() {
            // Get all visible rows (not hidden and not grand total)
            const visibleRows = document.querySelectorAll('#reportTableBodyInsprint tr:not(.grand-total-row):not(.hidden)');
            
            // Initialize totals
            let totalSum = 0;
            let insprintTotal = 0, insprintPass = 0, insprintFail = 0, insprintBlock = 0, insprintNA = 0, insprintNotRun = 0;
            
            // Sum up values from visible rows
            visibleRows.forEach(row => {
                const cells = row.querySelectorAll('td');
                // Column indices: 3=Total, 4=InsprintTotal, 5=Pass, 6=Fail, 7=Block, 8=NA, 9=NotRun
                totalSum += parseInt(cells[3].textContent) || 0;
                insprintTotal += parseInt(cells[4].textContent) || 0;
                insprintPass += parseInt(cells[5].textContent) || 0;
                insprintFail += parseInt(cells[6].textContent) || 0;
                insprintBlock += parseInt(cells[7].textContent) || 0;
                insprintNA += parseInt(cells[8].textContent) || 0;
                insprintNotRun += parseInt(cells[9].textContent) || 0;
            });
            
            // Calculate percentages
            const denomPass = insprintPass + insprintFail + insprintBlock;
            const passPct = denomPass > 0 ? (insprintPass / denomPass * 100) : 0;
            const denomExec = insprintTotal - insprintNA;
            const execPct = denomExec > 0 ? ((insprintPass + insprintFail + insprintBlock) / denomExec * 100) : 0;
            
            // Update grand total row
            const grandTotalRow = document.getElementById('grandTotalRowInsprint');
            if (grandTotalRow) {
                const cells = grandTotalRow.querySelectorAll('td');
                cells[1].textContent = totalSum;
                cells[2].textContent = insprintTotal;
                cells[3].textContent = insprintPass;
                cells[4].textContent = insprintFail;
                cells[5].textContent = insprintBlock;
                cells[6].textContent = insprintNA;
                cells[7].textContent = insprintNotRun;
                cells[8].textContent = execPct.toFixed(2) + '%';
                cells[9].textContent = passPct.toFixed(2) + '%';
            }
        }
        
        function applyFiltersInsprint() {
            const leadFilter = document.getElementById('leadFilterInsprint').value;
            const moduleFilter = document.getElementById('moduleFilterInsprint').value;
            const rows = document.querySelectorAll('#reportTableBodyInsprint tr:not(.grand-total-row)');
            
            let visibleCount = 0;
            
            rows.forEach(row => {
                const lead = row.getAttribute('data-lead');
                const module = row.getAttribute('data-module');
                
                const leadMatch = leadFilter === 'all' || lead === leadFilter;
                const moduleMatch = moduleFilter === 'all' || module === moduleFilter;
                
                if (leadMatch && moduleMatch) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });
            
            // Update visible count
            const totalRows = rows.length;
            document.getElementById('visibleCountInsprint').textContent = 
                `Showing ${visibleCount} of ${totalRows} rows`;
            
            // Update grand total based on visible rows
            updateGrandTotalInsprint();
        }
        
        function resetFiltersInsprint() {
            document.getElementById('leadFilterInsprint').value = 'all';
            document.getElementById('moduleFilterInsprint').value = 'all';
            updateModuleOptionsInsprint();
        }
        
        // Bug List Filter Functions
        function toggleStateDropdown() {
            const menu = document.getElementById('stateDropdownMenu');
            const button = document.getElementById('stateDropdownToggle');
            const isShowing = menu.classList.contains('show');
            
            if (isShowing) {
                menu.classList.remove('show');
            } else {
                // Reset to default position first
                menu.style.top = '100%';
                menu.style.bottom = 'auto';
                menu.style.marginTop = '2px';
                menu.style.marginBottom = '0';
                
                menu.classList.add('show');
                
                // Adjust dropdown position if it goes off-screen
                setTimeout(() => {
                    const buttonRect = button.getBoundingClientRect();
                    const menuRect = menu.getBoundingClientRect();
                    const viewportHeight = window.innerHeight;
                    
                    // Calculate space below and above
                    const spaceBelow = viewportHeight - buttonRect.bottom;
                    const spaceAbove = buttonRect.top;
                    const menuHeight = menuRect.height;
                    
                    // If not enough space below but enough space above, flip upward
                    if (spaceBelow < menuHeight + 20 && spaceAbove > menuHeight + 20) {
                        menu.style.top = 'auto';
                        menu.style.bottom = '100%';
                        menu.style.marginBottom = '2px';
                        menu.style.marginTop = '0';
                    }
                }, 10);
            }
            
            // Close dropdown when clicking outside
            document.addEventListener('click', function closeDropdown(e) {
                if (!e.target.closest('.custom-dropdown')) {
                    menu.classList.remove('show');
                    // Reset position
                    menu.style.top = '100%';
                    menu.style.bottom = 'auto';
                    menu.style.marginTop = '2px';
                    menu.style.marginBottom = '0';
                    document.removeEventListener('click', closeDropdown);
                }
            });
        }
        
        function toggleAllStates(checkbox) {
            const stateCheckboxes = document.querySelectorAll('.state-checkbox');
            stateCheckboxes.forEach(cb => {
                cb.checked = checkbox.checked;
            });
            updateStateFilter();
        }
        
        function updateStateFilter() {
            const allCheckbox = document.querySelector('input[value="all"]');
            const stateCheckboxes = document.querySelectorAll('.state-checkbox');
            const checkedCount = Array.from(stateCheckboxes).filter(cb => cb.checked).length;
            
            // Update "All States" checkbox
            allCheckbox.checked = checkedCount === stateCheckboxes.length;
            
            // Update dropdown button text
            const dropdownToggle = document.getElementById('stateDropdownToggle');
            if (checkedCount === 0) {
                dropdownToggle.textContent = 'No States Selected ▼';
            } else if (checkedCount === stateCheckboxes.length) {
                dropdownToggle.textContent = 'All States ▼';
            } else {
                dropdownToggle.textContent = `${checkedCount} State(s) Selected ▼`;
            }
            
            applyBugFilters();
        }
        
        function applyBugFilters() {
            const mpocFilter = document.getElementById('bugMpocFilter').value;
            const stateCheckboxes = document.querySelectorAll('.state-checkbox:checked');
            const selectedStates = Array.from(stateCheckboxes).map(cb => cb.value);
            const rows = document.querySelectorAll('#bugListTableBody tr');
            
            let visibleCount = 0;
            
            rows.forEach(row => {
                const mpoc = row.getAttribute('data-mpoc');
                const state = row.getAttribute('data-state');
                
                const mpocMatch = mpocFilter === 'all' || mpoc === mpocFilter;
                const stateMatch = selectedStates.length === 0 || selectedStates.includes(state);
                
                if (mpocMatch && stateMatch) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });
            
            // Update visible count
            const totalRows = rows.length;
            document.getElementById('bugVisibleCount').textContent = 
                `Showing ${visibleCount} of ${totalRows} bugs`;
        }
        
        function resetBugFilters() {
            document.getElementById('bugMpocFilter').value = 'all';
            
            // Re-enable and check all state checkboxes
            const allCheckbox = document.querySelector('input[value="all"]');
            const stateCheckboxes = document.querySelectorAll('.state-checkbox');
            allCheckbox.checked = true;
            allCheckbox.disabled = false;
            stateCheckboxes.forEach(cb => {
                cb.checked = true;
                cb.disabled = false;
                cb.parentElement.style.opacity = '1';
            });
            
            updateStateFilter();
        }

        // ============================================================================
        // INSPRINT DEFECTS TAB FILTER FUNCTIONS
        // ============================================================================
        
        // Store insprint defect data for dynamic filtering
        const insprintDefectData = [];
        
        // Initialize insprint defect data from table
        function initializeInsprintDefectData() {
            const rows = document.querySelectorAll('#insprintDefectTableBody tr');
            rows.forEach(row => {
                insprintDefectData.push({
                    mpoc: row.getAttribute('data-mpoc'),
                    state: row.getAttribute('data-state')
                });
            });
        }
        
        // Update state checkboxes based on selected MPOC for insprint defects
        function updateInsprintStateOptions() {
            const mpocFilter = document.getElementById('insprintMpocFilter').value;
            
            // Get unique states for selected MPOC
            let availableStates = new Set();
            
            if (mpocFilter === 'all') {
                // All states available
                const rows = document.querySelectorAll('#insprintDefectTableBody tr');
                rows.forEach(row => {
                    availableStates.add(row.getAttribute('data-state'));
                });
            } else {
                // Only states for selected MPOC
                const rows = document.querySelectorAll('#insprintDefectTableBody tr');
                rows.forEach(row => {
                    if (row.getAttribute('data-mpoc') === mpocFilter) {
                        availableStates.add(row.getAttribute('data-state'));
                    }
                });
            }
            
            // Enable/disable state checkboxes based on availability
            const stateCheckboxes = document.querySelectorAll('.insprint-state-checkbox');
            stateCheckboxes.forEach(cb => {
                const state = cb.value;
                if (availableStates.has(state)) {
                    cb.disabled = false;
                    cb.parentElement.style.opacity = '1';
                } else {
                    cb.disabled = true;
                    cb.checked = false;
                    cb.parentElement.style.opacity = '0.5';
                }
            });
            
            updateInsprintStateFilter();
        }
        
        function toggleInsprintStateDropdown() {
            const menu = document.getElementById('insprintStateDropdownMenu');
            const toggle = document.getElementById('insprintStateDropdownToggle');
            const isShowing = menu.classList.contains('show');
            
            if (isShowing) {
                menu.classList.remove('show');
                return;
            }
            
            menu.classList.add('show');
            
            // Check if dropdown would overflow bottom
            const rect = menu.getBoundingClientRect();
            const viewportHeight = window.innerHeight;
            
            if (rect.bottom > viewportHeight) {
                // Open upward
                menu.style.bottom = '100%';
                menu.style.top = 'auto';
                menu.style.marginBottom = '2px';
                menu.style.marginTop = '0';
            } else {
                // Open downward (default)
                menu.style.top = '100%';
                menu.style.bottom = 'auto';
                menu.style.marginTop = '2px';
                menu.style.marginBottom = '0';
            }
            
            // Close dropdown when clicking outside
            document.addEventListener('click', function closeDropdown(e) {
                if (!e.target.closest('.custom-dropdown')) {
                    menu.classList.remove('show');
                    menu.style.top = '100%';
                    menu.style.bottom = 'auto';
                    menu.style.marginTop = '2px';
                    menu.style.marginBottom = '0';
                    document.removeEventListener('click', closeDropdown);
                }
            });
        }
        
        function toggleAllInsprintStates(checkbox) {
            const stateCheckboxes = document.querySelectorAll('.insprint-state-checkbox');
            stateCheckboxes.forEach(cb => {
                if (!cb.disabled) {
                    cb.checked = checkbox.checked;
                }
            });
            updateInsprintStateFilter();
        }
        
        function updateInsprintStateFilter() {
            const allCheckbox = document.querySelector('#insprintStateDropdownMenu input[value="all"]');
            const stateCheckboxes = document.querySelectorAll('.insprint-state-checkbox');
            const enabledCheckboxes = Array.from(stateCheckboxes).filter(cb => !cb.disabled);
            const checkedCount = enabledCheckboxes.filter(cb => cb.checked).length;
            
            // Update "All States" checkbox
            if (allCheckbox) {
                allCheckbox.checked = checkedCount === enabledCheckboxes.length;
            }
            
            // Update dropdown button text
            const dropdownToggle = document.getElementById('insprintStateDropdownToggle');
            if (checkedCount === 0) {
                dropdownToggle.textContent = 'No States Selected ▼';
            } else if (checkedCount === enabledCheckboxes.length) {
                dropdownToggle.textContent = 'All States ▼';
            } else {
                dropdownToggle.textContent = `${checkedCount} State(s) Selected ▼`;
            }
            
            applyInsprintDefectFilters();
        }
        
        function applyInsprintDefectFilters() {
            const mpocFilter = document.getElementById('insprintMpocFilter').value;
            const stateCheckboxes = document.querySelectorAll('.insprint-state-checkbox:checked');
            const selectedStates = Array.from(stateCheckboxes).map(cb => cb.value);
            const rows = document.querySelectorAll('#insprintDefectTableBody tr');
            
            let visibleCount = 0;
            
            rows.forEach(row => {
                const mpoc = row.getAttribute('data-mpoc');
                const state = row.getAttribute('data-state');
                
                const mpocMatch = mpocFilter === 'all' || mpoc === mpocFilter;
                const stateMatch = selectedStates.length === 0 || selectedStates.includes(state);
                
                if (mpocMatch && stateMatch) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });
            
            // Update visible count
            const totalRows = rows.length;
            document.getElementById('insprintDefectVisibleCount').textContent = 
                `Showing ${visibleCount} of ${totalRows} defects`;
        }
        
        function resetInsprintDefectFilters() {
            document.getElementById('insprintMpocFilter').value = 'all';
            
            // Re-enable and check all state checkboxes
            const allCheckbox = document.querySelector('#insprintStateDropdownMenu input[value="all"]');
            const stateCheckboxes = document.querySelectorAll('.insprint-state-checkbox');
            if (allCheckbox) {
                allCheckbox.checked = true;
                allCheckbox.disabled = false;
            }
            stateCheckboxes.forEach(cb => {
                cb.checked = true;
                cb.disabled = false;
                cb.parentElement.style.opacity = '1';
            });
            
            updateInsprintStateFilter();
        }

        // Initialize on page load
        document.addEventListener('DOMContentLoaded', function() {
            initializeLeadModuleMap();
            updateModuleOptions();
            initializeBugData();
            updateStateOptions(); // Initialize state options on load
            
            // Initialize insprint defects if tab exists
            if (document.getElementById('insprintDefectTableBody')) {
                initializeInsprintDefectData();
                updateInsprintStateOptions();
            }
        });
    </script>
</body>
</html>
"""
        return html
    
    def calculate_leads_summary(self, organized_data):
        """Calculate summary by lead (manual tests only)"""
        leads_summary = defaultdict(lambda: {
            'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0
        })
        
        for lead in organized_data:
            for module in organized_data[lead]:
                manual = organized_data[lead][module]['manual']
                for key in leads_summary[lead]:
                    leads_summary[lead][key] += manual[key]
        
        return leads_summary
    
    def calculate_automation_leads_summary(self, organized_data):
        """Calculate summary by lead (automation tests only)"""
        leads_summary = defaultdict(lambda: {
            'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0
        })
        
        for lead in organized_data:
            for module in organized_data[lead]:
                automation = organized_data[lead][module]['automation']
                for key in leads_summary[lead]:
                    leads_summary[lead][key] += automation[key]
        
        return leads_summary
    
    def calculate_insprint_leads_summary(self, organized_data):
        """Calculate summary by lead (insprint tests only)"""
        leads_summary = defaultdict(lambda: {
            'total': 0, 'passed': 0, 'failed': 0, 'blocked': 0, 'na': 0, 'not_run': 0
        })
        
        for lead in organized_data:
            for module in organized_data[lead]:
                manual = organized_data[lead][module]['manual']
                for key in leads_summary[lead]:
                    leads_summary[lead][key] += manual[key]
        
        return leads_summary
    
    def generate_html_file(self, filename=None):
        """Generate and save HTML report to file"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"regression_execution_report_{timestamp}.html"
        
        html_content = self.generate_html()
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"\n✅ Report generated: {filename}")
        return filename

    @staticmethod
    def save_to_onedrive_sync(local_file, sharepoint_sync_folder):
        """Copy the HTML file to the OneDrive-synced SharePoint folder."""
        if not os.path.exists(sharepoint_sync_folder):
            print(f"\n❌ SharePoint sync folder does not exist: {sharepoint_sync_folder}")
            print("   Please sync the folder using OneDrive first.")
            return None
        dest_file = os.path.join(sharepoint_sync_folder, os.path.basename(local_file))
        shutil.copy2(local_file, dest_file)
        print(f"\n✅ Report also copied to SharePoint sync folder:\n   {dest_file}")
        print("   OneDrive will sync this file to SharePoint automatically.")
        return dest_file

# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main execution flow"""
    print("=" * 80)
    print("🚀 AZURE DEVOPS TEST EXECUTION REPORT GENERATOR")
    print("=" * 80)
    
    # Initialize client
    client = AzureDevOpsClient(ADO_CONFIG)
    
    # Test connection
    if not client.test_connection():
        print("\n❌ Connection failed. Please check your configuration.")
        return
    
    # Get test plan info
    plan_info = client.get_test_plan()
    if not plan_info:
        print("\n❌ Could not fetch test plan information.")
        return
    
    # Verify target suite exists
    suite_info = client.verify_suite_exists(ADO_CONFIG['suite_id'])
    if not suite_info:
        print(f"\n⚠️  Warning: Suite {ADO_CONFIG['suite_id']} not directly accessible")
        print("   Attempting to find suite in plan hierarchy...")
        
        all_suites = client.get_all_suites_in_plan()
        target_suite = None
        
        for suite in all_suites:
            if str(suite.get('id')) == str(ADO_CONFIG['suite_id']):
                target_suite = suite
                break
            elif suite.get('name', '').lower() == ADO_CONFIG['target_suite_name'].lower():
                target_suite = suite
                ADO_CONFIG['suite_id'] = suite.get('id')
                break
        
        if target_suite:
            suite_info = target_suite
            print(f"   ✅ Found suite: {suite_info.get('name')}")
        else:
            print(f"\n❌ Could not find suite '{ADO_CONFIG['target_suite_name']}'")
            print("\n📋 Available suites in this plan:")
            for suite in all_suites[:20]:
                print(f"   - {suite.get('name')} (ID: {suite.get('id')})")
            return
    
    suite_name = suite_info.get('name', ADO_CONFIG['target_suite_name'])
    suite_id = suite_info.get('id', ADO_CONFIG['suite_id'])
    
    # Collect test data
    test_data = client.get_all_test_data_from_suite(suite_id, suite_name)
    
    if not test_data:
        print("\n⚠️  No test data found in the suite.")
        return
    
    # Collect Insprint test data
    insprint_suite_id = ADO_CONFIG.get('insprint_suite_id')
    insprint_suite_name = ADO_CONFIG.get('insprint_suite_name', 'Insprint Execution')
    insprint_data = []
    
    if insprint_suite_id:
        print(f"\n📊 Collecting Insprint Test Data from Suite {insprint_suite_id}...")
        insprint_suite_info = client.verify_suite_exists(insprint_suite_id)
        
        if insprint_suite_info:
            insprint_data = client.get_all_test_data_from_suite(insprint_suite_id, insprint_suite_name)
            print(f"   ✅ Found {len(insprint_data)} Insprint test items")
        else:
            print(f"   ⚠️  Could not access Insprint suite {insprint_suite_id}")
    
    # Fetch bug data from query
    bug_query_id = 'bb2654f5-78c0-4188-a174-ec65039f226a'
    bug_data = client.get_bugs_from_query(bug_query_id)
    
    # Fetch defects with Insprint_Regression and Automation Regression tags created after Feb 12, 2026
    print(f"\n📊 Fetching Insprint_Regression and Automation Regression defects...")
    insprint_defects = client.get_defects_by_tag_and_date(
        tags=['Insprint_Regression', 'Automation Regression'],
        created_after_date='2026-02-12'
    )
    
    # Note: Keep insprint_defects separate for the dedicated tab
    # Don't merge with bug_data
    
    # Generate HTML report
    print(f"\n📝 Generating HTML Report...")
    report_gen = CustomHTMLReportGenerator(
        test_data, 
        plan_info, 
        suite_name, 
        bug_data, 
        insprint_data, 
        insprint_defects
    )
    report_file = report_gen.generate_html_file()
    
    # --- Add this block after report_file is generated ---
    # Set your local OneDrive sync folder path here:
    sharepoint_sync_folder = r"C:\Users\nandini.baskaran\Accenture\mySP Testing - Regression Execution Report"
    report_gen.save_to_onedrive_sync(report_file, sharepoint_sync_folder)
    # -----------------------------------------------------
    
    print("\n" + "=" * 80)
    print("✅ REPORT GENERATION COMPLETED")
    print("=" * 80)
    print(f"\n📄 Report Location: {report_file}")
    print(f"📊 Total Test Cases: {len(test_data)}")
    
    # Summary statistics
    manual_count = sum(1 for t in test_data if t['type'].lower() == 'manual')
    auto_count = sum(1 for t in test_data if t['type'].lower() == 'automation')
    
    print(f"   - Manual: {manual_count}")
    print(f"   - Automation: {auto_count}")
    
    # Outcome summary
    outcomes = {}
    for test in test_data:
        outcome = test['outcome']
        outcomes[outcome] = outcomes.get(outcome, 0) + 1
    
    print(f"\n📈 Outcome Summary:")
    for outcome, count in sorted(outcomes.items()):
        print(f"   - {outcome}: {count}")
    
    # Bug summary
    if bug_data:
        allowed_states = {
            'new', 'active', 'blocked', 'ready to deploy', 'resolved', 
            'ba clarification', 're-open', 'blocked in pt', 'blocked in uat', 'deferred'
        }
        filtered_bugs = [bug for bug in bug_data if bug['state'].lower() in allowed_states]
        
        print(f"\n🐛 Bug Summary:")
        print(f"   - Total Bugs from Query: {len(bug_data)}")
        print(f"   - Bugs in Report (New/Active/Blocked/etc.): {len(filtered_bugs)}")
        
        severity_counts = {}
        for bug in filtered_bugs:
            severity = bug['severity']
            severity_counts[severity] = severity_counts.get(severity, 0) + 1
        
        for severity, count in sorted(severity_counts.items()):
            print(f"   - {severity}: {count}")
    
    # Insprint Defects summary
    if insprint_defects:
        print(f"\n🏷️  Regression Defects Summary (Insprint_Regression & Automation Regression):")
        print(f"   - Total Defects (created after Feb 12, 2026): {len(insprint_defects)}")
        
        # State counts
        state_counts = {}
        for defect in insprint_defects:
            state = defect['state']
            state_counts[state] = state_counts.get(state, 0) + 1
        
        print(f"   - By State:")
        for state, count in sorted(state_counts.items()):
            print(f"      • {state}: {count}")
        
        # Severity counts
        severity_counts = {}
        for defect in insprint_defects:
            severity = defect['severity']
            severity_counts[severity] = severity_counts.get(severity, 0) + 1
        
        print(f"   - By Severity:")
        for severity, count in sorted(severity_counts.items()):
            print(f"      • {severity}: {count}")
    
    print("\n" + "=" * 80)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user")
    except Exception as e:
        print(f"\n\n❌ An error occurred: {e}")
        import traceback
        traceback.print_exc()