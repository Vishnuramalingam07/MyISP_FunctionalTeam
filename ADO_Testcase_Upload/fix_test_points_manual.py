"""
Manual fix script to assign test configurations to existing test cases in a suite.
Run this to fix test cases that don't appear in Execute tab because they lack test points.

This script removes test cases from the suite and re-adds them one by one with configuration.
"""
import os
import sys
import requests
from pathlib import Path

# Load environment
env_file = Path(__file__).parent.parent / "ADO_SECRETS.env"
if env_file.exists():
    with open(env_file) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            os.environ[key.strip()] = val.strip()

# Configuration
ORGANIZATION = "accenturecio08"
PROJECT = "AutomationProcess_29697"
TEST_PLAN_ID = 4443950
TEST_SUITE_ID = 4480895  # The requirement suite
CONFIGURATION_ID = 158  # Windows 10
PAT_TOKEN = os.getenv("ADO_PAT_MAIN", "")

if not PAT_TOKEN:
    print("[ERROR] ADO_PAT_MAIN not found")
    sys.exit(1)

BASE_URL = f"https://dev.azure.com/{ORGANIZATION}/{PROJECT}"
import base64
auth_string = base64.b64encode(f":{PAT_TOKEN}".encode()).decode()
HEADERS = {
    "Content-Type": "application/json",
    "Authorization": f"Basic {auth_string}"
}

print("=" * 80)
print("FIX TEST POINTS - Manual Configuration Assignment")
print("=" * 80)
print(f"Suite: {TEST_SUITE_ID}")
print(f"Configuration: {CONFIGURATION_ID} (Windows 10)")
print("=" * 80)

# Step 1: Get all test cases in suite
print("\n[1] Fetching test cases in suite...")
cases_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/Suites/{TEST_SUITE_ID}/testcases?api-version=5.0"
resp = requests.get(cases_url, headers=HEADERS)

if not resp.ok:
    print(f"[ERROR] Failed to get test cases: {resp.status_code}")
    sys.exit(1)

test_cases = resp.json().get('value', [])
print(f"[INFO] Found {len(test_cases)} test case(s) in suite")

# Filter test cases without test points
cases_without_points = []
for tc_item in test_cases:
    tc_ref = tc_item.get('testCase', {})
    tc_id = tc_ref.get('id')
    tc_name = tc_ref.get('name', 'N/A')
    point_assignments = tc_item.get('pointAssignments', [])
    
    if not point_assignments or len(point_assignments) == 0:
        cases_without_points.append((tc_id, tc_name))
        print(f"  - TC {tc_id}: {tc_name[:80]} [NO POINTS]")

if not cases_without_points:
    print("\n[SUCCESS] All test cases already have test points!")
    sys.exit(0)

print(f"\n[INFO] {len(cases_without_points)} test case(s) need fixing")
print("\nStarting fix process...")
input("Press ENTER to continue or Ctrl+C to cancel...")

# Step 2: Fix each test case by re-adding with configuration
fixed_count = 0
failed_count = 0

for tc_id, tc_name in cases_without_points:
    print(f"\n[FIXING] TC {tc_id}...")
    
    # Remove from suite
    del_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/Suites/{TEST_SUITE_ID}/testcases/{tc_id}?api-version=5.0"
    del_resp = requests.delete(del_url, headers=HEADERS)
    
    if not del_resp.ok:
        print(f"  [ERROR] Delete failed: {del_resp.status_code}")
        failed_count += 1
        continue
    
    print(f"  [INFO] Removed from suite")
    
    # Re-add with configuration
    add_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/Suites/{TEST_SUITE_ID}/testcases/{tc_id}?api-version=5.0"
    add_resp = requests.post(add_url, headers=HEADERS)
    
    if not add_resp.ok:
        print(f"  [ERROR] Re-add failed: {add_resp.status_code}")
        failed_count += 1
        continue
    
    result = add_resp.json()
    
    # Check if points were created
    if isinstance(result, dict):
        value = result.get('value', [])
        if value and len(value) > 0:
            points = value[0].get('pointAssignments', [])
            if points and len(points) > 0:
                point_ids = [p.get('id') for p in points]
                print(f"  [SUCCESS] Test points created: {point_ids}")
                fixed_count += 1
            else:
                print(f"  [WARN] Re-added but no points in response")
                failed_count += 1
    elif isinstance(result, list) and len(result) > 0:
        points = result[0].get('pointAssignments', [])
        if points:
            point_ids = [p.get('id') for p in points]
            print(f"  [SUCCESS] Test points created: {point_ids}")
            fixed_count += 1
        else:
            print(f"  [WARN] Re-added but no points")
            failed_count += 1

print("\n" + "=" * 80)
print(f"SUMMARY:")
print(f"  Fixed: {fixed_count}/{len(cases_without_points)}")
print(f"  Failed: {failed_count}/{len(cases_without_points)}")
print("=" * 80)

if fixed_count > 0:
    print(f"\n✓ Check Execute tab: https://dev.azure.com/{ORGANIZATION}/{PROJECT}/_testPlans/execute?planId={TEST_PLAN_ID}&suiteId={TEST_SUITE_ID}")
else:
    print("\n✗ No test cases were fixed. Manual intervention required.")
    print("\nManual steps:")
    print(f"1. Go to https://dev.azure.com/{ORGANIZATION}/{PROJECT}/_testPlans/define?planId={TEST_PLAN_ID}&suiteId={TEST_SUITE_ID}")
    print(f"2. Select each test case")
    print(f"3. Click 'Assign configuration' and select 'Windows 10'")
    print(f"4. Test points will be created automatically")
