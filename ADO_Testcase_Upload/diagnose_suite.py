"""
Diagnostic script to check suite configuration and test points.
Run this to see what's currently in your suite.
"""
import os
import sys
import requests
from pathlib import Path

# Load environment from ADO_SECRETS.env
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
TEST_SUITE_ID = 4480895  # The requirement-based suite
PAT_TOKEN = os.getenv("ADO_PAT_MAIN", "")

if not PAT_TOKEN:
    print("[ERROR] ADO_PAT_MAIN not found in ADO_SECRETS.env")
    sys.exit(1)

BASE_URL = f"https://dev.azure.com/{ORGANIZATION}/{PROJECT}"
import base64
auth_string = base64.b64encode(f":{PAT_TOKEN}".encode()).decode()
HEADERS = {
    "Content-Type": "application/json",
    "Authorization": f"Basic {auth_string}"
}

print("=" * 80)
print(f"DIAGNOSTIC: Suite {TEST_SUITE_ID} in Plan {TEST_PLAN_ID}")
print("=" * 80)

# 1. Get suite information
print("\n[1] Suite Information:")
suite_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/Suites/{TEST_SUITE_ID}?api-version=5.0"
resp = requests.get(suite_url, headers=HEADERS)
if resp.ok:
    suite = resp.json()
    print(f"    Name: {suite.get('name')}")
    print(f"    Type: {suite.get('suiteType')}")
    print(f"    Requirement ID: {suite.get('requirementId')}")
    print(f"    Test Case Count: {suite.get('testCaseCount', 0)}")
else:
    print(f"    ERROR: {resp.status_code} - {resp.text[:200]}")

# 2. Get plan configurations
print("\n[2] Plan Configurations:")
config_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/configurations?api-version=5.0"
resp = requests.get(config_url, headers=HEADERS)
if resp.ok:
    configs = resp.json().get('value', [])
    print(f"    Count: {len(configs)}")
    for cfg in configs:
        print(f"      - ID: {cfg.get('id')}, Name: {cfg.get('name')}, State: {cfg.get('state')}")
else:
    print(f"    ERROR: {resp.status_code} - {resp.text[:200]}")

# 3. Get test cases in suite
print("\n[3] Test Cases in Suite:")
cases_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/Suites/{TEST_SUITE_ID}/testcases?api-version=5.0"
resp = requests.get(cases_url, headers=HEADERS)
if resp.ok:
    cases = resp.json().get('value', [])
    print(f"    Count: {len(cases)}")
    for i, tc in enumerate(cases[:10]):  # Show first 10
        tc_ref = tc.get('testCase', {})
        print(f"      - TC {tc_ref.get('id')}: {tc_ref.get('name', 'N/A')[:60]}")
    if len(cases) > 10:
        print(f"      ... and {len(cases) - 10} more")
else:
    print(f"    ERROR: {resp.status_code} - {resp.text[:200]}")

# 4. Get test points in suite
print("\n[4] Test Points in Suite:")
points_url = f"{BASE_URL}/_apis/test/Plans/{TEST_PLAN_ID}/Suites/{TEST_SUITE_ID}/points?api-version=5.0"
resp = requests.get(points_url, headers=HEADERS)
if resp.ok:
    points = resp.json().get('value', [])
    print(f"    Count: {len(points)}")
    for i, pt in enumerate(points[:10]):  # Show first 10
        tc_ref = pt.get('testCase', {}) or pt.get('testCaseReference', {})
        cfg = pt.get('configuration', {})
        print(f"      - Point {pt.get('id')}: TC {tc_ref.get('id')} + Config {cfg.get('id', 'N/A')} ({cfg.get('name', 'N/A')})")
    if len(points) > 10:
        print(f"      ... and {len(points) - 10} more")
else:
    print(f"    ERROR: {resp.status_code} - {resp.text[:200]}")

# 5. Get requirement information
req_id = suite.get('requirementId') if 'suite' in locals() else None
if req_id:
    print(f"\n[5] User Story {req_id} Test Cases:")
    work_item_url = f"{BASE_URL}/_apis/wit/workitems/{req_id}?$expand=relations&api-version=7.0"
    resp = requests.get(work_item_url, headers=HEADERS)
    if resp.ok:
        wi = resp.json()
        relations = wi.get('relations', [])
        tested_by = [r for r in relations if r.get('rel') == 'Microsoft.VSTS.Common.TestedBy-Forward']
        print(f"    User Story has {len(tested_by)} 'TestedBy' relationship(s)")
        for i, rel in enumerate(tested_by[:10]):
            url = rel.get('url', '')
            tc_id = url.split('/')[-1] if '/' in url else 'N/A'
            print(f"      - Test Case {tc_id}")
        if len(tested_by) > 10:
            print(f"      ... and {len(tested_by) - 10} more")
    else:
        print(f"    ERROR: {resp.status_code}")

print("\n" + "=" * 80)
print("DIAGNOSIS COMPLETE")
print("=" * 80)
print("\nExpected for test cases to appear in Execute tab:")
print("  ✓ Suite Type should be 'RequirementTestSuite'")
print("  ✓ Test Cases should be listed in section [3]")
print("  ✓ Test Points should be listed in section [4] (one per case per config)")
print("  ✓ User Story should have 'TestedBy' relationships in section [5]")
print("\nIf Test Points count is 0, test cases won't appear in Execute tab!")
