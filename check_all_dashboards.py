import os
import re
from pathlib import Path
import glob

# Find all dashboard files in the correct locations (project-local only)
project_root = Path(__file__).resolve().parent

dashboard_files = []

# Main Release - find all timestamped dashboards
main_release_folder = project_root / "Main_Release_Daily_Status_Report"
if main_release_folder.exists():
    dashboard_files.extend(glob.glob(str(main_release_folder / "Daily_Status_Dashboard_*.html")))

# Hot Fix
hotfix_folder = project_root / "Hot_Fix_Daily_Status_Report"
hotfix_dashboard = hotfix_folder / "Daily_Status_Dashboard.html"
if hotfix_dashboard.exists():
    dashboard_files.append(str(hotfix_dashboard))

print("=" * 100)
print("CHECKING ALL DAILY_STATUS_DASHBOARD.HTML FILES")
print("=" * 100)

for file_path in dashboard_files:
    if os.path.exists(file_path):
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(5000)  # Read first 5000 chars
            
            # Check for milestone dates
            may_match = re.search(r'09 MAY 2026|May 09, 2026|May 9', content, re.IGNORECASE)
            april_match = re.search(r'18 APR 2026|April 18, 2026|April 18', content, re.IGNORECASE)
            
            file_size = os.path.getsize(file_path)
            modified = os.path.getmtime(file_path)
            from datetime import datetime
            mod_date = datetime.fromtimestamp(modified).strftime('%Y-%m-%d %H:%M:%S')
            
            if may_match:
                print(f"\n✅ MAY 9 DATES: {file_path}")
                print(f"   Last Modified: {mod_date}")
                print(f"   Size: {file_size:,} bytes")
            elif april_match:
                print(f"\n❌ APRIL 18 DATES: {file_path}")
                print(f"   Last Modified: {mod_date}")
                print(f"   Size: {file_size:,} bytes")
            else:
                print(f"\n⚠️  UNKNOWN DATES: {file_path}")
                print(f"   Last Modified: {mod_date}")
        except Exception as e:
            print(f"\n❌ ERROR reading {file_path}: {e}")
    else:
        print(f"\n⚠️  FILE NOT FOUND: {file_path}")

print("\n" + "=" * 100)
print("CORRECT FILE FOR MAIN RELEASE:")
print("c:\\Users\\vishnu.ramalingam\\MyISP_Tools\\Main_Release_Daily_Status_Report\\Daily_Status_Dashboard.html")
print("=" * 100)
