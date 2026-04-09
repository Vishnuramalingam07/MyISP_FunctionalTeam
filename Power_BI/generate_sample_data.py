"""
Generate Sample Excel Files for Testing Excel Compare Agent
============================================================
This script creates sample Excel files matching the requirement structure.
"""

import pandas as pd
from pathlib import Path
import random
from datetime import datetime

def generate_sample_release_files():
    """Generate sample release tracking Excel files"""
    
    print("Generating sample Excel files for testing...")
    
    # Sample data structure matching the requirement
    data_feb_21 = {
        'US_ID': ['US001', 'US002', 'US003', 'US004', 'US005', 'US006', 'US007', 'US008'],
        'Feature Name': [
            'User Authentication',
            'Dashboard Widgets',
            'Report Generation',
            'Data Export',
            'Search Functionality',
            'User Profile',
            'Email Notifications',
            'API Integration'
        ],
        'PT Status': [
            'In Progress',
            'Completed',
            'Not Started',
            'In Progress',
            'Completed',
            'In Progress',
            'Completed',
            'Not Started'
        ],
        'In sprint test case count': [15, 23, 0, 12, 18, 10, 25, 0],
        'If insprint YES - % of completion': [60.5, 100.0, 0.0, 45.0, 100.0, 75.0, 100.0, 0.0],
        'Assigned To': [
            'John Doe',
            'Jane Smith',
            'Bob Wilson',
            'Alice Brown',
            'Charlie Davis',
            'Eve Martin',
            'Frank White',
            'Grace Lee'
        ],
        'Priority': ['High', 'Medium', 'Low', 'High', 'Medium', 'High', 'Low', 'Medium']
    }
    
    # Create Feb 21st baseline
    df_feb_21 = pd.DataFrame(data_feb_21)
    
    # Create Mar 28th with some changes
    df_mar_28 = df_feb_21.copy()
    
    # Simulate changes for comparison testing
    # Change 1: US001 progressed
    df_mar_28.loc[df_mar_28['US_ID'] == 'US001', 'PT Status'] = 'Completed'
    df_mar_28.loc[df_mar_28['US_ID'] == 'US001', 'In sprint test case count'] = 18
    df_mar_28.loc[df_mar_28['US_ID'] == 'US001', 'If insprint YES - % of completion'] = 100.0
    
    # Change 2: US003 started
    df_mar_28.loc[df_mar_28['US_ID'] == 'US003', 'PT Status'] = 'In Progress'
    df_mar_28.loc[df_mar_28['US_ID'] == 'US003', 'In sprint test case count'] = 8
    df_mar_28.loc[df_mar_28['US_ID'] == 'US003', 'If insprint YES - % of completion'] = 30.0
    
    # Change 3: US004 updated progress
    df_mar_28.loc[df_mar_28['US_ID'] == 'US004', 'In sprint test case count'] = 15
    df_mar_28.loc[df_mar_28['US_ID'] == 'US004', 'If insprint YES - % of completion'] = 78.5
    
    # Change 4: US006 updated
    df_mar_28.loc[df_mar_28['US_ID'] == 'US006', 'If insprint YES - % of completion'] = 90.0
    
    # Change 5: US008 started
    df_mar_28.loc[df_mar_28['US_ID'] == 'US008', 'PT Status'] = 'In Progress'
    df_mar_28.loc[df_mar_28['US_ID'] == 'US008', 'In sprint test case count'] = 5
    df_mar_28.loc[df_mar_28['US_ID'] == 'US008', 'If insprint YES - % of completion'] = 15.0
    
    # Add one new user story in Mar 28th
    new_story = {
        'US_ID': 'US009',
        'Feature Name': 'Mobile App Support',
        'PT Status': 'In Progress',
        'In sprint test case count': 12,
        'If insprint YES - % of completion': 40.0,
        'Assigned To': 'Henry Kumar',
        'Priority': 'High'
    }
    df_mar_28 = pd.concat([df_mar_28, pd.DataFrame([new_story])], ignore_index=True)
    
    # Remove one story from Mar 28th (to test missing in B)
    df_mar_28 = df_mar_28[df_mar_28['US_ID'] != 'US007']
    
    # Save files
    output_dir = Path(r'GHC files\Daily status report')
    output_dir.mkdir(parents=True, exist_ok=True)
    
    file_feb_21 = output_dir / '21st Feb_Release.xlsx'
    file_mar_28 = output_dir / 'Mar 28th_Release.xlsx'
    
    # Save with 'Release' sheet name as per requirement
    with pd.ExcelWriter(file_feb_21, engine='openpyxl') as writer:
        df_feb_21.to_excel(writer, sheet_name='Release', index=False)
    
    with pd.ExcelWriter(file_mar_28, engine='openpyxl') as writer:
        df_mar_28.to_excel(writer, sheet_name='Release', index=False)
    
    print(f"✓ Created: {file_feb_21}")
    print(f"  - {len(df_feb_21)} rows, {len(df_feb_21.columns)} columns")
    print(f"  - Sheet: Release")
    
    print(f"✓ Created: {file_mar_28}")
    print(f"  - {len(df_mar_28)} rows, {len(df_mar_28.columns)} columns")
    print(f"  - Sheet: Release")
    
    print("\nExpected Changes:")
    print("  - US001: Status changed, test count increased, completion 100%")
    print("  - US003: Started (was Not Started)")
    print("  - US004: Progress updated")
    print("  - US006: Completion % updated")
    print("  - US008: Started (was Not Started)")
    print("  - US009: New story (missing in Feb 21st)")
    print("  - US007: Removed in Mar 28th (missing in Mar 28th)")
    
    return file_feb_21, file_mar_28


if __name__ == "__main__":
    feb_file, mar_file = generate_sample_release_files()
    print("\n" + "=" * 80)
    print("Sample files created successfully!")
    print("You can now run: python test_excel_compare.py")
    print("=" * 80)
