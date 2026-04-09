"""
Detailed analysis of column names in IMSVT to match with mappings
"""
import pandas as pd
import numpy as np

print("="*120)
print("ANALYZING IMSVT.XLSX COLUMN STRUCTURE")
print("="*120)

# Load IMSVT with multi-level headers (row 2 as header)
imsvt_df = pd.read_excel('IMSVT.xlsx', header=[2, 3])
print(f"✓ Loaded IMSVT.xlsx - Shape: {imsvt_df.shape}")
print(f"  Header structure: Multi-level (rows 2 & 3)")

print("\n" + "="*120)
print("SAMPLE OF IMSVT COLUMNS")
print("="*120)

# Show columns related to our test case
print("\nSearching for 'Offshore Ratio' columns:")
offshore_cols = []
for i, col in enumerate(imsvt_df.columns):
    if isinstance(col, tuple):
        main_header = str(col[0]).strip()
        sub_header = str(col[1]).strip() if len(col) > 1 else ""
        
        if 'Offshore Ratio' in main_header or 'Offshore Ratio' in sub_header:
            offshore_cols.append((i, main_header, sub_header))
            print(f"  Col {i}: Main='{main_header}' | Sub='{sub_header}'")
            # Show first value
            if len(imsvt_df) > 0:
                print(f"         First value: {imsvt_df[col].iloc[0]}")

print(f"\n  Found {len(offshore_cols)} columns with 'Offshore Ratio'")

# Now load the mapping file and check what it expects
print("\n" + "="*120)
print("CHECKING MAPPING EXPECTATIONS")
print("="*120)

mapping_df = pd.read_excel('IMS VT Automation Mappings.xlsx', header=None)

print("\nFor 'Offshore Ratio' example:")
print("  Mapping file expects (IMSVT side):")
for i in range(2, 5):
    imsvt_main = mapping_df.iloc[i, 3]
    imsvt_sub = mapping_df.iloc[i, 4]
    ims_main = mapping_df.iloc[i, 6]
    ims_sub = mapping_df.iloc[i, 7]
    print(f"    Row {i}: '{imsvt_main}' -> '{imsvt_sub}'")
    print(f"           Should map to IMS: '{ims_main}' -> '{ims_sub}'")
    
    # Check if this column exists in IMSVT
    found = False
    for col in imsvt_df.columns:
        if isinstance(col, tuple):
            if col[0].strip() == str(imsvt_main).strip() and col[1].strip() == str(imsvt_sub).strip():
                found = True
                break
    
    status = "✓ FOUND" if found else "✗ NOT FOUND"
    print(f"           Status in IMSVT: {status}")

# Check all columns in IMSVT that start with "Managed Security"
print("\n" + "="*120)
print("ALL 'Managed Security' COLUMNS IN IMSVT")
print("="*120)

mgd_security_cols = []
for i, col in enumerate(imsvt_df.columns):
    if isinstance(col, tuple):
        main_header = str(col[0]).strip()
        sub_header = str(col[1]).strip() if len(col) > 1 else ""
        
        if 'Managed Security' in main_header:
            mgd_security_cols.append((i, main_header, sub_header))

print(f"\nFound {len(mgd_security_cols)} 'Managed Security' columns")
print("\nFirst 20:")
for i, (col_idx, main, sub) in enumerate(mgd_security_cols[:20]):
    print(f"  {i+1}. Main: '{main}'")
    print(f"     Sub:  '{sub}'")

# Check what sub-column values exist
print("\n" + "="*120)
print("UNIQUE SUB-COLUMN NAMES IN IMSVT")
print("="*120)

sub_cols = set()
for col in imsvt_df.columns:
    if isinstance(col, tuple) and len(col) > 1:
        sub_cols.add(str(col[1]).strip())

print(f"\nFound {len(sub_cols)} unique sub-column names:")
for sub in sorted(sub_cols):
    if sub and 'Unnamed' not in sub:
        print(f"  - {sub}")

print("\n" + "="*120)
print("VERIFICATION SUMMARY")
print("="*120)
print("\nThe mapping file expects sub-columns named:")
print("  - Guidance")
print("  - AsPerSolution")  
print("  - Variance")
print("\nActual sub-columns in IMSVT include the above (if they appear in the list)")
