"""
Inspect the structure of both Excel files to understand their format
"""
import pandas as pd
import numpy as np

print("="*120)
print("INSPECTING IMSVT.xlsx")
print("="*120)

# Load IMSVT with multi-level headers
imsvt_df = pd.read_excel('IMSVT.xlsx', header=[0, 1])
print(f"Shape: {imsvt_df.shape}")
print(f"\nFirst 10 column names:")
for i, col in enumerate(imsvt_df.columns[:10]):
    print(f"  {i}: {col}")

print(f"\nSearching for 'Offshore Ratio' in IMSVT columns:")
for i, col in enumerate(imsvt_df.columns):
    if isinstance(col, tuple):
        if 'Offshore Ratio' in col[0]:
            print(f"  Found: {col}")
            print(f"    First value: {imsvt_df[col].iloc[0]}")

print("\n" + "="*120)
print("INSPECTING IMS Mgd Security_OCPKDA_0012026420.xlsm")
print("="*120)

# First, read without header to see structure
ims_raw = pd.read_excel('IMS Mgd Security_OCPKDA_0012026420.xlsm', sheet_name='KeyDealAttributes', header=None)
print(f"Raw shape: {ims_raw.shape}")

print(f"\nFirst 3 rows (to identify header structure):")
for i in range(min(3, len(ims_raw))):
    print(f"\nRow {i}:")
    for j in range(min(15, ims_raw.shape[1])):
        val = ims_raw.iloc[i, j]
        if pd.notna(val):
            print(f"  Col {j}: {val}")

# Try to find where 'Offshore Ratio' appears
print(f"\nSearching for 'Offshore Ratio' in IMS file:")
for i in range(min(10, len(ims_raw))):
    for j in range(ims_raw.shape[1]):
        val = str(ims_raw.iloc[i, j])
        if 'Offshore Ratio' in val:
            print(f"  Found at Row {i}, Col {j}: {val}")
            # Show surrounding context
            if i + 1 < len(ims_raw):
                print(f"    Next row value: {ims_raw.iloc[i+1, j]}")

# Now let's try to load with correct headers
print("\n" + "="*120)
print("ATTEMPTING TO LOAD WITH MULTI-LEVEL HEADERS")
print("="*120)

# Check if it has multi-level headers by looking at first 2 rows
has_multiheader = False
row0_vals = [str(v) for v in ims_raw.iloc[0] if pd.notna(v)]
row1_vals = [str(v) for v in ims_raw.iloc[1] if pd.notna(v)]

print(f"Row 0 non-null values: {row0_vals[:10]}")
print(f"Row 1 non-null values: {row1_vals[:10]}")

# Try loading with multi-header
try:
    ims_df = pd.read_excel('IMS Mgd Security_OCPKDA_0012026420.xlsm', sheet_name='KeyDealAttributes', header=[0, 1])
    print(f"✓ Successfully loaded with multi-level headers")
    print(f"  Shape: {ims_df.shape}")
    print(f"  First 5 columns:")
    for i, col in enumerate(ims_df.columns[:5]):
        print(f"    {i}: {col}")
except Exception as e:
    print(f"✗ Could not load with multi-level headers: {e}")
    print("\nTrying with single header...")
    try:
        ims_df = pd.read_excel('IMS Mgd Security_OCPKDA_0012026420.xlsm', sheet_name='KeyDealAttributes', header=0)
        print(f"✓ Successfully loaded with single header")
        print(f"  Shape: {ims_df.shape}")
        print(f"  First 10 columns:")
        for i, col in enumerate(ims_df.columns[:10]):
            print(f"    {i}: {col}")
    except Exception as e2:
        print(f"✗ Could not load with single header either: {e2}")
