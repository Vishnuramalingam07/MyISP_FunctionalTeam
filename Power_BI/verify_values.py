"""
Verify actual values from both files according to the mappings
"""
import pandas as pd
import numpy as np

# Load the mapping file
print("="*120)
print("LOADING FILES...")
print("="*120)

mapping_df = pd.read_excel('IMS VT Automation Mappings.xlsx', header=None)
print("✓ Loaded mapping file")

# Load IMSVT file with multi-level headers
imsvt_df = pd.read_excel('IMSVT.xlsx', header=[0, 1])
print(f"✓ Loaded IMSVT.xlsx - Shape: {imsvt_df.shape}")

# Load IMS Managed Security file
ims_files = [f for f in pd.ExcelFile('IMS Mgd Security_OCPKDA_0012026420.xlsm').sheet_names]
print(f"✓ Found sheets in IMS file: {ims_files}")

# Try to load with different header configurations
try:
    ims_df = pd.read_excel('IMS Mgd Security_OCPKDA_0012026420.xlsm', sheet_name=ims_files[0], header=[0, 1])
except:
    # Try single header
    try:
        ims_df = pd.read_excel('IMS Mgd Security_OCPKDA_0012026420.xlsm', sheet_name=ims_files[0], header=0)
    except:
        # Read without header to inspect
        ims_df = pd.read_excel('IMS Mgd Security_OCPKDA_0012026420.xlsm', sheet_name=ims_files[0], header=None)

print(f"✓ Loaded IMS Managed Security - Sheet: {ims_files[0]} - Shape: {ims_df.shape}")

print("\n" + "="*120)
print("IMSVT COLUMNS (Multi-level):")
print("="*120)
for i, col in enumerate(imsvt_df.columns[:20]):  # Show first 20
    print(f"{i}: {col}")

print("\n" + "="*120)
print("IMS COLUMNS (Multi-level):")
print("="*120)
for i, col in enumerate(ims_df.columns[:20]):  # Show first 20
    print(f"{i}: {col}")

print("\n" + "="*120)
print("VERIFYING VALUES FOR EACH MAPPING")
print("="*120)

results = []

# Parse mappings starting from row 2
for i in range(2, len(mapping_df)):
    imsvt_main = mapping_df.iloc[i, 3]
    imsvt_sub = mapping_df.iloc[i, 4]
    ims_main = mapping_df.iloc[i, 6]
    ims_sub = mapping_df.iloc[i, 7]
    
    if pd.notna(imsvt_main) and pd.notna(imsvt_sub) and pd.notna(ims_main) and pd.notna(ims_sub):
        imsvt_main = str(imsvt_main).strip()
        imsvt_sub = str(imsvt_sub).strip()
        ims_main = str(ims_main).strip()
        ims_sub = str(ims_sub).strip()
        
        # Try to find the columns in both dataframes
        imsvt_value = None
        ims_value = None
        
        # Search for IMSVT column
        for col in imsvt_df.columns:
            if isinstance(col, tuple):
                if col[0].strip() == imsvt_main and col[1].strip() == imsvt_sub:
                    imsvt_value = imsvt_df[col].iloc[0] if len(imsvt_df) > 0 else None
                    break
        
        # Search for IMS column
        for col in ims_df.columns:
            if isinstance(col, tuple):
                if col[0].strip() == ims_main and col[1].strip() == ims_sub:
                    ims_value = ims_df[col].iloc[0] if len(ims_df) > 0 else None
                    break
        
        match_status = "✓" if imsvt_value == ims_value else "✗"
        
        results.append({
            'row': i,
            'imsvt_col': f"{imsvt_main} -> {imsvt_sub}",
            'ims_col': f"{ims_main} -> {ims_sub}",
            'imsvt_value': imsvt_value,
            'ims_value': ims_value,
            'match': match_status
        })

# Display results
print(f"\n{'Row':<5} | {'Match':<6} | {'IMSVT Column':<70} | {'IMS Column':<70}")
print("="*120)
print(f"{'':5} | {'':6} | {'IMSVT Value':<70} | {'IMS Value':<70}")
print("="*120)

for r in results[:10]:  # Show first 10 mappings
    print(f"{r['row']:<5} | {r['match']:<6} | {r['imsvt_col']:<70} | {r['ims_col']:<70}")
    print(f"{'':5} | {'':6} | {str(r['imsvt_value']):<70} | {str(r['ims_value']):<70}")
    print("-"*120)

print(f"\n... showing first 10 of {len(results)} mappings")

# Summary
matches = sum(1 for r in results if r['match'] == '✓')
mismatches = sum(1 for r in results if r['match'] == '✗')

print("\n" + "="*120)
print("SUMMARY")
print("="*120)
print(f"Total mappings checked: {len(results)}")
print(f"✓ Matching values: {matches}")
print(f"✗ Mismatching values: {mismatches}")

# Show mismatches if any
if mismatches > 0:
    print("\n" + "="*120)
    print("MISMATCHES DETAILS:")
    print("="*120)
    for r in results:
        if r['match'] == '✗':
            print(f"\nRow {r['row']}:")
            print(f"  IMSVT: {r['imsvt_col']} = {r['imsvt_value']}")
            print(f"  IMS:   {r['ims_col']} = {r['ims_value']}")
