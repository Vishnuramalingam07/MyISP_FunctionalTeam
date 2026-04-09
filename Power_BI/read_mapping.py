"""
Read and understand the IMS VT Automation Mappings.xlsx structure
"""
import pandas as pd
import sys

# Read the mapping file
df = pd.read_excel(r'IMS VT Automation Mappings.xlsx', header=None)

print("="*80)
print("MAPPING FILE STRUCTURE")
print("="*80)
print(f"Shape: {df.shape}")
print(f"\nColumn Count: {df.shape[1]}")

print("\n" + "="*80)
print("FIRST 25 ROWS")
print("="*80)

for i in range(min(25, len(df))):
    row_data = []
    for j in range(df.shape[1]):
        val = df.iloc[i, j]
        if pd.notna(val):
            row_data.append(f"Col{j}='{val}'")
    if row_data:
        print(f"Row {i}: {', '.join(row_data)}")

print("\n" + "="*80)
print("IDENTIFYING MAPPING PATTERN")
print("="*80)

# Look for the main headers
for i in range(min(10, len(df))):
    for j in range(df.shape[1]):
        val = str(df.iloc[i, j]).strip()
        if 'IMSVT' in val or 'IMS Managed' in val or 'MainHeader' in val:
            print(f"Found at Row {i}, Col {j}: '{val}'")

print("\n" + "="*80)
print("EXTRACTING COLUMN MAPPINGS")
print("="*80)

# Try to extract mappings (assuming pattern based on what we saw)
mappings = []
for i in range(len(df)):
    imsvt_col = df.iloc[i, 3]  # Column 3 seems to have IMSVT column names
    ims_col = df.iloc[i, 6]    # Column 6 seems to have IMS Managed Security column names
    
    if pd.notna(imsvt_col) and pd.notna(ims_col):
        imsvt_str = str(imsvt_col).strip()
        ims_str = str(ims_col).strip()
        
        # Skip header rows
        if 'MainHeader' not in imsvt_str and 'Columns' not in imsvt_str and 'IMSVT' not in imsvt_str:
            if imsvt_str and ims_str and imsvt_str != 'nan' and ims_str != 'nan':
                mappings.append({
                    'IMSVT_Column': imsvt_str,
                    'IMS_Managed_Security_Column': ims_str,
                    'Row': i
                })

print(f"\nFound {len(mappings)} potential column mappings:")
for m in mappings[:20]:  # Show first 20
    print(f"  Row {m['Row']}: '{m['IMSVT_Column']}' -> '{m['IMS_Managed_Security_Column']}'")

if len(mappings) > 20:
    print(f"  ... and {len(mappings) - 20} more")

# Save mappings to CSV for easy reference
if mappings:
    df_mappings = pd.DataFrame(mappings)
    df_mappings.to_csv('column_mappings.csv', index=False)
    print(f"\n✓ Saved {len(mappings)} mappings to 'column_mappings.csv'")
