import pandas as pd

# Load the mapping file without headers
df = pd.read_excel('IMS VT Automation Mappings.xlsx', header=None)

print(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
print("\n" + "="*80)
print("FIRST 15 ROWS, ALL COLUMNS")
print("="*80)

for row_idx in range(min(15, len(df))):
    print(f"\nRow {row_idx}:")
    for col_idx in range(df.shape[1]):
        val = df.iloc[row_idx, col_idx]
        if pd.notna(val):
            print(f"  Col {col_idx}: {val}")

print("\n" + "="*80)
print("CHECKING COLUMN HEADERS AND STRUCTURE")
print("="*80)

# Check what's in specific rows that might be headers
print("\nRow 1 (possible headers):")
print(list(df.iloc[1, :]))

print("\nRow 2 (possible data start):")
print(list(df.iloc[2, :]))

# Look at the example mentioned: "Offshore Ratio (%)"
print("\n" + "="*80)
print("SEARCHING FOR 'Offshore Ratio' EXAMPLE")
print("="*80)

for row_idx in range(len(df)):
    for col_idx in range(df.shape[1]):
        val = str(df.iloc[row_idx, col_idx])
        if 'Offshore Ratio' in val:
            print(f"\nFound at Row {row_idx}, Col {col_idx}: {val}")
            # Print the surrounding context
            print(f"  Same row, other columns:")
            for c in range(df.shape[1]):
                if pd.notna(df.iloc[row_idx, c]):
                    print(f"    Col {c}: {df.iloc[row_idx, c]}")
