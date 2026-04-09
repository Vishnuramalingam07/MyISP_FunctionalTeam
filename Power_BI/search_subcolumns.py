import pandas as pd

# Load IMSVT file
df = pd.read_excel('IMSVT.xlsx', header=None)

print("="*80)
print("SEARCHING FOR 'Guidance', 'AsPerSolution', 'Variance' IN IMSVT")
print("="*80)

search_terms = ['Guidance', 'AsPerSolution', 'Variance', 'Solution Standards', 'Actual Value']

for term in search_terms:
    print(f"\nSearching for '{term}':")
    found = False
    for r in range(min(20, len(df))):
        for c in range(df.shape[1]):
            val = df.iloc[r, c]
            if pd.notna(val) and term.lower() in str(val).lower():
                print(f"  Found at Row {r}, Col {c}: {val}")
                found = True
                break
        if found:
            break
    if not found:
        print(f"  NOT FOUND in first 20 rows")

# Also check what columns exist around "Offshore Ratio"
print("\n" + "="*80)
print("COLUMNS AROUND 'Offshore Ratio (%)'")
print("="*80)

for c in range(df.shape[1]):
    if pd.notna(df.iloc[2, c]) and 'Offshore Ratio' in str(df.iloc[2, c]):
        print(f"\nFound 'Offshore Ratio' at column {c}")
        print(f"  Row 2 (main header): {df.iloc[2, c]}")
        print(f"  Row 3 (sub header): {df.iloc[3, c]}")
        print(f"  Row 4: {df.iloc[4, c]}")
        print(f"  Row 5: {df.iloc[5, c]}")
        
        # Check neighbor columns
        for offset in range(-2, 3):
            col_idx = c + offset
            if 0 <= col_idx < df.shape[1]:
                print(f"\n  Col {col_idx} (offset {offset}):")
                print(f"    Row 2: {df.iloc[2, col_idx]}")
                print(f"    Row 3: {df.iloc[3, col_idx]}")
                print(f"    Row 4: {df.iloc[4, col_idx]}")
