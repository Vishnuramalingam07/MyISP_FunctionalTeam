"""
Check where actual data values are in IMSVT
"""
import pandas as pd

# Load IMSVT
imsvt_df = pd.read_excel('IMSVT.xlsx', header=[2, 3])

print("="*120)
print("CHECKING DATA LOCATION IN IMSVT")
print("="*120)

# Find the Offshore Ratio column
for col in imsvt_df.columns:
    if isinstance(col, tuple):
        main = str(col[0]).strip()
        sub = str(col[1]).strip()
        
        if main == 'Managed Security - Offshore Ratio (%)' and sub == '%':
            print(f"\nColumn: {col}")
            print(f"  Main Header: '{main}'")
            print(f"  Sub Header: '{sub}'")
            print(f"\n  Data rows:")
            for i in range(min(10, len(imsvt_df))):
                val = imsvt_df[col].iloc[i]
                print(f"    Row {i}: {val}")
            break

# Also check IMS column
for col in imsvt_df.columns:
    if isinstance(col, tuple):
        main = str(col[0]).strip()
        sub = str(col[1]).strip()
        
        if main == 'Offshore Ratio (%)' and sub == '%':
            print(f"\n\nIMS Column: {col}")
            print(f"  Main Header: '{main}'")
            print(f"  Sub Header: '{sub}'")
            print(f"\n  Data rows:")
            for i in range(min(10, len(imsvt_df))):
                val = imsvt_df[col].iloc[i]
                print(f"    Row {i}: {val}")
            break

print("\n" + "="*120)
print("ANALYSIS")
print("="*120)
print("\nRow 0 (Index 0): Contains labels like 'Guidance', 'AsPerSolution', 'Variance'")
print("Row 1 (Index 1): Usually empty or contains additional metadata")
print("Row 2+ (Index 2+): Contains actual data values")
print("\nThe script should look at row index >= 2 for actual data values")
