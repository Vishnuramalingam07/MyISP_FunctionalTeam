"""
Display all mappings to verify correctness
"""
import pandas as pd

# Load the mapping file
df = pd.read_excel('IMS VT Automation Mappings.xlsx', header=None)

print("="*100)
print("ALL COLUMN MAPPINGS")
print("="*100)
print(f"{'IMSVT Main Column':<60} | {'IMSVT Sub':<20} | {'IMS Main Column':<50} | {'IMS Sub':<30}")
print("="*100)

# Start from row 2 (data rows)
for i in range(2, len(df)):
    imsvt_main = df.iloc[i, 3]
    imsvt_sub = df.iloc[i, 4]
    ims_main = df.iloc[i, 6]
    ims_sub = df.iloc[i, 7]
    
    if pd.notna(imsvt_main) and pd.notna(imsvt_sub) and pd.notna(ims_main) and pd.notna(ims_sub):
        print(f"{str(imsvt_main):<60} | {str(imsvt_sub):<20} | {str(ims_main):<50} | {str(ims_sub):<30}")

print("="*100)
print("\nCHECKING FOR CORRECT MAPPINGS:")
print("="*100)
print("\nExpected sub-column mappings:")
print("  - Guidance          -> Solution Standards")
print("  - AsPerSolution     -> Actual Value")
print("  - Variance          -> Variation From Standard")
print("\nVerifying each row...")
print("="*100)

correct_count = 0
incorrect_count = 0
issues = []

for i in range(2, len(df)):
    imsvt_main = df.iloc[i, 3]
    imsvt_sub = df.iloc[i, 4]
    ims_main = df.iloc[i, 6]
    ims_sub = df.iloc[i, 7]
    
    if pd.notna(imsvt_main) and pd.notna(imsvt_sub) and pd.notna(ims_main) and pd.notna(ims_sub):
        imsvt_sub = str(imsvt_sub).strip()
        ims_sub = str(ims_sub).strip()
        
        is_correct = False
        if imsvt_sub == "Guidance" and ims_sub == "Solution Standards":
            is_correct = True
        elif imsvt_sub == "AsPerSolution" and ims_sub == "Actual Value":
            is_correct = True
        elif imsvt_sub == "Variance" and ims_sub == "Variation From Standard":
            is_correct = True
        
        if is_correct:
            correct_count += 1
        else:
            incorrect_count += 1
            issues.append({
                'row': i,
                'imsvt_main': imsvt_main,
                'imsvt_sub': imsvt_sub,
                'ims_main': ims_main,
                'ims_sub': ims_sub
            })

print(f"\n✓ Correct mappings: {correct_count}")
print(f"✗ Incorrect mappings: {incorrect_count}")

if issues:
    print("\n" + "="*100)
    print("ISSUES FOUND:")
    print("="*100)
    for issue in issues:
        print(f"\nRow {issue['row']}:")
        print(f"  IMSVT: {issue['imsvt_main']} -> {issue['imsvt_sub']}")
        print(f"  IMS:   {issue['ims_main']} -> {issue['ims_sub']}")
        
        # Suggest correction
        if issue['imsvt_sub'] == 'Guidance':
            print(f"  ⚠️ Should be: '{issue['imsvt_sub']}' -> 'Solution Standards' (currently: '{issue['ims_sub']}')")
        elif issue['imsvt_sub'] == 'AsPerSolution':
            print(f"  ⚠️ Should be: '{issue['imsvt_sub']}' -> 'Actual Value' (currently: '{issue['ims_sub']}')")
        elif issue['imsvt_sub'] == 'Variance':
            print(f"  ⚠️ Should be: '{issue['imsvt_sub']}' -> 'Variation From Standard' (currently: '{issue['ims_sub']}')")
