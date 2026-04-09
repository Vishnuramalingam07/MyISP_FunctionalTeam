# If you need to filter or compare on 'Node Name', do so case-insensitively
# Example: to filter for a specific node name or list of node names
# node_name_filter = ["Some Node", "Another Node"]
# if node_name_filter:
#     df = df[df['Node Name'].str.lower().isin([n.lower() for n in node_name_filter])]
import pandas as pd
import os
import subprocess
import sys

# Always run the prerequisite script to ensure fresh data
print("="*80)
print("Running prerequisite script: Add POC for bugs.py")
print("="*80)

prerequisite_script = os.path.join(os.path.dirname(__file__), "Add POC for bugs.py")

try:
    result = subprocess.run(
        [sys.executable, prerequisite_script],
        check=True,
        capture_output=False
    )
    print("="*80)
    print("✅ Prerequisite script completed successfully!")
    print("="*80)
except subprocess.CalledProcessError as e:
    print(f"\n❌ ERROR: Failed to run prerequisite script")
    print(f"   Please check the script: {prerequisite_script}")
    exit(1)
except FileNotFoundError:
    print(f"\n❌ ERROR: Prerequisite script not found!")
    print(f"   Expected location: {prerequisite_script}")
    exit(1)

print("\nGenerating pivot tables...")
csv_path = os.path.join(os.path.dirname(__file__), "Final Bug list with POC added.csv")

# Read the CSV file
df = pd.read_csv(csv_path)


# State filter: set to a list of states to include, or None for all
state_filter = None  # Example: ['Active', 'Closed']
if state_filter:
    # Convert both DataFrame and filter values to lower case for case-insensitive comparison
    df = df[df['System.State'].str.lower().isin([s.lower() for s in state_filter])]

pivot_states_1 = [s.lower() for s in ["ready to deploy", "resolved"]]
df1 = df[df['System.State'].str.lower().isin(pivot_states_1)]
pivot1 = pd.pivot_table(
    df1,
    index=["AD POC", "SM POC"],
    columns=["Microsoft.VSTS.Common.Severity"],
    values="System.Title",
    aggfunc="count",
    fill_value=0,
    margins=True,
    margins_name="Grand Total"
)
# Sort by Grand Total in descending order
if "Grand Total" in pivot1.columns:
    pivot1 = pivot1.sort_values(by="Grand Total", ascending=False)

pivot_states_1 = [s.lower() for s in ["ready to deploy", "resolved"]]
df2 = df[~df['System.State'].str.lower().isin(pivot_states_1)]
pivot2 = pd.pivot_table(
    df2,
    index=["AD POC", "SM POC"],
    columns=["Microsoft.VSTS.Common.Severity"],
    values="System.Title",
    aggfunc="count",
    fill_value=0,
    margins=True,
    margins_name="Grand Total"
)
# Sort by Grand Total in descending order
if "Grand Total" in pivot2.columns:
    pivot2 = pivot2.sort_values(by="Grand Total", ascending=False)

def sort_pivot_with_grand_total_last(pivot):
    if "Grand Total" in pivot.index:
        grand_total = pivot.loc[["Grand Total"]]
        pivot_wo_total = pivot.drop("Grand Total")
        pivot_sorted = pivot_wo_total.sort_values(by="Grand Total", ascending=False)
        return pd.concat([pivot_sorted, grand_total])
    return pivot

pivot1 = sort_pivot_with_grand_total_last(pivot1)
pivot2 = sort_pivot_with_grand_total_last(pivot2)

def sort_pivot_ad_poc_grand_total(pivot):
    if "Grand Total" in pivot.index:
        grand_total = pivot.loc[["Grand Total"]]
        pivot_wo_total = pivot.drop("Grand Total")
        # Sort by AD POC, then by Grand Total descending
        pivot_wo_total = (
            pivot_wo_total
            .reset_index()
            .sort_values(by=["AD POC", "Grand Total"], ascending=[True, False])
            .set_index(["AD POC", "SM POC"])
        )
        return pd.concat([pivot_wo_total, grand_total])
    return pivot

pivot1 = sort_pivot_ad_poc_grand_total(pivot1)
pivot2 = sort_pivot_ad_poc_grand_total(pivot2)

# Write both pivots to a single Excel file with two sheets
output_excel = os.path.join(os.path.dirname(__file__), "Final Pivot to submit.xlsx")
with pd.ExcelWriter(output_excel) as writer:
    pivot1.to_excel(writer, sheet_name="Ready_Resolved")
    pivot2.to_excel(writer, sheet_name="Other_States")
print(f"Both pivots saved to: {output_excel}")
