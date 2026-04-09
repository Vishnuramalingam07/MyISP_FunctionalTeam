import pandas as pd
from openpyxl import load_workbook

# Path to your input Excel file
input_file = r"C:\Users\d.sampathkumar\GHC files\File to consolidate.xlsx"
# Path for the output consolidated file
output_file = r"C:\Users\d.sampathkumar\GHC files\Consolidated File.xlsx"

# Load workbook to check for hidden sheets
wb = load_workbook(input_file)

# Get list of visible sheet names only
visible_sheets = [sheet.title for sheet in wb.worksheets if sheet.sheet_state == 'visible']
wb.close()

print(f"Found {len(visible_sheets)} visible sheets to consolidate")

# Read only visible sheets into a dictionary of DataFrames
all_sheets = pd.read_excel(input_file, sheet_name=visible_sheets)

# Collect all DataFrames, ensuring headers are included only once
consolidated_df = pd.DataFrame()
for i, (sheet_name, df) in enumerate(all_sheets.items()):
    # Drop completely empty rows (optional, but often useful)
    df = df.dropna(how='all')
    
    # Add Tab Name column at the beginning
    df.insert(0, 'Tab Name', sheet_name)
    
    if i == 0:
        consolidated_df = df.copy()
    else:
        # Simply concatenate all sheets, aligning by column names
        # This will handle missing columns gracefully
        consolidated_df = pd.concat([consolidated_df, df], ignore_index=True, sort=False)

# Save the consolidated DataFrame to a new Excel file
consolidated_df.to_excel(output_file, index=False)
print(f"Consolidation complete! Output saved as: {output_file}")