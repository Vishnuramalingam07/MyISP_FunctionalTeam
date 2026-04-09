import pandas as pd
import re

# Load the CSV file
df = pd.read_csv(r'C:\Users\d.sampathkumar\GHC files\Nov 08 -2025 Defects - Action on PT team.csv')

# Combine states
combine_states = ['In Test', 'PT In Test', 'Ready to Test']
df['State'] = df['State'].apply(lambda x: 'PT in test' if x in combine_states else x)

# Extract only the value inside < > and remove '@accenture.com'
def extract_name(text):
    if pd.isna(text):
        return ''
    match = re.search(r'<([^@<>]+)', text)
    return match.group(1).strip() if match else text

df['Assigned To'] = df['Assigned To'].apply(extract_name)

# Prepare output DataFrame for main sheet, sorted by Assigned To within each state
output_rows = []
for state in df['State'].unique():
    subset = df[df['State'] == state][['ID', 'Assigned To']].sort_values(by='Assigned To')
    for _, row in subset.iterrows():
        output_rows.append({'State': state, 'ID': row['ID'], 'Assigned To': row['Assigned To']})

output_df = pd.DataFrame(output_rows, columns=['State', 'ID', 'Assigned To'])

# Prepare pivot table for second sheet
pivot_df = df.groupby('State').size().reset_index(name='Number of Defects')

# Save to Excel with two sheets
with pd.ExcelWriter(r'C:\Users\d.sampathkumar\GHC files\Output_Action with PT.xlsx') as writer:
    output_df.to_excel(writer, sheet_name='Sheet1', index=False)
    pivot_df.to_excel(writer, sheet_name='pivot', index=False)

print("Output saved to Output_Action with PT.xlsx with pivot tab.")