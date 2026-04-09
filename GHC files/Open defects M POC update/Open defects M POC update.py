import pandas as pd

# Absolute file paths
open_defects_path = r'C:\Users\d.sampathkumar\GHC files\Open defects M POC update\Open Defects.csv'
mapping_path = r'C:\Users\d.sampathkumar\GHC files\POC mapping\M POC Mapping.csv'
output_path = r'C:\Users\d.sampathkumar\GHC files\Open defects M POC update\Open Defects with M POC.csv'

# Read CSV files
open_defects = pd.read_csv(open_defects_path, dtype=str)
mapping = pd.read_csv(mapping_path, dtype=str)

# Clean whitespace and convert to lowercase for comparison
open_defects['Node Name_clean'] = open_defects['Node Name'].fillna('').str.strip().str.lower()
mapping['Node Name_clean'] = mapping['Node Name'].fillna('').str.strip().str.lower()
mapping['ExternalRef ID'] = mapping['ExternalRef ID'].fillna('').str.strip()
mapping['Contract ID Info'] = mapping['Contract ID Info'].fillna('').str.strip()

# Create mapping dictionaries (lowercase node names)
node_to_extref = dict(zip(mapping['Node Name_clean'], mapping['ExternalRef ID']))
node_to_contract = dict(zip(mapping['Node Name_clean'], mapping['Contract ID Info']))

# Populate ExternalRef ID only if Node Name is present in mapping (case-insensitive)
def get_extref(node):
    key = str(node).strip().lower() if node else ''
    return node_to_extref.get(key, '')

# Populate Contract ID Info based on Node Name (case-insensitive)
def get_contract(node):
    key = str(node).strip().lower() if node else ''
    return node_to_contract.get(key, '')

open_defects['ExternalRef ID'] = open_defects['Node Name'].apply(get_extref)
open_defects['Contract ID Info'] = open_defects['Node Name'].apply(get_contract)

# Select only required columns in specified order
output_df = open_defects[['ID', 'Title', 'Work Item Type', 'ExternalRef ID', 'Contract ID Info']]

# Save to new CSV
output_df.to_csv(output_path, index=False)

print(f"Output file saved with only required columns: {output_path}")