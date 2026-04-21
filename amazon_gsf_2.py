import pandas as pd
import glob
import os
from pathlib import Path

# Set the path to the folder containing the Excel files
folder_path = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs\PO Lines'
output_path = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs\2-PO Lines '

# Create a pattern to match all Excel files in the folder
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# Initialize an empty list to hold individual dataframes
dataframes = []

# Loop through each Excel file
for file in excel_files:
    try:
        # Read the Excel file into a dataframe
        df = pd.read_excel(file)

        # Extract PO number from file path
        po_number = Path(file).stem
        print(po_number)

        # Drop the first row after the headers
        df = df.iloc[1:].reset_index(drop=True)

        # Add file name (PO Number) as a column
        df['Customer WO#'] = po_number

        dataframes.append(df)
        print(f"Loaded {file} with {len(df)} rows.")
    except Exception as e:
        print(f"Failed to load {file}: {e}")

# Combine all dataframes into a single dataframe
combined_df = pd.concat(dataframes, ignore_index=True)

# Extract Available amount from description and convert to number
combined_df['Available Amount'] = combined_df['Item Amount'].str.extract(r'Available: \$(\d{1,3}(?:,\d{3})*(?:\.\d+)?)')[0].str.replace(',', '').astype(float)

# Extract Site Name from Description
combined_df['Site Name'] = combined_df['Description'].str.split('-', n=1).str[0].str.strip()
combined_df['Site Name'] = combined_df['Site Name'].str.split(' ', n=1).str[0].str.strip()

# Print out the final shape
print(f"\nCombined dataframe has {combined_df.shape[0]} rows and {combined_df.shape[1]} columns.")
print(combined_df)

# Optional: Save to a new Excel or CSV file
combined_df.to_excel(output_path+'combined_output.xlsx', index=False)
