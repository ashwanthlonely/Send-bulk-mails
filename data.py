import os
import pandas as pd
from tqdm import tqdm

# Set the directory path
dir_path = r'C:\Users\ashwa\Downloads\drive-download-20240918T115343Z-001'

# Get a list of all Excel files in the directory
excel_files = [f for f in os.listdir(dir_path) if f.endswith('.xlsx')]

# Create a list to store the dataframes
dfs = []

# Specify the column name to merge on
merge_column = 'Column_Name'  # Replace with the actual column name

# Iterate over the Excel files and read them into dataframes
for file in tqdm(excel_files, desc='Reading Excel files'):
    file_path = os.path.join(dir_path, file)
    df = pd.read_excel(file_path)
    dfs.append(df)

# Merge the dataframes based on the specified column
merged_df = pd.concat(dfs, ignore_index=True)
merged_df = merged_df.merge(merged_df, on=merge_column, how='outer')

# Save the merged dataframe to a new Excel file
output_file = os.path.join(dir_path, 'Merged_File.xlsx')
merged_df.to_excel(output_file, index=False)

print("Merged file saved to", output_file)