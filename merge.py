import os
import pandas as pd

# Define folder path containing Excel files
folder_path = r'F:\Non voice data'

# Define path for saving merged file
save_path = r'F:\Nonvoice_merged_data.xlsx'

# List to hold dataframes
all_dataframes = []

# Loop through all files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        file_path = os.path.join(folder_path, file_name)
        # Read the excel file
        df = pd.read_excel(file_path)
        all_dataframes.append(df)

# Concatenate all dataframes
merged_data = pd.concat(all_dataframes, ignore_index=True)

# Save the merged data to the specified path
merged_data.to_excel(save_path, index=False)

print(f"Files merged successfully. Saved at: {save_path}")
