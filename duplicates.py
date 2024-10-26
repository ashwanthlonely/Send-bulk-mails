import pandas as pd

# File path of the original Excel file
file_path = r'C:\Users\ashwa\OneDrive\Desktop\Merged_File.xlsx'

# Load the Excel file into a pandas DataFrame
df = pd.read_excel(file_path)

# Remove duplicate rows
df_cleaned = df.drop_duplicates()

# Save the cleaned DataFrame to a new Excel file
output_path = r'C:\Users\ashwa\OneDrive\Desktop\merged_file_latest.xlsx'
df_cleaned.to_excel(output_path, index=False)

print(f"Duplicate rows removed. The updated file has been saved as {output_path}")
