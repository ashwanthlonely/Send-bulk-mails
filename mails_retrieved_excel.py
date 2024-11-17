import pandas as pd

# Load email IDs from text file
with open("found_emails.txt", "r") as file:
    email_ids = set(line.strip() for line in file)

# Load the provided Excel file
input_excel = r"C:\Users\ashwa\OneDrive\Desktop\Merged_File.xlsx"  # Replace with your actual file name
df = pd.read_excel(input_excel)

# Assuming the email column has a standard name like 'Email', modify if necessary
email_column = "Email ID"  # Replace with the actual column name in your Excel

# Filter rows where email is in the loaded email IDs list
matched_rows = df[df[email_column].isin(email_ids)]
print("matched_rows")
# Save the matched rows to a new Excel file with the same header
output_excel = "filtered_output_extmails.xlsx"
matched_rows.to_excel(output_excel, index=False)

print(f"Filtered rows saved to {output_excel}")
