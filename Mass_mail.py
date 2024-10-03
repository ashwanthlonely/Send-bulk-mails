import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tqdm import tqdm
import imaplib
import yaml

# Load email accounts from YAML file
with open('email_accounts.yaml', 'r') as f:
    email_accounts_config = yaml.safe_load(f)
email_accounts = email_accounts_config['email_accounts']

# Load your DataFrame and add a "Status" column if it doesn't exist
excel_path = r"C:\Users\ashwa\OneDrive\Desktop\Merged_File.xlsx"
df = pd.read_excel(excel_path)
if 'Status' not in df.columns:
    df['Status'] = ''  # Create a new column for Status if it doesn't exist

# Set up the SMTP server details
smtp_server = 'smtpout.secureserver.net'
smtp_port = 587

# Set up the IMAP server details
imap_server = 'imap.secureserver.net'
imap_port = 993

# Email sending limit per account
email_limit_per_account = 500

# Function to connect to SMTP server
def connect_smtp(email, password):
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email, password)
    return server

# Function to connect to IMAP server
def connect_imap(email, password):
    mail = imaplib.IMAP4_SSL(imap_server, imap_port)
    mail.login(email, password)
    return mail

# Initialize account index and count for emails sent from each account
account_index = 0
emails_sent_from_current_account = 0

# Calculate the number of emails left to send
emails_left = df[df['Status'] != 'Sent'].shape[0]
print(f"Emails left to send: {emails_left}")

# Connect to the first email account's SMTP and IMAP
server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

# Initialize progress bar
total_emails = len(df)
pbar = tqdm(total=emails_left, desc='Sending emails', unit='email')

# Count of total emails sent
total_emails_sent = 0

# Iterate through each row in the DataFrame and send an email
for index, row in df.iterrows():
    # Check if the email was already sent (by checking the "Status" column)
    if row['Status'] == 'Sent':
        pbar.update(1)
        continue  # Skip this email since it was already sent

    name = row['Name']
    email = row['Email ID']
    cc_email = ''

    # Create the email content
    subject = '5-Day Workshop with Confirmed Job Opportunity – Exclusive Offer!'
    body = f"""Dear {name}, 

We are excited to invite you to a 5-day intensive workshop training designed exclusively for freshers. This program not only equips you with essential technical and professional skills but also provides a direct path to a confirmed job opportunity with one of our esteemed clients.

Details of the Workshop: Fresher( Java , Data analyst with AI , DevOps  ) 

Duration: 5 days
Training Focus: [Brief description of the skills/technologies covered]
Job Offer: Upon successful completion of the workshop and selection by our client, you will receive a confirmed job offer.
Salary Package: ₹5,00,000 to ₹6,00,000 per annum (depending on your performance and the client's evaluation).
Placement Fee Structure: Upon securing a job with our client, there will be a placement fee of ₹2,50,000. This fee is to be paid after you have been selected by the client and received the job offer.

Best Regards,
HR Department
"""

    message = MIMEMultipart()
    message.attach(MIMEText(body, 'plain'))

    message['From'] = email_accounts[account_index]['email']
    message['To'] = email
    message['Cc'] = cc_email
    message['Subject'] = subject

    try:
        # Send the email
        server.sendmail(email_accounts[account_index]['email'], [email], message.as_string())

        # Append the sent email to the 'Sent' mailbox
        mail.append('Sent', None, None, message.as_bytes())

        # Mark email as sent in the DataFrame
        df.at[index, 'Status'] = 'Sent'
        total_emails_sent += 1
        emails_sent_from_current_account += 1
    except Exception as e:
        # Mark the email as failed if there's an error
        df.at[index, 'Status'] = f'Failed: {str(e)}'

    # Update progress bar
    pbar.update(1)

    # Check if the current account has reached its limit
    if emails_sent_from_current_account >= email_limit_per_account:
        # Logout and switch to the next account
        server.quit()
        mail.logout()

        account_index += 1
        if account_index >= len(email_accounts):
            print("All email accounts have reached the limit.")
            break

        # Reset the counter for emails sent from the current account
        emails_sent_from_current_account = 0

        # Connect to the next email account
        server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
        mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

# Close the progress bar
pbar.close()

# Print total emails sent after the process is completed
print(f"Total emails sent: {total_emails_sent}")

# Save the updated Excel file
df.to_excel(excel_path, index=False)

# Quit the SMTP server and IMAP logout
server.quit()
mail.logout()
