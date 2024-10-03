import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tqdm import tqdm
import imaplib
import yaml
import time
from datetime import datetime, timedelta
import os

# Load email accounts and their daily limits from YAML
yaml_path = 'email_accounts.yaml'
if os.path.exists(yaml_path):
    with open(yaml_path, 'r') as f:
        email_accounts_config = yaml.safe_load(f)
    email_accounts = email_accounts_config['email_accounts']
    print(f"Loaded email accounts.")
else:
    print(f"YAML file not found at {yaml_path}")
    raise FileNotFoundError(f"YAML file not found at {yaml_path}")

# Load the Excel file and add a "Status" column if it doesn't exist
excel_path = r"C:\Users\ashwa\OneDrive\Desktop\Prismire.xlsx"
if os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    if 'Status' not in df.columns:
        df['Status'] = ''  # Create a new column for Status if it doesn't exist
else:
    print(f"Excel file not found at {excel_path}")
    raise FileNotFoundError(f"Excel file not found at {excel_path}")

# Set up the SMTP and IMAP server details
smtp_server = 'smtpout.secureserver.net'
smtp_port = 587
imap_server = 'imap.secureserver.net'
imap_port = 993

# Email sending limit per account
email_limit_per_account = 500
daily_reset_hours = 24  # Reset email count every 24 hours

# Function to connect to SMTP server
def connect_smtp(email, password):
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email, password)
        print(f"Connected to SMTP server for {email}")
        return server
    except smtplib.SMTPAuthenticationError as e:
        print(f"SMTP Authentication error for {email}: {e}")
        raise e
    except Exception as e:
        print(f"Error connecting to SMTP server: {e}")
        raise e

# Function to connect to IMAP server
def connect_imap(email, password):
    try:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.login(email, password)
        print(f"Connected to IMAP server for {email}")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"IMAP Authentication error for {email}: {e}")
        raise e
    except Exception as e:
        print(f"Error connecting to IMAP server: {e}")
        raise e

# Function to update email account counts in YAML
def update_account_counts():
    with open(yaml_path, 'w') as f:
        yaml.dump({'email_accounts': email_accounts}, f)
    print("Updated account counts in YAML")

# Function to check if the daily limit should be reset
def reset_daily_count(account):
    last_sent_date = datetime.strptime(account['last_sent_date'], "%Y-%m-%d")
    if datetime.now() - last_sent_date > timedelta(hours=daily_reset_hours):
        account['daily_count'] = 0
        account['last_sent_date'] = datetime.now().strftime("%Y-%m-%d")
        update_account_counts()
        print(f"Daily count reset for {account['email']}")

# Initialize progress bar for only emails that haven't been sent
emails_left = df[df['Status'] != 'Sent'].shape[0]

pbar = tqdm(total=emails_left, desc='Sending emails', unit='email')


# Count of total emails sent
total_emails_sent = 0

# Initialize account index and count for emails sent from each account
account_index = 0
server = None
mail = None

# Iterate through each row in the DataFrame and send an email
for index, row in df.iterrows():
    if row['Status'] == 'Sent':
        pbar.update(1)  # Skip if already sent
        continue

    name = row['Name']
    email = row['Email ID']
    cc_email = ''  # Update with actual CC if necessary

    # Check and reset the daily count if necessary
    reset_daily_count(email_accounts[account_index])

    # If the account's daily limit is reached, switch to the next account
    while email_accounts[account_index]['daily_count'] >= email_limit_per_account:
        print(f"Account {email_accounts[account_index]['email']} reached its limit. Switching accounts.")
        account_index += 1
        if account_index >= len(email_accounts):
            print("All email accounts have reached the limit for today.")
            break
        reset_daily_count(email_accounts[account_index])

    if account_index >= len(email_accounts):
        print("No more available accounts. Exiting.")
        break  # No more accounts available

    # Connect to SMTP and IMAP servers if not already connected
    if not server or not mail:
        server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
        mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

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
    if cc_email:
        message['Cc'] = cc_email
    message['Subject'] = subject

    try:
        # Send the email
        server.sendmail(email_accounts[account_index]['email'], [email], message.as_string())
        print(f"Email sent to {email} using {email_accounts[account_index]['email']}")

        # Append the email to the 'Sent' mailbox
        mail.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), message.as_bytes())

        # Mark email as sent and update count
        df.at[index, 'Status'] = 'Sent'
        total_emails_sent += 1
        email_accounts[account_index]['daily_count'] += 1
        update_account_counts()

    except Exception as e:
        print(f"Error sending email to {email}: {e}")
        df.at[index, 'Status'] = f'Failed: {str(e)}'

    # Update Excel file after each email to prevent duplicates on rerun
    df.to_excel(excel_path, index=False)
    print(f"Excel updated after sending email to {email}")

    # Update progress bar
    pbar.update(1)

    # Check if account has reached its daily limit and switch
    if email_accounts[account_index]['daily_count'] >= email_limit_per_account:
        server.quit()
        mail.logout()
        account_index += 1
        if account_index < len(email_accounts):
            server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
            mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

# Close the progress bar
pbar.close()

# Print total emails sent
print(f"Total emails sent: {total_emails_sent}")

# Quit SMTP and IMAP servers
if server:
    print("Closing SMTP server")
    server.quit()
if mail:
    print("Logging out from IMAP")
    mail.logout()
