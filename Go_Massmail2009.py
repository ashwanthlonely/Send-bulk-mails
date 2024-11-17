import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tqdm import tqdm
import imaplib
import yaml
import time
from datetime import datetime, timedelta

# Load email accounts from YAML file
with open('email_accounts.yaml', 'r') as f:
    email_accounts_config = yaml.safe_load(f)
email_accounts = email_accounts_config['email_accounts']

# Load your DataFrame and add a "Status" column if it doesn't exist
excel_path = r"C:\Users\ashwa\OneDrive\Desktop\Nonvoice_merged_data1.xlsx"
df = pd.read_excel(excel_path)
if 'Status' not in df.columns:
    df['Status'] = ''  # Create a new column for Status if it doesn't exist

# Set up the SMTP and IMAP server details
smtp_server = 'smtpout.secureserver.net'
smtp_port = 587
imap_server = 'imap.secureserver.net'
imap_port = 993

# Email sending limit per account
email_limit_per_account = 500
email_refresh_interval = timedelta(days=1)  # 24 hours

# Function to update email count and timestamp in the YAML file
def update_email_count(account_index, emails_sent):
    email_accounts_config['email_accounts'][account_index]['emails_sent'] = emails_sent
    email_accounts_config['email_accounts'][account_index]['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open('email_accounts.yaml', 'w') as f:
        yaml.safe_dump(email_accounts_config, f)

# Function to check if the limit should be reset (24 hours passed)
def check_reset_limit(account_index):
    # Initialize fields if they don't exist
    if 'emails_sent' not in email_accounts[account_index]:
        email_accounts[account_index]['emails_sent'] = 0
    if 'last_sent' not in email_accounts[account_index]:
        email_accounts[account_index]['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    last_sent_str = email_accounts[account_index]['last_sent']
    last_sent_time = datetime.strptime(last_sent_str, '%Y-%m-%d %H:%M:%S')
    
    # Check if the 24-hour limit reset is applicable
    if datetime.now() - last_sent_time >= email_refresh_interval:
        email_accounts[account_index]['emails_sent'] = 0  # Reset the email count
        with open('email_accounts.yaml', 'w') as f:
            yaml.safe_dump(email_accounts_config, f)
        print(f"Email count reset for account {email_accounts[account_index]['email']} due to 24-hour refresh.")

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

# Initialize account index and retrieve the email count and last sent time
account_index = 0
check_reset_limit(account_index)  # Reset the limit if 24 hours have passed
emails_sent_from_current_account = email_accounts[account_index]['emails_sent']

# Calculate the number of emails left to send
emails_left = df[df['Status'] != 'Sent'].shape[0]
print(f"Emails left to send: {emails_left}")

# Connect to the first email account's SMTP and IMAP
server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

# Initialize progress bar
total_emails = len(df)
pbar = tqdm(total=total_emails, desc='Sending emails', unit='email')

# Count of total emails sent
total_emails_sent = 0

# Retry mechanism for failed emails
max_retries = 3

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
    subject = 'Job Assured Training Program Invitation!'
    body = f"""<b>Dear {name}</b>, 
<p><b>We are pleased to invite you to our Job-Assured Training Program</b>, which will train you as a <b>Full Stack Java Developer, Data Analyst</b> with AI, Python, and SAP skills. You pay the program fees only after receiving an offer letter from one of our clients.</p>
<p><b>This is a golden opportunity for Freshers and working professionals who wants to switch into IT career.</b></p>
<b>Program Cost:</b> Rs. 2,50,000 + Taxes (loan options available)

<b>Program Details:</b>

<ul>
  <li><b>Program Name:</b> Job-Assured IT Programs</li>
  <li><b>Salary Package:</b> CTC of Rs. 4.5 to 5.5 lakhs per annum</li>
  <li><b>Program Duration:</b> 2 months</li>
  <li><b>Selection Process:</b> Initial Screening &gt;&gt; Assessment &gt;&gt; Interview &gt;&gt; Provisional Offer Letter &gt;&gt; Training &gt;&gt; Join Company</li>
</ul>

<b>Eligibility Criteria:</b>

<ul>
  <li>B.E/B.Tech graduates (<b>CS, IT, & Electronics graduates preferred</b>)</li>
  <li>Graduates from the <b>2015 to 2023 batch</b></li>
  <li><b>Minimum 55% marks</b> or equivalent across 10th, 12th, and UG</li>
</ul>

<b>Job Location:</b> Hyderabad<br>
<b>Program Cost:</b> Rs. 2,50,000 + Taxes (loan options available)

<p>With over <b>5+ years of experience</b> in successfully shaping candidates' futures, we ensure a smooth transition into your IT career.</p>

<p>Additionally, we request that <b>one of your family members speaks with us</b> before you join our program. For security purposes, you will be required to submit your <b>educational certificates</b> and a <b>cheque</b>.</p>

<b>Loan Processing Fees:</b> Rs. 20,000/-

<p>If you have any further questions, please feel free to contact us.</p>

<b>Regards<br> Hr-Team<br> Hyderabad, Hi-tech city,<br> Ph: +91-8121698002, +91-9030216038<b>

"""

    message = MIMEMultipart()
    message.attach(MIMEText(body, 'html'))

    message['From'] = email_accounts[account_index]['email']
    message['To'] = email
    message['Cc'] = cc_email
    message['Subject'] = subject

    retries = 0
    while retries < max_retries:
        try:
            # Send the email
            server.sendmail(email_accounts[account_index]['email'], [email], message.as_string())

            # Append the sent email to the 'Sent' mailbox
            mail.append('Sent', None, None, message.as_bytes())

            # Mark email as sent in the DataFrame
            df.at[index, 'Status'] = 'Sent'
            total_emails_sent += 1
            emails_sent_from_current_account += 1

            # Update email count and timestamp in the YAML file
            update_email_count(account_index, emails_sent_from_current_account)
            break
        except Exception as e:
            retries += 1
            if retries >= max_retries:
                # Mark the email as failed if retries are exhausted
                df.at[index, 'Status'] = f'Failed: {str(e)}'
                print(f"Failed to send email to {email}. Error: {str(e)}")

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

        # Reset the email count if 24 hours have passed for the new account
        check_reset_limit(account_index)

        # Get the updated count for the new account
        emails_sent_from_current_account = email_accounts[account_index]['emails_sent']

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