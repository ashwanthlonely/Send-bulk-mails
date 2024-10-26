import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tqdm import tqdm
import imaplib
import yaml
from datetime import datetime, timedelta

# Load email accounts from YAML file
with open('email_accounts.yaml', 'r') as f:
    email_accounts_config = yaml.safe_load(f)
email_accounts = email_accounts_config['email_accounts']

# Load DataFrame and add "Status" column if it doesn't exist
excel_path = r"C:\Users\ashwa\OneDrive\Desktop\Nonvoice_merged_data1.xlsx"
df = pd.read_excel(excel_path)
if 'Status' not in df.columns:
    df['Status'] = ''

# Server details
smtp_server = 'smtpout.secureserver.net'
smtp_port = 587
imap_server = 'imap.secureserver.net'
imap_port = 993

# Email sending limit per account
email_limit_per_account = 500
email_refresh_interval = timedelta(days=1)

# Function to update email count and timestamp in YAML file
def update_email_count(account_index, emails_sent):
    email_accounts_config['email_accounts'][account_index]['emails_sent'] = emails_sent
    email_accounts_config['email_accounts'][account_index]['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open('email_accounts.yaml', 'w') as f:
        yaml.safe_dump(email_accounts_config, f)

# Function to reset email count if 24 hours passed
def reset_email_count_if_needed(account):
    if 'emails_sent' not in account:
        account['emails_sent'] = 0
    if 'last_sent' not in account:
        account['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    last_sent_time = datetime.strptime(account['last_sent'], '%Y-%m-%d %H:%M:%S')
    if datetime.now() - last_sent_time >= email_refresh_interval:
        account['emails_sent'] = 0
        return True
    return False

# Function to find the next available account
def get_next_available_account():
    for index, account in enumerate(email_accounts):
        if reset_email_count_if_needed(account) or account['emails_sent'] < email_limit_per_account:
            return index
    return None

# Connect to SMTP
def connect_smtp(email, password):
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email, password)
    return server

# Connect to IMAP
def connect_imap(email, password):
    mail = imaplib.IMAP4_SSL(imap_server, imap_port)
    mail.login(email, password)
    return mail

# Initialize and get the first available account
account_index = get_next_available_account()
if account_index is None:
    print("No available accounts with remaining quota.")
    exit()

emails_sent_from_current_account = email_accounts[account_index]['emails_sent']

# Email sending preparation
server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
pbar = tqdm(total=len(df), desc='Sending emails', unit='email')
total_emails_sent = 0
max_retries = 3

for index, row in df.iterrows():
    if row['Status'] == 'Sent':
        pbar.update(1)
        continue

    name = row['Name']
    email = row['Email ID']
    cc_email = ''

    subject = "Job-Assured Internship Program Invitation"
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
            server.sendmail(email_accounts[account_index]['email'], [email], message.as_string())
            mail.append('Sent', None, None, message.as_bytes())

            df.at[index, 'Status'] = 'Sent'
            total_emails_sent += 1
            emails_sent_from_current_account += 1

            update_email_count(account_index, emails_sent_from_current_account)
            break
        except Exception as e:
            retries += 1
            if retries >= max_retries:
                df.at[index, 'Status'] = f'Failed: {str(e)}'
                print(f"Failed to send email to {email}. Error: {str(e)}")

    pbar.update(1)

    if emails_sent_from_current_account >= email_limit_per_account:
        server.quit()
        mail.logout()

        account_index = get_next_available_account()
        if account_index is None:
            print("No available accounts with remaining quota.")
            break

        emails_sent_from_current_account = email_accounts[account_index]['emails_sent']
        server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
        mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

pbar.close()
print(f"Total emails sent: {total_emails_sent}")

df.to_excel(excel_path, index=False)

server.quit()
mail.logout()
