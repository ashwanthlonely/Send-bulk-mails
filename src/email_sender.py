import smtplib
import imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime
import time
from tqdm import tqdm
import pandas as pd

from src.config import load_credentials, save_credentials
from src.utils import load_sent_emails, save_sent_email

MAX_EMAILS_PER_DAY = 400

class EmailSender:
    def __init__(self):
        self.credentials = load_credentials()
        self.current_account_index = 0
        self.emails_sent = 0
        self.sent_emails = load_sent_emails()
        self.df = pd.DataFrame()  # Initialize an empty DataFrame

    def connect_to_email(self, account):
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        imap_server = 'imap.gmail.com'
        imap_port = 993

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(account['email'], account['password'])

        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.login(account['email'], account['password'])

        return server, mail

    def send_emails(self, subject, body, signature, cc_email, attachment_path=None):
        if not subject or not body or not self.credentials:
            raise ValueError("Subject, message, and at least one credential are required.")

        if 'Email Ids' not in self.df.columns or 'STUDENTS NAMES' not in self.df.columns:
            raise ValueError("Excel must have 'Email Ids' and 'STUDENTS NAMES' columns.")

        server, mail = self.connect_to_email(self.credentials[self.current_account_index])

        with tqdm(total=len(self.df)) as pbar:
            for index, row in self.df.iterrows():
                email = row['Email Ids']
                name = row['STUDENTS NAMES']
                
                if email in self.sent_emails:
                    continue

                personalized_body = f"Dear {name},\n\n{body}\n\n{signature}"
                message = MIMEMultipart()
                message.attach(MIMEText(personalized_body, 'plain'))

                if attachment_path:
                    with open(attachment_path, 'rb') as attachment:
                        image_mime = MIMEImage(attachment.read())
                        image_mime.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                        message.attach(image_mime)

                message['From'] = self.credentials[self.current_account_index]['email']
                message['To'] = email
                message['Cc'] = cc_email
                message['Subject'] = subject

                try:
                    server.sendmail(self.credentials[self.current_account_index]['email'], [email, cc_email], message.as_string())
                    self.sent_emails.add(email)
                    save_sent_email(email)

                    self.df.at[index, 'Sent Status'] = 'Sent'
                    self.df.at[index, 'Sent Date'] = datetime.now().strftime('%Y-%m-%d')
                    self.df.at[index, 'Sent Time'] = datetime.now().strftime('%H:%M:%S')
                except smtplib.SMTPException as e:
                    print(f"Failed to send email to {email}: {e}")
                    self.df.at[index, 'Sent Status'] = f'Failed: {e}'

                self.emails_sent += 1
                pbar.update(1)

                if self.emails_sent >= MAX_EMAILS_PER_DAY:
                    server.quit()
                    mail.logout()
                    self.current_account_index += 1
                    self.emails_sent = 0

                    if self.current_account_index >= len(self.credentials):
                        print("All accounts have reached the limit. Waiting for 24 hours.")
                        time.sleep(86400)  # Wait for 24 hours
                        self.current_account_index = 0

                    server, mail = self.connect_to_email(self.credentials[self.current_account_index])

        server.quit()
        mail.logout()

        # Save the updated Excel file with statuses and dates
        self.df.to_excel('updated_email_status.xlsx', index=False)
