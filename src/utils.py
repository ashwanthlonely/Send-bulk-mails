import os

SENT_EMAILS_FILE = 'sent_emails.txt'

def load_sent_emails():
    sent_emails = set()
    if os.path.exists(SENT_EMAILS_FILE):
        with open(SENT_EMAILS_FILE, 'r') as f:
            sent_emails.update(line.strip() for line in f)
    return sent_emails

def save_sent_email(email):
    with open(SENT_EMAILS_FILE, 'a') as f:
        f.write(email + '\n')
