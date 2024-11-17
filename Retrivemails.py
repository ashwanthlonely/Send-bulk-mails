import yaml
import imaplib
import email
import re

# Email server configuration
IMAP_SERVER = 'imap.secureserver.net'  # Set your IMAP server here (e.g., 'imap.gmail.com' for Gmail)
IMAP_PORT = 993  # Default IMAP SSL port

# Load YAML file with account credentials
with open("email_accounts.yaml", "r") as file:
    accounts = yaml.safe_load(file)["email_accounts"]

# Define the regex pattern for "I am interested" in all variations
search_phrases = r'\b[Ii][ ]?[Aa][Mm][ ]?[Ii][Nn][Tt][Ee][Rr][Ee][Ss][Tt][Ee][Dd]\b'

found_emails = set()  # Use a set to avoid duplicates

# Function to connect to IMAP, search emails, and collect matching addresses
def search_emails(account):
    try:
        print(f"Connecting to account: {account['email']}")
        # Connect to the IMAP server with SSL
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(account["email"], account["password"])
        mail.select("inbox")

        # Search all emails in the inbox
        status, messages = mail.search(None, 'ALL')
        if status != "OK":
            print(f"Failed to search emails in {account['email']}")
            return

        email_ids = messages[0].split()
        print(f"Found {len(email_ids)} emails in inbox for {account['email']}")

        for eid in email_ids:
            _, msg_data = mail.fetch(eid, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            # Extract email body from both plain text and HTML
            body = ""
            for part in msg.walk():
                if part.get_content_type() in ["text/plain", "text/html"]:
                    try:
                        body += part.get_payload(decode=True).decode(errors="ignore")
                    except Exception as e:
                        print(f"Error decoding message part: {e}")

            # Search for variations of "I am interested" using regex
            if re.search(search_phrases, body, re.IGNORECASE):
                from_addr = msg.get("From")
                print(f"Match found from {from_addr}")
                found_emails.add(from_addr)

        mail.logout()
    except Exception as e:
        print(f"Error with account {account['email']}: {e}")

# Iterate over each account
for account in accounts:
    search_emails(account)

# Save found emails to a text file
with open("found_emails.txt", "w") as file:
    for email in found_emails:
        file.write(email + "\n")

print("found_emails.txt")
