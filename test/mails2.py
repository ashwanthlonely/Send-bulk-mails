import pandas as pd
import smtplib
import imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from tkinter import Tk, Label, Entry, Button, filedialog, Text, messagebox, Listbox, MULTIPLE
import yaml
import os
import time
from tqdm import tqdm
from tkinter import simpledialog
from datetime import datetime

# Global variables
credentials = []
current_account_index = 0
emails_sent = 0
max_emails_per_day = 400
sent_emails = set()
attachment_path = None
df = pd.DataFrame()  # Initialize an empty DataFrame

# Load credentials from a YAML file
def load_credentials():
    global credentials
    try:
        with open('credentials.yaml', 'r') as file:
            credentials = yaml.safe_load(file)['accounts']
    except FileNotFoundError:
        credentials = []

# Save credentials to a YAML file
def save_credentials():
    global credentials
    with open('credentials.yaml', 'w') as file:
        yaml.dump({'accounts': credentials}, file)

# GUI functions
def upload_excel():
    global df
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filepath:
        df = pd.read_excel(filepath)
        df['Sent Status'] = ''  # Initialize a column for email sending status
        df['Sent Date'] = ''  # Initialize a column for sent date
        df['Sent Time'] = ''  # Initialize a column for sent time
        messagebox.showinfo("File Uploaded", "Excel file has been uploaded successfully!")

def upload_attachment():
    global attachment_path
    attachment_path = filedialog.askopenfilename(filetypes=[("All files", "*.*")])
    if attachment_path:
        attachment_label.config(text=f"Attachment: {os.path.basename(attachment_path)}")

def add_credential():
    email = simpledialog.askstring("Input", "Enter email:")
    password = simpledialog.askstring("Input", "Enter password:", show='*')
    if email and password:
        credentials.append({'email': email, 'password': password})
        save_credentials()
        update_credentials_list()

def delete_selected_credentials():
    selected = credentials_listbox.curselection()
    if not selected:
        messagebox.showerror("Selection Error", "Please select one or more credentials to delete.")
        return

    if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected credentials?"):
        for index in reversed(selected):  # Delete from the end to avoid index shifting issues
            del credentials[index]
        save_credentials()
        update_credentials_list()

def edit_credential():
    selected = credentials_listbox.curselection()
    if not selected:
        messagebox.showerror("Selection Error", "Please select a credential to edit.")
        return
    
    index = selected[0]
    email = simpledialog.askstring("Edit Email", "Edit email:", initialvalue=credentials[index]['email'])
    password = simpledialog.askstring("Edit Password", "Edit password:", initialvalue=credentials[index]['password'], show='*')
    if email and password:
        credentials[index] = {'email': email, 'password': password}
        save_credentials()
        update_credentials_list()

def update_credentials_list():
    credentials_listbox.delete(0, 'end')
    for account in credentials:
        credentials_listbox.insert('end', account['email'])

def send_emails():
    global emails_sent, current_account_index, df

    subject = subject_entry.get()
    body = message_text.get("1.0", 'end-1c')
    signature = signature_text.get("1.0", 'end-1c')
    cc_email = cc_entry.get()

    if not subject or not body or not credentials:
        messagebox.showerror("Input Error", "Subject, message, and at least one credential are required.")
        return

    if 'Email Ids' not in df.columns or 'STUDENTS NAMES' not in df.columns:
        messagebox.showerror("File Error", "Excel must have 'Email Ids' and 'STUDENTS NAMES' columns.")
        return

    def connect_to_email(account):
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

    sent_emails_file = 'sent_emails.txt'
    if os.path.exists(sent_emails_file):
        with open(sent_emails_file, 'r') as f:
            sent_emails.update(line.strip() for line in f)

    server, mail = connect_to_email(credentials[current_account_index])

    with tqdm(total=len(df)) as pbar:
        for index, row in df.iterrows():
            email = row['Email Ids']
            name = row['STUDENTS NAMES']
            
            if email in sent_emails:
                continue

            personalized_body = f"Dear {name},\n\n{body}\n\n{signature}"
            message = MIMEMultipart()
            message.attach(MIMEText(personalized_body, 'plain'))

            if attachment_path:
                with open(attachment_path, 'rb') as attachment:
                    image_mime = MIMEImage(attachment.read())
                    image_mime.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    message.attach(image_mime)

            message['From'] = credentials[current_account_index]['email']
            message['To'] = email
            message['Cc'] = cc_email
            message['Subject'] = subject

            try:
                server.sendmail(credentials[current_account_index]['email'], [email, cc_email], message.as_string())
                
                # Log the sent email
                sent_emails.add(email)
                with open(sent_emails_file, 'a') as f:
                    f.write(email + '\n')

                # Update DataFrame with success status
                df.at[index, 'Sent Status'] = 'Sent'
                df.at[index, 'Sent Date'] = datetime.now().strftime('%Y-%m-%d')
                df.at[index, 'Sent Time'] = datetime.now().strftime('%H:%M:%S')

            except smtplib.SMTPException as e:
                print(f"Failed to send email to {email}: {e}")
                df.at[index, 'Sent Status'] = f'Failed: {e}'

            emails_sent += 1
            pbar.update(1)

            if emails_sent >= max_emails_per_day:
                server.quit()
                mail.logout()
                current_account_index += 1
                emails_sent = 0

                if current_account_index >= len(credentials):
                    print("All accounts have reached the limit. Waiting for 24 hours.")
                    time.sleep(86400)  # Wait for 24 hours
                    current_account_index = 0

                server, mail = connect_to_email(credentials[current_account_index])

    server.quit()
    mail.logout()
    pbar.close()

    # Save the updated Excel file with statuses and dates
    df.to_excel('updated_email_status.xlsx', index=False)

    messagebox.showinfo("Completed", f"Total emails sent: {len(sent_emails)}")

# Initialize the main window
root = Tk()
root.title("Bulk Email Sender")

# Excel upload
upload_btn = Button(root, text="Upload Excel", command=upload_excel)
upload_btn.grid(row=0, column=0, padx=10, pady=10)

# Subject entry
Label(root, text="Subject:").grid(row=1, column=0, sticky='e')
subject_entry = Entry(root, width=50)
subject_entry.grid(row=1, column=1, padx=10, pady=10)

# CC entry
Label(root, text="CC:").grid(row=2, column=0, sticky='e')
cc_entry = Entry(root, width=50)
cc_entry.grid(row=2, column=1, padx=10, pady=10)

# Message text box
Label(root, text="Message:").grid(row=3, column=0, sticky='ne')
message_text = Text(root, width=50, height=10)
message_text.grid(row=3, column=1, padx=10, pady=10)

# Signature text box
Label(root, text="Signature:").grid(row=4, column=0, sticky='ne')
signature_text = Text(root, width=50, height=5)
signature_text.grid(row=4, column=1, padx=10, pady=10)

# Attachment upload
attachment_label = Label(root, text="No attachment")
attachment_label.grid(row=5, column=0, padx=10, pady=10, sticky='w')
upload_attachment_btn = Button(root, text="Upload Attachment", command=upload_attachment)
upload_attachment_btn.grid(row=5, column=1, padx=10, pady=10, sticky='w')

# Start sending button
start_btn = Button(root, text="Start Sending Emails", command=send_emails)
start_btn.grid(row=6, column=1, padx=10, pady=10)

# Credentials list
Label(root, text="Credentials:").grid(row=7, column=0, padx=10, pady=10, sticky='nw')
credentials_listbox = Listbox(root, selectmode=MULTIPLE)
credentials_listbox.grid(row=7, column=1, padx=10, pady=10, sticky='w')

# Add, Edit, Delete buttons for credentials
add_cred_btn = Button(root, text="Add Credential", command=add_credential)
add_cred_btn.grid(row=8, column=1, padx=10, pady=10, sticky='w')
edit_cred_btn = Button(root, text="Edit Credential", command=edit_credential)
edit_cred_btn.grid(row=8, column=1, padx=10, pady=10, sticky='e')
delete_cred_btn = Button(root, text="Delete Selected Credential(s)", command=delete_selected_credentials)
delete_cred_btn.grid(row=9, column=1, padx=10, pady=10, sticky='e')

# Load existing credentials into the list
load_credentials()
update_credentials_list()

# Run the main loop
root.mainloop()
