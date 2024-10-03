import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import time
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import yaml
from tqdm import tqdm
import os

# Load email accounts from YAML file
def load_email_accounts():
    try:
        with open('email_accounts.yaml', 'r') as f:
            return yaml.safe_load(f)['email_accounts']
    except FileNotFoundError:
        return []

email_accounts = load_email_accounts()
email_limit_per_account = 500  # Limit emails per account

# Global variables for email data
df = None
excel_path = ""
subject = ""
message_body = ""

# Countdown mechanism (initially disabled)
countdown_time = None

# Create the main application window
root = tk.Tk()
root.title("Bulk Email Sender")

# Create a Notebook widget for tabs
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill='both')

### Tab 1: Email Sending Interface ###

# Function to load Excel file and populate dropdown menus
def load_excel():
    global df, excel_path
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_path:
        df = pd.read_excel(excel_path)
        name_combobox['values'] = df.columns.tolist()  # Populate dropdown with columns
        email_combobox['values'] = df.columns.tolist()
        messagebox.showinfo("Success", "Excel file loaded successfully!")
    else:
        messagebox.showerror("Error", "No file selected.")

# Function to set subject and message body
def set_email_data():
    global subject, message_body
    subject = subject_entry.get()
    message_body = body_text.get("1.0", tk.END)
    if not subject or not message_body:
        messagebox.showerror("Error", "Subject and message body cannot be empty.")
    else:
        messagebox.showinfo("Success", "Email data set successfully!")

# Function to send emails
def send_emails():
    global countdown_time
    
    # Check for countdown (prevent sending if within 24 hours)
    if countdown_time and time.time() < countdown_time:
        remaining_time = int(countdown_time - time.time())
        hours, minutes, seconds = remaining_time // 3600, (remaining_time % 3600) // 60, remaining_time % 60
        messagebox.showerror("Quota Exhausted", f"Please wait {hours}h {minutes}m {seconds}s before sending more emails.")
        return

    # Confirm email data, Excel sheet, and column selections are valid
    if df is None or not subject or not message_body:
        messagebox.showerror("Error", "Please load Excel file and set email data before sending.")
        return
    if not name_combobox.get() or not email_combobox.get():
        messagebox.showerror("Error", "Please select Name and Email fields from the dropdowns.")
        return
    
    threading.Thread(target=email_sending_process).start()

def email_sending_process():
    global countdown_time
    
    # Initialize progress bar
    progress_bar['value'] = 0
    progress_bar['maximum'] = len(df)

    account_index = 0
    emails_sent_from_current_account = 0
    total_emails_sent = 0

    # Set up SMTP connection
    server = smtplib.SMTP('smtpout.secureserver.net', 587)
    server.starttls()
    server.login(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

    # Iterate through DataFrame and send emails
    for index, row in df.iterrows():
        if emails_sent_from_current_account >= email_limit_per_account:
            server.quit()
            account_index += 1
            if account_index >= len(email_accounts):
                countdown_time = time.time() + 24 * 3600
                messagebox.showinfo("Info", "Quota exhausted. Emails will resume after 24 hours.")
                break

            # Reset counter and connect to next account
            emails_sent_from_current_account = 0
            server = smtplib.SMTP('smtpout.secureserver.net', 587)
            server.starttls()
            server.login(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

        name_field = name_combobox.get()  # Get selected Name field
        email_field = email_combobox.get()  # Get selected Email field
        name = row[name_field]
        email = row[email_field]

        message = MIMEMultipart()
        message.attach(MIMEText(message_body.format(name=name), 'plain'))
        message['From'] = email_accounts[account_index]['email']
        message['To'] = email
        message['Subject'] = subject

        try:
            server.sendmail(email_accounts[account_index]['email'], email, message.as_string())
            df.at[index, 'Status'] = 'Sent'
            emails_sent_from_current_account += 1
            total_emails_sent += 1
            progress_bar['value'] += 1
        except Exception as e:
            df.at[index, 'Status'] = f'Failed: {str(e)}'

    server.quit()
    df.to_excel(excel_path, index=False)

    messagebox.showinfo("Success", f"Total emails sent: {total_emails_sent}")

# Email Sending Tab (Tab 1)
send_frame = ttk.Frame(notebook)
notebook.add(send_frame, text='Send Emails')

# GUI Elements for Email Sending
upload_button = tk.Button(send_frame, text="Upload Excel", command=load_excel)
upload_button.pack(pady=10)

subject_label = tk.Label(send_frame, text="Subject")
subject_label.pack()
subject_entry = tk.Entry(send_frame, width=50)
subject_entry.pack(pady=5)

body_label = tk.Label(send_frame, text="Message Body")
body_label.pack()
body_text = tk.Text(send_frame, height=10, width=50)
body_text.pack(pady=5)

# Dropdown for selecting Name field
name_label = tk.Label(send_frame, text="Select Name Field")
name_label.pack()
name_combobox = ttk.Combobox(send_frame, state="readonly")
name_combobox.pack(pady=5)

# Dropdown for selecting Email ID field
email_label = tk.Label(send_frame, text="Select Email ID Field")
email_label.pack()
email_combobox = ttk.Combobox(send_frame, state="readonly")
email_combobox.pack(pady=5)

send_button = tk.Button(send_frame, text="Send Emails", command=send_emails)
send_button.pack(pady=20)

progress_bar = ttk.Progressbar(send_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
progress_bar.pack(pady=10)

### Tab 2: Manage Email Accounts ###
def load_accounts():
    email_list.delete(0, tk.END)
    for account in email_accounts:
        email_list.insert(tk.END, account['email'])

def add_account():
    new_email = email_entry.get()
    new_password = password_entry.get()
    if new_email and new_password:
        email_accounts.append({'email': new_email, 'password': new_password})
        save_accounts()
        load_accounts()
        email_entry.delete(0, tk.END)
        password_entry.delete(0, tk.END)
        messagebox.showinfo("Success", "Account added successfully!")
    else:
        messagebox.showerror("Error", "Email and password cannot be empty.")

def delete_account():
    selected = email_list.curselection()
    if selected:
        del email_accounts[selected[0]]
        save_accounts()
        load_accounts()
        messagebox.showinfo("Success", "Account deleted successfully!")
    else:
        messagebox.showerror("Error", "No account selected.")

def save_accounts():
    with open('email_accounts.yaml', 'w') as f:
        yaml.dump({'email_accounts': email_accounts}, f)

# Email Accounts Tab (Tab 2)
manage_frame = ttk.Frame(notebook)
notebook.add(manage_frame, text='Manage Accounts')

# GUI Elements for managing email accounts
email_label = tk.Label(manage_frame, text="Email")
email_label.pack(pady=5)
email_entry = tk.Entry(manage_frame, width=30)
email_entry.pack(pady=5)

password_label = tk.Label(manage_frame, text="Password")
password_label.pack(pady=5)
password_entry = tk.Entry(manage_frame, show='*', width=30)
password_entry.pack(pady=5)

add_button = tk.Button(manage_frame, text="Add Account", command=add_account)
add_button.pack(pady=10)

email_list = tk.Listbox(manage_frame, width=50, height=10)
email_list.pack(pady=5)

delete_button = tk.Button(manage_frame, text="Delete Account", command=delete_account)
delete_button.pack(pady=10)

load_accounts()

# Run the GUI loop
root.mainloop()
