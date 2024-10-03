from tkinter import *
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import pandas as pd
import yaml
import smtplib
import imaplib
from datetime import datetime, timedelta

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Sender")
        self.root.geometry("800x600")

        # Load email accounts from YAML file
        self.email_accounts_config = self.load_email_accounts()
        self.email_accounts = self.email_accounts_config['email_accounts']

        # Email sending limit per account
        self.email_limit_per_account = 500
        self.email_refresh_interval = timedelta(days=1)  # 24 hours

        # Initialize data
        self.df = None  # Placeholder for the DataFrame after Excel upload

        # Setup GUI elements
        self.setup_gui()

    def load_email_accounts(self):
        with open('email_accounts.yaml', 'r') as f:
            return yaml.safe_load(f)

    def save_email_accounts(self):
        with open('email_accounts.yaml', 'w') as f:
            yaml.safe_dump(self.email_accounts_config, f)

    def setup_gui(self):
        # Add tabs for navigation
        self.tab_control = ttk.Notebook(self.root)

        # Tab 1 - Email Configuration
        self.config_tab = Frame(self.tab_control)
        self.tab_control.add(self.config_tab, text="Email Config")

        self.upload_button = Button(self.config_tab, text="Upload Excel File", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.select_email_label = Label(self.config_tab, text="Select Email Field")
        self.select_email_label.pack()
        self.email_column_var = StringVar(self.config_tab)
        self.email_column_dropdown = OptionMenu(self.config_tab, self.email_column_var, "")
        self.email_column_dropdown.pack()

        self.select_name_label = Label(self.config_tab, text="Select Name Field")
        self.select_name_label.pack()
        self.name_column_var = StringVar(self.config_tab)
        self.name_column_dropdown = OptionMenu(self.config_tab, self.name_column_var, "")
        self.name_column_dropdown.pack()

        self.subject_label = Label(self.config_tab, text="Email Subject")
        self.subject_label.pack()
        self.subject_entry = Entry(self.config_tab, width=50)
        self.subject_entry.pack()

        self.body_label = Label(self.config_tab, text="Email Body")
        self.body_label.pack()
        self.body_text = Text(self.config_tab, height=10, width=50)
        self.body_text.pack(pady=10)

        self.send_button = Button(self.config_tab, text="Send Emails", command=self.send_emails)
        self.send_button.pack(pady=20)

        # Tab 2 - Edit/Delete Logins
        self.login_tab = Frame(self.tab_control)
        self.tab_control.add(self.login_tab, text="Manage Logins")

        self.login_listbox = Listbox(self.login_tab, width=60, height=10)
        self.login_listbox.pack(pady=10)
        self.populate_login_listbox()

        self.edit_button = Button(self.login_tab, text="Edit Selected", command=self.edit_login)
        self.edit_button.pack(pady=5)

        self.delete_button = Button(self.login_tab, text="Delete Selected", command=self.delete_login)
        self.delete_button.pack(pady=5)

        self.add_login_button = Button(self.login_tab, text="Add New Login", command=self.add_login)
        self.add_login_button.pack(pady=5)

        # Tab 3 - Progress Tracking
        self.progress_tab = Frame(self.tab_control)
        self.tab_control.add(self.progress_tab, text="Progress")

        self.progress_bar = ttk.Progressbar(self.progress_tab, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=20)

        self.progress_label = Label(self.progress_tab, text="Progress: 0%")
        self.progress_label.pack()

        self.status_text = Text(self.progress_tab, height=15, width=80)
        self.status_text.pack(pady=10)

        self.tab_control.pack(expand=1, fill="both")

    def upload_excel(self):
        # File dialog to upload Excel file
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        if file_path:
            self.excel_path = file_path
            self.df = pd.read_excel(file_path)
            self.populate_dropdowns()

            messagebox.showinfo("File Uploaded", "Excel file uploaded successfully.")

    def populate_dropdowns(self):
        # Populate dropdowns for selecting email and name fields
        columns = self.df.columns
        self.email_column_var.set(columns[0])
        self.name_column_var.set(columns[1])

        self.email_column_dropdown["menu"].delete(0, "end")
        self.name_column_dropdown["menu"].delete(0, "end")

        for col in columns:
            self.email_column_dropdown["menu"].add_command(label=col, command=lambda value=col: self.email_column_var.set(value))
            self.name_column_dropdown["menu"].add_command(label=col, command=lambda value=col: self.name_column_var.set(value))

    def send_emails(self):
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", END)

        email_column = self.email_column_var.get()
        name_column = self.name_column_var.get()

        if not subject or not body or not hasattr(self, 'excel_path'):
            messagebox.showerror("Missing Information", "Please provide subject, body, and upload an Excel file.")
            return

        # Send emails logic (implement email sending here)

    def populate_login_listbox(self):
        self.login_listbox.delete(0, END)
        for account in self.email_accounts:
            self.login_listbox.insert(END, f"{account['email']}")

    def edit_login(self):
        selected_index = self.login_listbox.curselection()
        if not selected_index:
            messagebox.showerror("Error", "No login selected")
            return

        selected_account = self.email_accounts[selected_index[0]]
        self.edit_login_window(selected_account, selected_index[0])

    def edit_login_window(self, account, index):
        edit_win = Toplevel(self.root)
        edit_win.title("Edit Login")

        Label(edit_win, text="Email:").pack()
        email_entry = Entry(edit_win, width=40)
        email_entry.insert(END, account['email'])
        email_entry.pack()

        Label(edit_win, text="Password:").pack()
        password_entry = Entry(edit_win, width=40, show="*")
        password_entry.insert(END, account['password'])
        password_entry.pack()

        def save_changes():
            updated_email = email_entry.get()
            updated_password = password_entry.get()
            self.email_accounts[index]['email'] = updated_email
            self.email_accounts[index]['password'] = updated_password
            self.save_email_accounts()
            self.populate_login_listbox()
            edit_win.destroy()

        Button(edit_win, text="Save", command=save_changes).pack(pady=10)

    def delete_login(self):
        selected_index = self.login_listbox.curselection()
        if not selected_index:
            messagebox.showerror("Error", "No login selected")
            return

        del self.email_accounts[selected_index[0]]
        self.save_email_accounts()
        self.populate_login_listbox()

    def add_login(self):
        self.edit_login_window({"email": "", "password": ""}, len(self.email_accounts))

# Running the application
if __name__ == "__main__":
    root = Tk()
    app = EmailSenderApp(root)
    root.mainloop()
