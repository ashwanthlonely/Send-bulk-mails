from tkinter import Tk, Label, Entry, Button, filedialog, Text, messagebox, Listbox, MULTIPLE, simpledialog
from tkinter import ttk
import pandas as pd
import os

from src.config import load_credentials, save_credentials, save_template, load_templates, delete_template
from src.email_sender import EmailSender

class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Sender")

        self.email_sender = EmailSender()
        self.attachment_path = None

        # Create Notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        # Create Tabs
        self.create_send_tab()
        self.create_credentials_tab()
        # self.create_templates_tab()

    def create_send_tab(self):
        self.send_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.send_tab, text="Send Emails")

        # Excel upload
        upload_btn = Button(self.send_tab, text="Upload Excel", command=self.upload_excel)
        upload_btn.grid(row=0, column=0, padx=10, pady=10)

        # Subject entry
        Label(self.send_tab, text="Subject:").grid(row=1, column=0, sticky='e')
        self.subject_entry = Entry(self.send_tab, width=50)
        self.subject_entry.grid(row=1, column=1, padx=10, pady=10)

        # CC entry
        Label(self.send_tab, text="CC:").grid(row=2, column=0, sticky='e')
        self.cc_entry = Entry(self.send_tab, width=50)
        self.cc_entry.grid(row=2, column=1, padx=10, pady=10)

        # Message text box
        Label(self.send_tab, text="Message:").grid(row=3, column=0, sticky='ne')
        self.message_text = Text(self.send_tab, width=50, height=10)
        self.message_text.grid(row=3, column=1, padx=10, pady=10)

        # Signature text box
        Label(self.send_tab, text="Signature:").grid(row=4, column=0, sticky='ne')
        self.signature_text = Text(self.send_tab, width=50, height=5)
        self.signature_text.grid(row=4, column=1, padx=10, pady=10)

        # Attachment upload
        self.attachment_label = Label(self.send_tab, text="No attachment")
        self.attachment_label.grid(row=5, column=0, padx=10, pady=10, sticky='w')
        upload_attachment_btn = Button(self.send_tab, text="Upload Attachment", command=self.upload_attachment)
        upload_attachment_btn.grid(row=5, column=0, padx=10, pady=10, sticky='w')
        # Start sending button
        start_btn = Button(self.send_tab, text="Start Sending Emails", command=self.send_emails)
        start_btn.grid(row=5, column=1, padx=10, pady=10)
        # Template management
        Label(self.send_tab, text="Templates:", font=("bold")).grid(row=6, column=0, padx=10, pady=10, sticky='nw')
        self.template_listbox = Listbox(self.send_tab)
        self.template_listbox.grid(row=7, column=1, padx=15, pady=20, sticky='w')
        self.update_template_list()

        save_template_btn = Button(self.send_tab, text="Save Template", command=self.save_template)
        save_template_btn.grid(row=5, column=2, padx=10, pady=10, sticky='w')

        load_template_btn = Button(self.send_tab, text="Load Template", command=self.load_template)
        load_template_btn.grid(row=8, column=0, padx=10, pady=10, sticky='w')

        edit_template_btn = Button(self.send_tab, text="Edit Template", command=self.edit_template)
        edit_template_btn.grid(row=8, column=1, padx=10, pady=10, sticky='w')

        delete_template_btn = Button(self.send_tab, text="Delete Template", command=self.delete_template)
        delete_template_btn.grid(row=8, column=2, padx=10, pady=10, sticky='w')

        

    def create_credentials_tab(self):
        self.credentials_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.credentials_tab, text="Credentials Management")

        # Credentials list
        Label(self.credentials_tab, text="Credentials:").grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        self.credentials_listbox = Listbox(self.credentials_tab, selectmode=MULTIPLE)
        self.credentials_listbox.grid(row=1, column=1, padx=10, pady=10, sticky='w')

        # Add, Edit, Delete buttons for credentials
        add_cred_btn = Button(self.credentials_tab, text="Add Credential", command=self.add_credential)
        add_cred_btn.grid(row=2, column=1, padx=10, pady=10, sticky='w')
        edit_cred_btn = Button(self.credentials_tab, text="Edit Credential", command=self.edit_credential)
        edit_cred_btn.grid(row=2, column=1, padx=10, pady=10, sticky='e')
        delete_cred_btn = Button(self.credentials_tab, text="Delete Selected Credential(s)", command=self.delete_selected_credentials)
        delete_cred_btn.grid(row=3, column=1, padx=10, pady=10, sticky='e')

        # Load existing credentials into the list
        self.update_credentials_list()

  

    def upload_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filepath:
            self.email_sender.df = pd.read_excel(filepath)
            self.email_sender.df['Sent Status'] = ''  # Initialize a column for email sending status
            self.email_sender.df['Sent Date'] = ''  # Initialize a column for sent date
            self.email_sender.df['Sent Time'] = ''  # Initialize a column for sent time
            messagebox.showinfo("File Uploaded", "Excel file has been uploaded successfully!")

    def upload_attachment(self):
        self.attachment_path = filedialog.askopenfilename(filetypes=[("All files", "*.*")])
        if self.attachment_path:
            self.attachment_label.config(text=f"Attachment: {os.path.basename(self.attachment_path)}")

    def add_credential(self):
        email = simpledialog.askstring("Input", "Enter email:")
        password = simpledialog.askstring("Input", "Enter password:", show='*')
        if email and password:
            self.email_sender.credentials.append({'email': email, 'password': password})
            save_credentials(self.email_sender.credentials)
            self.update_credentials_list()

    def delete_selected_credentials(self):
        selected = self.credentials_listbox.curselection()
        if not selected:
            messagebox.showerror("Selection Error", "Please select one or more credentials to delete.")
            return

        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected credentials?"):
            for index in reversed(selected):  # Delete from the end to avoid index shifting issues
                del self.email_sender.credentials[index]
            save_credentials(self.email_sender.credentials)
            self.update_credentials_list()

    def edit_credential(self):
        selected = self.credentials_listbox.curselection()
        if not selected:
            messagebox.showerror("Selection Error", "Please select a credential to edit.")
            return
        
        index = selected[0]
        email = simpledialog.askstring("Edit Email", "Edit email:", initialvalue=self.email_sender.credentials[index]['email'])
        password = simpledialog.askstring("Edit Password", "Edit password:", initialvalue=self.email_sender.credentials[index]['password'], show='*')
        if email and password:
            self.email_sender.credentials[index] = {'email': email, 'password': password}
            save_credentials(self.email_sender.credentials)
            self.update_credentials_list()

    def update_credentials_list(self):
        self.credentials_listbox.delete(0, 'end')
        for account in self.email_sender.credentials:
            self.credentials_listbox.insert('end', account['email'])

    def send_emails(self):
        subject = self.subject_entry.get()
        body = self.message_text.get("1.0", 'end-1c')
        signature = self.signature_text.get("1.0", 'end-1c')
        cc_email = self.cc_entry.get()
        
        self.email_sender.send_emails(subject, body, signature, cc_email, self.attachment_path)

    def save_template(self):
        template_name = simpledialog.askstring("Template Name", "Enter template name:")
        if not template_name:
            return
        
        subject = self.subject_entry.get()
        body = self.message_text.get("1.0", 'end-1c')
        signature = self.signature_text.get("1.0", 'end-1c')
        
        save_template(template_name, subject, body, signature)
        messagebox.showinfo("Template Saved", f"Template '{template_name}' saved successfully.")
        self.update_template_list()

    def load_template(self):
        selected_template = self.template_listbox.get(self.template_listbox.curselection())
        if not selected_template:
            messagebox.showerror("Selection Error", "Please select a template to load.")
            return
        
        templates = load_templates()
        template = templates.get(selected_template)
        
        self.subject_entry.delete(0, 'end')
        self.subject_entry.insert(0, template['subject'])
        self.message_text.delete("1.0", 'end')
        self.message_text.insert("1.0", template['body'])
        self.signature_text.delete("1.0", 'end')
        self.signature_text.insert("1.0", template['signature'])

    def edit_template(self):
        selected_template = self.template_listbox.get(self.template_listbox.curselection())
        if not selected_template:
            messagebox.showerror("Selection Error", "Please select a template to edit.")
            return
        
        templates = load_templates()
        template = templates.get(selected_template)
        
        self.subject_entry.delete(0, 'end')
        self.subject_entry.insert(0, template['subject'])
        self.message_text.delete("1.0", 'end')
        self.message_text.insert("1.0", template['body'])
        self.signature_text.delete("1.0", 'end')
        self.signature_text.insert("1.0", template['signature'])
        
        # Remove the old template and allow saving under a new name
        delete_template(selected_template)

    def delete_template(self):
        selected_template = self.template_listbox.get(self.template_listbox.curselection())
        if not selected_template:
            messagebox.showerror("Selection Error", "Please select a template to delete.")
            return
        
        delete_template(selected_template)
        messagebox.showinfo("Template Deleted", f"Template '{selected_template}' deleted successfully.")
        self.update_template_list()

    def update_template_list(self):
        self.template_listbox.delete(0, 'end')
        templates = load_templates()
        for template_name in templates.keys():
            self.template_listbox.insert('end', template_name)

if __name__ == "__main__":
    root = Tk()
    gui = EmailSenderGUI(root)
    root.mainloop()
