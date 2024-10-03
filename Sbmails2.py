import os
import tkinter as tk
from tkinter import filedialog, messagebox
import winshell  # For creating desktop shortcuts
import yaml
import pandas as pd
from pathlib import Path

# Function to get default folder location (C:\SB Mails)
def get_default_folder():
    default_path = Path('C:/SB Mails')
    if not default_path.exists():
        default_path.mkdir(parents=True)
    return default_path

# Function to let user choose custom folder
def choose_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_selected)

# Function to save email accounts to chosen location
def save_email_accounts():
    chosen_folder = folder_entry.get() or str(get_default_folder())
    yaml_path = os.path.join(chosen_folder, 'email_accounts.yaml')
    try:
        with open(yaml_path, 'w') as f:
            yaml.dump({'email_accounts': email_accounts}, f)
        messagebox.showinfo("Success", f"Email accounts saved in {yaml_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save email accounts: {str(e)}")

# Function to create a desktop shortcut
def create_shortcut():
    chosen_folder = folder_entry.get() or str(get_default_folder())
    exe_path = os.path.abspath("SB Mails.exe")  # Assuming the .exe is in the current working directory
    shortcut_path = os.path.join(winshell.desktop(), "SB Mails.lnk")

    try:
        winshell.CreateShortcut(
            Path=shortcut_path,
            Target=exe_path,
            Icon=(exe_path, 0),
            Description="Shortcut to SB Mails"
        )
        messagebox.showinfo("Success", f"Desktop shortcut created: {shortcut_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create desktop shortcut: {str(e)}")

# GUI for folder selection and saving
root = tk.Tk()
root.title("SB Mails Setup")

folder_frame = tk.Frame(root)
folder_frame.pack(pady=10)

folder_label = tk.Label(folder_frame, text="Choose folder for saving files:")
folder_label.pack(side=tk.LEFT, padx=5)

# Entry box to show selected folder
folder_entry = tk.Entry(folder_frame, width=50)
folder_entry.pack(side=tk.LEFT, padx=5)
folder_entry.insert(0, str(get_default_folder()))  # Set default location in entry box

# Button to open folder dialog
choose_folder_button = tk.Button(folder_frame, text="Browse", command=choose_folder)
choose_folder_button.pack(side=tk.LEFT, padx=5)

# Button to save email accounts
save_button = tk.Button(root, text="Save Email Accounts", command=save_email_accounts)
save_button.pack(pady=10)

# Option to create desktop shortcut
shortcut_button = tk.Button(root, text="Create Desktop Shortcut", command=create_shortcut)
shortcut_button.pack(pady=10)

root.mainloop()
