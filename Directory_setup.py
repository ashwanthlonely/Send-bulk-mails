import os

# Define the base directory
base_dir = "SEND-BULK-MAILS"

# Create the base directory (if it doesn't exist)
os.makedirs(base_dir, exist_ok=True)  # exist_ok prevents errors if already exists

# Create subdirectories (src, tests, docs)
sub_dirs = ["src", "tests", "docs"]
for sub_dir in sub_dirs:
    os.makedirs(os.path.join(base_dir, sub_dir), exist_ok=True)

# Create source files within src directory
source_files = ["__init__.py", "email_sender.py", "gui.py", "config.py", "utils.py"]
source_dir = os.path.join(base_dir, "src")
for file in source_files:
    open(os.path.join(source_dir, file), 'w').close()  # Create empty files

# Create additional files (README.md, requirements.txt, setup.py)
other_files = ["main.py", "credentials.yaml"]
for file in other_files:
    open(os.path.join(base_dir, file), 'w').close()  # Create empty files

# Test directory structure creation
if os.path.exists(base_dir):
    print(f"Directory structure created successfully at: {base_dir}")
else:
    print("Error creating directory structure!")