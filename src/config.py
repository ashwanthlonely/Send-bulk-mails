import yaml
import os

CREDENTIALS_FILE = 'credentials.yaml'

def load_credentials():
    try:
        with open(CREDENTIALS_FILE, 'r') as file:
            return yaml.safe_load(file)['accounts']
    except FileNotFoundError:
        return []

def save_credentials(credentials):
    with open(CREDENTIALS_FILE, 'w') as file:
        yaml.dump({'accounts': credentials}, file)

TEMPLATE_FILE = 'templates.yaml'

def load_templates():
    """Load all templates from the YAML file."""
    if os.path.exists(TEMPLATE_FILE):
        with open(TEMPLATE_FILE, 'r') as file:
            templates = yaml.safe_load(file)
            return templates if templates else {}
    return {}

def save_template(template_name, subject, body, signature):
    """Save a new template or update an existing one."""
    templates = load_templates()
    templates[template_name] = {
        'subject': subject,
        'body': body,
        'signature': signature
    }
    with open(TEMPLATE_FILE, 'w') as file:
        yaml.dump(templates, file)

def delete_template(template_name):
    """Delete a template from the YAML file."""
    templates = load_templates()
    if template_name in templates:
        del templates[template_name]
        with open(TEMPLATE_FILE, 'w') as file:
            yaml.dump(templates, file)