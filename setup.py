from setuptools import setup, find_packages

setup(
    name="bulk_email_sender",  # Project name
    version="0.1.0",  # Initial version
    author="Your Name",
    author_email="your.email@example.com",
    description="A bulk email sender with a GUI interface using Tkinter.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/bulk_email_sender",  # Your project URL
    packages=find_packages(where="src"),  # Include all packages under the src directory
    package_dir={"": "src"},  # Tell setuptools that packages are in the src directory
    include_package_data=True,
    install_requires=[  # List your project dependencies here
        "pandas",
        "tqdm",
        "pyyaml"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    entry_points={  # Entry point to run the main script
        'console_scripts': [
            'bulk-email-sender=src.gui:EmailSenderGUI',
        ],
    },
)
