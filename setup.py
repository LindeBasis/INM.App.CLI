from setuptools import setup, find_packages

setup(
    name="ticket-assigner",
    version="4.6.0",
    description="Daily Incident Assignment CLI Tool for SAP Basis Team",
    author="Your Name",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "pandas",
        "xlrd",
        "openpyxl",
        "pywin32",
        "beautifulsoup4"
    ],
    entry_points={
        "console_scripts": [
            "ticket-assigner = ticket_assigner.cli:cli",
            "ticket-unassigned = ticket_assigner.cli_un:cli_un"
        ]
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
)
