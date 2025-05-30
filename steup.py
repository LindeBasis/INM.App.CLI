from setuptools import setup, find_packages

setup(
    name='ticket_assigner',
    version='1.0.0',
    packages=find_packages(),
    install_requires=[
        'pandas',
    ],
    entry_points={
        'console_scripts': [
            'ticket-assigner = ticket_assigner.cli:cli',
        ],
    },
    author='Your Name',
    description='CLI tool to assign tickets to team members from CSVs and generate HTML summary.',
    classifiers=[
        'Programming Language :: Python :: 3',
    ],
)
