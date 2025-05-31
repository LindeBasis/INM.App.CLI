from setuptools import setup, find_packages

setup(
    name='ticket_assigner',
    version='2.1.0',
    packages=find_packages(),
    install_requires=['pandas'],
    entry_points={
        'console_scripts': [
            'ticket-assigner = ticket_assigner.cli:cli',
        ],
    },
    author='Your Name',
    description='Assigns tickets from INM.csv to available team members and generates HTML output.',
    classifiers=['Programming Language :: Python :: 3'],
)
