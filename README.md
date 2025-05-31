# INM.App.CLI
# Ticket assigner.exe CLI

## Description
Automates daily ticket assignment using INM.csv, TEAM_AVAIL.csv, and TEAM_MASTER.csv. Generates `TEAM_Assigned.csv` and an HTML version for email.

## Folder Structure

ticket_assigner/
│
├── ticket_assigner/
│   └── cli.py
│
├── data/
│   ├── INM.csv
│   ├── INM.normalized.csv  <-- auto-created
│   ├── TEAM_AVAIL.csv
│   ├── TEAM_MASTER.csv
│   ├── TEAM_Assigned.csv   <-- auto-created
│   ├── TEAM_Assigned.html  <-- auto-created
│   └── ticket_assign.db    <-- created by load-inm
│
├── setup.py
└── README.md (optional)


## Usage

```bash
# Install the CLI
pip uninstall ticket_assigner
pip install -e .

# Normalize column names
ticket-assigner normalize-inm

# Load data to DB
ticket-assigner load-inm --inm INM.normalized.csv

# Assign tickets and generate TEAM_Assigned.csv and TEAM_Assigned_Email.csv
ticket-assigner assign

# Convert TEAM_Assigned_Email.csv to HTML
ticket-assigner html

```


Or use:

```bash
./bin/run_assign.sh      # Linux/macOS
./bin/run_assign.bat     # Windows
```

