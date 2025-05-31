# INM.App.CLI
# Ticket assigner.exe CLI

## Description
Automates daily ticket assignment using INM.csv, TEAM_AVAIL.csv, and TEAM_MASTER.csv. Generates `TEAM_Assigned.csv` and an HTML version for email.

## Folder Structure

data/
INM.csv
TEAM_AVAIL.csv
TEAM_MASTER.csv`
TEAM_Assigned.csv

## Usage

```bash
# Install the CLI
pip install .

# Step 1: Normalize INM.csv
ticket-assigner.exe normalize-inm --input INM.csv

# Step 2: Load normalized version
ticket-assigner.exe load-inm --inm INM.normalized.csv

# Step 3: Assign tickets using normalized file
ticket-assigner.exe assign --inm INM.normalized.csv --avail TEAM_AVAIL.csv --team TEAM_MASTER.csv

# Step 4: Generate HTML
ticket-assigner.exe html --input TEAM_Assigned.csv
```


Or use:

```bash
./bin/run_assign.sh      # Linux/macOS
./bin/run_assign.bat     # Windows
```

