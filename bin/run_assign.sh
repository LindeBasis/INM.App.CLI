#!/bin/bash
# Step 1: Normalize INM.csv
ticket-assigner normalize-inm --input INM.csv

# Step 2: Load normalized version
ticket-assigner load-inm --inm INM.normalized.csv

# Step 3: Assign tickets using normalized file
ticket-assigner assign --inm INM.normalized.csv --avail TEAM_AVAIL.csv --team TEAM_MASTER.csv

# Step 4: Generate HTML
ticket-assigner html --input TEAM_Assigned.csv