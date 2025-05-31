#!/bin/bash
# Normalize column names
ticket-assigner normalize-inm

# Load data to DB
ticket-assigner load-inm --inm INM.normalized.csv

# Assign tickets and generate TEAM_Assigned.csv and TEAM_Assigned_Email.csv
ticket-assigner assign

# Convert TEAM_Assigned_Email.csv to HTML
# ticket-assigner html