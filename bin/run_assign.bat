@echo off
rem # Normalize column names
ticket-assigner normalize-inm

rem # Load data to DB
ticket-assigner load-inm --inm INM.normalized.csv

rem # Assign tickets and generate TEAM_Assigned.csv and TEAM_Assigned_Email.csv
ticket-assigner assign

rem # Convert TEAM_Assigned_Email.csv to HTML
rem ticket-assigner html
pause
