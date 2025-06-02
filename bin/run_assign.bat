@echo off
rem # Fetch Latest INM file
ticket-assigner fetch-latest

rem # Normalize column names
ticket-assigner normalize-inm

rem # Load data to DB
ticket-assigner load-inm 

rem # Assign tickets and generate TEAM_Assigned.csv and TEAM_Assigned_Email.csv
ticket-assigner assign

rem # Convert TEAM_Assigned_Email.csv to HTML
rem ticket-assigner html
python send_outlook_email.py
pause
