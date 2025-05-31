@echo off
rem # Step 1: Normalize INM.csv
ticket-assigner.exe normalize-inm --input INM.csv

rem # Step 2: Load normalized version
ticket-assigner.exe load-inm --inm INM.normalized.csv

rem # Step 3: Assign tickets using normalized file
ticket-assigner.exe assign --inm INM.normalized.csv --avail TEAM_AVAIL.csv --team TEAM_MASTER.csv

rem # Step 4: Generate HTML
ticket-assigner.exe html --input TEAM_Assigned.csv
pause
