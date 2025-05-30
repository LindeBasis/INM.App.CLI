    #!/bin/bash
ticket-assigner load-inm --inm INM.csv
ticket-assigner assign --inm INM.csv --avail TEAM_AVAIL.csv --team TEAM_MASTER.csv
ticket-assigner html --input TEAM_Assigned.csv --output TEAM_Assigned.html
