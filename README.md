# MaddenToCSV
Powershell script which allows the Madden Companion Application to export directly to CSV files on your local computer.

# Usage Notes
- Download MaddenToCSV.ps1
- Click to run it

# FYI
- If your system has disabled running local powershell scripts, you can enable using command "Set-ExecutionPolicy -Scope CurrentUser remotesigned"
- Team names are not included in the stats and roster files.  They must be mapped from the table in leagueInfo.csv using the TeamID.
- The app exports 8 stat tables for each week:  (schedules, defense, kicking, punting, passing, receiving, rushing, teamstats)
