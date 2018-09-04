# MaddenToCSV
Powershell script which allows the Madden Companion Application to export directly to CSV files on your local computer.

# Usage Notes
- Download MaddenToCSV.ps1
- Open a powershell window as administrator
- Type "Set-ExecutionPolicy Bypass"
- Click to run the script
- Enter the server address shown on the PowerShell window into your Madden Companion App, and Export
- .CSV files will be saved onto your PC

# FYI
- Team names are not included in the stats and roster files.  They must be mapped from the table in leagueInfo.csv using the TeamID.
- The app exports 8 stat tables for each week:  (schedules, defense, kicking, punting, passing, receiving, rushing, teamstats)
