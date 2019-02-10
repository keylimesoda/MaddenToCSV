# MaddenToCSV
Powershell script which allows the Madden Companion Application to export directly to CSV files on your local computer.

# Usage Notes
- Before 1st run, you must enable powershell scripts on your computer
    - Open a powershell window as administrator
    - Type "Set-ExecutionPolicy Bypass"
- Run the MaddenToCSV.ps1 script
- options (none are required)
    - -ipAddress 0.0.0.0  specify local IP address to use for listening to app
    - -outputAMP          output AMP Editor compatible roster file
    - -scrimTeam1         choose a team to swap with Bucs
    - -scrimTeam2         choose a team to swap with the Saints

- Enter the server address shown on the PowerShell window into your Madden Companion App, and Export
- .CSV files will be saved onto your PC
- The script will run indefinitely.  Close the window when you're done.

# FYI
- Team names are not included in the stats and roster files.  They must be mapped from the table in leagueInfo.csv using the TeamID.
- The app exports 8 stat tables for each week:  (schedules, defense, kicking, punting, passing, receiving, rushing, teamstats)
