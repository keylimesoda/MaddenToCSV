# MaddenToCSV
Powershell script which allows the Madden Companion Application to export directly to CSV files on your local computer.

# Usage Notes
- Must have local script execution enabled "set-executionpolicy remotesigned"
- Script must be run from an admin PowerShell to create the httplistener server object
- Must add a passthrough on (or temporarily disable) PC firewall for server to be visible to companion app
- Team names are not included in the rosters.csv file.  They must be mapped from the table in leagueInfo.csv
