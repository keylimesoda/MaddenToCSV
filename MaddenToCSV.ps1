<#
    .SYNOPSIS
        Function to export League Info, Schedules and Rosters to CSV files on-disk
        
    .DIRECTIONS
    	Before 1st run, you must enable powershell scripts on your computer
		Open a powershell window as administrator
		Type "Set-ExecutionPolicy Bypass"
	Run the MaddenToCSV.ps1 script
	Enter the server address shown on the PowerShell window into your Madden Companion App, and Export
	.CSV files will be saved onto your PC
	The script will run indefinitely.  Close the window when you're done.
	
    .NOTES
        Team names are not included in the stats and roster files. They must be mapped from the table in leagueInfo.csv using the TeamID.
	The app exports 8 stat tables for each week: (schedules, defense, kicking, punting, passing, receiving, rushing, teamstats)

#>

Add-Type -AssemblyName System.IO
$FormatEnumerationLimit=-1  #Enables larger tables for debug output
$enc = [system.Text.Encoding]::Default
$port = 8080

#//////////////////////////////////////////////////////////////////
#//This script must be run as admin to setup firewall and listener
#//This section creates an escalated window, following correct security practices
#//////////////////////////////////////////////////////////////////

# Get the ID and security principal of the current user account
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
 
# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
 
# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))
{
   # We are running "as Administrator" - so change the title and background color to indicate this
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
   $Host.UI.RawUI.BackgroundColor = "DarkBlue"
   Set-Location -Path $PSScriptRoot
   clear-host
}
else
{
   # We are not running "as Administrator" - so relaunch as administrator
   
   # Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   
   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;
   
   # Pass along the current working directory for any output
   $newProcess.WorkingDirectory = $PSScriptRoot
   
   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";
   
   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess) >$null
   
   # Exit from the current, unelevated, process
   exit
}

#//////////////////////////////////////////////////////////////////
#//Add MaddenToCSV to firewall to allow companion to talk to PC
#//////////////////////////////////////////////////////////////////


$firewallPort = $port
$firewallRuleName = "MaddenToCSV port $firewallPort"
    
if (-Not(Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue))
{
    write-host "Firewall rule for '$firewallRuleName' on port '$firewallPort' does not already exist, creating new rule now..."
    New-NetFirewallRule -DisplayName $firewallRuleName -Direction Inbound -Profile Domain,Private,Public -Action Allow -Protocol TCP -LocalPort $firewallPort -RemoteAddress Any >$null
    write-host "Firewall rule for '$firewallRuleName' on port '$firewallPort' created successfully"
    write-host ""
}


#//////////////////////////////////////////////////////////////////
#//Setup listener service
#//////////////////////////////////////////////////////////////////

$ipAddr = (Get-NetIPAddress | ?{ $_.AddressFamily -eq "IPv4"  -and !($_.IPAddress -match "169") -and !($_.IPaddress -match "127") }).IPAddress

$serverAddr = "http://"+$ipAddr+":$port/"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($serverAddr)

try
{
    $listener.Start()
}
catch
{
     Write-Host "Cannot start server.  Script already running in another window?"
     Write-Host -NoNewLine 'Press any key to continue...';
     $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
     exit
}

$serverAddr = $serverAddr.TrimEnd('/')
Write-Host "Server started at $serverAddr"


#//////////////////////////////////////////////////////////////////
#//Start the listen/response loop
#//////////////////////////////////////////////////////////////////

$counter = 1
Write-Host ""
Write-Host "Listening for Madden Companion App (close window to exit)"
Write-host ""
        
do {
        
    try{
        $context = $listener.GetContext() #script will pause here waiting for Companion App

        $response = $context.Response
        $request = $context.Request

        $requestUrl = $request.Url
        
        #// Enable this line to see what URL the app is posting to
        #Write-Host $requestUrl

        ### grab and decompress POST data
        $decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
        $readStream = [System.IO.StreamReader]::new($decompress)
        $content = $readStream.ReadToEnd()
        

        ### create JSON object
        $requestJson = $content | ConvertFrom-Json
        
        ### based on POST URL, select which kind of export is happening        
        switch -Wildcard ($requestUrl)
        {
            '*leagueteams' #LEAGUE INFO
            {
                $infoList = $requestJson.leagueTeamInfoList
                #Write-Host ($infoList | Format-Table -Property *| Out-String -Width 4096)
                $infoList | Export-Csv -Path "leagueInfo.csv" -NoTypeInformation
                Write-Host "LEAGUE INFO `t => leagueInfo.csv"
                
                break
            }

            '*standings' #LEAGUE INFO
            {
                $standingsList = $requestJson.teamStandingInfoList
                #Write-Host ($standingsList | Format-Table -Property *| Out-String -Width 4096)
                $standingsList | Export-Csv -Path "standingsInfo.csv" -NoTypeInformation
                Write-Host "STANDINGS `t => standingsInfo.csv"
                
                break
            }

            '*roster' #ROSTERS
            {
                
                Write-Host "ROSTER"
                
                ##grab the 1st team, since request already received
                Write-Host "Team 1"
                $teamList = $requestJson.rosterInfoList
                $readStream.Close()
                $quickResponse = "please connect with Madden App"
                $content = $enc.GetBytes($quickResponse)
                $response.ContentLength64 = $content.Length
                $response.OutputStream.Write($content, 0, $content.Length)
                $response.Close()

                for ($i=2; $i -le 32; $i++)
                {
                     ### We assume next 32 calls will be rosters so we run full loop, building roster file in memory
                     Write-Host "Team $i"
                     $context = $listener.GetContext()
                     $response = $context.Response
                     $request = $context.Request
                     $decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
                     $readStream = [System.IO.StreamReader]::new($decompress)
                     $content = $readStream.ReadToEnd()
                     $requestJson = $content | ConvertFrom-Json
                     $teamList += $requestJson.rosterInfoList
                   
                     $readStream.Close()
                     $quickResponse = "please connect with Madden App"
                     $content = $enc.GetBytes($quickResponse)
                     $response.ContentLength64 = $content.Length
                     $response.OutputStream.Write($content, 0, $content.Length)
                     $response.Close()
                } 

                ### Last roster call saves roster file to disk and skips closing stream and response objects, since the end of loop code below will handle closing
                Write-Host "Free Agents"
                $context = $listener.GetContext()
                $response = $context.Response
                $request = $context.Request
                $decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
                $readStream = [System.IO.StreamReader]::new($decompress)
                $content = $readStream.ReadToEnd()
                $requestJson = $content | ConvertFrom-Json
                $teamList += $requestJson.rosterInfoList
                   
                $teamList | Export-Csv -Path "rosters.csv" -NoTypeInformation
                Write-Host "Export to disk:  rosters.csv" 
		break
            }	    
	        '*week*' #WEEKLY STATISTICS
            {
                ### Get a week ID that we'll use as a prefix for the stats files
                $weekType = $requestUrl.Segments[4].TrimEnd('/')
                $weekNum  = ($requestUrl.Segments[5].TrimEnd('/'))
                $weekNum  = [int]$weekNum
                $week = "({0} {1:D2})" -f $weekType, $weekNum
                
                switch -Wildcard ($requestURL)
                {
                    '*schedules'
                    {
                        $statFilename = "$($week) scheduleInfo.csv"
                        $requestJson.gameScheduleInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "SCHEDULE `t => $statFileName"
                    }
                    '*defense'
                    {
                        $statFileName = "$($week) defensiveStats.csv"
                        $requestJson.playerDefensiveStatInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "DEFENSIVE STATS  => $statFilename"
                    }
                    '*kicking'
                    {
                        $statFileName = "$($week) kickingStats.csv"
                        $requestJson.playerKickingStatInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "KICKING STATS `t => $statFilename"
                    }
                    '*passing'
                    {
                        $statFileName = "$($week) passingStats.csv"
                        $requestJson.playerPassingStatInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "PASSING STATS `t => $statFilename"
                    }
                    '*punting'
                    {
                        $statFileName = "$($week) puntingStats.csv"
                        $requestJson.playerPuntingStatInfoList| Export-Csv -Path $statFilename  -NoTypeInformation
                        Write-Host "PUNTING STATS `t => $statFilename"
                    }
                    '*receiving'
                    {
                        $statFileName = "$($week) receivingStats.csv"
                        $requestJson.playerReceivingStatInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "RECEIVING STATS  => $statFilename"
                    }
                    '*rushing'
                    {
                        $statFileName = "$($week) rushingStats.csv"
                        $requestJson.playerRushingStatInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "RUSHING STATS `t => $statFilename"
                    }
                    '*teamstats'
                    {
                        $statFileName = "$($week) teamStats.csv"
                        $requestJson.teamStatInfoList| Export-Csv -Path $statFilename -NoTypeInformation
                        Write-Host "TEAM STATS `t => $statFilename"
                    }
                }
                break
            }            
        }
        $readStream.Close()
    } 
    catch {
        $_
        $content =  "$($_.InvocationInfo.MyCommand.Name) : $($_.Exception.Message)"
        $content +=  "$($_.InvocationInfo.PositionMessage)"
        $content +=  "    + $($_.CategoryInfo.GetMessage())"
        $content +=  "    + $($_.FullyQualifiedErrorId)"

        $content = [System.Text.Encoding]::UTF8.GetBytes($content)
        $response.StatusCode = 500
    }

    
    $quickResponse = "please connect with Madden Companion App"
    $content = $enc.GetBytes($quickResponse)
    
    $response.ContentLength64 = $content.Length
    $response.OutputStream.Write($content, 0, $content.Length)
    $response.Close()

    $responseStatus = $response.StatusCode

} while ($listener.IsListening)
