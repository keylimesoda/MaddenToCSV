<#
    .SYNOPSIS
        Function to export League Info, Schedules and Rosters to CSV files on-disk
        
    .DIRECTIONS
        The script will run indefinitely.  Close the window when you're done.

    .NOTES
        Does not yet handle stats
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
   clear-host
   }
else
   {
   # We are not running "as Administrator" - so relaunch as administrator
   
   # Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   
   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;
   
   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";
   
   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess);
   
   # Exit from the current, unelevated, process
   exit
   }


#//////////////////////////////////////////////////////////////////
#//Add MaddenToCSV to firewall to allow companion to talk to PC
#//////////////////////////////////////////////////////////////////

$firewallPort = $port
$firewallRuleName = "MaddenToCSV port $firewallPort"
    
if (-Not(Get-NetFirewallRule –DisplayName $firewallRuleName -ErrorAction SilentlyContinue))
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
     Write-Host "Cannot start server.  Script must be run from an admin PowerShell window"
     exit
}

$serverAddr = $serverAddr.TrimEnd('/')
Write-Host "Server started at $serverAddr"


#//////////////////////////////////////////////////////////////////
#//Start the listen/response loop
#//////////////////////////////////////////////////////////////////

Write-Host ""
Write-Host "Listening for Madden Companion App (close window to exit)"
Write-host ""
        
do {
        
    try{
        $context = $listener.GetContext() #script will pause here waiting for Companion App

        $response = $context.Response
        $request = $context.Request


<#
        ### show request headers for debug
        $requestHeaders = $request.Headers
        $stringH = $requestHeaders.AllKeys | 
            Select-Object @{ Name = "Key";Expression = {$_}},
            @{ Name = "Value";Expression={$requestHeaders.GetValues($_)}}
     
        foreach ($i in $stringH)
        {
            Write-Host $i
        }
#>
        $requestUrl = $request.Url
        #Write-Host $requestUrl

        ### grab and decompress POST data
        $decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
        $readStream = [System.IO.StreamReader]::new($decompress)
        $content = $readStream.ReadToEnd()
        #Write-Host "$content"
        

        ### create JSON object
        $requestJson = $content | ConvertFrom-Json
        
        ### based on POST URL, select which kind of export is happening        
        switch -Wildcard ($requestUrl)
        {
            '*leagueteams' #LEAGUE INFO
            {
                $infoList = $requestJson.leagueTeamInfoList
                #Write-Host ($infoList | Format-Table -Property *| Out-String -Width 4096)
                $infoList | Export-Csv -Path "leagueInfo.csv" 
                Write-Host "LEAGUE INFO saved to leagueInfo.csv"
                
                break
            }

            '*standings' #LEAGUE INFO
            {
                $standingsList = $requestJson.teamStandingInfoList
                #Write-Host ($standingsList | Format-Table -Property *| Out-String -Width 4096)
                $standingsList | Export-Csv -Path "standingsInfo.csv" 
                Write-Host "STANDINGS saved to standingsInfo.csv"
                
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


       		#Write-Host ($teamList | Format-Table -Property *| Out-String -Width 4096)
                for ($i=2; $i -le 32; $i++)
                {
                     ### We can safely assume next 32 calls will be rosters so we run full loop, building roster file in memory
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
                   
                $teamList | Export-Csv -Path "rosters.csv"
                Write-Host "Export to disk:  rosters.csv" 
	        break
            }
	    
	        '*week*' #WEEKLY STATISTICS
            {
               Write-Host "STATISTICS:  Not yet implemented"
               
               <#
                $scheduleList = $requestJson.gameScheduleInfoList
                Write-Host ($scheduleList | Format-Table -Property *| Out-String -Width 4096)
                $scheduleList | Export-Csv -Path "scheduleInfo.csv" 
                Write-Host "SCHEDULES saved to scheduleInfo.csv"
                #>

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
    
    #Write-Host "////////////////////////////////////////////////////////////////////////////RESPONSE: $responseStatus"
    #$listener.Stop()

} while ($listener.IsListening)
