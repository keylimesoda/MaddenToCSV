<#
    .SYNOPSIS
        Function to export League Info, Schedules and Rosters to CSV files on-disk
        
    .DIRECTIONS
    	Before 1st run, you must enable powershell scripts on your computer
		Open a powershell window as administrator
		Type "Set-ExecutionPolicy Bypass"
	Run the MaddenToCSV.ps1 script
        options (none are required)
             -ipAddress 0.0.0.0     specify local IP address to use for listening to app
	     -outputAMP             output AMP Editor compatible roster file
             -scrimTeam1	    choose a team to swap with Bucs
	     -scrimTeam2	    choose a team to swap with the Saints
	Enter the server address shown on the PowerShell window into your Madden Companion App, and Export
	.CSV files will be saved onto your PC
	The script will run indefinitely.  Close the window when you're done.
	
    .NOTES
        Team names are not included in the stats and roster files. They must be mapped from the table in leagueInfo.csv using the TeamID.
	The app exports 8 stat tables for each week: (schedules, defense, kicking, punting, passing, receiving, rushing, teamstats)
#>

param(
    [string] $ipAddress=$null,
    [switch] $outputAMP=$false,
    [string] $scrimTeam1=$null, #//will map this team to Bucs
    [string] $scrimTeam2=$null  #//will map this team to Saints
)

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
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell"
   
   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition
   
   # Add input parameters
   if ($ipAddress -ne "") {$newProcess.Arguments += " -ipAddress $ipAddress"}
   if ($outputAMP)        {$newProcess.Arguments += " -outputAmp"}
   
   if ($scrimTeam1 -ne ""){$newProcess.Arguments += " -scrimTeam1 $scrimTeam1"}
   if ($scrimTeam2 -ne ""){$newProcess.Arguments += " -scrimTeam2 $scrimTeam2"}
   
   
   # Pass along the current working directory for any output
   $newProcess.WorkingDirectory = $PSScriptRoot
   
   # Indicate that the process should be elevated
   $newProcess.Verb = "runas"
   
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

#if no ipAddress specified, it will look for active local IP4 address
if ($ipAddress.Length -eq 0){
    $ipAddress = (Get-NetIPAddress | ?{ $_.AddressFamily -eq "IPv4"  -and !($_.IPAddress -match "169") -and !($_.IPaddress -match "127") }).IPAddress
}

$serverAddr = "http://"+$ipAddress+":$port/"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($serverAddr)

try
{
    $listener.Start()
}
catch
{
     Write-Host "Cannot start server at $serverAddr  Script already running in another window, or bad IP address?"
     Write-Host -NoNewLine 'Press any key to continue...';
     $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
     exit
}

$serverAddr = $serverAddr.TrimEnd('/')
Write-Host "Server started at $serverAddr"

#//////////////////////////////////////////////////////////////////
#//Function to take a positionID string from Madden Companion App and 
#//convert it into a positionID as understood by AMP
#//////////////////////////////////////////////////////////////////
function GetPositionID ($positionString)
{
    switch ($positionString)
    {
        "QB"   {return 0}
        "HB"   {return 1}
        "FB"   {return 2}
        "WR"   {return 3}
        "TE"   {return 4}
        "LT"   {return 5}
        "LG"   {return 6}
        "C"    {return 7}
        "RG"   {return 8}
        "RT"   {return 9}
        "LE"   {return 10}
        "RE"   {return 11}
        "DT"   {return 12}
        "LOLB" {return 13}
        "MLB"  {return 14}
        "ROLB" {return 15}
        "CB"   {return 16}
        "FS"   {return 17}
        "SS"   {return 18}
        "K"    {return 19}
        "P"    {return 20}
    }
}

#//////////////////////////////////////////////////////////////////
#//Function to take a team name and
#//convert it into a teamID as understood by AMP
#//////////////////////////////////////////////////////////////////
function TeamNameToID ($teamName)
{
    switch ($teamName)
    {
        "49ers" {return "778043392"}
        "Bears" {return "778043394"}
        "Bengals" {return "778043395"}
        "Bills" {return "778043396"}
        "Broncos" {return "778043397"}
        "Browns" {return "778043398"}
        "Buccaneers" {return "778043399"}
        "Bucs" {return "778043399"}
        "Cardinals" {return "778043400"}
        "Chargers" {return "778043401"}
        "Cheifs" {return "778043402"}
        "Colts" {return "778043403"}
        "Cowboys" {return "778043404"}
        "Dolphins" {return "778043405"}
        "Eagles" {return "778043406"}
        "Falcons" {return "778043407"}
        "Giants" {return "778043409"}
        "Jaguars" {return "778043411"}
        "Jets" {return "778043412"}
        "Lions" {return "778043413"}
        "Packers" {return "778043416"}
        "Panthers" {return "778043417"}
        "Patriots" {return "778043418"}
        "Raiders" {return "778043419"}
        "Rams" {return "778043420"}
        "Ravens" {return "778043421"}
        "Redskins" {return "778043422"}
        "Saints" {return "778043423"}
        "Seahawks" {return "778043424"}
        "Steelers" {return "778043425"}
        "Texans" {return "778043426"}
        "Titans" {return "778043427"}
        "Vikings" {return "778043428"}
   }
   return $false
}

#//////////////////////////////////////////////////////////////////
#//Function to take a teamID value from Madden Companion App and 
#//convert it into a teamID as understood by AMP
#//////////////////////////////////////////////////////////////////
function GetTeamID2 ($teamAppVal)
{
    switch ($teamAppVal)
    {
        "778043392" {return 15}
        "778043394" {return 1}
        "778043395" {return 2}
        "778043396" {return 3}
        "778043397" {return 4}
        "778043398" {return 5}
        "778043399" {return 6}
        "778043400" {return 7}
        "778043401" {return 8}
        "778043402" {return 9}
        "778043403" {return 10}
        "778043404" {return 11}
        "778043405" {return 12}
        "778043406" {return 13}
        "778043407" {return 14}
        "778043409" {return 16}
        "778043411" {return 17}
        "778043412" {return 18}
        "778043413" {return 19}
        "778043416" {return 20}
        "778043417" {return 21}
        "778043418" {return 22}
        "778043419" {return 23}
        "778043420" {return 24}
        "778043421" {return 25}
        "778043422" {return 26}
        "778043423" {return 27}
        "778043424" {return 28}
        "778043425" {return 29}
        "778043426" {return 32}
        "778043427" {return 30}
        "778043428" {return 31}
    }
    return 1009
}

#//////////////////////////////////////////////////////////////////
#//Function to take a teamID value from Madden Companion App and 
#//convert it into a teamID as understood by AMP
#//////////////////////////////////////////////////////////////////
function GetTeamID ($teamAppVal)
{
    #//if scrim teams are set
    if (($scrimTeam1 -ne "") -and ($scrimTeam2 -ne ""))
    {
        if ($teamAppVal -eq $scrimTeam1)
            {return 6}                        #//return scrimTeam1 as Bucs

        if ($teamAppVal -eq $scrimTeam2)
            {return 27}                       #//return scrimTeam2 to be Saints

        if ($teamAppVal -eq "778043399")  
            {return GetTeamID2($scrimTeam1)} #//return Bucs as be scrimTeam1

        if ($teamAppVal -eq "778043423")
            {return GetTeamID2($scrimTeam2)} #//return Saints as to be scrimTeam2
    }
    return (GetTeamID2($teamAppVal))
}

#//////////////////////////////////////////////////////////////////
#//Function to take a teamList object and output it to AMP
#//View column mappings here:  https://1drv.ms/x/s!Ah8EhteTsIhUjIAQI-DRUd-fdZpJew
#//////////////////////////////////////////////////////////////////
function TeamListToAMP ($localTeamList)
{
  
  Write-Host ("converting for AMP")

  #//if scrim teams present, change string name to App Value
  if (($scrimTeam1 -ne "") -and ($scrimTeam2 -ne ""))
  {
    Write-Host ("Scrim Team 1:  $scrimTeam1 will be swapped with Bucs")
    $scrimTeam1 = TeamNameToID($scrimTeam1)
    
    Write-Host ("Scrim Team 2:  $scrimTeam2 will be swapped with Saints") 
    $scrimTeam2 = TeamnameToID($scrimTeam2)
  }


  [System.Collections.ArrayList]$playerListAMP = @()
  $i = 1

  $localTeamList | ForEach-Object {
    $curPlayer       = $null
    $convertedObject = $null
    $curPlayer = $_
    Write-Host -NoNewline "`rplayer: $i"

    $convertedObject = [PSCustomObject]@{
    
        PFNA = $curplayer.firstName
        PLNA = $curplayer.lastName
        TGID = GetTeamID($curplayer.teamID)
        PGID = $i #//PGID must be unique.
        POID = $i #//must match PGID
        PACC = $curplayer.accelRating
        PAGE = $curplayer.age
        PAGI = $curplayer.agilityRating
        PAWR = $curplayer.awareRating
        PBCV = $curplayer.bCVRating
        TRBH = $curplayer.bigHitTrait
        PLBD = ([int32]$curplayer.birthDay -shl 11) + ([int32]$curplayer.birthMonth -shl 7) + ([int32]$curplayer.birthYear-1940) #//Birthday appears to be stored as a single integer with offsets
        PBSG = $curplayer.blockShedRating
        PBSK = $curplayer.breakSackRating
        PBKT = $curplayer.breakTackleRating
        PCAR = $curplayer.carryRating
        PCTH = $curplayer.catchRating
        PLCI = $curplayer.cITRating
        TRCL = $curplayer.clutchTrait
        PCOL = 1 #// HACK this should say $curplayer.college, but string has issues
        PYCF = $curplayer.confRating
        PSBO = [int32]$curplayer.contractBonus/10000
        PCON = $curplayer.contractLength
        PTSA = [int32]$curplayer.contractSalary/10000
        PCYL = $curplayer.contractYearsLeft
        TRCB = $curplayer.coverBallTrait
        PROL = $curplayer.devTrait
        TRBR = $curplayer.dLBullRushTrait
        TRDS = $curplayer.dLSpinTrait
        TRSW = $curplayer.dLSwimTrait
        PDRO = $curplayer.draftPick
        PDPI = $curplayer.draftRound
        TRDO = $curplayer.dropOpenPassTrait
        PELU = $curplayer.elusiveRating
        TRFB = $curplayer.feetInBoundsTrait
        TRFY = $curplayer.fightForYardsTrait
        PFMS = $curplayer.finesseMovesRating
        TRFP = $curplayer.forcePassTrait
        PHGT = $curplayer.height
        TRHM = $curplayer.highMotorTrait
        PLHT = $curplayer.hitPowerRating
        PHSN = $curplayer.homeState
        PHTN = "Springville" #// HACK this should say $curplayer.homeTown, but some contain commas, which break parsing
        TRJR = $curplayer.hPCatchTrait
        PLIB = $curplayer.impactBlockRating
        PINJ = $curplayer.injuryRating
        PJEN = $curplayer.jerseyNum
        PLJM = $curplayer.jukeMoveRating
        PJMP = $curplayer.jumpRating
        PKAC = $curplayer.kickAccRating
        PKPR = $curplayer.kickPowerRating
        PKRT = $curplayer.kickRetRating
        PLBK = $curplayer.leadBlockRating
        PLMC = $curplayer.manCoverRating
        PPBF = $curplayer.passBlockFinesseRating
        PPBS = $curplayer.passBlockPowerRating
        PPBK = $curplayer.passBlockRating
        TRIC = $curplayer.penaltyTrait
        PPLA = $curplayer.playActionRating
        TRPB = $curplayer.playBallTrait
        POVR = $curplayer.playerBestOvr
        PLPR = $curplayer.playRecRating
        PSXP = $curplayer.portraitId
        TRCT = $curplayer.posCatchTrait
        PPOS = GetPositionID ($curplayer.position)
        PLPm = $curplayer.powerMovesRating
        PCMT = 8191 #//$curplayer.presentationId (mapping is incorrect, value can't be read)
        PLPE = $curplayer.pressRating
        PLPU = $curplayer.pursuitRating
        PQBS = $curplayer.qBStyleTrait
        PLRL = $curplayer.releaseRating
        PGHE = 250 #//HACK $i #//Face ID, unsure how used
        PDRR = $curplayer.routeRunDeepRating
        PMRR = $curplayer.routeRunMedRating
        SRRN = $curplayer.routeRunShortRating
        PRBF = $curplayer.runBlockFinesseRating
        PRBS = $curplayer.runBlockPowerRating
        PRBK = $curplayer.runBlockRating
        TRSP = $curplayer.sensePressureTrait
        PLSC = $curplayer.specCatchRating
        PSPD = $curplayer.speedRating
        PLSM = $curplayer.spinMoveRating
        PSTA = $curplayer.staminaRating
        PLSA = $curplayer.stiffArmRating
        PSTR = $curplayer.strengthRating
        TRSB = $curplayer.stripBallTrait
        PTAK = $curplayer.tackleRating
        PTAD = $curplayer.throwAccDeepRating
        PTAM = $curplayer.throwAccMidRating
        PTHA = $curplayer.throwAccRating
        PTAS = $curplayer.throwAccShortRating
        TRTA = $curplayer.throwAwayTrait
        PTOR = $curplayer.throwOnRunRating
        PTHP = $curplayer.throwPowerRating
        PTUP = $curplayer.throwUnderPressureRating
        TRTS = $curplayer.tightSpiralTrait
        PTGH = $curplayer.toughRating
        PLTR = $curplayer.truckRating
        PWGT = $curplayer.weight-160
        TRWU = $curplayer.yACCatchTrait
        PYRP = $curplayer.yearsPro
        PLZC = $curplayer.zoneCoverRating
        PLHY = -31 #//not sure why, must always be this
        PLPL = 100 #//not sure why, must alayws be this
        PLPO = 99  #//not sure why, must alayws be this
        PRL2 = 31  #//not sure why, must always be this
        PSHG = 6   #//not sure why, must always be this
    }
    $i++
    $playerListAMP.Add($convertedObject) | Out-Null
  }

  $outputAMPFileName = "rostersAMP.csv"
  

  #//add AMP header
  $AMPOutputString = "PLAY,2019,No,"
  
  #//add CSV header
  for ($i=3; $i -le 107; $i++){  #//THIS MUST MATCH ATTRIBUTE COUNT FOR CUSTOM OBJECT
    $AMPOutputString += ','
  }
  $AMPOutputString += "`n"

  #//add data
  $AMPOutputString += $playerListAMP | ConvertTo-Csv -NoTypeInformation -Delimiter ","
  
  #//add newlines
  $AMPOutputString = $AMPOutputString | % {$_ -replace '" "',"`"`n`""}
  
  #//remove quotes
  $AMPOutputString = $AMPOutputString | % {$_ -replace '"',''}
    
  #//to disk
  $AMPOutputString | Out-File $outputAMPFileName -Encoding ascii


  Write-Host "`nExport to disk:  $outputAMPFileName "
}


#//////////////////////////////////////////////////////////////////
#//Start the listen/response loop
#//////////////////////////////////////////////////////////////////

Write-Host ""
Write-Host "Listening for Madden Companion App (close window to exit)"
Write-host ""
$requestDump = $false; #set to $true to get a raw JSON output dump of the first incoming request
       
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
        
        
        if ($requestDump){
          $requestDump = $false
          Write-Host ("dumping JSON")
          $content | Out-File "team.json"

        }


        ### create JSON object
        $requestJson = $content | ConvertFrom-Json
        
        ### based on POST URL, select which kind of export is happening        
        switch -Wildcard ($requestUrl)
        {
            '*leagueteams' #LEAGUE INFO
            {
                $infoList = $requestJson.leagueTeamInfoList
                #Write-Host ($infoList | Format-Table -Property *| Out-String -Width 4096)  #//enable this line to see JSON table output to command line
                $infoList | Export-Csv -Path "leagueInfo.csv" -NoTypeInformation
                Write-Host "LEAGUE INFO `t => leagueInfo.csv"
                
                break
            }

            '*standings' #LEAGUE INFO
            {
                $standingsList = $requestJson.teamStandingInfoList
                #Write-Host ($standingsList | Format-Table -Property *| Out-String -Width 4096) #//enable this line to see JSON table output to command line
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
                
                if ($outputAMP -eq $true){TeamListToAmp($teamList)}               


		break
            }	    
	        '*week*' #WEEKLY STATISTICS
            {
                ### Get the week number that we'll use as a prefix for the stats files
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
