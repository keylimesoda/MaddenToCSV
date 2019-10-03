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
        "49ers" {return "1009778688"}
        "Bears" {return "1009778689"}
        "Bengals" {return "1009778690"}
        "Bills" {return "1009778691"}
        "Broncos" {return "1009778692"}
        "Browns" {return "1009778693"}
        "Buccaneers" {return "1009778694"}
        "Cardinals" {return "1009778695"}
        "Chargers" {return "1009778728"}
        "Cheifs" {return "1009778729"}
        "Colts" {return "1009778730"}
        "Cowboys" {return "1009778731"}
        "Dolphins" {return "1009778732"}
        "Eagles" {return "1009778733"}
        "Falcons" {return "1009778734"}
        "Giants" {return "1009778736"}
        "Jaguars" {return "1009778738"}
        "Jets" {return "1009778739"}
        "Lions" {return "1009778740"}
        "Packers" {return "1009778742"}
        "Panthers" {return "1009778743"}
        "Patriots" {return "1009778744"}
        "Raiders" {return "1009778745"}
        "Rams" {return "1009778746"}
        "Ravens" {return "1009778747"}
        "Redskins" {return "1009778748"}
        "Saints" {return "1009778749"}
        "Seahawks" {return "1009778750"}
        "Steelers" {return "1009778751"}
        "Texans" {return "1009778752"}
        "Titans" {return "1009778753"}
        "Vikings" {return "1009778754"}
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
        "1009778688" {return 15}
        "1009778689" {return 1}
        "1009778690" {return 2}
        "1009778691" {return 3}
        "1009778692" {return 4}
        "1009778693" {return 5}
        "1009778694" {return 6}
        "1009778695" {return 7}
        "1009778728" {return 8}
        "1009778729" {return 9}
        "1009778730" {return 10}
        "1009778731" {return 11}
        "1009778732" {return 12}
        "1009778733" {return 13}
        "1009778734" {return 14}
        "1009778736" {return 16}
        "1009778738" {return 17}
        "1009778739" {return 18}
        "1009778740" {return 19}
        "1009778742" {return 20}
        "1009778743" {return 21}
        "1009778744" {return 22}
        "1009778745" {return 23}
        "1009778746" {return 24}
        "1009778747" {return 25}
        "1009778748" {return 26}
        "1009778749" {return 27}
        "1009778750" {return 28}
        "1009778751" {return 29}
        "1009778752" {return 32}
        "1009778753" {return 30}
        "1009778754" {return 31}
    }
    return 1009
}

#//////////////////////////////////////////////////////////////////
#//Function to take a teamID value from Madden Companion App and 
#//convert it into a teamID as understood by AMP
#//////////////////////////////////////////////////////////////////
function GetCollegeID ($collegeAppVal)
{
    switch ($collegeAppVal)
    {
        "Abilene Chr." {return 1}
        "Air Force" {return 2}
        "Akron" {return 3}
        "Alabama" {return 4}
        "Alabama A&M" {return 5}
        "Alabama St." {return 6}
        "Albany" {return 280}
        "Albion College" {return 358}
        "Alcorn St." {return 7}
        "Appalach. St." {return 8}
        "Arizona" {return 9}
        "Arizona St." {return 10}
        "Arkansas" {return 11}
        "Arkansas P.B." {return 12}
        "Arkansas St." {return 13}
        "Army" {return 14}
        "Ashland" {return 360}
        "Assumption" {return 391}
        "Auburn" {return 15}
        "Augustana" {return 177}
        "Austin Peay" {return 16}
        "Azusa Pacific" {return 439}
        "Baker" {return 371}
        "Ball State" {return 17}
        "Baylor" {return 18}
        "Belhaven" {return 356}
        "Beloit College" {return 372}
        "Bemidji State" {return 394}
        "Bentley College" {return 295}
        "Beth Cookman" {return 19}
        "Bethel" {return 367}
        "Bloomsburg" {return 282}
        "Boise State" {return 20}
        "Boston College" {return 21}
        "Bowie State" {return 364}
        "Bowling Green St." {return 22}
        "Bridgewater St." {return 336}
        "Brown" {return 23}
        "Bucknell" {return 24}
        "Buffalo" {return 25}
        "Buffalo State" {return 285}
        "Butler" {return 26}
        "BYU" {return 27}
        "Cal Lutheran" {return 422}
        "Cal Poly SLO" {return 28}
        "Cal-Bakersfield" {return 293}
        "Calgary" {return 326}
        "California" {return 29}
        "California (PA)" {return 307}
        "California-Davis" {return 286}
        "Cal-Northridge" {return 30}
        "Cal-Sacramento" {return 31}
        "Campbell Univ." {return 410}
        "Canisius" {return 32}
        "Carroll College" {return 304}
        "Carson-Newman" {return 287}
        "Catawba College" {return 383}
        "Cent Conn St." {return 33}
        "Central Arkansas" {return 288}
        "Central Michigan" {return 283}
        "Central Missouri" {return 385}
        "Central Oklahoma" {return 361}
        "Central St Ohio" {return 35}
        "Central Wash." {return 294}
        "Centre College" {return 392}
        "Chadron State" {return 363}
        "Charleston S." {return 36}
        "Charlotte" {return 423}
        "Chattanooga" {return 289}
        "Cincinnati" {return 37}
        "Citadel" {return 38}
        "Clarion" {return 349}
        "Clemson" {return 39}
        "Clinch Valley" {return 40}
        "Coastal Carolina" {return 290}
        "Coe College" {return 329}
        "Colgate" {return 41}
        "Colorado" {return 42}
        "Colorado St." {return 43}
        "Columbia" {return 44}
        "Concordia" {return 368}
        "Connecticut" {return 228}
        "Cornell" {return 45}
        "CSU-Pueblo" {return 378}
        "Culver-Stockton" {return 46}
        "Cumberlands" {return 414}
        "Dartmouth" {return 47}
        "Davidson" {return 48}
        "Dayton" {return 49}
        "Delaware" {return 50}
        "Delaware St." {return 51}
        "Delta State" {return 298}
        "Drake" {return 52}
        "Dubuque" {return 441}
        "Duke" {return 53}
        "Duquesne" {return 54}
        "E. Illinois" {return 56}
        "E. Kentucky" {return 57}
        "E. Tenn. St." {return 58}
        "East Carolina" {return 59}
        "East Central Univ." {return 384}
        "East Stroudsburg" {return 292}
        "East Texas Baptist" {return 448}
        "Eastern Michigan" {return 34}
        "Eastern Oregon" {return 344}
        "Eastern Wash." {return 60}
        "Elon University" {return 61}
        "Emporia State" {return 276}
        "Fairfield" {return 62}
        "Fairmont State" {return 444}
        "FAU" {return 308}
        "Faulkner" {return 427}
        "Fayetteville State" {return 406}
        "Ferris State" {return 296}
        "Findlay" {return 418}
        "FIU" {return 297}
        "Florida" {return 63}
        "Florida A&M" {return 64}
        "Florida State" {return 65}
        "Florida Tech" {return 405}
        "Fordham" {return 66}
        "Fort Hays State" {return 373}
        "Fort Valley State" {return 299}
        "Franklin College" {return 275}
        "Fresno State" {return 67}
        "Frostburg State" {return 446}
        "Furman" {return 68}
        "Ga. Southern" {return 69}
        "Gardner-Webb" {return 300}
        "Georgetown" {return 70}
        "Georgia" {return 71}
        "Georgia State" {return 362}
        "Georgia Tech" {return 72}
        "Globe Tech NY" {return 428}
        "Grambling St." {return 73}
        "Grand Valley St." {return 74}
        "Greenville College" {return 429}
        "Hampton" {return 75}
        "Harding" {return 301}
        "Harvard" {return 76}
        "Hastings College" {return 268}
        "Hawaii" {return 77}
        "Heidelberg" {return 388}
        "Henderson St." {return 78}
        "Hillsdale" {return 354}
        "Hobart" {return 399}
        "Hofstra" {return 79}
        "Holy Cross" {return 80}
        "Houston" {return 81}
        "Howard" {return 82}
        "Humboldt State" {return 377}
        "Huntingdon" {return 375}
        "Idaho" {return 83}
        "Idaho State" {return 84}
        "Illinois" {return 85}
        "Illinois St." {return 86}
        "Incarnate Word" {return 419}
        "Indiana" {return 87}
        "Indiana St." {return 88}
        "Iona" {return 89}
        "Iowa" {return 90}
        "Iowa State" {return 91}
        "IUP" {return 273}
        "J. Madison" {return 92}
        "Jackson St." {return 93}
        "Jacksonv. St." {return 94}
        "Jacksonville Univ." {return 407}
        "John Carroll" {return 95}
        "Kansas" {return 96}
        "Kansas State" {return 97}
        "Kent State" {return 98}
        "Kentucky" {return 99}
        "Kentucky Wesleyan" {return 438}
        "Knoxville College" {return 345}
        "Kutztown" {return 100}
        "La Salle" {return 101}
        "LA. Tech" {return 102}
        "Lafayette" {return 302}
        "Lake Erie College" {return 430}
        "Lamar Univ." {return 403}
        "Lambuth" {return 103}
        "Lane" {return 303}
        "Laval" {return 431}
        "Lehigh" {return 104}
        "Liberty" {return 105}
        "Lindenwood" {return 324}
        "Louisiana College" {return 402}
        "Louisville" {return 106}
        "LSU" {return 107}
        "M. Valley St." {return 108}
        "Maine" {return 109}
        "Manitoba" {return 348}
        "Marian" {return 424}
        "Marist" {return 110}
        "Mars Hill" {return 401}
        "Marshall" {return 111}
        "Mary Hardin-Baylor" {return 381}
        "Maryland" {return 112}
        "Maryville College" {return 421}
        "Massachusetts" {return 113}
        "McGill Univ." {return 398}
        "McNeese St." {return 114}
        "Memphis" {return 115}
        "Merrimack" {return 203}
        "Mesa State" {return 306}
        "Miami" {return 116}
        "Miami Univ." {return 117}
        "Michigan" {return 118}
        "Michigan St." {return 119}
        "Michigan Tech" {return 365}
        "Mid Tenn St." {return 120}
        "Midwestern St." {return 269}
        "Minnesota" {return 121}
        "Minnesota State" {return 390}
        "Mississippi College" {return 432}
        "Mississippi St." {return 55}
        "Missouri" {return 123}
        "Missouri So. State" {return 309}
        "Missouri State" {return 310}
        "Missouri W State" {return 311}
        "Monmouth" {return 124}
        "Montana" {return 125}
        "Montana State" {return 126}
        "Montreal Univ." {return 389}
        "Morehead St." {return 127}
        "Morehouse" {return 128}
        "Morgan St." {return 129}
        "Morningside College" {return 417}
        "Morris Brown" {return 130}
        "Mount Union" {return 312}
        "Mt S. Antonio" {return 131}
        "Murray State" {return 132}
        "N. Alabama" {return 133}
        "N. Arizona" {return 134}
        "N. Colorado" {return 137}
        "N. Illinois" {return 138}
        "N.C. A&T" {return 135}
        "N.C. State" {return 139}
        "N/A" {return 0}
        "Navy" {return 140}
        "NC Central" {return 141}
        "Nebr.-Omaha" {return 142}
        "Nebraska" {return 143}
        "Nebraska-Kearney" {return 313}
        "Nevada" {return 144}
        "New Hampshire" {return 266}
        "New Mex. St." {return 145}
        "New Mexico" {return 146}
        "Newberry College" {return 359}
        "Nicholls St." {return 147}
        "No College" {return 265}
        "None" {return 314}
        "Norfolk State" {return 148}
        "North Carolina" {return 122}
        "North Dakota" {return 270}
        "North Dakota St." {return 315}
        "North Greenville" {return 386}
        "North Texas" {return 149}
        "Northeastern" {return 150}
        "Northeastern St." {return 396}
        "Northern Iowa" {return 151}
        "Northern Michigan" {return 411}
        "Northern State" {return 316}
        "Northwestern" {return 152}
        "Northwestern State" {return 420}
        "Northwood (MI)" {return 318}
        "Notre Dame" {return 153}
        "Notre Dame College" {return 443}
        "NW Missouri State" {return 317}
        "NW Oklahoma St." {return 154}
        "N'western St." {return 155}
        "Ohio" {return 156}
        "Ohio Northern" {return 319}
        "Ohio State" {return 157}
        "Oklahoma" {return 158}
        "Oklahoma St." {return 159}
        "Old Dominion" {return 380}
        "Ole Miss" {return 160}
        "Oregon" {return 161}
        "Oregon State" {return 162}
        "Ottawa" {return 320}
        "Ouachita Baptist" {return 370}
        "P. View A&M" {return 163}
        "Palomar College" {return 387}
        "Penn" {return 164}
        "Penn State" {return 165}
        "Pikeville College" {return 321}
        "Pittsburg St." {return 166}
        "Pittsburgh" {return 167}
        "Portland St." {return 168}
        "Presbyterian" {return 281}
        "Pretoria" {return 413}
        "Princeton" {return 169}
        "Purdue" {return 170}
        "Queen's Univ." {return 369}
        "Ramapo" {return 322}
        "Regina" {return 323}
        "Rensselaer Poly" {return 415}
        "Rhode Island" {return 171}
        "Rice" {return 172}
        "Richmond" {return 173}
        "Robert Morris" {return 174}
        "Rowan" {return 175}
        "Rutgers" {return 176}
        "S. Connecticut St." {return 347}
        "S. Dakota St." {return 178}
        "S. Illinois" {return 179}
        "S.C. State" {return 180}
        "Sacramento State" {return 325}
        "Sacred Heart" {return 183}
        "Saginaw Valley" {return 274}
        "Salisbury" {return 357}
        "Sam Houston" {return 184}
        "Samford" {return 185}
        "San Diego" {return 327}
        "San Diego St." {return 181}
        "San Jose St." {return 186}
        "Savannah St." {return 187}
        "Scottsbluff JC" {return 376}
        "SE Louisiana" {return 330}
        "SE Missouri" {return 188}
        "SE Missouri St." {return 189}
        "Seattle" {return 433}
        "Shepherd Univ." {return 395}
        "Shippensburg" {return 190}
        "Siena" {return 191}
        "Simon Fraser" {return 192}
        "Sioux Falls" {return 447}
        "Slippery Rock" {return 366}
        "SMU" {return 193}
        "Sonoma St." {return 264}
        "South Alabama" {return 379}
        "South Carolina" {return 136}
        "South Dakota" {return 328}
        "Southern" {return 194}
        "Southern Arkansas" {return 205}
        "Southern Miss" {return 195}
        "Southern Nazarene " {return 449}
        "Southern Oregon" {return 426}
        "Southern Utah" {return 196}
        "St. Augustine" {return 333}
        "St. Cloud State" {return 305}
        "St. Francis" {return 197}
        "St. John's" {return 198}
        "St. Mary's" {return 199}
        "St. Paul's" {return 340}
        "St. Peters" {return 200}
        "Stanford" {return 201}
        "Stephen F. Austin" {return 353}
        "Stetson" {return 434}
        "Stillman" {return 331}
        "Stony Brook" {return 202}
        "SW Miss St" {return 204}
        "SW Oklahoma State" {return 404}
        "Syracuse" {return 206}
        "T A&M K'ville" {return 207}
        "Tarleton State" {return 334}
        "TCU" {return 208}
        "Temple" {return 209}
        "Tenn. Tech" {return 210}
        "Tenn-Chat" {return 211}
        "Tennessee" {return 212}
        "Tennessee St." {return 213}
        "Tenn-Martin" {return 214}
        "Texas" {return 215}
        "Texas A&M" {return 216}
        "Texas A&M-Commerce" {return 400}
        "Texas South." {return 217}
        "Texas State" {return 332}
        "Texas Tech" {return 218}
        "Texas-Permian Basin" {return 435}
        "Tiffin" {return 351}
        "Toledo" {return 219}
        "Towson State" {return 220}
        "Trinity" {return 352}
        "Troy" {return 221}
        "Truman State" {return 335}
        "Tulane" {return 222}
        "Tulsa" {return 223}
        "Tusculum College" {return 291}
        "Tuskegee" {return 224}
        "UAB" {return 225}
        "UBC" {return 445}
        "UC Irvine" {return 397}
        "UCF" {return 226}
        "UCLA" {return 227}
        "UL Lafayette" {return 229}
        "UL Monroe" {return 230}
        "Union College" {return 249}
        "UNLV" {return 231}
        "USC" {return 232}
        "USF" {return 233}
        "Utah" {return 234}
        "Utah State" {return 235}
        "UTEP" {return 236}
        "UTSA" {return 355}
        "UW Lacrosse" {return 267}
        "UW Stevens Pt." {return 272}
        "UW-Milwaukee" {return 393}
        "Valdosta St." {return 237}
        "Valparaiso" {return 238}
        "Vanderbilt" {return 239}
        "Villanova" {return 240}
        "Virginia" {return 241}
        "Virginia Commonwealth" {return 436}
        "Virginia State" {return 442}
        "Virginia Tech" {return 242}
        "Virginia Union" {return 337}
        "Virginia-Lynchburg" {return 437}
        "VMI" {return 243}
        "W. Carolina" {return 244}
        "W. Illinois" {return 245}
        "W. Kentucky" {return 246}
        "W. Michigan" {return 247}
        "W. New Mexico" {return 279}
        "W. Texas A&M" {return 248}
        "Wagner College" {return 182}
        "Wake Forest" {return 250}
        "Walla Walla" {return 251}
        "Walsh" {return 374}
        "Wash. St." {return 252}
        "Washburn" {return 338}
        "Washington" {return 253}
        "Wayne State" {return 271}
        "Weber State" {return 254}
        "Wesley College" {return 408}
        "West Alabama" {return 382}
        "West Chester" {return 416}
        "West Georgia" {return 342}
        "West Virginia" {return 255}
        "Western Ontario" {return 346}
        "Western Oregon" {return 350}
        "Western State Colorado " {return 440}
        "Western Wash." {return 339}
        "Westminster" {return 256}
        "Wheaton" {return 278}
        "Whitworth" {return 284}
        "William & Mary" {return 257}
        "William Penn" {return 341}
        "Wingate" {return 277}
        "Winston Salem" {return 258}
        "Wisconsin" {return 259}
        "Wisconsin-Eau Claire" {return 409}
        "Wisconsin-Oshkosh" {return 412}
        "Wisc-Whitewater" {return 343}
        "Wofford" {return 260}
        "Wyoming" {return 261}
        "Yale" {return 262}
        "Youngstown St." {return 263}

    }
    return 1
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
        PCOL = GetCollegeID($curplayer.college)
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
     
        ### GZIP compression was apparently disabled for requests with 3 OCT 2019 update
        #$decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
        #$readStream = [System.IO.StreamReader]::new($decompress)
        
        $readStream = [System.IO.StreamReader]::new($request.InputStream)
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
                     
                     ### GZIP compression was apparently disabled for requests with 3 OCT 2019 update
                     #$decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
                     #$readStream = [System.IO.StreamReader]::new($decompress)
        
                     $readStream = [System.IO.StreamReader]::new($request.InputStream)
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
                ### GZIP compression was apparently disabled for requests with 3 OCT 2019 update
                #$decompress = [System.IO.Compression.GZipStream]::new($request.InputStream, [IO.Compression.CompressionMode]::Decompress)
                #$readStream = [System.IO.StreamReader]::new($decompress)
        
                $readStream = [System.IO.StreamReader]::new($request.InputStream)
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
