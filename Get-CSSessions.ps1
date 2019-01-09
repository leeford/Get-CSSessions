<# 
.SYNOPSIS
 
    Get-CSSessions - PowerShell script to get ALL user (inc. call queue) sessions using Get-CSUserSession
 
.DESCRIPTION

    Author: Lee Ford

    Using this script you can use Get-CSUserSession to gather ALL user sessions for ALL users between two dates. This will keep retrieving sessions,
    not just the first 1000 like Get-CSUserSession. You can filter on a particular user, a particular URI, all sessions or just specific (Audio, Conference, IM and Video) sessions, 
    include/exclude incomplete sessions etc. You can get sessions for a single user, a list of users in a CSV file or all users (priority in that order).
    
    For more details go to https://wp.me/p97Bkx-ec

.LINK

    Blog: https://www.lee-ford.co.uk
    Twitter: http://www.twitter.com/lee_ford
    LinkedIn: https://www.linkedin.com/in/lee-ford/
 
.EXAMPLE 
    
    .\Get-CSSessions.ps1 -SessionType Audio -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10
    Will retrieve all user audio sessions for the last 10 days from now and save as a CSV file.

    .\Get-CSSessions.ps1 -SessionType All -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10
    Will retrieve all user sessions for the last 10 days from now and save as a CSV file.

    .\Get-CSSessions.ps1 -SessionType Audio -OutputType GridView -DaysToSearch 10 -User user@domain.com
    Will retrieve user audio sessions for user@domain.com for the last 10 days from now and output to GridView.

    \Get-CSSessions.ps1 -SessionType Audio -OutputType GridView -DaysToSearch 10 -ImportUserCSV c:\temp\users.csv
    Will retrieve user audio sessions for all enabled users in c:\temp\users.csv for the last 10 days from now and output to GridView.

    .\Get-CSSessions.ps1 -SessionType All -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10 -EndDate "04/24/2018 18:00"
    Will retrieve all user sessions between 14th April 2018 18:00 to 24th April 2018 18:00 and save as a CSV file.
    
    .\Get-CSSessions.ps1 -SessionType Audio -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10 -IncludeIncomplete
    Will retrieve all user audio sessions for the last 10 days from now, including incomplete sessions and save as a CSV file.

    .\Get-CSSessions.ps1 -SessionType Audio -OutputType GridView -DaysToSearch 10 -URI +441234
    Will retrieve all user audio sessions for the last 10 days from now that contain a To or From URI of +441234 and output to a GridView

    .\Get-CSSessions.ps1 -SessionType All -OutputType GridView -DaysToSearch 10 -ClientVersion CPE
    Will retrieve all  audio sessions for the last 10 days from now that contain a To or From ClientVersion of CPE (Lync Phone Edition) and output to a GridView

    .\Get-CSSessions.ps1 -SessionType Audio -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10 -User user@domain.com -URI +441234 -EndDate "04/24/2018 18:00"
    Will retrieve user audio sessions for user@domain.com between 14th April 2018 18:00 to 24th April 2018 18:00 that contain a To or From URI of +441234 and save to a CSV file.

    .\Get-CSSessions.ps1 -SessionType Audio -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10 -Credential $credential
    Will retrieve all user audio sessions for the last 10 days from now, using PSCredential $credential specified and save as a CSV file.

    .\Get-CSSessions.ps1 -SessionType Conference -OutputType CSV -CSVSavePath C:\Temp\Sessions.csv -DaysToSearch 10 -AllInformation
    Will retrieve all user conference sessions for the last 10 days from now including full session information (e.g. QoE report) and save as a CSV file.

.NOTES
    v1.0 - Initial release
    v1.1 - Added ability to specify group of users from CSV file
    v1.2 - Added ClientVersion filter
     
#>

param(

    [Parameter(mandatory=$true)][Int32]$DaysToSearch,
    [Parameter(mandatory=$true)][ValidateSet('All', 'Audio', 'Conference','IM', 'Video')][string]$SessionType,
    [Parameter(mandatory=$true)][ValidateSet('CSV', 'GridView')][string]$OutputType,
    [Parameter(mandatory=$false)][string]$CSVSavePath,
    [Parameter(mandatory=$false)][switch]$AllInformation,
    [Parameter(mandatory=$false)][switch]$IncludeIncomplete,
    [Parameter(mandatory=$false)][string]$User,
    [Parameter(mandatory=$false)][string]$URI,
    [Parameter(mandatory=$false)][string]$ClientVersion,
    [Parameter(mandatory=$false)][datetime]$EndDate,
    [Parameter(mandatory=$false)][PSCredential]$Credential,
    [Parameter(mandatory=$false)][string]$ImportUserCSV
)

# Check prerequisites
function CheckPrereq {

    # Do you have Skype Online module installed?
    Write-Host "`nChecking Skype Online Module installed..."

    if (Get-Module -ListAvailable -Name SkypeOnlineConnector) {
    
        Write-Host "Skype Online Module installed." -ForegroundColor Green

    } else {

        Write-Error -Message "Skype Online Module not installed, please install and try again."
        
        break

    }

    # Is OutputType CSV?
    if($OutputType -eq "CSV") {

        # Is the path specified?
        if($CSVSavePath) {
        
            # Can you write to it?
            try {

                Write-Host "`nChecking CSV can be written to $CSVSavePath..."

                    "Can I write?" | Export-CSV -Path $CSVSavePath -NoTypeInformation

            } catch {

                Write-Error -Message "Unable to create CSV file - Is the file in use? Does the path exist?"

                break

            }

        } else {

            Write-Error -Message "OutputType set to CSV but no path specified using -CSVSavePath"

            break

        }

    }

}


# Connect to SfBO
function ConnectSkypeOnline {

    # If you have PSCredentials provided via -Credential set, use them and create a session
    if ($Credential) {

        $global:SfBOPSSession = New-CsOnlineSession -Credential $Credential

    # Else create a session and ask for details (including MFA)
    } else {

        $global:SfBOPSSession = New-CsOnlineSession

    }
    
    # Connect to SfBO
    Write-Host "`nConnecting to Skype for Business Online..."
    Import-PSSession $global:SfBOPSSession -AllowClobber | Out-Null

    # Start Session Time
    $global:SfBOPSSessionStartTime = Get-Date

    Write-Host "Connected to Skype for Business Online" -ForegroundColor Green

    Write-Host "`nAttempting to set WSMan Network Delay to 60 seconds (to help with timeouts)..."

    Set-WinRMNetworkDelayMS 60000 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 

}

# Go and find all enabled users for SfBO and return them
function FindEnabledSkypeOnlineUsers {

    Write-Host "`nGathering all enabled users in Skype for Business Online..."

    # Find all users and only keep their DisplayName and SipAddress
    $EnabledUsers = Get-CsOnlineUser -Filter "(Enabled -eq '$True')" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select-Object SipAddress, DisplayName

    # Count all users found
    $global:UserCount = $EnabledUsers.Count
    
    # If no users found, quit
    if (!$EnabledUsers) {

        Write-Warning -Message "No enabled Skype Online users found. Qutting..."
        
        Remove-PSSession $global:SfBOPSSession

        break  
        
    }

    Write-Host "Found" $global:UserCount "enabled users." -ForegroundColor Green

    Return $EnabledUsers
        
}

# Go through each user and grab sessions
function ProcessUsers($Users) {

    # All Sessions array
    $AllSessions = @()

    # Set total sessions to 0
    $global:TotalSessions = 0

    # Put users in alphabetical order
    $Users = $Users | Sort-Object -Property SipAddress

    # Cycle through each user
    foreach ($User in $Users) {

        # Progress counter
        $Counter++

        # Get user sessions
        $AllSessions += ProcessUserSessions $User $Counter

    }

    # Remove any empty items from array you have may have been given
    $AllSessions = $AllSessions | Where-Object {$_}
        
    # Elapsed time
    $ElapsedTime = ((Get-Date)-$global:Runtime)

    Write-Host "`nFrom $($global:UserCount) users: $($global:TotalSessions) sessions found in total, of which there are $(@($AllSessions).Count) matching sessions. `nElapsed time: $($ElapsedTime.TotalMinutes) minutes ($($global:TotalSessions/$ElapsedTime.TotalSeconds) sessions per second)." -ForegroundColor Green

    # If you have some matching sessions to output
    if($AllSessions) {

        # Output depending on type
        switch ($OutputType) {

            "CSV" { ExportSessionsToCSV $AllSessions }

            "GridView" { $AllSessions | Out-GridView }

        }

    }

}

function ProcessUserSessions($User, $Counter) {

    # Check session timer before you start
    CheckSessionTimer

    # Update Progress
    UpdateProgress $Counter $User.SipAddress
    
    # Strip sip: from SIP Address
    $SipURI = $User.SipAddress.Replace("sip:","")
    $SipURI = $SipURI.Replace("SIP:","")
    
    Write-Host "`nChecking $SipURI for sessions..."

    # Reset sessions for next query
    $UserSessions = @()
    $UserMatchedSessions = @()

    # Get initial 1000 user sessions
    $UserSessions = GetSessions $SipURI $User.DisplayName $global:StartTime

    # User sessions returned from current query for user (including IM, conference etc)
    $UserSessionsReturned = $UserSessions.Count

    # Total user sessions returned across all queries for user (including IM, conference etc)
    $TotalUserSessionsReturned = $UserSessions.Count

    # Filter out matched sessions
    $UserMatchedSessions = FilterSessions $UserSessions

    # Sort so the oldest record returned is first - if 1000 sessions are returned you then know where to start the next search from
    $UserSessions = $UserSessions | Sort-Object -Property EndTime
    
    # While 1000 sessions are still being returned, you need to go back for more, starting from the oldest record you have
    while ($UserSessionsReturned -ge 1000) {

        # Check session timer before next pass
        CheckSessionTimer

        # Update Progress
        UpdateProgress $Counter $User.SipAddress
        
        # Reset count for additional sessions
        $UserSessionsReturned = 0

        # Make a new start time based on oldest record you have
        $NewStartTime = $UserSessions[-1].EndTime
        
        # Reset sessions for next query
        $UserSessions = $null
            
        Write-Host "$TotalUserSessionsReturned sessions returned for $SipURI, going back for more."

        # Get additional user sessions
        $UserSessions = GetSessions $SipURI $user.DisplayName $NewStartTime

        # User sessions returned from current query (including IM, conference etc)
        $UserSessionsReturned = $UserSessions.Count

        # Total user sessions returned across all queries for user (including IM, conference etc)
        $TotalUserSessionsReturned += $UserSessions.Count

        # Add matched sessions to existing sessions
        $UserMatchedSessions += FilterSessions $UserSessions

        # Reorder list again for next pass
        $UserSessions = $UserSessions | Sort-Object -Property EndTime
        
    }
    
    # Found x sessions for user y
    Write-Host "Found $TotalUserSessionsReturned sessions for $SipURI, of which there are $(@($UserMatchedSessions).Count) matching sessions." -ForegroundColor Green
    
    return $UserMatchedSessions

}

function ExportSessionsToCSV($AllSessions) {

    Write-Host "`nExporting to CSV file to $CSVSavePath..."

    try {

        # Export Sessions to CSV
        $AllSessions | Export-CSV -Path $CSVSavePath -NoTypeInformation


    } catch {

        Write-Error -Message "Unable to create CSV file - Is the file in use? Does the path exist?"

        Remove-PSSession $global:SfBOPSSession

        break

     }

    Write-Host "Done!" -ForegroundColor Green

}

function GetSessions($SipURI, $DisplayName, $StartTime) {

    try {
    
        # Get users sessions and also add the SipURI and DisplayName of the user you are querying (makes it easier to know who's session details it was for)
        $UserSessions = Get-CsUserSession -User $SipURI -StartTime $StartTime.ToUniversalTime() -EndTime $global:EndTime.ToUniversalTime() -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object *,@{label=”SipURI”; Expression= {$SipURI}},@{label=”DisplayName”; Expression= {$DisplayName}}
    
    # Catch a timeout
    } catch {

          # Let's try again before moving on
          Write-Error -Message "Error getting sessions, creating new session and trying again."

          # Close all PSSessions
          $global:SfBOPSSession | Remove-PSSession

          # Wait 10 seconds
          Start-Sleep -Seconds 10
            
          # Reconnect (hopefully using cached creds from -Credential)
          ConnectSkypeOnline

          # Try again, but don't stop this time
          Get-CsUserSession -User $SipURI -StartTime $StartTime.ToUniversalTime() -EndTime $global:EndTime.ToUniversalTime() -WarningAction SilentlyContinue | Select-Object *,@{label=”SipURI”; Expression= {$SipURI}},@{label=”User”; Expression= {$DisplayName}}

    }

    # If not all information is required, leave the most useful (in my opinion)
    if (!$AllInformation) {

        $UserSessions = $UserSessions | Select-Object DialogId, StartTime, EndTime, FromURI, ToURI, FromTelNumber, ToTelNumber, ReferrredBy, FromClientVersion, ToClientVersion, MediaTypesDescription, SipURI, DisplayName

    }

    # Add to running sessions count
    $global:TotalSessions += $UserSessions.Count
    
    return $UserSessions

}

function CheckSessionTimer() {

    # Calculate current session elapsed time
    $PSSessionTimer = ((Get-Date) - $global:SfBOPSSessionStartTime)    
    
    # Check if session is over 45 minutes old
    if ($global:SfBOPSSession.State -eq "Opened" -and $PSSessionTimer.TotalMinutes -ge "45") {

        Write-Warning -Message "`nPowerShell session time at $($PSSessionTimer.TotalMinutes) minutes. Closing session and creating a new session..."
            
        # Close all PSSessions
        $global:SfBOPSSession| Remove-PSSession

        # Wait 10 seconds
        Start-Sleep -Seconds 10
            
        # Reconnect (hopefully using cached creds)
        ConnectSkypeOnline
     
    }
}

function FilterSessions($UserSessions) {

    # If all not all session types, 
    if($SessionType -ne "All") {
    
        # Remove sessions that aren't the session type specified
        $UserSessions = $UserSessions | Where-Object {$_.MediaTypesDescription -like "*$SessionType*"}
    
    }

    # Specfic URI set?
    if ($URI) {

        $UserSessions = $UserSessions | Where-Object {$_.FromURI -like "*$URI*" -or $_.ToURI -like "*$URI*"}

    }

    # Specfic ClientVersion set?
    if ($ClientVersion) {

        $UserSessions = $UserSessions | Where-Object {$_.FromClientVersion -like "*$ClientVersion*" -or $_.ToClientVersion -like "*$ClientVersion*"}

    }

    # Remove Incomplete sessions from list if required by removing sessions with no endtime
    if (!$IncludeIncomplete) {

        $UserSessions = $UserSessions | Where-Object {$_.EndTime}

    }

    return $UserSessions

}

function UpdateProgress($counter, $UserSipAddress) {

    # If more than one user, track progress:        
    if ($global:UserCount -gt 1) {

        Write-Progress -Activity "Processing Users... SfBO Session Timer: $((Get-Date) - $global:SfBOPSSessionStartTime) Runtime: $((Get-Date)-$global:Runtime) Processed Sessions: $($global:TotalSessions)" -Status "Processing User $counter of $global:UserCount" -CurrentOperation $UserSipAddress.Replace("sip:","") -PercentComplete (($counter/$global:UserCount) * 100)

    }

}


# Start
Write-Host "`n----------------------------------------------------------------------------------------------
            `n Get-CSSessions.ps1 v1.2 - Lee Ford 2018 - https://www.lee-ford.co.uk
            `n----------------------------------------------------------------------------------------------" -ForegroundColor Yellow

# Set end date/time for (now or specified paramter EndDate)
if($EndDate) {

    $global:EndTime = $EndDate

} else {

    $global:EndTime = Get-Date

}

# Set start date/time (end date/time minus how many days)
$global:StartTime = $global:EndTime.AddDays(-$DaysToSearch)

# Getting sessions based on...
Write-Host "`nGetting sessions based on:
`r- Between $($global:StartTime) and $($global:EndTime)
`r- Session type: $SessionType"
    
if($IncludeIncomplete) { Write-Host "- Including incomplete sessions" }
if($User) { Write-Host "- Only sessions for: $User" }
if($URI) { Write-Host "- Only sessions containing URI: $URI" }
if($ImportUserCSV) { Write-Host "- Users contained in CSV file: $ImportUserCSV" }

# Check you have the correct module installed and can write to the path (before you start at least)
CheckPrereq

# Start the clock
$global:Runtime = (Get-Date)

# Connect to SfBO
ConnectSkypeOnline

$EnabledUsers = @()

# If single user specified
if ($User) {

    # If sip: from SIP address is missing, add it
    if (!$User.StartsWith("sip:")) {

        $User = "sip:$User"

    }

    $EnabledUsers = Get-CsOnlineUser -Filter "(Enabled -eq '$True') -and (SipAddress -eq '$User')" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select-Object SipAddress, DisplayName
    
    $global:UserCount = 1

    if(!$EnabledUsers) {

        Write-Warning -Message "$User not found in Skype Online users. Qutting..."
        
        break
        
        Remove-PSSession $global:SfBOPSSession
    }

$global:UserCount = 1

# If CSV Import specified
} elseif ($ImportUserCSV) {

    $ImportedUsers = Import-CSV -Path $ImportUserCSV

    # Check Imported Users are enabled

    foreach ($ImportedUser in $ImportedUsers) {

        # If sip: from SIP address is missing, add it
        if (!$ImportedUser.User.StartsWith("sip:")) {

            $ImportedUser.User = "sip:$($ImportedUser.User)"

        }

        # Reset Result
        $ImportedUserResult = $null

        # Get Result
        $ImportedUserResult = Get-CsOnlineUser -Filter "(Enabled -eq '$True') -and (SipAddress -eq '$($ImportedUser.User)')" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select-Object SipAddress, DisplayName    

        # User exists and is enabled for SfB
        if ($ImportedUserResult) {

            Write-Host "$($ImportedUser.User) exists and is enabled for SfB/Teams"

            # Add to EnabledUsers
            $EnabledUsers += $ImportedUserResult

        } else {

            Write-Warning "$($ImportedUser.User) does not exist or is not enabled for SfB/Teams"

        }
    
    }

    # Count Users
    $global:UserCount = $EnabledUsers.Count
    Write-Host "$($EnabledUsers.Count) Enabled Users found in $ImportUserCSV" -ForegroundColor Green

}
    
# No speficied user(s) so, find all users
else {

    $EnabledUsers = FindEnabledSkypeOnlineUsers

}

ProcessUsers $EnabledUsers

# Close Sessions with SfBO
Remove-PSSession $global:SfBOPSSession