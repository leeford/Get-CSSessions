# Get-CSSessions
PowerShell script to get ALL SfB user sessions

For Microsoft Teams PSTN call records, please look at [Get-TeamsPSTNCallRecords](https://github.com/leeford/Get-TeamsPSTNCallRecords)

## SYNOPSIS
 
Get-CSSessions - PowerShell script to get ALL SfB user (inc. call queue) sessions using Get-CSUserSession
 
## DESCRIPTION

Author: Lee Ford

Using this script you can use Get-CSUserSession to gather ALL user sessions for ALL users between two dates. This will keep retrieving sessions, not just the first 1000 like Get-CSUserSession. You can filter on a particular user, a particular URI, all sessions or just specific (Audio, Conference, IM and Video) sessions, include/exclude incomplete sessions etc. You can get sessions for a single user, a list of users in a CSV file or all users (priority in that order).
   
For more details go to https://wp.me/p97Bkx-ec

## LINK

Blog: https://www.lee-ford.co.uk

Twitter: http://www.twitter.com/lee_ford

LinkedIn: https://www.linkedin.com/in/lee-ford/
 
## EXAMPLE
   
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


## NOTES
v1.0 - Initial release

v1.1 - Added ability to specify group of users from CSV file

v1.2 - Added ClientVersion filter
