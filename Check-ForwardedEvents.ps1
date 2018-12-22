[CmdletBinding()]
param()

#region (Functions)
### email alert function ###
function sendMailAlert 
{
    [CmdletBinding()]
    param([string] $subject,
          [string] $body,
          $attachment
         )

    $MailMessage = @{
        From = "rdi-us-sv0111v_SecurityAudit@landesa.org"
        To = "alexo@landesa.org"
        Subject = $subject
        Body = $body
        SmtpServer = "10.0.0.10"
    }

    Send-MailMessage @MailMessage -BodyAsHtml:$true -Attachments $attachment -Priority:High
}

### old files cleanup function ###
### change to 30 days minimum when ready ###
function removeOldXMLs
{
    param ([string] $path)

    $files = Get-ChildItem -Path $path -Filter *_export.xml
    foreach($file in $files)
    {
        if($file.LastWriteTime -le ((Get-Date).AddDays(-1))) {
            Remove-Item -Path $file.FullName # -WhatIf
        }
    }
}

### optional coundown timer ###
function countDown 
{
    param($seconds)
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining 0 -Completed
}
#endregion (end Functions)

Clear-Host

# $xpath = "*[System[TimeCreated[timediff(@SystemTime) <= 60000]]]"
# $xpath = "*[System[EventID=4625 and TimeCreated[timediff(@SystemTime) <= 3600000]]]"
$xpath = "*[System[band(Keywords,4503599627370496) and TimeCreated[timediff(@SystemTime) <= 60000]]]"
# $xpath = "*[System[band(Keywords,4503599627370496) and TimeCreated[timediff(@SystemTime) <= 120000]]]" # audit failure(s) in the past 2 minutes
# $xpath = "*[System[band(Keywords,4503599627370496)]]" # audit failure(s) in the past 2 minutes
# $xpath = "*[System[TimeCreated[timediff(@SystemTime) <= 60000]]]" 

do {
    # Measure-Command {
    try {
       #  $events = $null # reset variable
        # $events = Get-WinEvent -LogName "ForwardedEvents" -FilterXPath $xpath -MaxEvents 25 -ErrorAction Stop
        $events = Get-WinEvent -LogName 'ForwardedEvents' -FilterXPath $xpath -ErrorAction Stop
    } # end try
    catch [Exception] {
        Write-Host -ForegroundColor Yellow "No events found matching the specified query.`r`n"
    }
    finally 
    {
        # clean up task here
    }

    # location where to save xml files
    $filePath = "C:\_PowerShell\Xml\"

    # delete old XMLs
    removeOldXMLs -path $filePath

    $fileName = $filePath + (Get-Date -format "MMddyyyy-hhmmss.fff_") + "export.xml"
    $maxNumEvents = 2 # total number of events before generating an ALERT!

    # exclude groups of "eventsIDs" with less than 5 logged events
    $groups = $events | group Id | sort count -Descending # | Select -ExpandProperty Count,Name,Group
    # $groups | Where-Object { $_.Count -ge 0 } | ft -auto

    foreach ($group in $groups)
    {
        # filter event groups with (n) total events
        switch ($($group.Count))
        {
            {($group.Count) -ge $maxNumEvents} {                
                $group # | ft                
            }
            default {
                Write-Host -ForegroundColor Yellow " Group $($group.Name) has less than $maxNumEvents events."
                # default script block
            }
        }
    }

    # "Total events: {0}" -f $events.Count
    # "Total groups: {0}`r`n" -f $groups.Length
    $groups | Export-Clixml -LiteralPath $fileName -Force

    <#
    foreach($group in $groups)
    {
        # skip if less than (n) failed events 
        if($group.Count -ge 5) 
        {
            # "> {0} - Total Events: {1}" -f $($group.Name), $($group.Count)
            # export events to XML file
            # $group | Export-Clixml -LiteralPath 'C:\_PowerShell\export.xml' -Force

            # loop through individual events
            foreach ($grpEvent in ($group.group))
            {    
                # $grpEvent | fl       
                # $grpEventXML = [xml] $grpEvent.ToXml()
                # $grpEventXML.Event.System
                # $grpEventXML.Event.EventData.Data | ft -auto
                # "`r"
            } # end foreach
            # "`r`n"
        } # end if
    } # end foreach
    #>

    # } # end Measure-Command

    # Remove-Variable events
    # Remove-Variable groups

    Write-Host "`r`nPaused for 15 seconds, please wait...`r`n"
    # Start-Sleep -Milliseconds 60000 # pause execution for 1 minute
    Start-Sleep -Milliseconds 14500 # pause execution for 1 minutes (less 500 ms)
    # Start-Sleep (60*1); write-host ("`a"*4)
    # countDown -seconds 60

} While($true) # end do|While
