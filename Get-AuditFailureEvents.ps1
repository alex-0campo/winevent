[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$suppressEvents = @(5447, 4662, 4674, 4768, 4776)
)

# Clear-Host

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
} 

# determine the current script's location (file path)
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

#region functions

    #region Export-EventsToCsv
    function Export-EventsToCsv
    {
        [CmdletBinding()]
        param($reportDate,$events,$machineName)

        $reportDate = Get-Date
        $yyyy = $reportDate.Year
        $mm = "{0:00}" -f $reportDate.Month
        $dd = "{0:00}" -f $reportDate.Day
        $hour = "{0:00}" -f $reportDate.Hour
        $min = "{0:00}" -f $reportDate.Minute
        $sec = "{0:00}" -f $reportDate.Second

        # build CSV file (email report attachment)
        [string]$filepath = "L:\Logs\CsvReports\$yyyy-$mm-$dd-$hour$min$sec-$machineName-EventID-$($events.id | Select -Unique).csv"
    
        # string builder
        $sb = New-Object -TypeName System.Text.StringBuilder

        # number of events found
        $eventsCount = $events.Count

        for ( $i = 0; $i -lt $eventsCount; $i++ )
        {
            $timeCreated = $events[$i].TimeCreated #.ToString()
            $eventID = $events[$i].Id #.ToString()

            # do for all events
            $eventXML = [xml]$events[$i].ToXml()

            $elements = $(($eventXML.Event.EventData.Data.Count))

            # skipt last 5-columns on event id 4662
            # $elements = $(($eventXML.Event.EventData.Data.Count) - 5)

            # if $i = 0 (first event), then do these tasks
            if ( $i -eq 0 )
            {
                # add TimeCreated and EventID header labels
                $sb.Append("TimeCreated,EventID,") | Out-Null

                # add headers to string builder $sb
                for ( $x = 0; $x -lt $elements; $x++ )
                {
                    # iterate each element of xml data
                    if ( $x -lt $($elements - 1) )
                    {
                        $sb.Append( ($eventXML.Event.EventData.Data[$x].Name) + "," ) | Out-Null
                    }
                    else
                    {
                        $sb.AppendLine( ($eventXML.Event.EventData.Data[$x].Name) ) | Out-Null
                    }
                } # end add headers to string builder $sb

                #$sb.AppendLine("`r`n") | Out-Null

                # add TimeCreated and EventID values for the first event        
                $sb.Append("$timeCreated,$eventID,") | Out-Null

                # iterate each element of xml data
                # add first event values to string builder $sb
                for ( $y = 0; $y -lt $elements; $y++ )
                {
                    if ( $y -lt $($elements - 1) )
                    {
                        $sb.Append( ($eventXML.Event.EventData.Data[$y].'#text') + "," ) | Out-Null
                    }
                    else
                    {
                        $sb.AppendLine( ($eventXML.Event.EventData.Data[$y].'#text') ) | Out-Null
                    }
                } # end add first events values to string builder $sb
            }
            else
            {

                # add TimeCreated and EventID values for the first event
                $sb.Append("$timeCreated,$eventID,") | Out-Null

                # iterate each element of xml data
                # add remaining events' values to string builder $sb
                for ( $z = 0; $z -lt $elements; $z++ )
                {
                    if ( $z -lt $($elements - 1) )
                    {
                        $sb.Append( ($eventXML.Event.EventData.Data[$z].'#text') + "," ) | Out-Null
                    }
                    else
                    {
                        $sb.AppendLine( ($eventXML.Event.EventData.Data[$z].'#text') ) | Out-Null
                    }
                } # end add remaining events values to string builder $sb
            }

            # $sb.ToString()

        } # end for $i $events loop

        $str = $($sb.ToString())
        Write-Debug $str
    
        New-Item -Path $filepath -Value $str -Force | Out-Null

        # return the CSV's file path
        return $filepath

    } # end of Export-EvensToCsv function
    #endregion Export-EventsToCsv

    #region Add-ProgressBar
    function Add-ProgressBar
    {
        [CmdletBinding()]
        param($sleepDurationMinutes)

        $sleepDurationSeconds = [Math]::Round($sleepDurationMinutes*60)

        for($i=1;$i -le $sleepDurationSeconds; $i++)
        {
            Write-Progress -Activity "Resume task in $sleepDurationSeconds seconds, please wait..." `
            -Status "$($sleepDurationSeconds-$i) seconds remaining" `
            -PercentComplete ($i/$($sleepDurationSeconds)*100)
            Start-Sleep -Seconds 1 # (60/6)*1000 = 10000 ms
        }
    } 
    #endregion Add-ProgressBar

    #region sendMailAlert
    function sendMailAlert 
    {
        [CmdletBinding()]
        param([string]$subject,
              [string]$body,
              [string[]]$attachments
             )
        
        if ( $attachments )
        {
            $MailMessage = @{
                From = "securityAlert@landesa.org"
                To = "Alex Ocampo <alexo@landesa.org>" # "Landesa Global IT Team <Landesa_GBL_IT@rdiland.org>"
                Subject = "TEST: " + $subject
                Body = $body
                SmtpServer = "10.0.0.10"
                Attachments = $attachments
            }
        }
        else
        {
            $MailMessage = @{
                From = "securityAlert@landesa.org"
                To = "Alex Ocampo <alexo@landesa.org>" # "Landesa Global IT Team <Landesa_GBL_IT@rdiland.org>"
                Subject = $subject
                Body = $body
                SmtpServer = "10.0.0.10"
            }
        }

        Send-MailMessage @MailMessage -Priority High -DeliveryNotificationOption onSuccess, onFailure -BodyAsHtml:$true 
    }
    #endregion sendMailAlert

#endregion functions

Write-Host "`r`n`r`n`r`n"
# simulate execution every 15 minutes
# do {
    
#################################
#  script execution starts here #
#################################

#region task details
    <# Task: Search for failed events in the past 15 minutes)

    Workflow:

    1. If the last archived ForwardedEvents log is less than 15 minutes old 
       then Get-FailedEvents from the last archived log then add to failed events 
       from the active ForwardedEvents log.

    2. If the last archived ForwardedEvents log is older than 15 minutes then 
       Get-FailedEvents from the active ForwardedEvents log only.

    
    Suppress Events: 4662, 4674, 5447, ?


    $xml = "<QueryList>
        <Query Id='0' Path='ForwardedEvents'>
        <Select Path='ForwardedEvents'>*[System[EventID=4625 and band(Keywords,4503599627370496) 
        and 
        TimeCreated[timediff(@SystemTime) &lt;=" + $ms + "]]]</Select>
        <Suppress Path='ForwardedEvents'>*[System[(EventID=5447)]]</Suppress>
        </Query>
    </QueryList>" 
    
    3. Skipped monitoring event 4768 every 6- hours. - Control (temporary) email sent
       until a solution to block inactive India staff from requesting Kerberos 
       authentication ticket (TGT).
    4. Send event id count, plus any other valuable information on the email alert

    5. Group by EventID then (inspect) and group by computer or user object.
    
    #> 

    #endregion task details

#region hide
# search for failed events in the past 16 minutes scheduled to run every 15 minutes 
# overlapping by 1-minute from last run to take run time into consideration.

#########################################################################################
#####         set task duration for 16-minutes or other duration in minutes         #####
#########################################################################################
$searchRangeMinutes = 16
$searchRangeMilliseconds = $(New-TimeSpan -Minutes $searchRangeMinutes).TotalMilliseconds

# $flag = "$ScriptDir\lastrun-4768.txt"
$skipDuration = 6 # skip monitoring event id 4768 for 6-hours
$sleepDuration = 900 # 900 # (n) seconds
$eventGroupCountThreshold = 10
$eventCountThreshold = 5 # bump 5 to 10
#########################################################################################

# Using local time did not work, UTC worked
$start = $((Get-Date).AddMilliseconds(-$searchRangeMilliseconds).ToUniversalTime()) 

### start timing execution ###
"Report Range: {0} -to- {1}" -f $start.ToLocalTime(), $start.AddMinutes(16).ToLocalTime()
"`r`n"
$timer = New-Object System.Diagnostics.Stopwatch
$timer.Start()
    
# current value of $start
Write-Debug $("StartSearchDateTime '{0}'" -f $start.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")) #.ToLocalTime())

# set search range "TO datetime" 
$end = $start.AddMilliseconds($searchRangeMilliseconds)

# current value of $end
Write-Debug $("EndSearchDateTime '{0}'" -f $end.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")) #.ToLocalTime())
Write-Debug "`r`n"

# format start and end datetime for the xml filter    
$strStart = "'" + $start.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "'"
$strEnd = "'" + $end.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "'"


# set xml filter

# how to dynamically build xml filter excluding events based on passed parameter(s)
# exclude event ids (5447, 4662, 4674, 4768, and 4776) and filter with less than 11 events

$sb = New-Object -TypeName System.Text.StringBuilder

for ($i=0; $i -le $(($suppressEvents.Length)-1); $i++)
{
    if ( $i -lt $(($suppressEvents.Length)-1) )
    {
        $sb.Append("EventID=" + $suppressEvents[$i] + " or ") | Out-Null
    }
    else
    {
        $sb.Append("EventID=" + $suppressEvents[$i]) | Out-Null
    }    
}

$xml = @"
<QueryList>
    <Query Id='0' Path='ForwardedEvents'>
    <Select Path='ForwardedEvents'>
        *[
            System[band(Keywords,4503599627370496) 
            and
            (TimeCreated[@SystemTime&gt;=$strStart and @SystemTime&lt;=$strEnd])]
        ]
    </Select>
    <Suppress Path='ForwardedEvents'>
	    *[System[($($sb -join ''))]]
	</Suppress>
    </Query>
</QueryList>
"@

#endregion hide 

Write-Debug $xml  

<#################################################################### 
India inactive staff requesting Kerberos authentication ticket (TGT),
causing excessive audit failures. Temporarily control the frequency
of email alerts for event id 4768.
####################################################################>

$failedEvents = Get-WinEvent -FilterXml $xml

# ignore audit failure (group by id) with 9 or less events
$failedEventsGroup = $failedEvents | group id | Where-Object { $_.Count -ge $eventGroupCountThreshold } | sort count -desc

# $timer = New-Object -TypeName System.Diagnostics.Stopwatch
# $timer.Start()

# container (array) CSVs created (attachment => email alert)
$attachments = @()

# string builder for the alert message
$sbBody = New-Object -TypeName System.Text.StringBuilder

$failedEventsGroup | ForEach-Object {
    # "Audit Failure Events:  {0}" -f $_.Name
    # "`r`n"
    $_.Group | group MachineName | ForEach-Object {
        # ignore MachineName if reported audit failure with 9 or less events
        if ($_.Count -ge $eventCountThreshold)
        {
            # export to csv...
            Write-Debug $("Reported by MachineName: {0}" -f $($_.Name).Split(".")[0])
            $events = $_.Group
            $attachments += Export-EventsToCsv -reportDate $start -events $events -machineName $(($_.Name).Split(".")[0])

            # append failed event info into string builder
            $sbBody.Append($(($_.Name).Split(".")[0]) + " reported " + $($_.Count) + " event IDs " + $($_.Group.Id | Select -First 1 -Unique) + " audit failures in the past 15-minutes.<br />") | Out-Null
            # $sbBody.ToString()

            Write-Debug $($events | Out-String)
            Write-Debug $("Event count: {0}`r`n" -f $_.Group.count)
        } # end MachineName group with 10 or more events

            


    } # end group by MachineName    
} # end $failedEventsGroup

# send alert to SystemsAdmin and attach all CSVs 
# skip sendMailAlert if there is no $failedEventsGroup to report
if ( $($attachments.Length) -ge 1)
{
    sendMailAlert -subject "AUDIT FAILURE" -body "$($sbBody.ToString())<br />See attached CSV(s) for more information." -attachments $attachments
}
else
{
    Write-Host "`r`n  There are no `$failedEventsGroup with more than ($eventCountThreshold) audit failures."
    Write-Host "`r  Skip sending alert..."
    # sendMailAlert -subject "AUDIT FAILURE" -body "No audit failure events match report criteria in the past 15-minutes."
}

"`r`n"
# $timer.Stop()
# $timer.Elapsed



$failedEventsGroup = $null 

# Write-Host "`r`n  End of do/while loop; pause for $($sleepDuration/60) minute(s),`r`n  please wait...`r`n" -ForegroundColor Yellow
# Add-ProgressBar -sleepDurationMinutes $($sleepDuration/60)
# Start-Sleep -Seconds $sleepDuration

# } while ( $true ) # end of do loop
