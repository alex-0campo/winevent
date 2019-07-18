[CmdletBinding()]
param()

Clear-Host

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
} 

# determine the current script's location (file path)
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

Push-Location
Set-Location $ScriptDir

#region functions

    #region Export-EventsToCsv
    function Export-EventsToCsv
    {
        [CmdletBinding()]
        param($reportDate,$events)

        $reportDate = Get-Date
        $yyyy = $reportDate.Year
        $mm = "{0:00}" -f $reportDate.Month
        $dd = "{0:00}" -f $reportDate.Day
        $hour = "{0:00}" -f $reportDate.Hour
        $min = "{0:00}" -f $reportDate.Minute
        $sec = "{0:00}" -f $reportDate.Second

        # build CSV file (email report attachment)
        [string]$filepath = "L:\Logs\CsvReports\$yyyy-$mm-$dd-$hour$min$sec-EventID-$($events.id | Select -Unique).csv"
    
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
                To = "Landesa Global IT Team <Landesa_GBL_IT@rdiland.org>"
                Subject = $subject
                Body = $body
                SmtpServer = "10.0.0.10"
                Attachments = $attachments
            }
        }
        else
        {
            $MailMessage = @{
                From = "securityAlert@landesa.org"
                To = "Landesa Global IT Team <Landesa_GBL_IT@rdiland.org>"
                Subject = $subject
                Body = $body
                SmtpServer = "10.0.0.10"
            }
        }

        Send-MailMessage @MailMessage -Priority High -DeliveryNotificationOption onSuccess, onFailure -BodyAsHtml:$true 
    }
    #endregion sendMailAlert

#endregion functions

# simulate execution every 15 minutes

# do {

Get-Date

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
    
    3. Control (temporary) email sent until a solution to block inactive India staff
       from requesting Kerberos authentication ticket (TGT).
    4. Send event id count, plus any other valuable information on the email alert
    
    #> 

#endregion

# search for failed events in the past 15 minutes ran at 15-minutes (no overlap monitoring)
# set task duration for 15-minutes or other duration in minutes
$searchRangeMilliseconds = $(New-TimeSpan -Minutes 30).TotalMilliseconds

# Using local time did not work, UTC worked
$start = $((Get-Date).AddMilliseconds(-$searchRangeMilliseconds).ToUniversalTime()) 

### start timing execution ###
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
    </Query>
</QueryList>
"@

Write-Debug $xml   


<#################################################################### 
India inactive staff requesting Kerberos authentication ticket (TGT),
causing excessive audit failures. Temporarily control the frequency
of email alerts for event id 4768.
####################################################################>

$flag = "$ScriptDir\lastrun-4768.txt"

#############################
$skipDuration = 3 # (n) hours
$sleepDuration = 900 # (n) seconds
$eventCountThreshold = 5
#############################

# is this the script's first run?
# test if file does not exist, create a new file with current's date and time

if ( !$(Test-Path -Path $flag) )
{
    Write-Debug "`r`nThe file $flag not found..." # -ForegroundColor Red
    Write-Debug "Create missing file and set value to current date..." # -ForegroundColor Red
    New-Item -Path $flag -Value $(Get-Date) | Out-Null

    ###########################################################

    # run initial tasks here...
    Write-Debug "Get failed events here excluding events 5447, 4662, and 4674..." # -ForegroundColor Red
    
    # exclude event ids (5447, 4662, 4674) and filter with less than 11 events
    # too much events 4768, and 4776
    $failedEventsGroup = Get-WinEvent -FilterXml $xml | group id | Where-Object { 
        ($_.name -ne 5447 -and $_.name -ne 4662 -and $_.name -ne 4674) -and $_.count -gt $eventCountThreshold
    } | sort count -Descending

    ###########################################################    
}
else
{
    # $flag file exists; therefore not the initial run of the script
    # run more test to meet alternate tasks (Get failed event groups excluding
    # initial event ids plus event id 4768

    Write-Debug "The file $flag exists..." # -ForegroundColor Green  
    Write-Debug "Do more tests here..." # -ForegroundColor Green

    # is last run date time (n) minutes in the past?
    if ($(New-TimeSpan -Start $(Get-Date $(Get-Content -Path $flag)) -End $(Get-Date)).TotalHours -ge $skipDuration )
    {
        Write-Debug "  Last run date time is more than ($skipDuration) hours ago..." # -ForegroundColor Green
        Write-Debug "  Rerun initial tasks here..." # -ForegroundColor Green

        #########################################################
        # rerun initial tasks but exclude event id 4768 here... #

        $failedEventsGroup = Get-WinEvent -FilterXml $xml | group id | Where-Object { 
            ( $_.name -ne 5447 -and $_.name -ne 4662 -and $_.name -ne 4674 ) -and $_.count -gt $eventCountThreshold
        } | sort count -Descending

        #########################################################

        # update $flag last run time value
        Write-Debug "  Create a new file and set value to current date..." # -ForegroundColor Green
        New-Item -Path $flag -Value $(Get-Date) -Force | Out-Null

        # repeat more tasks here
        # Write-Host "  Repeat more tasks here..." -ForegroundColor Magenta
    }
    else 
    {
        # skip for tasks for the next (n) hours.
        # Write-Host "  $(Get-Date)`r`n  Do nothing in the next ($skipDuration) hours..." -ForegroundColor Yellow

        Write-Debug "  Rerun initial tasks but exclude event id 4768 here......" # -ForegroundColor Green

        #########################################################
        # rerun initial tasks but exclude event id 4768 here... #

        $failedEventsGroup = Get-WinEvent -FilterXml $xml | group id | Where-Object { 
            ( $_.name -ne 5447 -and $_.name -ne 4662 -and $_.name -ne 4674 -and $_.name -ne 4768 ) -and $_.count -gt $eventCountThreshold
        } | sort count -Descending

        #########################################################
    }
}

######################## old code no changes from here ########################

$failedEventsGroup | ft -auto

$attachments = @()

foreach ( $group in $failedEventsGroup )
{
    # export each group of events to CSV and add the filename to attachments (array)
    $attachments += $(Export-EventsToCsv -reportDate $start -events $($group.Group))
     
}

# send alert to SystemsAdmin and attach all CSVs 
# skip sendMailAlert if there is no $failedEventsGroup to report
if ( $($attachments.Length) -ge 1)
{
    sendMailAlert -subject "AUDIT FAILURE" -body "There are audit failures reported in the past 15-minutes.<br />See attachments for details." -attachments $attachments
}
else
{
    Write-Debug "`r`n  No `$failedEventsGroup with events reported exceeding ($eventCountThreshold)."
    # sendMailAlert -subject "AUDIT FAILURE" -body "No audit failure events match report criteria in the past 15-minutes."
}

Write-Debug $("Duration: {0} seconds" -f $(($timer.ElapsedMilliseconds)/1000))
Write-Debug "`r`n"

# set the current value of $end to next loop's $start
# $start = $end
Write-Debug $("Next loop's SearchStartDateTime '{0}'" -f $start.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))
Write-Debug $("Next loop's SearchEndDateTime '{0}'" -f $start.AddMilliseconds($searchRangeMilliseconds).ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))
Write-Debug "`r`n"
Write-Debug $("Next loop will start at '{0}'" -f $start.AddMilliseconds($searchRangeMilliseconds).ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))

$timer.Stop()
# $timerElapsedMilliseconds = $($timer.ElapsedMilliseconds)

"`r`n  Duration: {0} (ElapsedSeconds)" -f $(($timer.ElapsedMilliseconds)/1000)

# set $failedEventsGroup variable to $null
$failedEventsGroup = $null
