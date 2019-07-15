﻿[CmdletBinding()]
param()

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
} 

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
        [string]$filepath = "L:\Logs\CsvReports\$yyyy-$mm-$dd $hour$min$sec EventID $($events.id | Select -Unique).csv"
    
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
              [string]$attachment
             )
        
        if ( $attachment )
        {
            $MailMessage = @{
                From = "securityAlert@landesa.org"
                To = "alexo@landesa.org"
                Subject = $subject
                Body = $body
                SmtpServer = "10.0.0.10"
                Attachments = $attachment
            }
        }
        else
        {
            $MailMessage = @{
                From = "securityAlert@landesa.org"
                To = "alexo@landesa.org"
                Subject = $subject
                Body = $body
                SmtpServer = "10.0.0.10"
            }
        }

        Send-MailMessage @MailMessage -Priority High -DeliveryNotificationOption onSuccess, onFailure -BodyAsHtml:$true 
    }
    #endregion sendMailAlert

#endregion functions

#################################
#  script execution starts here #
#################################

Clear-Host

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

# test here
# Get-WinEvent -FilterXml $xml

Write-Debug $xml    

$failedEventsGroup = Get-WinEvent -FilterXml $xml | group id | sort count -desc | Where-Object { 
    $_.count -gt 5 
} 

foreach ( $group in $failedEventsGroup )
{
    <# $($group.Name)
    

    
    # loop through each event in the group
    $group.group | ForEach-Object {
        $eventXml = [xml]$_.ToXml() 

        "`r`nSTART OF EVENT`r`n"
        $eventXml.Event.System
        $eventXml.Event.EventData.Data
        "`r`nEND OF EVENT`r`n" 
    }
     "`r`nEND OF GROUP`r`n" #>

     # start with EventID 4625: Failed login, account lockout
     if ( $($group.Name) -eq 4625 )
     {
        # create string builder object to store each events information
        # export string builder to CSV file

        # note $start in UTC
        Export-EventsToCsv -reportDate $start -events $($group.Group)
    
     }
     else
     {
        Write-Host "The event group EventID:$($group.Name) is not included in critical events to monitor. No action required." -ForegroundColor Yellow
     }
    
    "`r`nEND OF Group`r`n"


    # add to list of CSVs
}

# send alert to SystemsAdmin and attach all CSVs


Write-Debug $("Duration: {0} seconds" -f $(($timer.ElapsedMilliseconds)/1000))
Write-Debug "`r`n"

# set the current value of $end to next loop's $start
$start = $end
Write-Debug $("Next loop's SearchStartDateTime '{0}'" -f $start.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))
Write-Debug $("Next loop's SearchEndDateTime '{0}'" -f $start.AddMilliseconds($searchRangeMilliseconds).ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))
Write-Debug "`r`n"
Write-Debug $("Next loop will start at '{0}'" -f $start.AddMilliseconds($searchRangeMilliseconds).ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))

$timer.Stop()
$timerElapsedMilliseconds = $($timer.ElapsedMilliseconds)

# $sleepDurationMinutes = $(($searchRangeMilliseconds - $timerElapsedMilliseconds)/60000)

# Add-ProgressBar -sleepDurationMinutes $sleepDurationMinutes