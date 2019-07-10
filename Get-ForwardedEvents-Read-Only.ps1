[CmdletBinding()]
param()

# $DebugPreference = "Continue" # debug
$DebugPreference = "SilentlyContinue" # live

Clear-Host

#region functions
function sendMailAlert {
    [CmdletBinding()]
    param([string] $subject,
          [string] $body
         )

    $MailMessage = @{
        From = "securityAlert@landesa.org"
        To = "Landesa_GBL_IT@landesa.org"
        Subject = $subject
        Body = $body
        SmtpServer = "10.0.0.10"
    }

    Send-MailMessage @MailMessage -Priority High -DeliveryNotificationOption onSuccess, onFailure -BodyAsHtml:$true 
}
#endregion

#region filterxpath attributes ###  
    # Time created filter (milliseconds)
    $ms = 60*60*1000 # events in the past 3-hour(s)
    # $ms = 15*60*1000 # events in the past 15-minutes

    # search for EventID 4625 in the past (n) milliseconds.
    # $xpath = "*[System[(EventID=4625) and TimeCreated[timediff(@SystemTime) <= $ms]]]"
    $xpath = "*[System[TimeCreated[timediff(@SystemTime) <= $ms]]]"
#endregion

#region variables
$events = $null
$message = $null
$msg = $null
$groups = $null
$group = $null
$msgBody = $null
$alertSendTime = $null
$lastSentDateTime = $null
$lastEmailMessage = $null
$prevMessage = $null

# the minumum audit failures to trigger alert
$groupCount = 10

$alertNum = 1
$maxAlertSent = 3 # max repeated alerts if no additional failed logins reported

# execution timer
$getEventsInterval = 600 # interval number of seconds (300) between get-winevent pass
$skipAlertMinutes = 60 # (30) minutes before resending email alert
#endregion

#region main script
do {
    # "`r`n`r`n`r`n`r`n`r`n`r`n`r`n"
    "`r`n"
    Write-Debug "$(Get-Date)"

    try    
    {
        #region Get-WinEvent
        $events = Get-WinEvent -LogName 'ForwardedEvents' -FilterXPath $xpath -ErrorAction Stop
        # $events = Get-WinEvent -path 'L:\Logs\ForwardedEvents\Archive-ForwardedEvents-2019-01-14-18-33-55-830.evtx' -FilterXPath $xpath -ErrorAction Stop

        ### get high failed logins group by computers
        $groups = $events | group MachineName | sort count -Desc

        ### email alert message        
        $message = foreach ($group in $groups) 
        {
            # iterate each group with (6 or more failed logins)
            if($group.Count -ge $groupCount)
            {
                # if group meets condition add it to message
                "{0,5} failed logins on computer {1}<br />" -f $group.Count, $group.Name
            } # end if
        } # end foreach $group
        #endregion
        
        ### set email subject
        $subj = "ALERT: High Failed Logins (EventID 4625)" # "TEST ALERT: High Failed Logins" 

        
        ### first run ($message -eq$null and $lastSentDateTime -eq $null)
        # if (($message -eq $null) -or ($lastSentDateTime -eq $null))
        if ($lastSentDateTime -eq $null)
        {
            # no previous alert message (initial run) 
            ### if($lastSentDateTime -eq $null)
            ### {
                # no previous alert sent                
                Write-Debug "   Send first alert: #$alertNum."

                $lastSentDateTime = Get-Date

                Write-Debug "Message:`r`n$($message.Replace("`<br />","`n"))"

                $msg = $message | Out-String
                sendMailAlert -subject $subj -body $msg
                Write-Debug "   Next alert after $($lastSentDateTime.AddMinutes($skipAlertMinutes))"

                # save the value of $msg to $prevMessage
                $prevMessage = $msg
                $alertNum++
                # Write-Debug "`$alertNum: $alertNum"
            ### } # end if($lastSentDateTime -eq $null)

            # $lastSentDateTime not equal to $null
        }

        # next pass
        else
        {
            # is there a change in the reported events or have not resend the alert in the past (x) minutes
            if (  (($lastSentDateTime -le $((Get-Date).AddMinutes(-$skipAlertMinutes))) -and ($alertNum -le $maxAlertSent)) -or  ($($message | Out-String) -ne $prevMessage) )
            {

                # reset the value of $alertNum to 1 on new alerts
                if($($message | Out-String) -ne $prevMessage)
                {
                    $alertNum = 1
                }
                
                Write-Debug "   Resend previous alert or new alerts: #$alertNum."
                
                $lastSentDateTime = Get-Date

                Write-Debug "Message:`r`n$($message.Replace("`<br />","`n"))"
                
                $msg = $message | Out-String
                sendMailAlert -subject $subj -body $msg
                Write-Debug "   Next alert after $($lastSentDateTime.AddMinutes($skipAlertMinutes))"

                # save the value of $msg to $prevMessage
                $prevMessage = $msg

                # reset the value of $alertNum to 1 on new alerts
                $alertNum++
                # Write-Debug "`$alertNum: $alertNum"
            }
            else
            {
                # skip sending same reported message
                Write-Debug "   Last alert sent at $lastSentDateTime."
                Write-Debug "   Skip alert for $skipAlertMinutes minutes, same events."
                Write-Debug "   Re-send next alert after $($lastSentDateTime.AddMinutes($skipAlertMinutes))"                
            }
        }


    } # end try

    catch [Exception]
    {
    if ($_.Exception -match "No events were found that match the specified selection criteria") 
        {
            Write-Host "No events were found that match the specified selection criteria." -ForegroundColor Yellow
        }
    }

    finally
    {
        # Clean up tasks
    } # end finally

    Write-Debug "End of failed logins search, re-starting in $getEventsInterval seconds, please wait.`r`n`r`n"
    Write-Debug "Next run at $((Get-Date).AddSeconds($getEventsInterval))"
    Start-Sleep -Seconds $getEventsInterval 

    #region display display bar while waiting for the next run
    <#
    for($i=1;$i -le $getEventsInterval; $i++)
    {
        Write-Progress -Activity "End of failed logins search, re-starting in $getEventsInterval seconds, please wait." `
            -Status "Seconds remaining: $($getEventsInterval-$i)" `
            -PercentComplete ($i/$getEventsInterval*100)
        Start-Sleep -Seconds 1  # pause execution for 5-minutes (300) seconds
    } #>
    #endregion 
} while ($true)
#endregion

#remove variables
Remove-Variable message
Remove-Variable msgBody
Remove-Variable groupCount
Remove-Variable alertNum
Remove-Variable lastSentDateTime
Remove-Variable lastEmailMessage
Remove-Variable prevMessage

#region future additions
### group events by machineName ###

### switch: machineName events count ###

    ### switch emailAlertLastSent ###
#endregion
