[CmdletBinding()]
param()

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
} 

function Get-AuditFailureEvents
{
    [CmdletBinding()]
    param($filter)

    Get-WinEvent -FilterXml $filter
}

function Group-AuditFailureEvents
{
    [CmdletBinding()]
    param($events)

    $events | group id | Where-Object { $_.count -ge 9 }
}

function Add-ProgressBar
{
    [CmdletBinding()]
    param($duration)

    for($i=1;$i -le $duration; $i++)
    {
        Write-Progress -Activity "Resume task in $([Math]::Round($($duration/6))) minutes, please wait..." `
        -Status "$([Math]::Round($(($i/$($duration)*100))))%" `
        -PercentComplete ($i/$($duration)*100)
        Start-Sleep -MilliSeconds 10000 # (60/6)*1000 = 10000 ms
    }
}

#################################
#  script execution starts here #
#################################

Clear-Host

# set search range "FROM <datetime>"
$start = $((Get-Date).AddMinutes(-$searchRangeMinutes)).ToUniversalTime()

# set task duration in minutes
$searchRangeMinutes = 15
$searchRangeMilliseconds = $(New-TimeSpan -Minutes $searchRangeMinutes).TotalMilliseconds

do {
    
    ### start timing execution ###
    $timer = New-Object System.Diagnostics.Stopwatch
    $timer.Start()
    
    # current value of $start
    Write-Debug $("{0} (Current loop's StartSearchDateTime)" -f $start.ToLocalTime())

    # set search range "TO datetime" 
    $end = $start.AddMilliseconds($searchRangeMilliseconds)

    # current value of $end
    Write-Debug $("{0} (Current loop's EndSearchDateTime)" -f $end.ToLocalTime())
    
    $strStart = "'" + $start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "'"
    $strEnd = "'" + $end.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "'"

    $xml = @"
    <QueryList>
      <Query Id='0' Path='ForwardedEvents'>
        <Select Path='ForwardedEvents'>
          *[
            System[band(Keywords,4503599627370496) and (TimeCreated[@SystemTime&gt;=$strStart and @SystemTime&lt;=$strEnd])]
          ]
        </Select>
      </Query>
    </QueryList>
"@

    

    
    # Get-WinEvent -FilterXml $xml | group id | Where-Object { $_.count -ge 9 } | sort count -desc | ft -auto

    $failedEvents = Get-AuditFailureEvents -filter $xml # | group id | Where-Object { $_.count -ge 9 } | sort count -desc | ft -auto
    $(Group-AuditFailureEvents -events $failedEvents) | sort count -desc | ft -auto

    $timer.Stop()
    $timerElapsedMilliseconds = $($timer.ElapsedMilliseconds)

    Write-Debug $("{0} Utc (start)" -f $strStart)
    #"'{0}' (now)" -f $(((Get-Date).ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ss.fffZ"))
    Write-Debug $("{0} Utc (end)" -f $strEnd)

    "`r`nElapsedMilliseconds: {0}`r`n" -f $timer.ElapsedMilliseconds

    # set the current value of $end to next loop's $start
    $start = $end
    Write-Debug $("{0} (Next loop's SearchStartDateTime)" -f $start.ToLocalTime())
    Write-Debug $("{0} (Next loop's SearchEndDateTime)" -f $start.AddMilliseconds($searchRangeMilliseconds).ToLocalTime())

    Write-Debug $("{0} (Next loop's StartAt)" -f $start.AddMilliseconds($searchRangeMilliseconds).ToLocalTime())

    $sleepMilliseconds = $($searchRangeMilliseconds - $timerElapsedMilliseconds)
    # Write-Host $("Paused for {0:N2} minutes, please wait...`r`n" -f $($sleepMilliseconds/60000)) -ForegroundColor Yellow


    ### Start progress bar
    $sleepTimer = [math]::Round($($sleepMilliseconds / 10000)) # convert to 1/6 minutes

    # If ($PSBoundParameters['Debug']) {
        Add-ProgressBar -duration $sleepTimer
    # }
    
} while ( $true )

