[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$suppressEvents = @(5447, 4662, 4674, 4768, 4776)
)

#region Test new feature

# search for failed events in the past 16 minutes scheduled to run every 15 minutes 
# overlapping by 1-minute from last run to take run time into consideration.

# set task duration for 16-minutes or other duration in minutes
$searchRangeMilliseconds = $(New-TimeSpan -Minutes 60).TotalMilliseconds

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

# how to dynamically build xml filter excluding events based on passed parameter(s)
# exclude event ids (5447, 4662, 4674, 4768, and 4776) and filter with less than 11 events

$sb = New-Object -TypeName System.Text.StringBuilder

# build suppress string
<# $suppressEvents | ForEach-Object {
    # $_
    $sb.Append("EventID=" + $_ + " -and ") | Out-Null
} #>

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

# $($sb -join "")

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

Write-Host $xml  

#endregion Test new feature

Clear-Host

$flag = "$ScriptDir\lastrun-4768.txt"
$skipDuration = 6 # skip monitoring event id 4768 for 6-hours
$sleepDuration = 900 # (n) seconds
$eventGroupCountThreshold = 10
$eventCountThreshold = 5 # bump 5 to 10

$failedEvents = Get-WinEvent -FilterXml $xml
$failedEventsGroup = $failedEvents | group id | Where-Object { $_.Count -ge $eventGroupCountThreshold } | sort count -desc

$timer = New-Object -TypeName System.Diagnostics.Stopwatch
$timer.Start()

Clear-Host
$failedEventsGroup | ForEach-Object {
    "Inspect  {0}" -f $_.Name
    "`r`n"
    $_.Group | group MachineName | ForEach-Object {
        if ($_.Count -ge $eventCountThreshold)
        {
            "MachineName: {0}" -f $_.Name
            $_.Group | ft -auto
            "`r`n"
        }
    }
    "End of failedEventsGroup`r`n"
}

"`r`n"
$timer.Stop()
$timer.Elapsed
