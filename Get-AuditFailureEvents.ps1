[CmdletBinding()]
param()

[System.GC]::Collect()

Push-Location
Set-Location 'C:\_PowerShell\_git\_dev'

#region sendMailAlert
function sendMailAlert {
    [CmdletBinding()]
    param([string]$subject,
          [string]$body,
          [string]$attachment
         )

    $MailMessage = @{
        From = "securityAlert@landesa.org"
        To = "alexo@landesa.org"
        Subject = $subject
        Body = $body
        SmtpServer = "10.0.0.10"
        Attachments = $attachment
    }

    Send-MailMessage @MailMessage -Priority High -DeliveryNotificationOption onSuccess, onFailure -BodyAsHtml:$true 
}
#endregion sendMailAlert


#region Get-FailedEvents
function Get-FailedEvents {
    [CmdletBinding()]
    param($ms,$searchRangeMinutes)

    <#
    Task: Search for failed events in the past 15 minutes)

    Workflow:

    1. If the last archived ForwardedEvents log is less than 15 minutes old 
       then Get-FailedEvents from the last archived log then add to failed events 
       from the active ForwardedEvents log.

    2. If the last archived ForwardedEvents log is older than 15 minutes then 
       Get-FailedEvents from the active ForwardedEvents log only.

    
    Suppress Events: 5447, ?


               $xml = "<QueryList>
          <Query Id='0' Path='ForwardedEvents'>
            <Select Path='ForwardedEvents'>*[System[EventID=4625 and band(Keywords,4503599627370496) 
            and 
            TimeCreated[timediff(@SystemTime) &lt;=" + $ms + "]]]</Select>
            <Suppress Path='ForwardedEvents'>*[System[(EventID=5447)]]</Suppress>
          </Query>
        </QueryList>"
    #>

    # duration in milliseconds to search for events
    # $ms = $(New-TimeSpan -Minutes 60).TotalMilliseconds # events in the past 1-hour(s)

    ################################
    ######   set xml filter   ######
    ################################
    

    $xml = @"
    <QueryList>
      <Query Id='0' Path='ForwardedEvents'>
        <Select Path='ForwardedEvents'>
          *[
            System[
              band(Keywords,4503599627370496) and TimeCreated[timediff(@SystemTime) &lt;=$ms]
            ]
          ]
        </Select>
      </Query>
    </QueryList>
"@

    ####################################################
    ### is the ForwardedEvents log recently archived ###
    ####################################################

    $searchPath = 'L:\Logs\ForwardedEvents'
    $lastArchivedLog = Get-ChildItem -Filter "Archive-ForwardedEvents*.evtx" -Path $searchPath | Select -Last 1
    $lastArchivedLogLastWriteTime = $lastArchivedLog.LastWriteTime

    # $xpath = "*[System[TimeCreated[timediff(@SystemTime) <= $ms] and band(Keywords,4503599627370496)]]"
    # $xpath = "*[System[EventID=4625] and EventData[Data[@Name='WorkstationName']='RDI-US-LP1537']]"
    $xpath = "*[System[TimeCreated[timediff(@SystemTime) <= $ms] and band(Keywords,4503599627370496)]]"

    if ( $lastArchivedLogLastWriteTime -ge $((Get-Date).AddMinutes(-$searchRangeMinutes)) )
    {
        Write-Host "Search for failed events from the last archived log." -ForegroundColor Cyan    
        "Search Range (sec): {0:n2}" -f $($searchRangeMilliSeconds/1000)

        $events = Get-WinEvent -Path $($lastArchivedLog.FullName) -FilterXPath $xpath -ea SilentlyContinue
        "ArchivedLogEvents: {0}`r`n" -f $events.Count
    }
    else
    {
        # get failed events from the ForwardedEvents log instead

        Write-Host "Search for failed events from ForwardedEvents log." -ForegroundColor Yellow

        # add previously archived log event to ForwardeEvents events
        $events = Get-WinEvent -Path $lastArchivedLog.FullName -FilterXPath $xpath -ea SilentlyContinue
        "ArchivedLogEvents: {0}`r`n" -f $events.Count

        $events += Get-WinEvent -FilterXml $xml -ea SilentlyContinue 
        "Archived + ForwardedEvents: {0}`r`n" -f $events.Count
    }

    return $events

} # end function
#endregion Get-FailedEvents 

#region Group-FailedEvents
function Group-FailedEvents
{
    [CmdletBinding()]
    param()


}
#endregion Group-FailedEvents

#region Export-EventsToCsv function
function Export-EvensToCsv
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

    # create a container for events email report$
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

        # skip last 3-columns of xml data
        $elements = $(($eventXML.Event.EventData.Data.Count) - 5)



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

    # $sb.ToString()

    
    # NOT WORKING
    $str = $($sb.ToString())
    # $str
    

    New-Item -Path $filepath -Value $str -Force | Out-Null

    # return $filepath

    } # end function
#endregion Export-EventsToCsv function



################################
# script execution starts here #
################################

$searchRangeMinutes = 30 # $(0.25*60)
$searchRangeMilliSeconds = $(New-TimeSpan -Minutes $searchRangeMinutes).TotalMilliseconds
"Search Range (min): {0:n0}" -f $($searchRangeMilliSeconds / 60000)

Clear-Host

do
{
    $reportDate = Get-Date
    $reportDate

    $timer = New-Object System.Diagnostics.Stopwatch
    $timer.Start()

    $failedEventGroups = Get-FailedEvents -ms $searchRangeMilliSeconds -searchRangeMinutes $searchRangeMinutes | group id | sort count -Descending

    $timer.Stop()
    "Elapsed Seconds: {0:n2}" -f $(($timer.ElapsedMilliseconds) / 1000)

    $failedEventGroups | ft count,name

    $failedEventGroups | ForEach-Object {

        if ( $_.group.Count -ge 10 )
        {
            Export-EvensToCsv -reportDate $reportDate -events $_.Group
            # $filepath = Export-EvensToCsv -reportDate $reportDate -events $_.Group
        }

        # $filepath
    }

    $sleepSeconds = $((New-TimeSpan -Minutes $($searchRangeMinutes - 29)).TotalSeconds) - 14.5

    Write-Host "Pause for $($sleepSeconds/60) minutes`r`n" -ForegroundColor Red
    Start-Sleep -Seconds $sleepSeconds # pause for 1-minute
} while( $true )

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#


<# pre-convert to function

$reportDate = Get-Date
$yyyy = $reportDate.Year
$mm = "{0:00}" -f $reportDate.Month
$dd = "{0:00}" -f $reportDate.Day
$hour = "{0:00}" -f $reportDate.Hour
$min = "{0:00}" -f $reportDate.Minute
$sec = "{0:00}" -f $reportDate.Second

# create a container for events email report$
[string]$filepath = "L:\Logs\CsvReports\$yyyy-$mm-$dd $hour$min$sec EventID$($events.id | Select -Unique).csv"
    

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
    $elements = $eventXML.Event.EventData.Data.Count



    # if $i = 0 (first event), then do these tasks
    if ( $i -eq 0 )
    {
        # add TimeCreated and EventID header labels
        $sb.Append("TimeCreated,EventID,") | Out-Null

        # add headers to string builder $sb
        for ( $x = 0; $x -lt $elements; $x++ )
        {
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


# $sb.ToString()


New-Item -Path $filepath -Value $($sb.ToString()) -Force | Out-Null

# check number of events; skip sending report if no events matched criteria
if ( $eventsCount -ge 1 )
{
    sendMailAlert -subject 'Failed logins' -body 'Excessive failed logins reported. See attached file for details?' -attachment $filepath
}
#>

Pop-Location




<#
Name                      #text          
----                      -----          
SubjectUserSid            S-1-0-0        
SubjectUserName           -              
SubjectDomainName         -              
SubjectLogonId            0x0            
TargetUserSid             S-1-0-0        
TargetUserName            leonardr5      
TargetDomainName          RDISEA         
Status                    0xc000006d     
FailureReason             %%2313         
SubStatus                 0xc0000064     
LogonType                 3              
LogonProcessName          NtLmSsp        
AuthenticationPackageName NTLM           
WorkstationName           RDI-US-LP1400SH
TransmittedServices       -              
LmPackageName             -              
KeyLength                 0              
ProcessId                 0x0            
ProcessName               -              
IpAddress                 10.0.0.1       
IpPort                    49170  
#>


