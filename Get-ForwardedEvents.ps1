[CmdletBinding()]
param()

Clear-Host
[System.GC]::Collect()

Push-Location
Set-Location 'C:\_Powershell\_git\_dev'



#region sendMailAlert functions
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
#endregion sendMailAlert functions

################################
# script execution starts here #
################################

# duration in milliseconds to search for events
$ms = 12*60*60*1000 # events in the past 1-hour(s)

$xml = "<QueryList>
  <Query Id='0' Path='ForwardedEvents'>
    <Select Path='ForwardedEvents'>*[System[EventID=4625 and band(Keywords,4503599627370496) 
    and 
    TimeCreated[timediff(@SystemTime) &lt;=" + $ms + "]]]</Select>
    <Suppress Path='ForwardedEvents'>*[System[(EventID=5447)]]</Suppress>
  </Query>
</QueryList>"

$xpath = "*[System[TimeCreated[timediff(@SystemTime) <= $ms] and band(Keywords,4503599627370496)]]"



# $events = Get-WinEvent -FilterXml $xml -ea SilentlyContinue | Where-Object { $_.Id -eq 4625 }
# $events = Get-WinEvent -LogName 'ForwardedEvents' -FilterXPath $xpath -ea SilentlyContinue | Where-Object { $_.Id -eq 4625 }


$searchPath = 'L:\Logs\ForwardedEvents'
$lastArchivedLog = Get-ChildItem -Filter "Archive-ForwardedEvents*.evtx" -Path $searchPath | Select -Last 1

$events = Get-WinEvent -Path $lastArchivedLog.FullName -FilterXPath $xpath | Where-Object { $_.Id -eq 4625 }
"ForwardedEvents: {0}`r`n" -f $events.Count


if ( $events.count -eq 0 )
{
    $searchPath = 'L:\Logs\ForwardedEvents'
    $lastArchivedLog = Get-ChildItem -Filter "Archive-ForwardedEvents*.evtx" -Path $searchPath | Select -Last 1

    $events = Get-WinEvent -Path $lastArchivedLog.FullName -FilterXPath $xpath | Where-Object { $_.Id -eq 4625 }
    "ArchiveEvents: {0}`r`n" -f $events.Count
}

# auditEvents -events $events

#region function auditEvents
<# function auditEvents
{
    [CmdletBinding()]
    param($events) #>

    $reportDT = Get-Date
    $yyyy = $reportDT.Year
    $mm = "{0:00}" -f $reportDT.Month
    $dd = "{0:00}" -f $reportDT.Day
    $hour = "{0:00}" -f $reportDT.Hour
    $min = "{0:00}" -f $reportDT.Minute
    $sec = "{0:00}" -f $reportDT.Second

    # create a container for events email report$
    [string]$filepath = "..\$yyyy-$mm-$dd_$hour-$min-$sec - EventID - $($events.id | Select -Unique).csv"
    

    # string builder
    $sb = New-Object -TypeName System.Text.StringBuilder

    # number of events found
    $eventsCount = $events.Count

    for ( $i = 0; $i -lt $eventsCount; $i++ )
    {
        $timeCreated = $events[$i].TimeCreated #.ToString()
        $eventID = $events[$i].Id #.ToString()

        # do for all events
        $eventXML = $events[$i].ToXml()
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
    } # end for $i $events loop


    $sb.ToString()
    New-Item -Path $filepath -Value $($sb.ToString()) -Force | Out-Null

    # check number of events; skip sending report if no events matched criteria
    if ( $eventsCount -ge 1 )
    {
        sendMailAlert -subject 'Failed logins' -body 'Excessive failed logins reported. See attached file for details?' -attachment $filepath
    }

# } # end auditEvents function
#endregion function auditEvents


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


