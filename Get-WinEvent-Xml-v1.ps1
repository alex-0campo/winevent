[CmdletBinding()]
param()

Clear-Host

#region functions

function getForwardedEvents {
    [CmdletBinding()]
    param()

    try
    {
        Get-WinEvent -FilterXml $xml -ComputerName 'rdi-us-sv0111v.rdiland.org' `
            -ErrorAction Stop
         
    }
    catch [Exception]
    {
        if ($_.Exception -match "No events were found*")
        {
            Write-Host "Oops, no events found matching the query..." -ForegroundColor Cyan
        }
    }
    finally
    {
        # cleanup here...  
        # Write-Host "end of query...`r`n"
    }        
}

function groupByID {
    [CmdletBinding()]
    param($events)

    $events | group id # | sort count -desc | ft count,name -auto) 
}

function groupByUser {
    [CmdletBinding()]
    param($group)

    $group.group # | fl *
}

#endregion

$milSec = (1*60*1000) # number of minutes x 60 sec x 1000 ms
$excludeID = 4624

#region xml query
$xml = "<QueryList>
          <Query Id='0' Path='ForwardedEvents'>
            <Select Path='ForwardedEvents'>
                *[
                    System[TimeCreated[timediff(@SystemTime) &lt;= $milSec] and band(Keywords,4503599627370496)] 
                    <!--System[TimeCreated[timediff(@SystemTime) &lt;= $milSec] and EventID=$eventID]--> 
                    <!--and--> 
                    <!--EventData[Data[@Name='SubjectUserName'] and (Data=$user)]-->
                ]
            </Select>
            <Suppress Path='ForwardedEvents'>
                *[System[(EventID=$excludeID)]]
            </Suppress>
          </Query>
        </QueryList>"
#endregion

#region execution starts here

do
{
    "Date: {0}" -f $(Get-Date)
    # group events by Event ID
    foreach ($group in (groupByID -events (getForwardedEvents)))
    {
        ##########################################
        ### work in progress #####################
        ### loop through each group of events ####
        ##########################################
        foreach($event in (groupByUser $group))
        {
            $eventXML =  [xml]$event.ToXml()
            $eventXML.Event.System
            $eventXML.Event.EventData.Data | ft -auto
            "`nx-x-x-x-x-x-x-x-x-x-x-x-x-x`n"
        }
    }
    # "end of run...`r`n"
    # "`r`n"
    Start-Sleep -Milliseconds ($milSec-($milSec*0.05))   
}
While ($true)

#endregion
