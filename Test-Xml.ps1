[CmdletBinding()]
param()

Clear-Host

$ms = $(New-Timespan -Hours 1).Totalmilliseconds

$xml = "<QueryList>
  <Query Id='0' Path='ForwardedEvents'>
    <Select Path='ForwardedEvents'>
	  *[System[(EventID=4625) and TimeCreated[timediff(@SystemTime) &lt;=$ms]]]
	  and
	  *[EventData[Data[@Name='ProcessName'] and (Data='C:\Program Files\Dell\SupportAssistAgent\PCDr\SupportAssist\6.0.7033.2285\pcdrsysinfosoftware.p5x')]]
    </Select>
  </Query>
</QueryList>"


$events = Get-WinEvent -FilterXml $xml
