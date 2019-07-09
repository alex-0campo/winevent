[CmdletBinding()]
param()

Clear-Host

$xml = '
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">*[System[band(Keywords,4503599627370496)]]</Select>
  </Query>
</QueryList>
'

$events = Get-WinEvent -FilterXml $xml # | Select -First 10

$events
