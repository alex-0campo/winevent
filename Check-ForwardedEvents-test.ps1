<#
# add forwarded events
# to dynamic array example
#>

Clear-Host
$xpath = "*[System[band(Keywords,4503599627370496) and TimeCreated[timediff(@SystemTime) <= 300000]]]"
$events = Get-WinEvent -LogName 'ForwardedEvents' -FilterXPath $xpath
$groups = $events | group id

$sortedGroups = $groups | sort count -desc # | select -First 1



foreach($group in $sortedGroups)
{
    $tempVar = @() # temporary variable to hold events in an array

    foreach ($event in ($group.Group))
    {
        $tempVar += $event
    } # end foreach event

    $varName = $group.Name
    if(Get-Variable -Name $varName)
    {
        Remove-Variable -Name ($varName)
        New-Variable -Name ($varName) -Value $tempVar
    } 
    else
    {
        New-Variable -Name ($varName) -Value $tempVar # sets variable to the value of temporary array
    }

    Remove-Variable tempVar # clean up
    Invoke-Expression "`$$varName`r`n" # return the value of the array
    "`r`n### End of array `$$varName ###`r`n"
    
} # end foreach group
