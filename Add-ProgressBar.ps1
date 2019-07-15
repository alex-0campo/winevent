Clear-Host
"`r`n`r`n`r`n`r`n`r`n`r`nProgress bar task started at: {0}" -f (Get-Date -Format "MM/dd/yyyy hh:mm:ss tt")

$sleepMilliSeconds = 1*60000 # n (min) x 60000 sec/min

$sleepTimer = $sleepMilliSeconds / 1000

for($i=1;$i -le $sleepTimer; $i++)
{
    Write-Progress -Activity "Start task in $sleepTimer seconds, please wait..." `
    -Status "`$i equals $i" `
    -PercentComplete ($i/$sleepTimer*100)
    Start-Sleep -Seconds 1
}

"`r`nProgress bar task completed at: {0}`r`n`r`n" -f (Get-Date -Format "MM/dd/yyyy hh:mm:ss tt")