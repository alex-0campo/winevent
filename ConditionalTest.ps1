[CmdletBinding()]
param()

Push-Location

Set-Location -Path 'C:\Temp'

Clear-Host

#region sendMailAlert
function sendMailAlert 
{
    [CmdletBinding()]
    param([string]$subject,
            [string]$body,
            [string[]]$attachments
            )
        
    if ( $attachments )
    {
        $MailMessage = @{
            From = "securityAlert@landesa.org"
            To = "Alex Ocampo <alexo@rdiland.org>"
            Subject = $subject
            Body = $body
            SmtpServer = "10.0.0.10"
            Attachments = $attachments
        }
    }
    else
    {
        $MailMessage = @{
            From = "securityAlert@landesa.org"
            To = "Alex Ocampo <alexo@rdiland.org>"
            Subject = $subject
            Body = $body
            SmtpServer = "10.0.0.10"
        }
    }

    Send-MailMessage @MailMessage -Priority High -DeliveryNotificationOption onSuccess, onFailure -BodyAsHtml:$true 
}
#endregion sendMailAlert

function TestFileName
{
    [CmdletBinding()]
    param($fileName)

    [bool]$result = Test-Path -Path $fileName -ErrorAction SilentlyContinue
    
    return $result
}

function doStuff
{
    Write-Host Write-Host "Doing stuff here...`r`n" -ForegroundColor Cyan
}

###################### start here ######################

$fileName = 'C:\Temp\lastSend.txt'



# TestFileName -fileName $fileName

do
{
    Get-Date
    
    [bool]$doesFileExists = TestFileName -fileName $fileName

    $doesFileExists

    if ( !($doesFileExists) )
    {
    
        Write-Host "$fileName not found, creating file now..." -ForegroundColor Red
        New-Item -Path $fileName -Value $(Get-Date).ToString() | Out-Null    

        # run doStuff function
        doStuff
        sendMailAlert -subject "First email..." -body "Doing stuff here...`r`n"
    } 
    else 
    {
        $lastChangeDate = $(Get-ChildItem -Path $fileName).LastWriteTime
        if ( $lastChangeDate -le $((Get-Date).AddMinutes(-15)) )
        {
            Write-Host "$fileName found, do some other stuff here...`r`n" -ForegroundColor Cyan
            sendMailAlert -subject "Next email..." -body "$fileName found, do some other stuff here...`r`n"
            
            New-Item -Path $fileName -Value $(Get-Date).ToString() -ErrorAction SilentlyContinue -Force | Out-Null
        }
        else
        {
            Write-Host "Do nothing...`r`n" -ForegroundColor Yellow
        }
    }

    Write-Host "Task will resume in 5 minutes, please wait...`r`n"

    Start-Sleep -Seconds 300
} while ( $true )

# $doesFileExists

TestFileName -fileName $fileName

Pop-Location


# if Get-Date is older than create date of a file

# then do this and this

# else do nothing


