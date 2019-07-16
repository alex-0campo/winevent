[CmdletBinding()]
param()

Clear-Host

function TestFileName
{
    [CmdletBinding()]
    param($fileName)

    [bool]$result = Test-Path -Path $fileName -ErrorAction SilentlyContinue

    <# if ( Test-Path -Path $fileName -ErrorAction SilentlyContinue )
    {
        Write-Host "$fileName file exists." -ForegroundColor Yellow
    }
    else 
    {
        Write-Host "$fileName file missing." -ForegroundColor Red
    } #>

    return $result
}

$fileName = 'C:\Temp\lastSend.txt'

[bool]$doesFileExists = TestFileName -fileName $fileName

$doesFileExists

if ( !($doesFileExists) )
{
    
    Write-Host "$fileName not found, creating file now..." -ForegroundColor Red
    New-Item -Path $fileName -Value $(Get-Date).ToString() | Out-Null
} 
else 
{
    Write-Host "$fileName found, do nothing..." -ForegroundColor Yellow
}

$doesFileExists

Clear-Host

# if Get-Date is older than create date of a file

# then do this and this

# else do nothing


