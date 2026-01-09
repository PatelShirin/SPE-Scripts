Clear-Host

# Dot-source the Get-SitecoreSession function so it is available
. "$PSScriptRoot\Get-SitecoreSession.ps1"
function Get-RemoteCsvFiles {
    param($Session)
    Invoke-RemoteScript -Session $Session -ScriptBlock {
        Get-ChildItem "C:\inetpub\wwwroot\app_data\packages" -Filter *.csv |
        Select-Object -ExpandProperty FullName
    }
}

function Remove-RemoteFile {
    param($Session, $RemotePath)
    Invoke-RemoteScript -Session $Session -ScriptBlock {
        if (Test-Path $Using:remotePath) {
            Remove-Item -Path $Using:remotePath -Force
            Write-Host "Deleted remote file: $Using:remotePath" -ForegroundColor Green
        }
        else {
            Write-Host "File not found: $Using:remotePath" -ForegroundColor Yellow
        }
    } -Arguments @{ remotePath = $RemotePath }
}

$sitecoreInfo = Get-SitecoreSession
if (-not $sitecoreInfo) { return }

Write-Host "Fetching CSV list from remote server..." -ForegroundColor Cyan
$csvFiles = Get-RemoteCsvFiles -Session $sitecoreInfo.Session
if (-not $csvFiles) {
    Write-Host "No CSV files found on remote server." -ForegroundColor Yellow
    return
}

Write-Host "`nAvailable CSV files:" -ForegroundColor Green
$index = 1
$csvFiles | ForEach-Object {
    Write-Host "[$index] $_"
    $index++
}


$selection = Read-Host "`nEnter the number of the file(s) to delete (comma-separated, or * for all)"
if ($selection -eq '*') {
    $selectedFiles = $csvFiles
}
else {
    $selectedIndexes = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
    if (-not $selectedIndexes) {
        Write-Host "No valid selections made." -ForegroundColor Yellow
        return
    }
    $selectedFiles = foreach ($i in $selectedIndexes) {
        $csvFiles[[int]$i - 1]
    }
}

Write-Host "`nDeleting selected files..." -ForegroundColor Cyan
foreach ($remotePath in $selectedFiles) {
    $fileName = Split-Path $remotePath -Leaf
    Remove-RemoteFile -Session $sitecoreInfo.Session -RemotePath $remotePath
    $fileName
}
Write-Host "`nFile deletion process completed." -ForegroundColor Green