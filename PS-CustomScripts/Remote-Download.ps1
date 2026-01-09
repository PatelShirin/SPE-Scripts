
Clear-Host
Import-Module -Name SPE
. "$PSScriptRoot\Get-SitecoreSession.ps1"

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Get-RemoteCsvFiles {
    param($Session)
    Invoke-RemoteScript -Session $Session -ScriptBlock {
        Get-ChildItem "C:\inetpub\wwwroot\app_data\packages" -Filter *.csv |
        Select-Object -ExpandProperty FullName
    }
}

function Get-RemoteFile {
    param($Session, $RemotePath, $LocalPath)

    try {
        $base64 = Invoke-RemoteScript -Session $Session -ScriptBlock {
            if (-not (Test-Path $Using:remotePath)) {
                throw "File not found: $Using:remotePath"
            }
            $bytes = [IO.File]::ReadAllBytes($Using:remotePath)
            [Convert]::ToBase64String($bytes)
        } -Arguments @{ remotePath = $RemotePath }
        if (-not $base64) {
            Write-Host "Failed to download $RemotePath" -ForegroundColor Red
            return $false
        }
        $clean = ($base64 | Where-Object { $_ -is [string] }) -join ""
        [IO.File]::WriteAllBytes($LocalPath, [Convert]::FromBase64String($clean))
        Write-Host "Saved to: $LocalPath" -ForegroundColor Cyan
        return $true
    }
    catch {
        Write-Host "Error downloading '$RemotePath'" -ForegroundColor Red
        Write-Host $_ -ForegroundColor DarkRed
        return $false
    }
}

# Main script
$sitecoreInfo = Get-SitecoreSession
if (-not $sitecoreInfo) { return }
$session = $sitecoreInfo.Session
$downloadFolder = Join-Path $PSScriptRoot "Downloads"
New-Directory $downloadFolder

Write-Host "Fetching CSV list from remote server..." -ForegroundColor Cyan
$csvFiles = Get-RemoteCsvFiles -Session $session
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
$selection = Read-Host "`nEnter the number(s) of the files to download (comma-separated)"
$selectedIndexes = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
if (-not $selectedIndexes) {
    Write-Host "No valid selections made." -ForegroundColor Yellow
    return
}
if ($csvFiles.Count -eq 1) {
    $selectedFiles = $csvFiles
}
else {
    $selectedFiles = foreach ($i in $selectedIndexes) {
        $csvFiles[[int]$i - 1]
    }
}
Write-Host "`nDownloading selected files..." -ForegroundColor Cyan
foreach ($remotePath in $selectedFiles) {
    $fileName = Split-Path $remotePath -Leaf
    $localPath = Join-Path $downloadFolder $fileName
    $fileName, $remotePath, $localPath
    Get-RemoteFile -Session $session -RemotePath $remotePath -LocalPath $localPath | Out-Null
}

Write-Host "`nAll selected downloads completed." -ForegroundColor Magenta
try {
    Stop-ScriptSession -Session $session
}
catch {
    Write-Host "Failed to stop remote session." -ForegroundColor Yellow
}

