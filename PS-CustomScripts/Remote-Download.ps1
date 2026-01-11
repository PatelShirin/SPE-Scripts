<#
.SYNOPSIS
    Downloads selected CSV files from a remote Sitecore server using SPE Remoting.

.DESCRIPTION
    This script connects to a remote Sitecore instance using SPE Remoting, lists available CSV files in the 
    "C:\inetpub\wwwroot\app_data\packages" directory, and allows the user to select and download one or more files 
    to a local "Downloads" folder. It handles session management, directory creation, and file transfer using 
    base64 encoding.

.PARAMETER None
    The script does not accept parameters; it prompts the user for input during execution.

.FUNCTIONS
    New-Directory
        Ensures a directory exists at the specified path, creating it if necessary.

    Get-RemoteCsvFiles
        Retrieves the full paths of all CSV files in the remote packages directory.

    Get-RemoteFile
        Downloads a specified file from the remote server by encoding it in base64 and saving it locally.

.NOTES
    - Requires the SPE (Sitecore PowerShell Extensions) module and a helper script "Get-SitecoreSession.ps1".
    - User must have appropriate permissions on the remote Sitecore server.
    - The script handles errors gracefully and provides colored output for status messages.

.EXAMPLE
    .\Remote-Download.ps1
    # Connects to the remote Sitecore server, lists available CSV files, and downloads selected files to the local "Downloads" folder.

#>

Clear-Host
Import-Module -Name SPE
. "$PSScriptRoot\Load-Env.ps1"
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
    Get-RemoteFile -Session $session -RemotePath $remotePath -LocalPath $localPath | Out-Null
}

Write-Host "`nAll selected downloads completed." -ForegroundColor Magenta
try {
    Stop-ScriptSession -Session $session
}
catch {
    Write-Host "Failed to stop remote session." -ForegroundColor Yellow
}

