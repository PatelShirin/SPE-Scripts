<#
.SYNOPSIS
    Clears CSV files from the Sitecore App_Data/packages directory on a remote server.

.DESCRIPTION
    This script connects to a remote Sitecore server, lists all CSV files in the App_Data/packages directory,
    and allows the user to select files to delete. The user can choose specific files by their index or delete all files.
    The script uses remote PowerShell sessions and custom helper functions for session management and remote execution.

.NOTES
    - Requires the Get-SitecoreSession.ps1 script in the same directory for session management.
    - Uses Invoke-RemoteScript for executing commands on the remote server.
    - Designed for Sitecore environments hosted on Windows/IIS.

.PARAMETER Session
    The remote PowerShell session object used for executing commands on the Sitecore server.

.PARAMETER RemotePath
    The full path of the remote file to be deleted.

.FUNCTIONS
    Get-RemoteCsvFiles
        Retrieves a list of CSV files from the remote Sitecore App_Data/packages directory.

    Remove-RemoteFile
        Deletes a specified file from the remote server and provides feedback on the operation.

.EXAMPLE
    # Run the script to list and delete CSV files from the remote Sitecore server.
    .\Clear-sitecore-App_Data-contents.ps1

.LINK
    https://doc.sitecore.com/
#>
Clear-Host

# Dot-source the Load-Env script so it is available
. "$PSScriptRoot\Load-Env.ps1"

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