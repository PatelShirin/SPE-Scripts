
param(
    [string]$localFilePath
)

# Ensure the local file path has .xlsx extension if missing
if (-not ([System.IO.Path]::GetExtension($localFilePath))) {
    $localFilePath += ".xlsx"
}

# Prompt for remote SSH details
[string]$remoteUser = Read-Host "Enter remote SSH username"
[string]$remoteHost = Read-Host "Enter remote SSH host (e.g., example.com / remote.ip.address)"
[string]$remotePort = Read-Host "Enter remote SSH port (default 22)"
if ([string]::IsNullOrWhiteSpace($remotePort)) { $remotePort = 22 }
$fileName = $localFilePath | Split-Path -Leaf
Write-Host $fileName
[string]$remotePath = Read-Host "Enter remote destination path (e.g., c:/Users/shirin/repos/git-repos/personal-git-repo/SPE-Scripts/PS-CustomScripts/suite-pricing/excel/2026/Jan/$fileName)"

# Ensure remote directory exists before copying
$remoteDir = [System.IO.Path]::GetDirectoryName($remotePath)
if ($remoteDir -match '^[a-zA-Z]:') {
    # Windows path: use backslashes and wrap in double quotes
    $remoteDirWin = $remoteDir -replace '/', '\'
    $sshMkdirCommand = 'ssh -p ' + $remotePort + ' ' + $remoteUser + '@' + $remoteHost + ' mkdir "' + $remoteDirWin + '"'
}
else {
    # Linux/Unix path: use forward slashes and -p
    $remoteDirUnix = $remoteDir -replace '\\', '/'
    $sshMkdirCommand = 'ssh -p ' + $remotePort + ' ' + $remoteUser + '@' + $remoteHost + ' mkdir -p ' + $remoteDirUnix
}
Write-Host "Ensuring remote directory exists using: $sshMkdirCommand"
try {
    & $env:COMSPEC /c $sshMkdirCommand
    Write-Host "Remote directory ensured."
}
catch {
    Write-Warning "Failed to create remote directory via ssh. The copy may fail if the directory does not exist."
}
$scpCommand = 'scp -P ' + $remotePort + ' "' + $localFilePath + '" ' + $remoteUser + '@' + $remoteHost + ':' + '"' + $remotePath + '"'
Write-Host "Copying file to remote server using: $scpCommand"
try {
    & $env:COMSPEC /c $scpCommand
    Write-Host "Remote copy completed."
}
catch {
    Write-Warning "Failed to copy file to remote server via scp. Please ensure scp is installed and accessible."
}
