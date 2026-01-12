. "$PSScriptRoot\Load-Env.ps1"


# Prompt for directory path, using env as default if available, and auto-select the only file
function Get-LocalFilePath {
    param([string]$defaultDir)
    while ($true) {
        if ($defaultDir) {
            $inputDir = Read-Host "Enter directory path to select file from [$defaultDir]"
            if ([string]::IsNullOrWhiteSpace($inputDir)) {
                $dirPath = $defaultDir

            }
            else {
                $dirPath = Read-Host "Enter directory path to select file from"
            }
            if ([string]::IsNullOrWhiteSpace($dirPath) -or -not (Test-Path $dirPath) -or -not (Get-Item $dirPath).PSIsContainer) {
                Write-Warning "Please enter a valid directory path."
                continue
            }
            # Get files in directory
            $files = Get-ChildItem -Path $dirPath -File
            if ($files.Count -eq 0) {
                Write-Warning "No files found in directory. Please choose another directory."
                continue
            }
            elseif ($files.Count -gt 1) {
                Write-Warning "More than one file found in directory. Please ensure only one file is present."
                continue
            }
            return $files[0].FullName
        }
    }
}
$localFilePath = Get-LocalFilePath $env:LOCAL_WINDOWS_UPLOAD_FILE_PATH

# Function to prompt for directory, list files, and prompt for file name
# function Get-ValidLocalFilePathFromDir {
#     param([string]$defaultDir)
#     while ($true) {
#         if ($defaultDir) {
#             $inputDir = Read-Host "Enter directory path to select file from [$defaultDir]"
#             if ([string]::IsNullOrWhiteSpace($inputDir)) {
#                 $dirPath = $defaultDir
#             }
#             else {
#                 $dirPath = $inputDir
#             }
#         }
#         else {
#             $dirPath = Read-Host "Enter directory path to select file from"
#         }
#         if ([string]::IsNullOrWhiteSpace($dirPath) -or -not (Test-Path $dirPath) -or -not (Get-Item $dirPath).PSIsContainer) {
#             Write-Warning "Please enter a valid directory path."
#             continue
#         }
#     }
# }
        
# $localFilePath = Get-LocalFilePath $env:LOCAL_WINDOWS_FILES_PATH


# Ensure the local file path has .xlsx extension if missing
if (-not ([System.IO.Path]::GetExtension($localFilePath))) {
    $localFilePath += ".xlsx"
}

# Get the local file name and append it to $localFilePath if not already present
$localFileName = $localFilePath | Split-Path -Leaf
if (-not ($localFilePath -like "*$localFileName")) {
    if ($localFilePath -notmatch '[\\/]$') {
        if ($localFilePath -match '^[a-zA-Z]:') {
            $localFilePath += '\'
        }
        else {
            $localFilePath += '/'
        }
    }
    $localFilePath += $localFileName
}


# Use environment variables as defaults for SSH details
$defaultRemoteUser = $env:REMOTE_SSH_USERNAME
$defaultRemoteHost = $env:REMOTE_SSH_HOST
$defaultRemotePort = $env:REMOTE_SSH_PORT

# Prompt for remote SSH details, using defaults if available
if ($defaultRemoteUser) {
    $remoteUser = Read-Host "Enter remote SSH username [$defaultRemoteUser]"
    if ([string]::IsNullOrWhiteSpace($remoteUser)) { $remoteUser = $defaultRemoteUser }
}
else {
    $remoteUser = Read-Host "Enter remote SSH username"
}

if ($defaultRemoteHost) {
    $remoteHost = Read-Host "Enter remote SSH host (e.g., example.com / remote.ip.address) [$defaultRemoteHost]"
    if ([string]::IsNullOrWhiteSpace($remoteHost)) { $remoteHost = $defaultRemoteHost }
}
else {
    $remoteHost = Read-Host "Enter remote SSH host (e.g., example.com / remote.ip.address)"
}

if ($defaultRemotePort) {
    $remotePort = Read-Host "Enter remote SSH port (default 22) [$defaultRemotePort]"
    if ([string]::IsNullOrWhiteSpace($remotePort)) { $remotePort = $defaultRemotePort }
}
else {
    $remotePort = Read-Host "Enter remote SSH port (default 22)"
    if ([string]::IsNullOrWhiteSpace($remotePort)) { $remotePort = 22 }
}
$fileName = $localFilePath | Split-Path -Leaf
Write-Host $fileName -ForegroundColor Green
[string]$remotePath = Read-Host "Enter remote destination directory (e.g., c:/Users/shirin/repos/git-repos/personal-git-repo/SPE-Scripts/PS-CustomScripts/suite-pricing/excel/2026/Jan/)"
# Ensure remotePath ends with a separator
if ($remotePath -notmatch '[\\/]$') {
    if ($remotePath -match '^[a-zA-Z]:') {
        $remotePath += '\'
    }
    else {
        $remotePath += '/'
    }
}
# Append fileName if not already present
if (-not ($remotePath -like "*$fileName")) {
    $remotePath += $fileName
}

# Ensure remote directory exists before copying
$remoteDir = [System.IO.Path]::GetDirectoryName($remotePath)
if ($null -ne $remoteDir -and $remoteDir -ne '' -and $remoteDir -ne '/' -and $remoteDir -ne '\\') {
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
}
else {
    Write-Host "No remote directory to create or directory is root. Skipping mkdir."
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
