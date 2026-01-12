. "$PSScriptRoot\Load-Env.ps1"

function Get-LocalFilePath {
    param([string]$defaultDir)
    while ($true) {
        $promptDir = if ($defaultDir) { "$HOME\$defaultDir" } else { "" }
        $inputDir = Read-Host "Enter directory path to select file from [$promptDir]"
        $dirPath = if ([string]::IsNullOrWhiteSpace($inputDir) -and $defaultDir) { "$HOME\$defaultDir" } else { $inputDir }
        if ([string]::IsNullOrWhiteSpace($dirPath) -or -not (Test-Path $dirPath) -or -not (Get-Item $dirPath).PSIsContainer) {
            Write-Warning "Please enter a valid directory path."
            continue
        }
        $files = Get-ChildItem -Path $dirPath -File
        if ($files.Count -eq 0) {
            Write-Warning "No files found in directory. Please choose another directory."
            continue
        }
        if ($files.Count -gt 1) {
            Write-Warning "More than one file found in directory. Please ensure only one file is present."
            continue
        }
        return $files[0].FullName
    }
}

function Get-RemoteValue {
    param(
        [string]$envVar,
        [string]$prompt,
        [string]$default
    )
    if ($envVar) {
        $userInput = Read-Host "$prompt [$envVar]"
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return $envVar
        }
        else {
            return $userInput
        }
    }
    elseif ($default) {
        $userInput = Read-Host "$prompt [$default]"
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return $default
        }
        else {
            return $userInput
        }
    }
    else {
        return Read-Host $prompt
    }
}

$localFilePath = Get-LocalFilePath $env:LOCAL_UPLOAD_FILE_PATH

# Ensure the local file path has .xlsx extension if missing
if (-not ([System.IO.Path]::GetExtension($localFilePath))) {
    $localFilePath += ".xlsx"
}

$localFileName = $localFilePath | Split-Path -Leaf
if (-not ($localFilePath -like "*$localFileName")) {
    if ($localFilePath -notmatch '[\\/]$') {
        $localFilePath += ($localFilePath -match '^[a-zA-Z]:') ? '\' : '/'
    }
    $localFilePath += $localFileName
}

$remoteUser = Get-RemoteValue $env:REMOTE_SSH_USERNAME "Enter remote SSH username" ""
$remoteHost = Get-RemoteValue $env:REMOTE_SSH_HOST "Enter remote SSH host (e.g., example.com / remote.ip.address)" ""
$remotePort = Get-RemoteValue $env:REMOTE_SSH_PORT "Enter remote SSH port (default 22)" "22"

$fileName = $localFilePath | Split-Path -Leaf
Write-Host $fileName $env:REMOTE_SSH_DESTINATION_PATH -ForegroundColor Green

$remotePath = Get-RemoteValue $env:REMOTE_SSH_DESTINATION_PATH "Enter remote destination directory where file will be copied to" ""
if ($remotePath -notmatch '[\\/]$') {
    $remotePath += ($remotePath -match '^[a-zA-Z]:') ? '\' : '/'
}
if (-not ($remotePath -like "*$fileName")) {
    $remotePath += $fileName
}

$remoteDir = [System.IO.Path]::GetDirectoryName($remotePath)
if ($null -ne $remoteDir -and $remoteDir -ne '' -and $remoteDir -ne '/' -and $remoteDir -ne '\\') {
    $isWindowsPath = $remoteDir -match '^[a-zA-Z]:'
    $remoteDirFinal = $isWindowsPath ? ($remoteDir -replace '/', '\') : ($remoteDir -replace '\\', '/')
    $mkdirCmd = $isWindowsPath ? "mkdir `"$remoteDirFinal`"" : "mkdir -p $remoteDirFinal"
    $sshMkdirCommand = "ssh -p $remotePort $remoteUser@$remoteHost $mkdirCmd"
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

$scpCommand = "scp -P $remotePort `"$localFilePath`" $remoteUser@${remoteHost}:`"${remotePath}`""
Write-Host "Copying file to remote server using: $scpCommand"
try {
    & $env:COMSPEC /c $scpCommand
    Write-Host "Remote copy completed."
}
catch {
    Write-Warning "Failed to copy file to remote server via scp. Please ensure scp is installed and accessible."
}