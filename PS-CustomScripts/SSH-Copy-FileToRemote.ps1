Clear-Host
. "$PSScriptRoot\Load-Env.ps1"

function Get-LocalFilePath {
    param([string]$defaultDir)
    while ($true) {
        # Use correct separator for current OS
        $sep = [IO.Path]::DirectorySeparatorChar
        $promptDir = if ($defaultDir) { "$HOME$sep$defaultDir" } else { "" }
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
        Write-Host "Files in directory:"
        for ($i = 0; $i -lt $files.Count; $i++) {
            Write-Host "$($i+1): $($files[$i].Name)"
        }
        $fileChoice = Read-Host "Enter the number of the file to upload"
        if ($fileChoice -match '^[0-9]+$' -and $fileChoice -ge 1 -and $fileChoice -le $files.Count) {
            return $files[$fileChoice - 1].FullName
        }
        else {
            Write-Warning "Invalid selection. Please try again."
        }
    }
}
function Get-RemoteValue {
    param(
        [string]$envVar,
        [string]$prompt,
        [string]$default
    )
    $userInput = if ($envVar) { Read-Host "$prompt [$envVar]" } elseif ($default) { Read-Host "$prompt [$default]" } else { Read-Host $prompt }
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        return $envVar ? $envVar : $default
    }
    return $userInput
}


# Inject year and first 3 letters of month into local file path
$year = (Get-Date).Year
$monthShort = (Get-Date).ToString("MMM")
$basePath = $env:LOCAL_UPLOAD_FILE_PATH
# Normalize base path first
if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
    Write-Host "Windows OS detected, Normalizing path for Windows."
    $basePath = $basePath -replace '/', '\'
    $dynamicLocalPath = "$basePath\$year\$monthShort"
    $dynamicLocalPath = $dynamicLocalPath -replace '/', '\'
}
else {
    Write-Host "Non-Windows OS detected, normalizing path for Unix-like system."
    $basePath = $basePath -replace '\\', '/'
    $dynamicLocalPath = "$basePath/$year/$monthShort"
    $dynamicLocalPath = $dynamicLocalPath -replace '\\', '/'
}
$localFilePath = Get-LocalFilePath $dynamicLocalPath
if (-not ([System.IO.Path]::GetExtension($localFilePath))) {
    $extension = Read-Host "No file extension detected. Please enter the file extension (e.g., .txt, .csv, .xlsx)"
    if ($extension -and $extension -notmatch '^\.') { $extension = ".$extension" }
    $localFilePath += $extension
}
$localFileName = $localFilePath | Split-Path -Leaf
if (-not ($localFilePath -like "*$localFileName")) {
    if ($localFilePath -notmatch '[\\/]$') { $localFilePath += ($localFilePath -match '^[a-zA-Z]:') ? '\\' : '/' }
    $localFilePath += $localFileName
}

# Get remote connection info
$remoteUser = Get-RemoteValue $env:REMOTE_SSH_USERNAME "Enter remote SSH username" ""
$remoteHost = Get-RemoteValue $env:REMOTE_SSH_HOST "Enter remote SSH host (e.g., example.com / remote.ip.address)" ""
$remotePort = Get-RemoteValue $env:REMOTE_SSH_PORT "Enter remote SSH port (default 22)" "22"

# Get remote destination path
$fileName = $localFilePath | Split-Path -Leaf
Write-Host $fileName $env:REMOTE_SSH_DESTINATION_PATH/excel/$year/$monthShort/ -ForegroundColor Green
$remotePath = Get-RemoteValue $env:REMOTE_SSH_DESTINATION_PATH/excel/$year/$monthShort/ "Enter remote destination directory where file will be copied to" ""
if ($remotePath -notmatch '[\\/]$') { $remotePath += ($remotePath -match '^[a-zA-Z]:') ? '\' : '/' }
if (-not ($remotePath -like "*$fileName")) { $remotePath += $fileName }

# Get remote destination path
$defaultRemoteDir = "$($env:REMOTE_SSH_DESTINATION_PATH)/excel/$year/$monthShort/"
$defaultRemotePath = if ($defaultRemoteDir -and ($defaultRemoteDir -notmatch '[\\/]$')) {
    $defaultRemoteDir += ($defaultRemoteDir -match '^[a-zA-Z]:') ? '\' : '/'
    $defaultRemoteDir
}
else {
    $defaultRemoteDir
}
$defaultRemotePath = $defaultRemotePath + $fileName
$remotePath = Read-Host "Confirm or edit remote destination path [$defaultRemotePath]"
if ([string]::IsNullOrWhiteSpace($remotePath)) { $remotePath = $defaultRemotePath }


# Always attempt to create remote directory, treat 'already exists' as success
$remoteDir = [System.IO.Path]::GetDirectoryName($remotePath)
$remoteDir = [System.IO.Path]::GetDirectoryName($remotePath)
if ($null -ne $remoteDir -and $remoteDir -ne '' -and $remoteDir -ne '/' -and $remoteDir -ne '\') {
    $isWindowsPath = $remoteDir -match '^[a-zA-Z]:'
    $remoteDirFinal = $isWindowsPath ? ($remoteDir -replace '/', '\\') : ($remoteDir -replace '\\', '/')
    if ($isWindowsPath) {
        $mkdirCmd = 'mkdir "' + $remoteDirFinal + '"'
    }
    else {
        $mkdirCmd = "mkdir -p '$remoteDirFinal'"
    }
    Write-Host "Ensuring remote directory exists..."
    try {
        $mkdirOutput = & ssh -p $remotePort "$remoteUser@${remoteHost}" $mkdirCmd 2>&1
        if ($mkdirOutput -match 'already exists|File exists') {
            Write-Host "Remote directory already exists."
        }
        else {
            Write-Host "Remote directory ensured."
        }
    }
    catch {
        Write-Warning "Failed to create remote directory via ssh. The copy may fail if the directory does not exist."
    }
}
else {
    Write-Host "No remote directory to create or directory is root. Skipping mkdir."
}

# Copy file to remote server
Write-Host "Copying file to remote server..."
try {
    & scp -P $remotePort $localFilePath "$remoteUser@${remoteHost}:${remotePath}" 2>&1
    Write-Host "Remote copy completed."
}
catch {
    Write-Warning "Failed to copy file to remote server via scp. Please ensure scp is installed and accessible."
}