# Load-Env.ps1
# Loads environment variables from a .env file in the script's directory

$envFile = Join-Path -Path $PSScriptRoot -ChildPath ".env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^(\s*#|\s*$)') { return } # Skip comments and empty lines
        if ($_ -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            [System.Environment]::SetEnvironmentVariable($key, $value, 'Process')
            ${env:$key} = $value
        }
    }
    Write-Host ".env variables loaded into session." -ForegroundColor Green
}
else {
    Write-Host ".env file not found in script directory." -ForegroundColor Yellow
}
