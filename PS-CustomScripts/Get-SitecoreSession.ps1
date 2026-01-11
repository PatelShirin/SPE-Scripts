<#
.SYNOPSIS
    Establishes a Sitecore PowerShell Remoting session for a specified environment and returns session details.

.DESCRIPTION
    The Get-SitecoreSession function prompts the user to select a Sitecore environment (dev, cwqa, or cwprod), 
    establishes a remoting session using predefined credentials and connection URIs, and returns a custom object 
    containing session information, environment details, and output file path for further operations.

.PARAMETER packageDir
    The directory path where package files are stored. Defaults to 'c:\inetpub\wwwroot\app_data\packages'.

.OUTPUTS
    PSCustomObject
        Session        - The established Sitecore PowerShell Remoting session object.
        Environment    - The selected environment name.
        EnvMap         - The hashtable containing environment and credential mappings.
        PackageDir     - The directory path for packages.
        Url            - The base URL for the selected environment.
        DateVariables  - Hashtable containing formatted date components.
        OutputFilePath - The generated output file path for the session.

.EXAMPLE
    PS> $sessionInfo = Get-SitecoreSession
    Prompts for environment, establishes a session, and returns session details.

.NOTES
    - Requires Sitecore PowerShell Remoting modules and appropriate permissions.
    - Credentials and connection URIs are hardcoded for demonstration purposes.
    - The function will prompt until a valid environment is entered.
#>
function Get-SitecoreSession {
    param(
        [string]$packageDir = 'c:\inetpub\wwwroot\app_data\packages'
    )
    #Clear-Host
    if ($null -ne $session) {
        try { Stop-ScriptSession -Session $session } catch {}
    }

    $now = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), [System.TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time'))

    $dateVariables = @{
        ShortMonth = $now.ToString('MMM')
        Day        = $now.Day.ToString('00')
        Year       = $now.Year
    }
    $envMap = @{
        "dev"    = @{ 
            url           = "https://dev.chartwell.com"
            connectionUri = if ($env:SITECORE_URI_DEV) { $env:SITECORE_URI_DEV } else { "https://xmc-chartwellmaa139-chartwellsxa-dev.sitecorecloud.io/" }
        }
        "cwqa"   = @{ 
            url           = "https://cwqa.chartwell.com"
            connectionUri = if ($env:SITECORE_URI_CWQA) { $env:SITECORE_URI_CWQA } else { "https://xmc-chartwellmabc8a-chartwellsxa-cwqa.sitecorecloud.io/" }
        }
        "cwprod" = @{ 
            url           = "https://chartwell.com"
            connectionUri = if ($env:SITECORE_URI_CWPROD) { $env:SITECORE_URI_CWPROD } else { "https://xmc-chartwellma83cf-chartwellsxa-cwprod.sitecorecloud.io/" }
        }
    }

    $sitecoreUser = $env:SITECORE_USER
    $sitecoreSecret = $env:SITECORE_SECRET
    if ([string]::IsNullOrWhiteSpace($sitecoreUser) -or [string]::IsNullOrWhiteSpace($sitecoreSecret)) {
        throw "Environment variables SITECORE_USER and SITECORE_SECRET must be set."
    }


    $validEnvs = $envMap.Keys | Where-Object { $_ -ne 'Creds' } | ForEach-Object { $_.ToLower() }    
    $sitecoreEnv = $null
    while ($true) {
        $inputEnv = Read-Host "Enter environment: dev / cwqa / cwprod (CTRL+C to exit)"
        $inputEnv = $inputEnv.Trim()
        if (![string]::IsNullOrWhiteSpace($inputEnv) -and ($validEnvs -contains $inputEnv)) {
            $sitecoreEnv = $inputEnv
            break
        }
        else {
            Write-Host "Invalid environment. Valid options: dev, cwqa, cwprod." -ForegroundColor Red
        }
    }

    $outputFilePath = "$packageDir\PropertyMaster-List-sxastarter-$($sitecoreEnv.ToUpper())-$($dateVariables.ShortMonth)-$($dateVariables.Day)-$($dateVariables.Year).csv"

    $connectionUri = $envMap[$sitecoreEnv].connectionUri
    $session = New-ScriptSession -ConnectionUri $connectionUri `
        -Username $sitecoreUser `
        -SharedSecret $sitecoreSecret

    Write-Host "session connected to: $($session.Connection.Uri)" -ForegroundColor Green  

    return [PSCustomObject]@{
        Session        = $session
        Environment    = $sitecoreEnv
        PackageDir     = $packageDir
        Url            = $envMap[$sitecoreEnv].url
        DateVariables  = $dateVariables
        OutputFilePath = $outputFilePath
    }
}

