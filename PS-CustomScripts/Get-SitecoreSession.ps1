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
        "Creds"  = @{ UserName = "sitecore\speremoting"; SharedSecret = "A345256D29924333A975FC96AFC46DE87B8F0F1B85705283B4F81AD502C3A50A" }
        "dev"    = @{ url = "https://dev.chartwell.com"; connectionUri = "https://xmc-chartwellmaa139-chartwellsxa-dev.sitecorecloud.io/" }
        "cwqa"   = @{ url = "https://cwqa.chartwell.com"; connectionUri = "https://xmc-chartwellmabc8a-chartwellsxa-cwqa.sitecorecloud.io/" }
        "cwprod" = @{ url = "https://chartwell.com"; connectionUri = "https://xmc-chartwellma83cf-chartwellsxa-cwprod.sitecorecloud.io/" }
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
        -Username $envMap["Creds"].UserName `
        -SharedSecret $envMap["Creds"].SharedSecret 

    Write-Host "session connected to: $($session.Connection.Uri)" -ForegroundColor Green  

    return [PSCustomObject]@{
        Session        = $session
        Environment    = $sitecoreEnv
        EnvMap         = $envMap
        PackageDir     = $packageDir
        Url            = $envMap[$sitecoreEnv].url
        DateVariables  = $dateVariables
        OutputFilePath = $outputFilePath
    }
}

