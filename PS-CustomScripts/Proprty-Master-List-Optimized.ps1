<#
.SYNOPSIS
    Generates a master list of property items from Sitecore and exports the data to a CSV file.

.DESCRIPTION
    This script connects to a remote Sitecore instance using SPE (Sitecore PowerShell Extensions), retrieves property items from specified regions and languages, processes their data (including bilingual status, address, URLs, and other fields), and exports the results to a CSV file. It also removes any existing property master list files in the output directory before exporting the new data.

.PARAMETER None
    All required parameters are retrieved from the Sitecore session information via the imported Get-SitecoreSession.ps1 script.

.NOTES
    - Requires the SPE module and a valid Sitecore session.
    - The script is intended to be run locally but executes the main logic remotely on the Sitecore server.
    - The script handles bilingual properties and generates URLs for both English and French versions as needed.
    - The output CSV file is saved to the path specified in the Sitecore session information.

.EXAMPLE
    .\Proprty-Master-List-Optimized.ps1

    Runs the script, connects to Sitecore, retrieves property data, and exports it to a CSV file.

#>
Clear-Host
Import-Module -Name SPE
. "$PSScriptRoot\Load-Env.ps1"
. "$PSScriptRoot\Get-SitecoreSession.ps1"
$sitecoreInfo = Get-SitecoreSession

if (-not $sitecoreInfo) { return }

$session = $sitecoreInfo.Session
$packageDir = $sitecoreInfo.PackageDir
$envUrl = $sitecoreInfo.Url

# Output file path
$outputFilePath = $sitecoreInfo.OutputFilePath 

if (-not $session.Connection.Uri) {
    Write-Host "Failed to connect to remote Sitecore session." -ForegroundColor Red
    return
}

Write-Host "Running script on remote server: $($session.Connection.Uri) - $($sitecoreInfo.DateVariables.ShortMonth) $($sitecoreInfo.DateVariables.Day) $($sitecoreInfo.DateVariables.Year)" -ForegroundColor Green
$scriptResults = Invoke-RemoteScript -Session $session -ScriptBlock {
    function Build-PropertyData {
        param(
            $property, $bilingual, $defaultLanguage, $envUrl
        )

        $urlOptions = [Sitecore.Links.UrlOptions]::DefaultOptions
        $urlOptions.LanguageEmbedding = [Sitecore.Links.LanguageEmbedding]::AsNeeded
        $urlOptions.Language = "fr"
        $urlOptions.AlwaysIncludeServerUrl = $false

        $city = "Missing City"
        $province = "Missing Province"
        if ($property.Fields["City"] -and $property.Fields["City"].Value) {
            $cityItem = Get-Item -Path master: -Id $property.Fields["City"].Value
            if ($cityItem -and $cityItem.Fields["City Name"] -and $cityItem.Fields["City Name"].Value) {
                $city = $cityItem.Fields["City Name"].Value
            }
        }
        if ($property.Fields["Province"] -and $property.Fields["Province"].Value) {
            $provinceItem = Get-Item -Path master: -Id $property.Fields["Province"].Value
            if ($provinceItem -and $provinceItem.Fields["Province Name"] -and $provinceItem.Fields["Province Name"].Value) {
                $province = $provinceItem.Fields["Province Name"].Value
            }
        }

        $itemUrlEn = ""
        $itemUrlFr = ""
        if ($bilingual) {
            $itemUrlEn = [Sitecore.Links.LinkManager]::GetItemUrl($property)
            $itemUrlFr = [Sitecore.Links.LinkManager]::GetItemUrl($property, $urlOptions)
        }
        elseif ($defaultLanguage -eq "en") {
            $itemUrlEn = [Sitecore.Links.LinkManager]::GetItemUrl($property)
        }
        else {
            $itemUrlFr = $Using:envUrl + [Sitecore.Links.LinkManager]::GetItemUrl($property, $urlOptions) -replace ("/sxastarter/sxastarter/accueil", "")
        }
        $priorityValue = if ($property.Fields["isPriorityProperty"].Value -eq "1") { "True" } else { "False" }

        return [ordered]@{
            "Bilingual"        = $bilingual
            "PropertyItemId"   = $property.Id
            "PropertyId"       = $property.Fields["PropertyID"].Value.ToString()
            "PropertyItemName" = $property.Name
            "PropertyName"     = $property.Fields["Property Name"].Value
            "PropertyAddress"  = $property.Fields["StreetNameAndNumber"].Value + ", " + $city + ", " + $province + " " + $property.Fields["Postal Code"].Value
            "PropertyPhone"    = $property.Fields["Contact Number"].Value
            "Language"         = $defaultLanguage
            "PriorityProperty" = $priorityValue
            "City"             = $city
            "Province"         = $province
            "URL En"           = $itemUrlEn
            "URL Fr"           = $itemUrlFr
        }
    }

    # $propertyPages = Get-Item -Path master: -ID "{8A4FBFEA-94DC-41CF-B6B8-CECA339490B1}" | Get-ItemReferrer | `
    #     Where-Object { $_.TemplateID -eq "{8A4FBFEA-94DC-41CF-B6B8-CECA339490B1}" `
    #         -and $_.Name -ne "__Standard Values" }

    $languages = @( "en", "fr")
    $regions = @( "ab", "bc", "on", "qc")
    $basePath = "master:/sitecore/content/sxastarter/sxastarter/home"
    $templateId = "{8A4FBFEA-94DC-41CF-B6B8-CECA339490B1}" # Use ID for faster filtering

    $propertyPages = foreach ($lang in $languages) {
        foreach ($region in $regions) {
            $regionPath = "$basePath/$region"
            Get-ChildItem -Path $regionPath -Language $lang -Recurse |
            Where-Object { $_.TemplateID.ToString() -eq $templateId }
        }
    }

    $uniqueProperties = $propertyPages | Group-Object ID | ForEach-Object { $_.Group[0] }

    $propertyCSVData = @()
    foreach ($property in $uniqueProperties) {
        $languages = $property.Versions.GetVersions($true).Language
        $bilingual = $languages.Count -gt 1 -and $property.Fields["Bilingual"].Value -eq "1"
        $defaultLanguage = if ($bilingual) { "en" } elseif ($property.Fields["Title"].Value -eq "") { "fr" } else { $languages[0].Name }
        $propData = Build-PropertyData $property $bilingual $defaultLanguage $envUrl
        $propertyCSVData += [PSCustomObject]$propData
    }
    # Remove existing files starting with PropertyMaster-List or Property-Master-List
    $filesToRemove = @(Get-ChildItem -Path $Using:packageDir -Filter 'PropertyMaster-List*' -File)
    $filesToRemove += @(Get-ChildItem -Path $Using:packageDir -Filter 'Property-Master-List*' -File)
    foreach ($file in $filesToRemove) {
        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
    }

    $propertyCSVData | Export-Csv -Path $Using:outputFilePath -NoTypeInformation -Encoding UTF8 -Force
    $propertyCSVData 
}-Arguments @{ outputFilePath = $outputFilePath; envUrl = $envUrl }
$scriptResults | Format-Table -Property  Bilingual, Language, PropertyItemId, PropertyItemName -AutoSize
Write-Host "Script completed successfully on remote server: $($session.Connection.Uri)" -ForegroundColor Green
Write-Host "CSV file uploaded to $outputFilePath ($($scriptResults.Count) row(s))" -ForegroundColor Magenta

try {
    Stop-ScriptSession -Session $session
}
catch {
    Write-Host "Failed to stop remote session." -ForegroundColor Yellow
}
