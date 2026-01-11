<#
.SYNOPSIS
    Fixes bilingual property items in Sitecore by removing unnecessary language versions from child items based on a CSV input.

.DESCRIPTION
    This script connects to a Sitecore instance using a remote session, locates a PropertyMaster CSV file for the selected environment, and processes property items marked as non-bilingual.
    For each non-bilingual property, it removes the alternate language version (French or English) from all child items that have more than one language version.
    The script outputs the actions taken for each item and skips items where no update is necessary.

.PARAMETER None
    The script does not accept parameters directly; it retrieves environment and session information from an imported module.

.INPUTS
    None. The script reads from a PropertyMaster CSV file located in the package directory.

.OUTPUTS
    System.String
        Outputs messages indicating which language versions were removed from which items, or if no update was necessary.

.NOTES
    - Requires the Get-SitecoreSession.ps1 module to be present in the parent directory.
    - Assumes the presence of a PropertyMaster CSV file matching the current environment in the package directory.
    - Requires Sitecore PowerShell Extensions (SPE) remoting capabilities.

.EXAMPLE
    .\fix-bilingual-properties.ps1
    # Connects to Sitecore, processes the PropertyMaster CSV, and removes unnecessary language versions from child items.
#>
Clear-Host
# Environment selection
. "$PSScriptRoot\..\Load-Env.ps1"

Import-Module "$PSScriptRoot\..\Get-SitecoreSession.ps1"
$sitecoreInfo = Get-SitecoreSession
if (-not $sitecoreInfo) { return }
$session = $sitecoreInfo.session
$sitecoreEnv = $sitecoreInfo.sitecoreEnv
$packageDir = $sitecoreInfo.packageDir

$scriptResults = Invoke-RemoteScript -Session $session  -ScriptBlock {
    $propertyMaster = Get-ChildItem -Path $Using:packageDir -File | Where-Object {
        $_.Name -like '*PropertyMaster*' -and
        $_.Name -like "*$sitecoreEnv*"
    } | Select-Object -First 1
    Write-Output "Processing file: $($propertyMaster.Name)"
    $properties = Import-Csv -Path "$($Using:packageDir)\$($propertyMaster.Name)"
    foreach ($property in $properties | Where-Object { $_.Bilingual -eq "False" }) {
        $scProperty = Get-Item -Path master: -Id $property.PropertyItemId -Language $property.Language
        $propertyChildren = $scProperty | Get-ChildItem -Recurse -WithParent | `
            Where-Object { $_.Versions.GetVersions($true).Language.Count -gt 1 }

        if (-not $propertyChildren) {
            Write-Output ("Nothing to update for {0}" -f $scProperty.Fields["Property Name"].Value)
            continue
        }

        switch ($property.Language) {
            "en" {
                $propertyChildren | ForEach-Object {
                    $_ | Remove-ItemVersion -Language "fr"
                    Write-Output ("Removed FR version from {0}" -f $_.Paths.ContentPath)
                }
            }
            "fr" {
                $propertyChildren | ForEach-Object {
                    $_ | Remove-ItemVersion -Language "en"
                    Write-Output ("Removed EN version from {0}" -f $_.Paths.ContentPath)
                }
            }
            default {
                $propertyChildren | ForEach-Object {
                    Write-Output ("Nothing to update for {0}" -f $_.Paths.ContentPath)
                }
            }
        }
    }
}-Arguments @{packageDir = $packageDir; sitecoreEnv = $sitecoreEnv }
$scriptResults
Stop-ScriptSession -Session $session