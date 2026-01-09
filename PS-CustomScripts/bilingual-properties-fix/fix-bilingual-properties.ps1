# Environment selection
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