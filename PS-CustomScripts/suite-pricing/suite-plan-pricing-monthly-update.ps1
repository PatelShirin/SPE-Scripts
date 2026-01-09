# Dot-source the Get-SitecoreSession function so it is available
. "$PSScriptRoot\..\Get-SitecoreSession.ps1"
#Clear-Host
if ($null -ne $session) {
    try { Stop-ScriptSession -Session $session } catch { }
}
$sitecoreInfo = Get-SitecoreSession
if (-not $sitecoreInfo) { return }
$sitecoreSessionInfo = $sitecoreInfo

$session = $sitecoreSessionInfo.Session

Write-Output "Running remote script to update suite plans pricing - $($sitecoreSessionInfo.DateVariables.ShortMonth) $($sitecoreSessionInfo.DateVariables.Day) $($sitecoreSessionInfo.DateVariables.Year) - Please wait for the script to finish..."

$envKey = ""
$propertyMasterCsv = ""
$suitePricingCsv = ""
if ($session.Connection.Uri -match "chartwellsxa-dev") { $envKey = "dev" }
elseif ($session.Connection.Uri -match "chartwellsxa-cwqa") { $envKey = "cwqa" }
elseif ($session.Connection.Uri -match "chartwellsxa-cwprod") { $envKey = "cwprod" }
if ($sitecoreSessionInfo.EnvMap.ContainsKey($envKey)) {
    $propertyMasterCsv = $sitecoreSessionInfo.OutputFilePath
    $fileName = Get-ChildItem -Path "$PSScriptRoot\export\$($sitecoreSessionInfo.DateVariables.Year)\$($sitecoreSessionInfo.DateVariables.ShortMonth.ToUpper())" -Filter "*.csv" | Select-Object -First 1 | ForEach-Object { $_.Name }    
    $suitePricingCsv = Join-Path -Path $PSScriptRoot -ChildPath "export" -AdditionalChildPath $sitecoreSessionInfo.DateVariables.Year, $sitecoreSessionInfo.DateVariables.ShortMonth.ToUpper(), ("$fileName")
}
else {
    Write-Host "Environment not recognized. Please check the connection URI." -ForegroundColor Red
    continue
}

$suitePricingCsvData = Import-Csv -Path $suitePricingCsv |
Where-Object { $_.'Marketing Care Level' -ne "" -and $_.'Web Suite Title' -ne "" }

$pricingUpdateResults = Invoke-RemoteScript -Session $session -Verbose -ScriptBlock {
    if (-not (Test-Path $Using:propertyMasterCsv)) {
        Write-Output "Property Master CSV file not found at $($Using:propertyMasterCsv). Exiting script." 
        Write-Output "Run the Property-Master-List-Optimized.ps1 script before executing this script."
        return
    }
    # --- Helper Functions ---
    function Get-SuiteItemName {
        param([string]$suiteTitle)
        $suiteItemName = @{
            "studio"               = "studios"
            "1bed"                 = "1bed"
            "1bed-and-den"         = "1bed-and-den"
            "2bed"                 = "2beds"
            "2beds-and-den"        = "2beds-and-den"
            "3beds"                = "3beds"
            "townhouse-1bed-2beds" = "townhouse-1bed-2beds"
            "deluxe-studio"        = "deluxe-studio"
        }
        $key = $suiteTitle.Trim().ToLower().Replace(' ', '-')
        return $suiteItemName[$key]
    }

    function Update-SuitePromo {
        param(
            $suitePromoItem, $suitePlan, $propertyName, $language, $isFrench = $false
        )
        $langSuffix = if ($isFrench) { ' FR' } else { '' }
        $suitePromoItem.Editing.BeginEdit()
        if ($suitePromoItem.Fields["Regular SuitePrice"].Value -ne $suitePlan."Regular SuitePrice") {
            Write-Output ("Updating Regular Price for {0} {1}{2}" -f $suitePromoItem.Name, $propertyName, $langSuffix)
            $suitePromoItem.Fields["Regular SuitePrice"].Value = $suitePlan."Regular SuitePrice"
        }
        else {
            Write-Output ("Regular Price not updated for {0} {1}{2}" -f $suitePromoItem.Name, $propertyName, $langSuffix)
        }
        $suitePromoItem.Fields["Promotion Price"].Value = $suitePlan.PromotionPrice -replace '[^\d]', ''
        $suitePromoItem.Fields["Start Date"].Value = if (![string]::IsNullOrWhiteSpace($suitePlan.'Start Date')) {
            [datetime]::ParseExact($suitePlan.'Start Date', 'yyyy-MM-dd', $null).ToString("yyyyMMddTHHmmss")
        }
        else { "" }
        $suitePromoItem.Fields["End Date"].Value = if (![string]::IsNullOrWhiteSpace($suitePlan.'End Date')) {
            [datetime]::ParseExact($suitePlan.'End Date', 'yyyy-MM-dd', $null).ToString("yyyyMMddTHHmmss")
        }
        else { "" }
        $suitePromoItem.Editing.EndEdit()
        Write-Output ("{0} {1} - {2}{3} Promo updated for {4}" -f $suitePlan.'Marketing Care Level', $suitePromoItem.Name, $language.ToUpper(), $langSuffix, $propertyName)
    }

    if (-not (Test-Path $Using:propertyMasterCsv)) {
        Write-Error "Property Master CSV file not found at $Using:propertyMasterCsv"
        Write-Information "Run the property master export script before executing this script."
        continue
    }
    $properties = Import-Csv -Path $Using:propertyMasterCsv 
    $suitePricingCsvData = $Using:suitePricingCsvData

    foreach ($property in $properties) {
        $scProperty = Get-Item -Path master: -Id $property.PropertyItemId -Language $property.Language
        $suitePricePromos = $suitePricingCSVData | Where-Object { $_.'Web Property ID - Yardi Number' -eq $property.PropertyId }
        if (-not $suitePricePromos) {
            Write-Output ("{0} No Promos" -f $scProperty.Fields["Property Name"].Value)
            continue
        }
        Write-Output ("Update: {0} Bilingual: {1}" -f $property.PropertyItemName, ([bool]::Parse($property.Bilingual)))
        $suitePlanItems = ($scProperty.Children | Where-Object { $_.Name -eq "Data" }).Children | Where-Object { $_.Name -eq "SuitePlans" }

        foreach ($suitePlan in $suitePricePromos) {
            $suiteTitleKey = $suitePlan.'Web SuiteTitle'
            if ($null -eq $suiteTitleKey -or [string]::IsNullOrWhiteSpace($suiteTitleKey)) {
                Write-Output ("Skipping empty suite item key for {0}" -f $property.PropertyItemName)
                continue
            }
            $suiteItemValue = Get-SuiteItemName $suiteTitleKey
            if (-not $suiteItemValue) {
                Write-Output ("Suite key '{0}' not found in mapping for {1}" -f $suiteTitleKey, $property.PropertyItemName)
                continue
            }
            $updateSuitePromo = $suitePlanItems.Children.Children | Where-Object {
                $_.Name.Trim().Replace(" ", "") -eq $suiteItemValue -and $_.Parent.Name -eq $suitePlan.'Marketing Care Level'.ToLower()
            }
            if (-not $updateSuitePromo) {
                Write-Output ("{0} - {1} does not exist" -f $suitePlan.'Marketing Care Level', $suitePlan.'Web SuiteTitle')
                continue
            }
            Update-SuitePromo -suitePromoItem $updateSuitePromo -suitePlan $suitePlan -propertyName $property.PropertyItemName -language $property.Language

            if ([bool]::Parse($property.Bilingual)) {
                $frVersion = Get-Item -Path master: -Id $updateSuitePromo.ID -Language "fr"
                Update-SuitePromo -suitePromoItem $frVersion -suitePlan $suitePlan -propertyName $property.PropertyItemName -language "fr" -isFrench $true
            }
        }
    }
}-Arguments @{ propertyMasterCsv = $propertyMasterCsv; suitePricingCsvData = $suitePricingCsvData }
$pricingUpdateResults | Format-Table
Stop-ScriptSession -Session $session
Write-Host "Script execution completed successfully." -ForegroundColor Green