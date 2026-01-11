<#
.SYNOPSIS
    Imports data from a specified Excel file, processes it, and exports the result as a CSV file with a structured directory output.

.DESCRIPTION
    This script allows the user to select or specify an Excel file containing market rate and business case data.
    It processes the data from a specific worksheet, converts date columns to ISO 8601 format, and exports the cleaned data to a CSV file.
    The output CSV is saved in a directory structure organized by year and month.
    The script ensures required modules are installed, handles file and directory creation, and skips export if the output file already exists.

.PARAMETER excelPath
    The path to the Excel file to import. Can be selected via a file picker on Windows or entered manually.

.NOTES
    - Requires the ImportExcel PowerShell module.
    - Designed for use on Windows, but supports manual file path entry on other platforms.
    - Excludes certain columns ('Customer Promotional Offers', 'Notes:', 'Region') from the output.
    - Only exports rows where 'PropertyID - Yardi Number' is not null or "TBD".

.EXAMPLE
    # Run the script and follow prompts to select or enter the Excel file path.
    .\suite-plan-pricing-excel-file-import.ps1

#>
Clear-Host

$month = (Get-Date).ToString('MMM').ToUpper()
$year = (Get-Date).Year

# Excel file selection
$excelPath = $null
if ($IsWindows) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = "$HOME/repos/git-repos\personal-git-repo\SPE-Scripts\PS-CustomScripts"
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select the Excel file to import"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $excelPath = $openFileDialog.FileName
        }
        else {
            Write-Host "No file selected. Exiting."
            exit 1
        }
    }
    catch {
        Write-Host "Could not load Windows.Forms. Please enter the Excel file path manually."
        $defaultExcelPath = Join-Path $PSScriptRoot "excel\$year Market Rate & Business Case Tracker (C2C & Web).xlsx"
        $excelPath = Read-Host "Enter the full path to the Excel file [$defaultExcelPath]"
        if ([string]::IsNullOrWhiteSpace($excelPath)) { $excelPath = $defaultExcelPath }
    }
}
else {
    Write-Host "File picker is only available on Windows. Please enter the Excel file path manually."
    $defaultExcelPath = Join-Path $PSScriptRoot "excel\$year Market Rate & Business Case Tracker (C2C & Web).xlsx"
    $excelPath = Read-Host "Enter the full path to the Excel file [$defaultExcelPath]"
    if ([string]::IsNullOrWhiteSpace($excelPath)) { $excelPath = $defaultExcelPath }
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host 'ImportExcel module not found. Installing...'
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

if (-not (Test-Path $excelPath)) {
    Write-Error "Excel file not found at $excelPath"
    exit 1
}

# File name for output
$fileName = ([System.IO.Path]::GetFileNameWithoutExtension($excelPath) -replace '[^a-zA-Z0-9]+', '-' -replace '-+', '-').ToLower().TrimEnd('-')
$localCsvPath = Join-Path -Path $PSScriptRoot -ChildPath "output" -AdditionalChildPath $year, $month, ("$fileName.csv")
if (-not (Test-Path $(Split-Path $localCsvPath -Parent))) {
    Write-Host "Creating directory structure at $(Split-Path $localCsvPath -Parent)"
    New-Item -ItemType Directory -Path (Split-Path $localCsvPath -Parent) -Force | Out-Null
}
else {
    Write-Host "Directory structure already already exists at $(Split-Path $localCsvPath -Parent) - Skipping."
}

if (Test-path $localCsvPath) {
    Write-Host "$(Split-Path $localCsvPath -Leaf) file already exists at $(Split-Path $localCsvPath -Parent) - Skipping export."
}
else {
    # Import Excel and process data
    $data = Import-Excel -Path $excelPath -HeaderRow 9 -WorksheetName "Market Rates and Business Cases" -ErrorAction Stop

    # Convert StartDate and EndDate columns to ISO 8601 string format before CSV conversion
    $convertedData = foreach ($row in $data) {
        $obj = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) {
            $name = $property.Name
            $value = $property.Value
            if ($name -in @('Customer Promotional Offers', 'Notes:', 'Region')) { continue }
            function Convert-DateValue($value) {
                if ($value -is [datetime]) { return $value.ToString('yyyy-MM-dd') }
                elseif ($value -is [double]) { return [datetime]::FromOADate($value).ToString('yyyy-MM-dd') }
                elseif ($null -ne $value) { return $value }
                else { return "" }
            }
            if ($name -in @('Start Date', 'End Date')) {
                $obj[$name] = Convert-DateValue $value
            }
            else {
                $obj[$name] = $value #if ($null -ne $value) { $value } else { "" }
            }
        }
        [PSCustomObject]$obj
    }

    Write-Host ("Excel file converted/exported to " + $localCsvPath)
    $convertedData | Where-Object { $_.'PropertyID - Yardi Number' -ne $null -and $_.'PropertyID - Yardi Number' -ne "TBD" } | Export-Csv -Path $localCsvPath -NoTypeInformation -Encoding UTF8 -Force
}

