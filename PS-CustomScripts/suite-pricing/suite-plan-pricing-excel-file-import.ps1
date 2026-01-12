Clear-Host

$month = (Get-Date).ToString('MMM').Substring(0, 1).ToUpper() + (Get-Date).ToString('MMM').Substring(1).ToLower()
$year = (Get-Date).Year

# Excel file selection
$excelPath = $null
if ($IsWindows) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = "$HOME/repos/rm-local/PS-CustomScripts"
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select the Excel file to import"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $excelPath = $openFileDialog.FileName
            $fileName = Split-Path $excelPath -Leaf
            Write-Host ("Selected file: " + (Split-Path $excelPath -Leaf)) -ForegroundColor Green
            # Copy the selected file to $defaultExcelPath
            $defaultExcelPath = Join-Path $PSScriptRoot "excel\$year\$month\$fileName"
            if ($excelPath -ne $defaultExcelPath) {
                Write-Host "Copying $excelPath to $defaultExcelPath ..."
                $destDir = Split-Path $defaultExcelPath -Parent
                if (-not (Test-Path $destDir)) {
                    Write-Host "Creating directory $destDir ..."
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }
                Copy-Item -Path $excelPath -Destination $defaultExcelPath -Force
                # Prompt for remote SSH details
                $copyRemote = Read-Host "Do you want to copy the file to a remote server via SSH? (y/n)"
                if ($copyRemote -eq 'y') {
                    $copyScript = Join-Path $PSScriptRoot '..\Copy-FileToRemote.ps1'
                    & $copyScript -localFilePath $defaultExcelPath
                }
            }
        }
        else {
            Write-Host "No file selected. Exiting."
            exit 1
        }
    }
    catch {
        Write-Host "Could not load Windows.Forms. Please enter the Excel file path manually."
        $defaultExcelPath = Join-Path $PSScriptRoot "excel\$year\$month\$year Market Rate & Business Case Tracker (C2C & Web).xlsx"
        $excelPath = Read-Host "Enter the full path to the Excel file [$defaultExcelPath]"
        if ([string]::IsNullOrWhiteSpace($excelPath)) { $excelPath = $defaultExcelPath }
    }
}
else {
    Write-Host "File picker is only available on Windows. Please enter the Excel file path manually."
    $defaultExcelPath = Join-Path $PSScriptRoot "excel\$year\$month\$year Market Rate & Business Case Tracker (C2C & Web).xlsx"
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
$localCsvPath = Join-Path -Path $PSScriptRoot -ChildPath "export" -AdditionalChildPath $year, $month, ("$fileName.csv")
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
# Call the monthly update script after export
$monthlyUpdateScript = Join-Path $PSScriptRoot 'suite-plan-pricing-monthly-update.ps1'
if (Test-Path $monthlyUpdateScript) {
    Write-Host "Calling suite-plan-pricing-monthly-update.ps1..."
    & $monthlyUpdateScript
}
else {
    Write-Warning "suite-plan-pricing-monthly-update.ps1 not found at $monthlyUpdateScript"
}

