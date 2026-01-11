<#
.SYNOPSIS
    Exports CSV files from Sitecore Media Library to a remote sitecore env packages directory.

.DESCRIPTION
    This script connects to a Sitecore instance using SPE (Sitecore PowerShell Extensions), locates CSV files in a specified media library folder, reads their contents, decodes them from UTF-8 (removing BOM if present), converts them to PowerShell objects, and exports them as CSV files to a local directory. The script ensures file extensions are handled correctly and provides verbose output for each exported file.

.PARAMETER None
    All required parameters are retrieved from the Sitecore session information.

.NOTES
    - Requires SPE module and a valid Sitecore session.
    - The target media library folder is: master:/sitecore/media library/project/sxastarter/sxastarter/files/upload
    - Only media items with the "file" template and "csv" extension are processed.
    - Output files are saved in the directory specified by $sitecoreInfo.PackageDir.

.EXAMPLE
    .\Remote-Upload.ps1
    # Exports all CSV files from the specified Sitecore media library folder to the local package directory.

#>
Clear-Host
Import-Module -Name SPE 
. "$PSScriptRoot\Load-Env.ps1"
. "$PSScriptRoot\Get-SitecoreSession.ps1"
$sitecoreInfo = Get-SitecoreSession

if (-not $sitecoreInfo) { return }

$session = $sitecoreInfo.Session
$packageDir = $sitecoreInfo.PackageDir

$results = Invoke-RemoteScript -Session $session -Verbose -ScriptBlock {
    $filestoUpload = get-item -path "master:/sitecore/media library/project/sxastarter/sxastarter/files/upload"

    # filter media items that are csvs
    $csvitems = get-childitem -path $filestoUpload.paths.fullpath | where-object {
        $_.template.name -eq "file" -and $_.fields["extension"].value -eq "csv"
    }

    foreach ($mediaitem in $csvitems) {
        $filename = $mediaitem.name

        # read the blob stream from media library
        [system.io.stream]$body = $mediaitem.fields["blob"].getblobstream()
        try {
            $contents = new-object byte[] $body.length
            $body.read($contents, 0, $body.length) | out-null
        }
        finally {
            $body.close()
        }

        # decode utf8 content
        $decoded = [system.text.encoding]::utf8.getstring($contents)

        # remove bom manually if present
        if ($decoded.startswith([char]0xfeff)) {
            $decoded = $decoded.substring(1)
        }

        # convert to csv object
        $csv = $decoded | convertfrom-csv -delimiter ","

        # remove extension if already present, then append .csv
        $basename = [system.io.path]::getfilenamewithoutextension($filename)
        # define export path
        $outputfilepath = join-path -path $using:packageDir -childpath "$basename.csv"

        # export
        $csv | export-csv -path $outputfilepath -notypeinformation -encoding utf8 -force

        write-output "✅ exported $filename to $outputfilepath"
    }
}-Arguments @{ packageDir = $packageDir }
$results
Stop-ScriptSession -Session $session