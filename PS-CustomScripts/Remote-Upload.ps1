# get the folder containing csv files in the media library
# https://xmc-chartwellma83cf-chartwellsxa-cwprod.sitecorecloud.io/
# https://xmc-chartwellmabc8a-chartwellsxa-cwqa.sitecorecloud.io/
# https://xmc-chartwellmaa139-chartwellsxa-dev.sitecorecloud.io/
Clear-Host
Import-Module -Name SPE 
$session = New-ScriptSession -ConnectionUri "https://xmc-chartwellmaa139-chartwellsxa-dev.sitecorecloud.io/" -Username "sitecore\speremoting" -SharedSecret "A345256D29924333A975FC96AFC46DE87B8F0F1B85705283B4F81AD502C3A50A"
$results = Invoke-RemoteScript -Session $session -Verbose -ScriptBlock {
    $backupfolder = get-item -path "master:/sitecore/media library/project/sxastarter/sxastarter/files/test"

    # filter media items that are csvs
    $csvitems = get-childitem -path $backupfolder.paths.fullpath | where-object {
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
        $outputfilepath = join-path -path "c:\inetpub\wwwroot\app_data\packages" -childpath "$basename.csv"

        # export
        $csv | export-csv -path $outputfilepath -notypeinformation -encoding utf8 -force

        write-output "✅ exported $filename to $outputfilepath"
    }
}
$results
Stop-ScriptSession -Session $session