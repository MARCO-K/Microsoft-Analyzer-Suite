function Import-SupportlistFromFolder {
    param(
        # Parameter help description
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Container' })]
        [string]$Folder,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateSet("Blacklist", "Whitelist")]
        [string]$Listtype
    )


    $files = Get-ChildItem -Path $Folder -Filter "*.csv"
    if ($files) {
        
        $files | ForEach-Object {
            $file = $_.FullName
            $content = Import-SupportlistFromCSV -listFile $file
            $content

        }
    }
}


