function Import-SupportlistCSV {

    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]
        $blacklistFile
    )
    begin {
        $File = Get-Item -Path $blacklistFile -Filter '*.csv'
    }

    process {
        if ($File) {
            $table = Import-Csv -Path $File.FullName
        }
    }
    

    end {
        if ($table.count -gt 0) {
            $table
        } else {
            Write-Host "No data found in the file"
            $false
        }
    }

}