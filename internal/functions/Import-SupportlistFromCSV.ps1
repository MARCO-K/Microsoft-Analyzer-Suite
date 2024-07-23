function Import-SupportlistFromCSV {

    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]
        $listFile
    )
    begin {
        $File = Get-Item -Path $listFile -Filter '*.csv'
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