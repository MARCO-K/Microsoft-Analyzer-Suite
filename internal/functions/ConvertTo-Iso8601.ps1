function ConvertTo-Iso8601 {
    <#
    .SYNOPSIS
    Convert datetime to ISO 8601 format
    
    .DESCRIPTION
    Convert datetime to  ISO 8601 format
    
    .PARAMETER InputObject
    DateTime object
    
    .INPUTS
    InputObject
    
    .OUTPUTS
    System.String
    
    .EXAMPLE
    (get-date) | ConvertTo-Iso8601
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [DateTime]$InputObject
    )

    begin {
    }

    process {
        Write-PSFMessage -Level 'Verbose' -Message 'Converting datetime to ISO 8601 format' -FunctionName $MyInvocation.MyCommand.Name -ModuleName $MyInvocation.MyCommand.ModuleName
        $result = $InputObject.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
    }
    
    end {
        $result
    }
}