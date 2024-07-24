function ConvertTo-Iso8601DateTime {
    <#
    .SYNOPSIS
    Converts a DateTime object to an ISO 8601 compliant UTC DateTime or string.
    
    .DESCRIPTION
    This function takes a DateTime object and converts it to either:
    1. A UTC DateTime object that complies with the ISO 8601 standard.
    2. An ISO 8601 formatted string representation of the UTC DateTime.
    The output type is controlled by the -AsString switch parameter.
    
    .PARAMETER InputObject
    The DateTime object to convert. This parameter is mandatory and can be
    passed through the pipeline.
    
    .PARAMETER AsString
    Switch parameter to return the result as an ISO 8601 formatted string
    instead of a DateTime object.
    
    .INPUTS
    System.DateTime
    
    .OUTPUTS
    System.DateTime or System.String
    
    .EXAMPLE
    Get-Date | ConvertTo-Iso8601DateTime
    Returns the current date and time as a UTC DateTime object.

    .EXAMPLE
    ConvertTo-Iso8601DateTime -InputObject (Get-Date).AddDays(-1) -AsString
    Returns yesterday's date and time as an ISO 8601 formatted string.
    
    .NOTES
    The function always converts the input to UTC.
    #>
    [CmdletBinding()]
    [OutputType([DateTime], [string])]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNull()]
        [DateTime]$InputObject,

        [Parameter()]
        [switch]$AsString
    )

    process {
        Write-PSFMessage -Level Verbose -Message 'Converting DateTime to ISO 8601 compliant UTC DateTime' -FunctionName $MyInvocation.MyCommand.Name

        try {
            $utcDateTime = $InputObject.ToUniversalTime()
            
            # Ensure the Kind property is set to UTC
            $iso8601DateTime = [DateTime]::SpecifyKind($utcDateTime, [DateTimeKind]::Utc)
            
            if ($AsString) {
                $result = $iso8601DateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffK", [System.Globalization.CultureInfo]::InvariantCulture)
                Write-PSFMessage -Level Verbose -Message "Converted $InputObject to ISO 8601 string: $result" -FunctionName $MyInvocation.MyCommand.Name
            } else {
                $result = $iso8601DateTime
                Write-PSFMessage -Level Verbose -Message "Converted $InputObject to UTC DateTime: $result" -FunctionName $MyInvocation.MyCommand.Name
            }
            
            return $result
        } catch {
            Write-PSFMessage -Level Error -Message "Failed to convert DateTime to ISO 8601 format" -FunctionName $MyInvocation.MyCommand.Name -Exception $_
            throw
        }
    }
}