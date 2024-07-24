function Get-OutputFolder {
    <#
    .SYNOPSIS
    Creates or verifies the existence of a specified output folder.

    .DESCRIPTION
    The Get-OutputFolder function takes a folder path as input, checks if it exists,
    and creates it if it doesn't. It uses PSFramework for logging and error handling.

    .PARAMETER Path
    The path of the folder to create or verify. This parameter is mandatory and can be
    passed through the pipeline.

    .EXAMPLE
    Get-OutputFolder -Path "C:\Temp\MyOutputFolder"
    Creates the folder "C:\Temp\MyOutputFolder" if it doesn't exist, or returns the existing folder.

    .EXAMPLE
    "C:\Temp\AnotherFolder" | Get-OutputFolder
    Creates the folder "C:\Temp\AnotherFolder" if it doesn't exist, or returns the existing folder.

    .OUTPUTS
    System.IO.DirectoryInfo
    Returns the DirectoryInfo object of the created or existing folder.

    .NOTES
    This function requires the PSFramework module for logging.
    It uses Write-PSFMessage for verbose output and error logging.
    #>
    
    [CmdletBinding()]
    [OutputType([System.IO.DirectoryInfo])]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    process {
        try {
            $fullPath = [System.IO.Path]::GetFullPath($Path)
            Write-PSFMessage -Level Verbose -Message "Processing path: $fullPath" -FunctionName $MyInvocation.MyCommand.Name

            if (-not (Test-Path -Path $fullPath -PathType Container)) {
                Write-PSFMessage -Level Verbose -Message "Creating directory: $fullPath" -FunctionName $MyInvocation.MyCommand.Name
                $directory = New-Item -Path $fullPath -ItemType Directory -Force -ErrorAction Stop
                Write-PSFMessage -Level Verbose -Message "Directory created successfully: $fullPath" -FunctionName $MyInvocation.MyCommand.Name
            } else {
                Write-PSFMessage -Level Verbose -Message "Directory already exists: $fullPath" -FunctionName $MyInvocation.MyCommand.Name
                $directory = Get-Item -Path $fullPath -Force
            }

            return $directory
        } catch {
            Write-PSFMessage -Level Error -Message "Failed to process directory: $fullPath" -FunctionName $MyInvocation.MyCommand.Name -Exception $_
            Write-PSFMessage -Level Verbose -Message $_.Exception.Message -FunctionName $MyInvocation.MyCommand.Name
            throw
        }
    }
}