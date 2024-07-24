function Clear-OutputFolder {
    <#
    .SYNOPSIS
    Clears the contents of a specified output folder.

    .DESCRIPTION
    The Clear-OutputFolder function removes all items (files and subdirectories) from a specified folder.
    If the folder doesn't exist, it creates it. This function uses PSFramework for logging.

    .PARAMETER Path
    Specifies the path of the folder to be cleared. This parameter is mandatory and can be passed through the pipeline.

    .EXAMPLE
    Clear-OutputFolder -Path "C:\Temp\OutputFolder"
    Clears all contents of the "C:\Temp\OutputFolder" directory.

    .EXAMPLE
    "C:\Temp\OutputFolder" | Clear-OutputFolder
    Clears all contents of the "C:\Temp\OutputFolder" directory using pipeline input.

    .NOTES
    This function uses the Write-PSFMessage cmdlet for logging, which requires the PSFramework module.

    .INPUTS
    System.String

    .OUTPUTS
    System.IO.DirectoryInfo
    Returns the DirectoryInfo object of the cleared folder.

    .LINK
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
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
            } else {
                $directory = Get-Item -Path $fullPath -Force
            }

            if ($PSCmdlet.ShouldProcess($fullPath, "Clear folder contents")) {
                Get-ChildItem -Path $fullPath -Force | ForEach-Object {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                    Write-PSFMessage -Level Verbose -Message "Removed: $($_.FullName)" -FunctionName $MyInvocation.MyCommand.Name
                }
                Write-PSFMessage -Level Verbose -Message "Cleared all contents of: $fullPath" -FunctionName $MyInvocation.MyCommand.Name
            }

            return $directory
        } catch {
            Write-PSFMessage -Level Error -Message "Failed to clear directory: $fullPath" -FunctionName $MyInvocation.MyCommand.Name -Exception $_
            Write-PSFMessage -Level Verbose -Message $_.Exception.Message -FunctionName $MyInvocation.MyCommand.Name
            throw
        }
    }
}