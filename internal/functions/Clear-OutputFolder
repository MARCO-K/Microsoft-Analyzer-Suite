function Clear-OutputFolder {
    <#
    .SYNOPSIS
    Clears the contents of a specified output folder.

    .DESCRIPTION
    The Clear-OutputFolder function removes all items (files and subdirectories) from a specified folder. If the folder doesn't exist, it creates it.

    .PARAMETER OUTPUT_FOLDER
    Specifies the path of the folder to be cleared. This parameter is mandatory and can be passed through the pipeline.

    .EXAMPLE
    Clear-OutputFolder -OUTPUT_FOLDER "C:\Temp\OutputFolder"
    Clears all contents of the "C:\Temp\OutputFolder" directory.

    .EXAMPLE
    "C:\Temp\OutputFolder" | Clear-OutputFolder
    Clears all contents of the "C:\Temp\OutputFolder" directory using pipeline input.

    .NOTES
    This function uses the Write-PSFMessage cmdlet for logging, which requires the PSFramework module.

    .INPUTS
    System.IO.FileInfo

    .OUTPUTS
    System.Boolean
    Returns $true upon completion.

    .LINK

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][String]$OUTPUT_FOLDER
    )
    
    begin {
        if ( -Not (Test-Path $OUTPUT_FOLDER -PathType Container) ) {
            Write-PSFMessage -Level 'Verbose' -Message 'Folder does not exist' -FunctionName $MyInvocation.MyCommand.Name
            $OUTPUT_FOLDER = New-Item -Path $OUTPUT_FOLDER -ItemType 'Directory' -Force
            Write-PSFMessage -Level 'Verbose' -Message "Folder $output_Folder created" -FunctionName $MyInvocation.MyCommand.Name
        }
    }
    
    process {
        Write-PSFMessage -Level 'Verbose' -Message "Deleting all files & Folders under: $output_Folder" -FunctionName $MyInvocation.MyCommand.Name
        $files = (Get-ChildItem -Path $OUTPUT_FOLDER -Recurse).FullName 
        if ($files) { 
            $files | ForEach-Object { Remove-Item $_ -Force -Recurse -ErrorAction SilentlyContinue
                Write-PSFMessage -Level 'Verbose' -Message "File deleted: $_" -FunctionName $MyInvocation.MyCommand.Name
            }
        }
    }
    
    end {
        Write-PSFMessage -Level 'Verbose' -Message "Output folder: $output_Folder" -FunctionName $MyInvocation.MyCommand.Name
        $OUTPUT_FOLDER
    }
}