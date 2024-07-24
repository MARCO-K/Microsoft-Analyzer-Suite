function Import-SupportlistFromFolder {
    <#
    .SYNOPSIS
    Imports support lists from CSV files in a specified folder.

    .DESCRIPTION
    The Import-SupportlistFromFolder function reads CSV files from a specified folder and its subfolders,
    importing them as support lists (either whitelist or blacklist). Each file is processed and its content
    is stored in a custom object along with metadata such as the list name and type.

    .PARAMETER Folder
    The path to the folder containing the CSV files to import. This parameter is mandatory and accepts
    pipeline input.

    .PARAMETER Filter
    Specifies the file filter to use when searching for CSV files. The default is "*.csv".

    .PARAMETER ListType
    Specifies whether the imported lists should be categorized as 'Whitelist' or 'Blacklist'. This parameter
    is mandatory.

    .EXAMPLE
    Import-SupportlistFromFolder -Folder "C:\SupportLists" -ListType Whitelist

    This example imports all CSV files from the "C:\SupportLists" folder and its subfolders as whitelists.

    .EXAMPLE
    "C:\SupportLists" | Import-SupportlistFromFolder -Filter "support-*.csv" -ListType Blacklist

    This example uses pipeline input to specify the folder and imports only CSV files that start with "support-"
    as blacklists.

    .INPUTS
    System.String
    You can pipe a string that contains the path to the folder to Import-SupportlistFromFolder.

    .OUTPUTS
    System.Management.Automation.PSCustomObject[]
    Returns an array of custom objects, each representing an imported support list.

    .NOTES
    This function requires the PSFramework module for advanced logging capabilities.
    Ensure you have this module installed before using the function.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript({
                if (-not (Test-Path $_ -PathType 'Container')) {
                    throw "The path '$_' is not a valid directory."
                }
                $true
            })]
        [string]$Folder,

        [Parameter()]
        [string]$Filter = "*.csv",

        [Parameter(Mandatory = $true)]
        [ValidateSet('Whitelist', 'Blacklist')]
        [string]$ListType
    )

    begin {
        Write-PSFMessage -Level Verbose -Message "Starting import from folder: $Folder"
        $supportLists = @()
    }

    process {
        $files = Get-ChildItem -Path $Folder -Filter $Filter -Recurse -File

        if ($files) {
            Write-PSFMessage -Level Verbose -Message "Found $($files.Count) files matching the filter"
            
            foreach ($file in $files) {
                Write-PSFMessage -Level Verbose -Message "Processing file: $($file.FullName)"
                
                try {
                    $content = Import-Csv -Path $file.FullName -ErrorAction Stop
                    
                    $supportLists += [PSCustomObject]@{
                        Name     = $file.BaseName.Split("-")[0]
                        ListName = $file.BaseName
                        Type     = $ListType
                        Content  = $content
                    }
                    
                    Write-PSFMessage -Level Verbose -Message "Successfully imported $($content.Count) entries from $($file.Name)"
                } catch {
                    Write-PSFMessage -Level Error -Message "Failed to import file $($file.Name)" -Exception $_

                }
            }
        } else {
            Write-PSFMessage -Level Error -Message "No files found matching the filter in the specified folder"
        }
    }

    end {
        Write-PSFMessage -Level Verbose -Message "Import complete. Total lists imported: $($supportLists.Count)"
        $supportLists
    }
}