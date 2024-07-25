function Test-IPinfo {
    <#
    .SYNOPSIS
    Tests if IPinfo CLI is installed and optionally installs it.

    .DESCRIPTION
    The Test-IPinfo function checks if the IPinfo CLI tool is installed in the expected location.
    If the -Install switch is used and IPinfo is not found, it will attempt to install it.

    .PARAMETER Install
    If specified and IPinfo is not found, the function will attempt to install it.

    .EXAMPLE
    Test-IPinfo

    Checks if IPinfo is installed and returns $true if it is, $false otherwise.

    .EXAMPLE
    Test-IPinfo -Install

    Checks if IPinfo is installed. If not, it attempts to install it and returns $true if successful.

    .OUTPUTS
    System.Boolean
    Returns $true if IPinfo is installed (or successfully installed), $false otherwise.

    .NOTES
    This function requires the PSFramework module for logging.
    It also depends on the Install-IPinfo function when the -Install switch is used.

    .LINK
    https://github.com/ipinfo/cli
    #>

    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [switch]$Install
    )

    begin {
        $InstallPath = "$env:LOCALAPPDATA\ipinfo"
        $IPinfoExe = Join-Path -Path $InstallPath -ChildPath "ipinfo.exe"
        Write-PSFMessage -Level Verbose -Message "Checking for IPinfo at: $IPinfoExe" -FunctionName $MyInvocation.MyCommand.Name
    }

    process {
        if (Test-Path -Path $IPinfoExe -PathType Leaf) {
            Write-PSFMessage -Level Verbose -Message "IPinfo is installed at: $IPinfoExe" -FunctionName $MyInvocation.MyCommand.Name
            return $true
        } elseif ($Install) {
            Write-PSFMessage -Level Verbose -Message "IPinfo not found. Attempting installation..." -FunctionName $MyInvocation.MyCommand.Name
            try {
                Install-IPinfo -ErrorAction Stop
                if (Test-Path -Path $IPinfoExe -PathType Leaf) {
                    Write-PSFMessage -Level Verbose -Message "IPinfo successfully installed at: $IPinfoExe" -FunctionName $MyInvocation.MyCommand.Name
                    return $true
                } else {
                    throw "Installation completed but IPinfo executable not found at expected location."
                }
            } catch {
                Write-PSFMessage -Level Error -Message "Failed to install IPinfo: $_" -FunctionName $MyInvocation.MyCommand.Name
                return $false
            }
        } else {
            Write-PSFMessage -Level Warning -Message "IPinfo is not installed at: $IPinfoExe" -FunctionName $MyInvocation.MyCommand.Name
            return $false
        }
    }
}