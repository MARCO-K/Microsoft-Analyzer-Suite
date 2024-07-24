function Test-ElevationStatus {
    <#
    .SYNOPSIS
    Checks if the PowerShell console is running in elevated mode.

    .DESCRIPTION
    This function determines if the PowerShell console is running with elevated privileges (as administrator).
    It can provide either a simple boolean result or detailed information about the PowerShell environment.

    .PARAMETER Detailed
    Switch parameter to return detailed information about the PowerShell environment.

    .EXAMPLE
    Test-ElevationStatus
    Returns a boolean indicating whether the console is elevated.

    .EXAMPLE
    Test-ElevationStatus -Detailed
    Returns a hashtable with detailed information about the PowerShell environment, including elevation status.

    .OUTPUTS
    System.Boolean or System.Collections.Specialized.OrderedDictionary
    Returns a boolean when -Detailed is not used, or an OrderedDictionary when -Detailed is specified.
    #>
    [CmdletBinding()]
    [OutputType([bool], [System.Collections.Specialized.OrderedDictionary])]
    param (
        [switch]$Detailed
    )

    begin {
        Write-PSFMessage -Level Verbose -Message "Checking if PowerShell console is running in elevated mode" -FunctionName $MyInvocation.MyCommand.Name
    }

    process {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
        $isElevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    end {
        if ($Detailed) {
            Write-PSFMessage -Level Verbose -Message 'Returning detailed info about PowerShell console' -FunctionName $MyInvocation.MyCommand.Name
            [ordered]@{
                IsElevated       = $isElevated
                Name             = $Host.Name
                OS               = $PSVersionTable.OS
                Platform         = $PSVersionTable.Platform
                PSEdition        = $PSVersionTable.PSEdition
                PSVersion        = $PSVersionTable.PSVersion
                CurrentCulture   = $Host.CurrentCulture
                CurrentUICulture = $Host.CurrentUICulture
                DebuggerEnabled  = $Host.DebuggerEnabled
                IsRunspacePushed = $Host.IsRunspacePushed
            }
        } else {
            Write-PSFMessage -Level Verbose -Message 'Returning only IsElevated status' -FunctionName $MyInvocation.MyCommand.Name
            $isElevated
        }
    }
}