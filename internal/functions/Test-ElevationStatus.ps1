function Test-ElevationStatus {
    <#
      .SYNOPSIS
      Check if Powershell Console is running in elevated mode.

      .DESCRIPTION
      This function is checking if Powershell Console is running in elevated mode(as administrator).

      .EXAMPLE
      Test-ElevationStatus

      .PARAMETER detailed
      The parameter will switch to a detailed output version.
    #>
    [cmdletBinding()]

    param
    (
        [switch]
        $detailed
    )

    process {

        Write-PSFMessage -Level Verbose -Message "Checking if Powershell Console is running in elevated mode" -FunctionName $MyInvocation.MyCommand.Name
        $IsElevated =     
        if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
        { $true }      
        else
        { $false }   
    }
    end {
        if ($detailed) {       
            Write-PSFMessage -Level 'Verbose' -Message 'Returning detailed info about Powershell Console' -FunctionName $MyInvocation.MyCommand.Name
            $Props = [ordered]@{
                IsElevated       = $IsElevated
                Name             = $($Host.Name)
                OS               = $($PSVersionTable.OS)
                Platform         = $($PSVersionTable.Platform)
                PSEdition        = $($PSVersionTable.PSEdition)
                PSVersion        = $($Host.Version)
                CurrentCulture   = $($Host.CurrentCulture)
                CurrentUICulture = $($Host.CurrentUICulture)
                DebuggerEnabled  = $($Host.DebuggerEnabled)
                IsRunspacePushed = $($Host.IsRunspacePushed)
            }
            $Props
        } else {
            Write-PSFMessage -Level 'Verbose' -Message 'Returning only IsElevated status' -FunctionName $MyInvocation.MyCommand.Name
            $IsElevated
        }
    }
}