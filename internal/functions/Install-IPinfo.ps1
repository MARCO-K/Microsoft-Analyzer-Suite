function Install-IPinfo {
    <#
    .SYNOPSIS
    Installs the IPinfo CLI tool on a Windows system.

    .DESCRIPTION
    The Install-IPinfo function downloads, extracts, and installs the IPinfo CLI tool
    for Windows. It also adds the installation directory to the system PATH.

    .PARAMETER Version
    The version of the IPinfo CLI tool to install. Defaults to "3.3.1".

    .EXAMPLE
    Install-IPinfo

    Installs the default version (3.3.1) of the IPinfo CLI tool.

    .EXAMPLE
    Install-IPinfo -Version "3.4.0"

    Installs version 3.4.0 of the IPinfo CLI tool.

    .NOTES
    This function requires administrative privileges to modify the system PATH.
    It uses the PSFramework module for logging.

    .LINK
    https://github.com/ipinfo/cli
    #>

    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [ValidatePattern('^\d+\.\d+\.\d+$')]
        [string]$Version = "3.3.1"
    )

    begin {
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'

        $FileName = "ipinfo_$($Version)_windows_amd64"
        $ZipFileName = "$FileName.zip"
        $TempPath = [System.IO.Path]::GetTempPath()
        $OutFile = Join-Path -Path $TempPath -ChildPath $ZipFileName
        $InstallPath = "$env:LOCALAPPDATA\ipinfo"

        Write-PSFMessage -Level Verbose -Message "Starting IPinfo CLI installation. Version: $Version" -FunctionName $MyInvocation.MyCommand.Name
    }

    process {
        try {
            # Download and extract zip
            $DownloadUrl = "https://github.com/ipinfo/cli/releases/download/ipinfo-$Version/$ZipFileName"
            Write-PSFMessage -Level Verbose -Message "Downloading IPinfo CLI from: $DownloadUrl" -FunctionName $MyInvocation.MyCommand.Name
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $OutFile -UseBasicParsing

            Write-PSFMessage -Level Verbose -Message "Unblocking downloaded file" -FunctionName $MyInvocation.MyCommand.Name
            Unblock-File $OutFile

            Write-PSFMessage -Level Verbose -Message "Extracting IPinfo CLI to: $InstallPath" -FunctionName $MyInvocation.MyCommand.Name
            Expand-Archive -Path $OutFile -DestinationPath $InstallPath -Force

            # Rename executable
            $OldPath = Join-Path -Path $InstallPath -ChildPath "$FileName.exe"
            $NewPath = Join-Path -Path $InstallPath -ChildPath "ipinfo.exe"
            if (Test-Path $NewPath) {
                Write-PSFMessage -Level Verbose -Message "Removing existing IPinfo executable" -FunctionName $MyInvocation.MyCommand.Name
                Remove-Item $NewPath -Force
            }
            Write-PSFMessage -Level Verbose -Message "Renaming IPinfo executable" -FunctionName $MyInvocation.MyCommand.Name
            Rename-Item -Path $OldPath -NewName "ipinfo.exe" -Force

            # Update PATH
            $UserPath = [Environment]::GetEnvironmentVariable("PATH", "User")
            if ($UserPath -notlike "*$InstallPath*") {
                Write-PSFMessage -Level Verbose -Message "Adding IPinfo path to User PATH" -FunctionName $MyInvocation.MyCommand.Name
                $NewUserPath = $UserPath + ";$InstallPath"
                [Environment]::SetEnvironmentVariable("PATH", $NewUserPath, "User")
                $env:PATH += ";$InstallPath"
            } else {
                Write-PSFMessage -Level Verbose -Message "IPinfo path already in User PATH" -FunctionName $MyInvocation.MyCommand.Name
            }
        } catch {
            Write-PSFMessage -Level Error -Message "Error installing IPinfo CLI: $_" -FunctionName $MyInvocation.MyCommand.Name
            throw $_
        }
    }

    end {
        # Clean up
        if (Test-Path $OutFile) {
            Write-PSFMessage -Level Verbose -Message "Removing temporary files" -FunctionName $MyInvocation.MyCommand.Name
            Remove-Item -Path $OutFile -Force
        }
        Write-PSFMessage -Level Verbose -Message "IPinfo CLI installation complete. You can now use 'ipinfo' from the command line." -FunctionName $MyInvocation.MyCommand.Name
    }
}