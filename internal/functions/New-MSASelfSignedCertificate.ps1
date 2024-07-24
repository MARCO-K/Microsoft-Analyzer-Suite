function New-MSASelfSignedCertificate {
    <#
    .SYNOPSIS
    Creates a new self-signed certificate with specified parameters.

    .DESCRIPTION
    The New-MSASelfSignedCertificate function creates a new self-signed certificate
    with customizable parameters such as key usage, duration, algorithm, and more.
    This function requires elevated permissions to run.

    .PARAMETER KeyUsage
    Specifies the key usage flags for the certificate. Default is 'None'.

    .PARAMETER CertificateType
    Specifies the type of certificate to create. Default is 'Custom'.

    .PARAMETER DurationDays
    Specifies the number of days the certificate will be valid. Range: 1-365. Default is 365.

    .PARAMETER KeyLength
    Specifies the length of the key in bits. Valid values: 2048, 4096. Default is 2048.

    .PARAMETER KeyAlgorithm
    Specifies the key algorithm. Valid values: 'RSA', 'ECDSA'. Default is 'RSA'.

    .PARAMETER HashAlgorithm
    Specifies the hash algorithm. Valid values: 'SHA256', 'SHA512'. Default is 'SHA256'.

    .PARAMETER CertificateFormat
    Specifies the certificate format. Default is 'Pfx'.

    .PARAMETER FriendlyName
    Specifies the friendly name for the certificate. Default is 'MSA Self-Signed Certificate'.

    .PARAMETER CertStoreLocation
    Specifies the certificate store location. Default is 'Cert:\CurrentUser\My'.

    .PARAMETER KeyExportPolicy
    Specifies the key export policy. Valid values: 'Exportable', 'ExportableEncrypted', 'NonExportable'. Default is 'Exportable'.

    .PARAMETER KeySpec
    Specifies the key spec. Valid values: 'None', 'KeyExchange', 'Signature'. Default is 'Signature'.

    .PARAMETER CertName
    Specifies the name of the certificate. Default is "Invictus_IR-App".

    .PARAMETER NotBefore
    Specifies the start date of the certificate's validity period. Must be greater than or equal to the current date. Default is the current date.

    .PARAMETER CertOnly
    If specified, returns only the certificate object instead of a custom object with certificate details.

    .EXAMPLE
    New-MSASelfSignedCertificate -CertName "MyApp" -DurationDays 180

    Creates a new self-signed certificate named "MyApp" that is valid for 180 days.

    .EXAMPLE
    New-MSASelfSignedCertificate -KeyLength 4096 -HashAlgorithm SHA512 -CertOnly

    Creates a new self-signed certificate with a 4096-bit key length and SHA512 hash algorithm, returning only the certificate object.

    .OUTPUTS
    PSCustomObject or System.Security.Cryptography.X509Certificates.X509Certificate2

    .NOTES
    This function requires elevated permissions to run.
    Ensure you have the necessary rights before executing this function.
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([PSCustomObject])]
    param (
        [Parameter()]
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags[]]$KeyUsage = 'None',

        [Parameter()]
        [Microsoft.CertificateServices.Commands.CertificateType]$CertificateType = 'Custom',

        [Parameter()]
        [ValidateRange(1, 365)]
        [int]$DurationDays = 365,

        [Parameter()]
        [ValidateSet(2048, 4096)]
        [int]$KeyLength = 2048,

        [Parameter()]
        [ValidateSet('RSA', 'ECDSA')]
        [string]$KeyAlgorithm = 'RSA',

        [Parameter()]
        [ValidateSet('SHA256', 'SHA512')]
        [System.Security.Cryptography.HashAlgorithmName]$HashAlgorithm = 'SHA256',

        [Parameter()]
        [System.Security.Cryptography.X509Certificates.X509ContentType]$CertificateFormat = 'Pfx',

        [Parameter()]
        [string]$FriendlyName = 'MSA Self-Signed Certificate',

        [Parameter()]
        [string]$CertStoreLocation = 'Cert:\CurrentUser\My',

        [Parameter()]
        [ValidateSet('Exportable', 'ExportableEncrypted', 'NonExportable')]
        [Microsoft.CertificateServices.Commands.KeyExportPolicy]$KeyExportPolicy = 'Exportable',

        [Parameter()]
        [ValidateSet('None', 'KeyExchange', 'Signature')]
        [Microsoft.CertificateServices.Commands.KeySpec]$KeySpec = 'Signature',

        [Parameter()]
        [string]$CertName = "Invictus_IR-App",

        [Parameter()]
        [ValidateScript({ $_ -ge (Get-Date) })]
        [datetime]$NotBefore = (Get-Date),

        [Parameter()]
        [switch]$CertOnly
    )

    begin {
        if (-not (Test-ElevationStatus)) {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    "This function requires elevated permissions",
                    "ElevatedPermissionsRequired",
                    [System.Management.Automation.ErrorCategory]::PermissionDenied,
                    $null
                )
            )
        }
        
        $NotAfter = $NotBefore.AddDays($DurationDays)

        $certificateParams = @{
            Subject           = "CN=$CertName"
            FriendlyName      = $FriendlyName
            KeyLength         = $KeyLength
            KeyUsage          = $KeyUsage
            KeyExportPolicy   = $KeyExportPolicy
            KeySpec           = $KeySpec
            NotBefore         = $NotBefore
            NotAfter          = $NotAfter
            HashAlgorithm     = $HashAlgorithm
            KeyAlgorithm      = $KeyAlgorithm
            Type              = $CertificateType
            CertStoreLocation = $CertStoreLocation
        }

        # Assuming Remove-FalsyHashtableEntry is a custom function
        $certificateParams = Remove-FalsyHashtableEntry $certificateParams
    }

    process {
        try {
            if ($PSCmdlet.ShouldProcess($($certificateParams.CertStoreLocation) , "Creating self-signed certificate")) {
                Write-PSFMessage -Level Verbose -Message "Creating self-signed certificate" -FunctionName $MyInvocation.MyCommand.Name
                $newCertificate = New-SelfSignedCertificate @certificateParams
            }
        } catch {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    $_.Exception.Message,
                    "CertificateCreationFailed",
                    [System.Management.Automation.ErrorCategory]::OperationStopped,
                    $null
                )
            )
        }

        $result = [PSCustomObject]@{
            Thumbprint         = $newCertificate.Thumbprint
            Subject            = $newCertificate.Subject
            NotBefore          = $newCertificate.NotBefore
            NotAfter           = $newCertificate.NotAfter
            FriendlyName       = $newCertificate.FriendlyName
            HasPrivateKey      = $newCertificate.HasPrivateKey
            SignatureAlgorithm = $newCertificate.SignatureAlgorithm.FriendlyName
            KeyAlgorithm       = $newCertificate.PublicKey.Key.KeyExchangeAlgorithm
            Certificate        = $newCertificate
        }

        if ($CertOnly) {
            $result.Certificate
        } else {
            $result
        }
    }
}