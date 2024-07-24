function New-MSASelfSignedCertificate {
    [cmdletBinding()]
    param (
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags[]]$KeyUsage = 'None',
        [Microsoft.CertificateServices.Commands.CertificateType]$CertificateType = 'Custom',
        [Parameter()][ValidateRange(1, 365)][int]$DurationDays = 365,
        [Parameter()][ValidateSet(2048, 4096)][int]$KeyLength = 2048,
        [Parameter()][ValidateSet('RSA', 'ECDSA')][string]$KeyAlgorithm = 'RSA',
        [Parameter()][ValidateSet('SHA256', '512')][System.Security.Cryptography.HashAlgorithmName]$HashAlgorithm = 'SHA256',
        [System.Security.Cryptography.X509Certificates.X509ContentType]$CertificateFormat = 'Pfx',
        [string]$FriendlyName = 'MSA Self-Signed Certificate',
        [string]$CertStoreLocation = 'Cert:\CurrentUser\My',
        [Parameter()][ValidateSet('Exportable', 'ExportableEncrypted', 'NonExportable')][Microsoft.CertificateServices.Commands.KeyExportPolicy]$KeyExportPolicy = 'Exportable',
        [Parameter()][ValidateSet('None', 'KeyExchange', 'Signature')][Microsoft.CertificateServices.Commands.KeySpec]$KeySpec = 'Signature',
        [Microsoft.CertificateServices.Commands.KeyUsageProperty]$EnhancedKeyUsage = 'All',
        [string]$CertName = "Invictus_IR-App",
        [Parameter()][ValidateScript({ $_ -ge (Get-Date) })][datetime]$NotBefore = (Get-Date),
        [switch]$certOnly
    )

    begin {
        if ( -not(Test-ElevationStatus)) {
            Write-PSFMessage -Level Error -Message "This function requires elevated permissions" -FunctionName $MyInvocation.MyCommand.Name
            break
        }
        
        $NotAfter = (Get-Date).AddDays($DurationDays)

        $certificate = @{
            Subject           = "CN=$CertName"
            FriendlyName      = $FriendlyName
            KeyLength         = $KeyLength
            KeyUsage          = $keyUsage
            KeyExportPolicy   = $KeyExportPolicy
            KeySpec           = $KeySpec
            NotBefore         = $NotBefore
            NotAfter          = $NotAfter
            HashAlgorithm     = $HashAlgorithm
            KeyAlgorithm      = $KeyAlgorithm
            Type              = $CertificateType
            CertStoreLocation = $CertStoreLocation
        }

        $cert_param = Remove-FalsyHashtableEntry $certificate
    }

    process {
        try {
            Write-PSFMessage -Level Verbose -Message "Creating self-signed certificate" -FunctionName $MyInvocation.MyCommand.Name
            $mycert = New-SelfSignedCertificate @cert_param
        } catch {
            Write-PSFMessage -Level Error -Message "Failed to create self-signed certificate" -FunctionName $MyInvocation.MyCommand.Name -Exception $_
        }

        $result = [PSCustomObject][Ordered]@{
            Thumbprint         = $mycert.Thumbprint
            Subject            = $mycert.Subject
            NotBefore          = $mycert.NotBefore
            NotAfter           = $mycert.NotAfter
            FriendlyName       = $mycert.FriendlyName
            HasPrivateKey      = $mycert.HasPrivateKey
            SignatureAlgorithm = $mycert.SignatureAlgorithm.FriendlyName
            KeyAlgorithm       = $mycert.PublicKey.Key.KeyExchangeAlgorithm
            Certificate        = $mycert
        }
    }

    end {
        if ($certOnly) {
            $result.Certificate
        } else {
            $result
        }	
    }
}