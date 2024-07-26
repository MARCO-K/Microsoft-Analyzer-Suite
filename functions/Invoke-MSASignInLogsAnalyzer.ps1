﻿function Invoke-MSASigninLogs {
    <#
.SYNOPSIS
  ADSignInLogsGraph-Analyzer v0.1 - Automated Processing of Microsoft Entra ID Sign-In Logs for DFIR

.DESCRIPTION
  ADSignInLogsGraph-Analyzer.ps1 is a PowerShell script utilized to simplify the analysis of Microsoft Entra ID Sign-In Logs extracted via "Microsoft Extractor Suite" by Invictus Incident Response.

  https://github.com/invictus-ir/Microsoft-Extractor-Suite (Microsoft-Extractor-Suite v1.3.2)

  https://microsoft-365-extractor-suite.readthedocs.io/en/latest/functionality/AzureSignInLogsGraph.html

.EXAMPLE
  PS> .\ADSignInLogsGraph-Analyzer.ps1

.NOTES
  Author - Martin Willing

.LINK
  https://lethal-forensics.com/
#>

    #############################################################################################################################################################################################
    #############################################################################################################################################################################################

    # How long does Microsoft Entra ID store the Sign-ins data?

    # Microsoft Entra ID Free      7 days
    # Microsoft Entra ID P1       30 days
    # Microsoft Entra ID P2       30 days

    # Note: You must have a Microsoft Entra ID P1 or P2 license to download sign-in logs using the Microsoft Graph API.

    #############################################################################################################################################################################################
    #############################################################################################################################################################################################



    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript({ if (Test-Path $_) { $true } { throw "File not found: $_" } } )]
        [Object]$inputobject,
        [Parameter(Mandatory = $true)]
        [string]$output,
        $Blacklist,
        $Whitelist,
        ## TODO: paramtersets
        [switch]$IPinfo,
        $Token,
        [Parameter()]
        [ValidateSet('CSV', 'XLSX', 'JSON')]	
        [string]$outputFormat,
        [switch]$cleanup,
        [switch]$Transcript

    )

    #region begin
    begin {
        # Set Progress Preference to Silently Continue
        $OriginalProgressPreference = $Global:ProgressPreference
        $Global:ProgressPreference = 'SilentlyContinue'

        # Check if the PowerShell script is being run with admin rights
        if ( -not(Test-ElevationStatus)) {
            Write-PSFMessage -Level Error -Message "Script is not running with admin rights. Restarting script with admin rights." -FunctionName $MyInvocation.MyCommand.Name
            Exit
        }

        # open file and process JSON
        Write-PSFMessage -Level Verbose -Message "Starting JSON file processing" -FunctionName $MyInvocation.MyCommand.Name
        $result = $InputObject | Open-JsonFile
        if ($null -ne $result) {
            Write-PSFMessage -Level Verbose -Message "JSON content successfully retrieved" -FunctionName $MyInvocation.MyCommand.Name
            $Data = $result | Sort-Object { $_.CreatedDateTime -as [datetime] } -Descending

        } else {
            Write-PSFMessage -Level Error -Message "No JSON content retrieved" -FunctionName $MyInvocation.MyCommand.Name
            Exit
        }
        Write-PSFMessage -Level Verbose -Message "JSON file processing completed" -FunctionName $MyInvocation.MyCommand.Name

        # Output Directory
        Write-PSFMessage -Level 'Verbose' -Message "Output Directory: $output" -FunctionName $MyInvocation.MyCommand.Name
        $OUTPUT_FOLDER = Get-OutputFolder -Path $output
        if ($cleanup) {
            Write-PSFMessage -Level 'Verbose' -Message "Clearing Output Directory: $($OUTPUT_FOLDER.FullName)" -FunctionName $MyInvocation.MyCommand.Name
            $null = Clear-OutputFolder -Path $OUTPUT_FOLDER -Confirm:$false
        }

        # IPinfo CLI
        if ($IPinfo) {
            if (Test-IPinfo) {
                Write-PSFMessage -Level 'Verbose' -Message "IPinfo CLI installed" -FunctionName $MyInvocation.MyCommand.Name
                $InstallPath = "$env:LOCALAPPDATA\ipinfo"
                $IPinfoExe = Join-Path -Path $InstallPath -ChildPath "ipinfo.exe"
            } else {
                Write-PSFMessage -Level 'Error' -Message "IPinfo CLI not installed" -FunctionName $MyInvocation.MyCommand.Name
                throw "IPinfo CLI not installed"
            }

            # IPinfo CLI - Access Token
            $request = Invoke-WebRequest -Uri "ipinfo.io?token=$token" -Method 'Get'
            if ($request.StatusCode -eq 200) {
                Write-PSFMessage -Level 'Verbose' -Message "IPinfo CLI Access Token is valid" -FunctionName $MyInvocation.MyCommand.Name
            } else {
                Write-PSFMessage -Level 'Error' -Message "IPinfo CLI Access Token invalid" -FunctionName $MyInvocation.MyCommand.Name
                throw "IPinfo CLI Access Token invalid"
            }
        }
        # xsv
        #$script:xsv = "$PSScriptRoot\Tools\xsv\xsv.exe"

        # TODO: useful?
        # Function Get-FileSize
        Function Get-FileSize() {
            Param ([long]$Length)
            If ($Length -gt 1TB) { [string]::Format("{0:0.00} TB", $Length / 1TB) }
            ElseIf ($Length -gt 1GB) { [string]::Format("{0:0.00} GB", $Length / 1GB) }
            ElseIf ($Length -gt 1MB) { [string]::Format("{0:0.00} MB", $Length / 1MB) }
            ElseIf ($Length -gt 1KB) { [string]::Format("{0:0.00} KB", $Length / 1KB) }
            ElseIf ($Length -gt 0) { [string]::Format("{0:0.00} Bytes", $Length) }
            Else { "" }
        }


    }
    #endregion begin

    #region processing
    process {
        #region header
        # Create a record of your PowerShell session to a text file
        if ($Transcript) {
            $TranscriptPath = Join-Path -Path $OUTPUT_FOLDER.FullName -ChildPath 'Transcript.txt'
            Write-PSFMessage -level 'Verbose' -message "Transcript will be saved to: $TranscriptPath" -functionname $MyInvocation.MyCommand.Name
            Start-Transcript -Path $TranscriptPath | out-null
        }
        

        # Get Start Time
        $startTime = (Get-Date)

        #region  Logo
        Write-Logo
        #endregion header

        #region import Blacklist Whitelist Status
        # import Blacklists
        if ($Blacklist) {
            Write-PSFMessage -Level 'Verbose' -Message "Importing Blacklist from: $Blacklist" -FunctionName $MyInvocation.MyCommand.Name
            Import-SupportlistFromFolder -Folder $Blacklist -ListType 'Blacklist'
        }

        # import Whitelists
        if ($Whitelist) {
            Write-PSFMessage -Level 'Verbose' -Message "Importing Blacklist from: $Whitelist" -FunctionName $MyInvocation.MyCommand.Name
            Import-SupportlistFromFolder -Folder $Whitelist -ListType 'Whitelist'
        }

        # Create StatusTable from 'Status.csv'
        $statusfile = Join-Path -path $PSScriptRoot -ChildPath 'internal\Config\Status.csv'
        Write-PSFMessage -Level Verbose -Message "Importing  $statusfile" -FunctionName $MyInvocation.MyCommand.Name
        $StatusTable = @()
        if (Test-Path  $statusfile) {
            $StatusTable = Import-Csv  $statusfile -Delimiter "," -Encoding UTF8
            if ($StatusTable.Count -gt 0) {
                Write-PSFMessage -Level Verbose -Message "Status: $StatusTable.Count rows imported." -FunctionName $MyInvocation.MyCommand.Name
            }
        } else {
            Write-PSFMessage -Level Error -Message "Status table NOT found." -FunctionName $MyInvocation.MyCommand.Name
            Break
        }

        # IP addresses and details
        $UniqueIP = $Results.IpAddress | Select-Object -Unique
        Write-PSFMessage -Level Verbose -Message "$($UniqueIP.Count) unique IP addresses found (total: $($Results.Count))" -FunctionName $MyInvocation.MyCommand.Name
        $IPDetails = Get-Ipinfo -Token $token -AddressList $Unique
        Write-PSFMessage -Level Verbose -Message "IPInfo details retrieved" -FunctionName $MyInvocation.MyCommand.Name

        #endregion import Blacklist Whitelist Status
        
        #region Analysis

        Function Start-Processing {

            $StartTime_Processing = (Get-Date)

            # Input Size
            $InputSize = Get-FileSize((Get-Item $inputobject).Length)
            Write-PSFMessage -Level Important -Message "[Info]  Total Input Size: $InputSize"

            # Count rows of JSON (w/ thousands separators)
            $Count = 0
            switch -File $inputobject { default { ++$Count } }
            $Rows = '{0:N0}' -f $Count
            Write-Output "[Info]  Total Lines: $Rows"

            # Prpare output folder for OutputFormats
            switch ($outputFormat) {
                'CSV' {
                    Write-PSFMessage -Level Verbose -Message "[Info] Output Format: CSV"
                    New-Item "$OUTPUT_FOLDER\SignInLogsGraph\CSV" -ItemType Directory -Force | Out-Null
                }
                'XLSX' {
                    Write-PSFMessage -Level Verbose -Message "[Info] Output Format: XLSX"
                    New-Item "$OUTPUT_FOLDER\SignInLogsGraph\XLSX" -ItemType Directory -Force | Out-Null
                }
                'JSON' {
                    Write-PSFMessage -Level Verbose -Message "[Info] Output Format: JSON"
                    New-Item "$OUTPUT_FOLDER\SignInLogsGraph\JSON" -ItemType Directory -Force | Out-Null
                }
            }

            Write-PSFMessage -Level Verbose -Message "[Info] Processing Microsoft Entra ID Sign-In Logs ..."

            # Time Frame
            $StartDate = (($Data | Sort-Object { $_.CreatedDateTime -as [datetime] } -Descending | Select-Object -Last 1).CreatedDateTime).ToString("yyyy-MM-dd HH:mm:ss")
            $EndDate = (($Data | Sort-Object { $_.CreatedDateTime -as [datetime] } -Descending | Select-Object -First 1).CreatedDateTime).ToString("yyyy-MM-dd HH:mm:ss")
            Write-PSFMessage -Level Verbose -Message "Log data from $StartDate UTC until $EndDate UTC"

            #region Data Processing
            # https://learn.microsoft.com/en-us/powershell/module/Microsoft.Graph.Beta.Reports/Get-MgBetaAuditLogSignIn?view=graph-powershell-beta
            # https://learn.microsoft.com/nb-no/graph/api/resources/signin?view=graph-rest-beta

            # CSV
            $Results = @()
            $Results = 
            ForEach ($Record in $Data) {
                [PSCustomObject]@{
                    "Id"                             = $Record.Id # The identifier representing the sign-in activity.
                    "CreatedDateTime"                = ($Record | Select-Object -ExpandProperty CreatedDateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    "UserDisplayName"                = $Record.UserDisplayName # The display name of the user.
                    "UserPrincipalName"              = $Record.UserPrincipalName # The UPN of the user.
                    "UserId"                         = $Record.UserId # The identifier of the user.
                    "AppDisplayName"                 = $Record.AppDisplayName # The application name displayed in the Microsoft Entra admin center.
                    "AppId"                          = $Record.AppId # The application identifier in Microsoft Entra ID.
                    "ClientAppUsed"                  = $Record.ClientAppUsed # The legacy client used for sign-in activity.
                    "IpAddress"                      = $Record.IpAddress # The IP address of the client from where the sign-in occurred.
                    "ASN"                            = $Record.AutonomousSystemNumber # The Autonomous System Number (ASN) of the network used by the actor.
                    "IPAddressFromResourceProvider"  = $Record.IPAddressFromResourceProvider # The IP address a user used to reach a resource provider, used to determine Conditional Access compliance for some policies. For example, when a user interacts with Exchange Online, the IP address that Microsoft Exchange receives from the user can be recorded here. This value is often null.
                    "City"                           = $Record | Select-Object -ExpandProperty Location | Select-Object -ExpandProperty City # The city from where the sign-in occurred.
                    "State"                          = $Record | Select-Object -ExpandProperty Location | Select-Object -ExpandProperty State # The state from where the sign-in occurred.
                    "CountryOrRegion"                = $Record | Select-Object -ExpandProperty Location | Select-Object -ExpandProperty CountryOrRegion # The two letter country code from where the sign-in occurred.
                    "Latitude"                       = $Record | Select-Object -ExpandProperty Location | Select-Object -ExpandProperty GeoCoordinates | Select-Object -ExpandProperty Latitude
                    "Longitude"                      = $Record | Select-Object -ExpandProperty Location | Select-Object -ExpandProperty GeoCoordinates | Select-Object -ExpandProperty Longitude
                    "AuthenticationRequirement"      = $Record.AuthenticationRequirement # This holds the highest level of authentication needed through all the sign-in steps, for sign-in to succeed.
                    "SignInEventTypes"               = $Record | Select-Object -ExpandProperty SignInEventTypes # Indicates the category of sign in that the event represents.
                    "AuthenticationMethodsUsed"      = $Record | Select-Object -ExpandProperty AuthenticationMethodsUsed # The authentication methods used.

                    # Status - The sign-in status. Includes the error code and description of the error (for a sign-in failure).
                    # https://learn.microsoft.com/nb-no/graph/api/resources/signinstatus?view=graph-rest-beta
                    "ErrorCode"                      = $Record | Select-Object -ExpandProperty Status | Select-Object -ExpandProperty ErrorCode # Provides the 5-6 digit error code that's generated during a sign-in failure.
                    "FailureReason"                  = $Record | Select-Object -ExpandProperty Status | Select-Object -ExpandProperty FailureReason # Provides the error message or the reason for failure for the corresponding sign-in activity.
                    "AdditionalDetails"              = $Record | Select-Object -ExpandProperty Status | Select-Object -ExpandProperty AdditionalDetails # Provides additional details on the sign-in activity.

                    # AuthenticationDetails - The result of the authentication attempt and more details on the authentication method.
                    # https://learn.microsoft.com/nb-no/graph/api/resources/authenticationdetail?view=graph-rest-beta
                    "AuthenticationMethod"           = ($Record | Select-Object -ExpandProperty AuthenticationDetails | Select-Object -ExpandProperty AuthenticationMethod -Unique) -join ", " # The type of authentication method used to perform this step of authentication.
                    "AuthenticationMethodDetail"     = ($Record | Select-Object -ExpandProperty AuthenticationDetails | Select-Object -ExpandProperty AuthenticationMethodDetail -Unique) -join ", " # Details about the authentication method used to perform this authentication step.
                    "AuthenticationStepDateTime"     = ($Record | Select-Object -ExpandProperty AuthenticationDetails | Select-Object -ExpandProperty AuthenticationStepDateTime -Unique | ForEach-Object { ($_).ToString("yyyy-MM-dd HH:mm:ss") }) -join ", " # Represents date and time information using ISO 8601 format and is always in UTC time.
                    "AuthenticationStepRequirement"  = ($Record | Select-Object -ExpandProperty AuthenticationDetails | Select-Object -ExpandProperty AuthenticationStepRequirement -Unique) -join ", " # The step of authentication that this satisfied. 
                    "AuthenticationStepResultDetail" = ($Record | Select-Object -ExpandProperty AuthenticationDetails | Select-Object -ExpandProperty AuthenticationStepResultDetail -Unique) -join ", " # Details about why the step succeeded or failed. 
                    "Succeeded"                      = ($Record | Select-Object -ExpandProperty AuthenticationDetails | Select-Object -ExpandProperty Succeeded -Unique) -join ", " # Indicates the status of the authentication step.

                    # AuthenticationProcessingDetails - More authentication processing details, such as the agent name for PTA and PHS, or a server or farm name for federated authentication.
                    "Domain Hint Present"            = ($Record | Select-Object -ExpandProperty AuthenticationProcessingDetails | Where-Object { $_.Key -eq 'Domain Hint Present' }).Value
                    "Is CAE Token"                   = ($Record | Select-Object -ExpandProperty AuthenticationProcessingDetails | Where-Object { $_.Key -eq 'Is CAE Token' }).Value
                    "Login Hint Present"             = ($Record | Select-Object -ExpandProperty AuthenticationProcessingDetails | Where-Object { $_.Key -eq 'Login Hint Present' }).Value
                    "Oauth Scope Info"               = ($Record | Select-Object -ExpandProperty AuthenticationProcessingDetails | Where-Object { $_.Key -eq 'Oauth Scope Info' }).Value
                    "Root Key Type"                  = ($Record | Select-Object -ExpandProperty AuthenticationProcessingDetails | Where-Object { $_.Key -eq 'Root Key Type' }).Value

                    "ClientCredentialType"           = $Record.ClientCredentialType # Describes the credential type that a user client or service principal provided to Microsoft Entra ID to authenticate itself. You can review this property to track and eliminate less secure credential types or to watch for clients and service principals using anomalous credential types.
                    "ConditionalAccessStatus"        = $Record.ConditionalAccessStatus # The status of the conditional access policy triggered.
                    "CorrelationId"                  = $Record.CorrelationId # The identifier that's sent from the client when sign-in is initiated.
                    "IncomingTokenType"              = $Record.IncomingTokenType # Indicates the token types that were presented to Microsoft Entra ID to authenticate the actor in the sign in. 
                    "OriginalRequestId"              = $Record.OriginalRequestId # The request identifier of the first request in the authentication sequence.
                    "IsInteractive"                  = $Record.IsInteractive # Indicates whether a user sign in is interactive. In interactive sign in, the user provides an authentication factor to Microsoft Entra ID. These factors include passwords, responses to MFA challenges, biometric factors, or QR codes that a user provides to Microsoft Entra ID or an associated app. In non-interactive sign in, the user doesn't provide an authentication factor. Instead, the client app uses a token or code to authenticate or access a resource on behalf of a user. Non-interactive sign ins are commonly used for a client to sign in on a user's behalf in a process transparent to the user.
                    "ProcessingTimeInMilliseconds"   = $Record.ProcessingTimeInMilliseconds # The request processing time in milliseconds in AD STS.
                    "ResourceDisplayName"            = $Record.ResourceDisplayName # The name of the resource that the user signed in to.
                    "ResourceId"                     = $Record.ResourceId # The identifier of the resource that the user signed in to.
                    "ResourceServicePrincipalId"     = $Record.ResourceServicePrincipalId # The identifier of the service principal representing the target resource in the sign-in event.
                    "ResourceTenantId"               = $Record.ResourceTenantId # The tenant identifier of the resource referenced in the sign in.
                    "RiskDetail"                     = $Record.RiskDetail # The reason behind a specific state of a risky user, sign-in, or a risk event.
                    "RiskEventTypes_v2"              = $Record | Select-Object -ExpandProperty RiskEventTypes_v2 # The list of risk event types associated with the sign-in.
                    "RiskLevelAggregated"            = $Record.RiskLevelAggregated # The aggregated risk level. The value hidden means the user or sign-in wasn't enabled for Microsoft Entra ID Protection.
                    "RiskLevelDuringSignIn"          = $Record.RiskLevelDuringSignIn # The risk level during sign-in. The value hidden means the user or sign-in wasn't enabled for Microsoft Entra ID Protection.
                    "RiskState"                      = $Record.RiskState # The risk state of a risky user, sign-in, or a risk event.
                    "SignInTokenProtectionStatus"    = $Record.SignInTokenProtectionStatus # oken protection creates a cryptographically secure tie between the token and the device it is issued to. This field indicates whether the signin token was bound to the device or not.
                    "TokenIssuerName"                = $Record.TokenIssuerName # The name of the identity provider.
                    "TokenIssuerType"                = $Record.TokenIssuerType # The type of identity provider.
                    "UniqueTokenIdentifier"          = $Record.UniqueTokenIdentifier # A unique base64 encoded request identifier used to track tokens issued by Microsoft Entra ID as they're redeemed at resource providers.
                    "UserAgent"                      = $Record.UserAgent # The user agent information related to sign-in.
                    "UserType"                       = $Record | Select-Object -ExpandProperty UserType | ForEach-Object { $_.Replace("member", "Member") } | ForEach-Object { $_.Replace("guest", "Guest") } # Identifies whether the user is a member or guest in the tenant.
                    "AuthenticationProtocol"         = $Record.AuthenticationProtocol # Lists the protocol type or grant type used in the authentication.
                    "OriginalTransferMethod"         = $Record.OriginalTransferMethod # Transfer method used to initiate a session throughout all subsequent request.

                    # MfaDetail - This property is deprecated.
                    "AuthMethod"                     = $Record | Select-Object -ExpandProperty MfaDetail | Select-Object -ExpandProperty AuthMethod
                    "AuthDetail"                     = $Record | Select-Object -ExpandProperty MfaDetail | Select-Object -ExpandProperty AuthDetail

                    # DeviceDetail - The device information from where the sign-in occurred. Includes information such as deviceId, OS, and browser.
                    # https://learn.microsoft.com/nb-no/graph/api/resources/devicedetail?view=graph-rest-beta
                    "DeviceId"                       = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty DeviceId # Refers to the UniqueID of the device used for signing-in.
                    "DisplayName"                    = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty DisplayName # Refers to the name of the device used for signing-in.
                    "OperatingSystem"                = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty OperatingSystem # Indicates the OS name and version used for signing-in.
                    "Browser"                        = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty Browser # Indicates the browser information of the used for signing-in.
                    "IsCompliant"                    = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty IsCompliant # Indicates whether the device is compliant or not.
                    "IsManaged"                      = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty IsManaged # Indicates if the device is managed or not.
                    "TrustType"                      = $Record | Select-Object -ExpandProperty DeviceDetail | Select-Object -ExpandProperty TrustType # Indicates information on whether the signed-in device is Workplace Joined, AzureAD Joined, Domain Joined.
    
                    # NetworkLocationDetails - The network location details including the type of network used and its names.
                    # https://learn.microsoft.com/nb-no/graph/api/resources/networklocationdetail?view=graph-rest-beta
                    "NetworkType"                    = $Record | Select-Object -ExpandProperty NetworkLocationDetails | Select-Object -ExpandProperty NetworkType # Provides the type of network used when signing in.
                    "NetworkNames"                   = $Record | Select-Object -ExpandProperty NetworkLocationDetails | Select-Object -ExpandProperty NetworkNames # Provides the name of the network used when signing in.
                }
            }

            #endregion Data Processing

            # export results
            if ($Results.Count -gt 0) {
                Write-PSFMessage -Level Verbose -Message "$($Result.Count) rows to process" -FunctionName $MyInvocation.MyCommand.Name
                switch ($outputFormat) {
                    'CSV' { 
                        Write-PSFMessage -Level Verbose -Message "Exporting CSV" -FunctionName $MyInvocation.MyCommand.Name
                        $Results | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Untouched.csv" -NoTypeInformation -Encoding UTF8
                    }
                    'XLSX' {                 
                        Write-PSFMessage -Level Verbose -Message "Exporting XLSX" -FunctionName $MyInvocation.MyCommand.Name
                        $Results | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\XLSX\Untouched.xlsx" -NoNumberConversion * -FreezePane 2, 5 -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "SignInLogsGraph" -CellStyleSB {
                            param($WorkSheet)
                            # BackgroundColor and FontColor for specific cells of TopRow
                            $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                            Set-Format -Address $WorkSheet.Cells["A1:BP1"] -BackgroundColor $BackgroundColor -FontColor White
                            # HorizontalAlignment "Center" of columns A-BP
                            $WorkSheet.Cells["A:BP"].Style.HorizontalAlignment = "Center"
                        }
                    } 
                    'JSON' { 
                        Write-PSFMessage -Level Verbose -Message "Exporting JSON" -FunctionName $MyInvocation.MyCommand.Name
                        $Results | ConvertTo-Json | Out-File "$OUTPUT_FOLDER\SignInLogsGraph\JSON\Untouched.json" -Encoding UTF8
                    }
                }
            } else {
                Write-PSFMessage -Level Error -Message "No data to process" -FunctionName $MyInvocation.MyCommand.Name
                Exit
            } 



            # UserId
            $UserId = ($Results | Select-Object UserId -Unique).Count
            Write-PSFMessage -Level Verbose -Message "[Info]  $UserId UserId(s) found" -FunctionName $MyInvocation.MyCommand.Name

            # Member
            $Member = ($Results | Where-Object { $_.UserType -eq 'Member' } | Select-Object UserId -Unique).Count
            Write-PSFMessage -Level Verbose -Message "[Info]  $Member UserId(s) found" -FunctionName $MyInvocation.MyCommand.Name

            # Guest
            $Guest = ($Results | Where-Object { $_.UserType -eq 'Guest' } | Select-Object UserId -Unique).Count
            Write-PSFMessage -Level Verbose -Message "[Info]  $Guest UserId(s) found" -FunctionName $MyInvocation.MyCommand.Name

            # DeviceId
            $DeviceId = ($Results | Select-Object DeviceId -Unique).Count
            Write-PSFMessage -Level Verbose -Message "[Info]  $DeviceId UserId(s) found" -FunctionName $MyInvocation.MyCommand.Name


            # Microsoft Entra ID P2
            # https://www.microsoft.com/en-us/security/business/microsoft-entra-pricing
            $RiskLevelDuringSignIn = ($Results | Select-Object RiskLevelDuringSignIn -Unique).RiskLevelDuringSignIn
            if ( -not($RiskLevelDuringSignIn -eq "hidden")) {
                $EntraType = "Microsoft Entra ID P2"
                Write-PSFMessage -Level Verbose -Message "[Info] Microsoft Entra ID P2 detected" -FunctionName $MyInvocation.MyCommand.Name
            }
            # Identity Protection
            # - Risk-based Conditional Access (sign-in risk, user risk)
            # - Authentication context (step-up authentication)
            # - Device and application filters for Conditional Access
            # - Token protection
            # - Vulnerabilities and risky accounts
            # - Risk event investigation

            $EndTime_Processing = (Get-Date)
            $Time_Processing = ($EndTime_Processing - $StartTime_Processing)
            ('ADSignInLogsGraph Processing duration: {0} h {1} min {2} sec' -f $Time_Processing.Hours, $Time_Processing.Minutes, $Time_Processing.Seconds) >> "$OUTPUT_FOLDER\Stats.txt"

        
    
    #region Hunt

    # Authorization Error Codes (AADSTS) aka Entra ID Sign-in Error Codes
    # https://learn.microsoft.com/en-us/azure/active-directory/develop/reference-error-codes
    # https://login.microsoftonline.com/error
    # https://blog.icewolf.ch/archive/2021/02/04/hunting-for-basic-authentication-in-azuread/
    $Records =
    ForEach ($Record in $Results) {
        # Status
        $statusrecord = $StatusTable | Where-Object { $_.ErrorCode -eq $record.ErrorCode }
        if ($statusrecord) {
               $ErrorCode = $statusrecord.ErrorCode
               $Status = $statusrecord.Status
               $Message = $statusrecord.Message
            }
        else {
                $ErrorCode = $record.ErrorCode
                $Status = 'Interrupted'
                $Message = 'unknown error'
            }
        # ip dateils
        $iprecord =  $IPDetails | Where-Object { $_.IP -eq $Record.IpAddress }
        if ($iprecord) {
                $City = $iprecord.City
                $Region = $iprecord.Region
                $Country = $iprecord.Country
                $CountryName = $iprecord.CountryName
                $Location = $iprecord.Location
                $ASN = $iprecord.ASN
                $OrgName = $iprecord.Organization
                $PostalCode = $iprecord.PostalCode
                $Timezone = $iprecord.Timezone
            }
        else {
                $City        = ""
                $Region      = ""
                $Country     = ""
                $CountryName = ""
                $EU          = ""
                $Location    = ""
                $ASN         = ""
                $OrgName     = ""
                $PostalCode  = ""
                $Timezone    = ""
        }
        
        
        [PSCustomObject]@{
            "Id"                           = $Record.Id
            "CreatedDateTime"              = $Record.CreatedDateTime
            "UserDisplayName"              = $Record.UserDisplayName
            "UserPrincipalName"            = $Record.UserPrincipalName
            "UserId"                       = $Record.UserId
            "AppId"                        = $Record.AppId
            "AppDisplayName"               = $Record.AppDisplayName
            "ClientAppUsed"                = $Record.ClientAppUsed
            "CorrelationId"                = $Record.CorrelationId
            "ConditionalAccessStatus"      = $Record.ConditionalAccessStatus
            "OriginalRequestId"            = $Record.OriginalRequestId
            "IsInteractive"                = $Record.IsInteractive
            "TokenIssuerName"              = $Record.TokenIssuerName
            "TokenIssuerType"              = $Record.TokenIssuerType
            "ProcessingTimeInMilliseconds" = $Record.ProcessingTimeInMilliseconds
            "RiskDetail"                   = $Record.RiskDetail
            "RiskLevelAggregated"          = $Record.RiskLevelAggregated
            "RiskLevelDuringSignIn"        = $Record.RiskLevelDuringSignIn
            "RiskState"                    = $Record.RiskState
            "RiskEventTypesV2"             = $Record.RiskEventTypesV2
            "ResourceDisplayName"          = $Record.ResourceDisplayName
            "ResourceId"                   = $Record.ResourceId
            "AuthenticationMethodsUsed"    = $Record.AuthenticationMethodsUsed
            "ErrorCode"                    = $ErrorCode
            'ErrorMessage'                 = $Message
            "FailureReason"                = $Record.FailureReason
            "AdditionalDetails"            = $Record.AdditionalDetails
            "Status"                       = $Status
            "DeviceId"                     = $Record.DeviceId
            "DisplayName"                  = $Record.DisplayName
            "OperatingSystem"              = $Record.OperatingSystem
            "Browser"                      = $Record.Browser
            "IsCompliant"                  = $Record.IsCompliant
            "IsManaged"                    = $Record.IsManaged
            "TrustType"                    = $Record.TrustType
            "AuthMethod"                   = $Record.AuthMethod
            "AuthDetail"                   = $Record.AuthDetail
            "AuthenticationProtocol"       = $Record.AuthenticationProtocol
            "IpAddress"                    = $record.IpAddress
            "City"                         = $City
            "Region"                       = $Region
            "Country"                      = $Country
            "Country Name"                 = $CountryName
            "EU"                           = $EU
            "Location"                     = $Location
            "ASN"                          = $ASN
            "OrgName"                      = $OrgName
            "Postal Code"                  = $PostalCode
            "Timezone"                     = $Timezone
            "UserType"                     = $Record.UserType
        }
    }

    $Hunt = $Records | Sort-Object { $_.IP -as [Version] } 
    #endregion Hunt

    #region Stats

    Function Get-Stats {

        $StartTime_Stats = (Get-Date)

        # AppDisplayName (Stats)
        $AppCount = '{0:N0}' -f ($Results | Select-Object AppDisplayName -Unique).Count
        $AppDisplayName = $Results | Group-Object AppDisplayName, AppId | Select-Object @{Name = 'AppDisplayName'; Expression = { $_.Values[0] } }, @{Name = 'AppId'; Expression = { $_.Values[1] } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending
        Write-Output "[Info]  $AppCount Applications found"

        # Export
        if ($AppDisplayName.Count -gt 0) {
            switch ($outputFormat) {
                'CSV' {
                    $AppDisplayName | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\AppDisplayName.csv" -NoTypeInformation -Encoding UTF8
                }
                'XLSX' {
                    Export-MSAExcel -exportdata $AppDisplayName -name 'AppDisplayName'
                }
                'JSON' {
                    $AppDisplayName | ConvertTo-Json | Out-File "$OUTPUT_FOLDER\SignInLogsGraph\JSON\AppDisplayName.json" -Encoding UTF8
                }
            }
        }
            

        # ASN / Status

            if (($Hunt | Select-Object ASN).Count -gt 0) {
                $ASN = $Hunt | Select-Object ASN, OrgName, Status | Where-Object { $_.ASN -ne '' } | Where-Object { $null -ne ($_.PSObject.Properties | ForEach-Object { $_.Value }) } | Group-Object ASN, OrgName, Status | Select-Object @{Name = 'ASN'; Expression = { $_.Values[0] } }, @{Name = 'OrgName'; Expression = { $_.Values[1] } }, @{Name = 'Status'; Expression = { $_.Values[2] } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending 

            }


        # XLSX (Stats)

                if ($ASN.Count -gt 0) {
                    $ASN | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\XLSX\ASN.xlsx" -NoNumberConversion * -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "ASN" -CellStyleSB {
                        param($WorkSheet)
                        # BackgroundColor and FontColor for specific cells of TopRow
                        $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                        Set-Format -Address $WorkSheet.Cells["A1:E1"] -BackgroundColor $BackgroundColor -FontColor White
                        # HorizontalAlignment "Center" of columns A-E
                        $WorkSheet.Cells["A:E"].Style.HorizontalAlignment = "Center"

                        # Iterating over the ASN-Blacklist HashTable
                        foreach ($ASN in $AsnBlacklist_HashTable.Keys) {
                            $ConditionValue = 'NOT(ISERROR(FIND("{0}",$A1)))' -f $ASN
                            Add-ConditionalFormatting -Address $WorkSheet.Cells["A:E"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                        }

                    }
                }


        # AuthenticationProtocol (Stats)
        $Total = ($Results | Select-Object AuthenticationProtocol | Measure-Object).Count
        $Results | Group-Object AuthenticationProtocol | Select-Object @{Name = 'AuthenticationProtocol'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\AuthenticationProtocol.csv" -NoTypeInformation

        # XLSX
        if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\AuthenticationProtocol.csv") {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\AuthenticationProtocol.csv" -Delimiter ","
            if ($IMPORT.Count -gt 0) {
                $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\AuthenticationProtocol.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "AuthenticationProtocol" -CellStyleSB {
                    param($WorkSheet)
                    # BackgroundColor and FontColor for specific cells of TopRow
                    $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                    Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                    # HorizontalAlignment "Center" of columns B-C
                    $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
                    # ConditionalFormatting - AuthenticationProtocol
                    Add-ConditionalFormatting -Address $WorkSheet.Cells["A:C"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("deviceCode",$A1)))' -BackgroundColor Red # Device Code Authentication
                }
            }
        }

        # AuthenticationRequirement (Stats)
        $Total = ($Results | Select-Object AuthenticationRequirement | Measure-Object).Count
        $Results | Group-Object AuthenticationRequirement | Select-Object @{Name = 'AuthenticationRequirement'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\AuthenticationRequirement.csv" -NoTypeInformation

        # XLSX
        if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\AuthenticationRequirement.csv") {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\AuthenticationRequirement.csv" -Delimiter ","
            if ($IMPORT.Count -gt 0) {
                $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\AuthenticationRequirement.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "AuthenticationRequirement" -CellStyleSB {
                    param($WorkSheet)
                    # BackgroundColor and FontColor for specific cells of TopRow
                    $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                    Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                    # HorizontalAlignment "Center" of columns B-C
                    $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
                }
            }
        }

        # AuthMethod (Stats)
        $Total = ($Results | Select-Object AuthMethod | Measure-Object).Count
        $AuthMethod = $Results | Group-Object AuthMethod | Select-Object @{Name = 'AuthMethod'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending

        # XLSX
        if ($AuthMethod.Count -gt 0) {
            $AuthMethod | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\AuthMethod.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "AuthMethod" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
            }
        }

        # Browser (Stats)
        $Total = ($Results | Select-Object Browser | Measure-Object).Count
        $Browser = $Results | Group-Object Browser | Select-Object @{Name = 'Browser'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending

        # XLSX
        if ($Browser.Count -gt 0) {
            $Browser | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\Browser.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "Browser" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
            }
        }
    

    # ClientAppUsed (Stats)
    $Total = ($Results | Select-Object ClientAppUsed | Measure-Object).Count
    $ClientAppUsed = $Results | Group-Object ClientAppUsed | Select-Object @{Name = 'ClientAppUsed'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ClientAppUsed.csv") {
        if ($ClientAppUsed.Count -gt 0) {
            $ClientAppUsed | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\ClientAppUsed.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "ClientAppUsed" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
                # ConditionalFormatting - Modern Authentication Clients
                $Green = [System.Drawing.Color]::FromArgb(0, 176, 80)
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:C"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Browser",$A1)))' -BackgroundColor $Green
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:C"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Mobile Apps and Desktop clients",$A1)))' -BackgroundColor $Green
                # ConditionalFormatting - Legacy Authentication Clients
                # https://learn.microsoft.com/en-us/azure/active-directory/conditional-access/block-legacy-authentication
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:C"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Authenticated SMTP",$A1)))' -BackgroundColor Red
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:C"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Other clients",$A1)))' -BackgroundColor Red
            }
        }
    }

    # ClientAppUsed / Status 
                        
    # CSV (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") {
        if ([int](& $xsv count "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") -gt 0) {
            $Total = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object ClientAppUsed | Measure-Object).Count
            Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," -Encoding UTF8 | Select-Object @{Name = 'ClientAppUsed'; Expression = { if ($_.ClientAppUsed) { $_.ClientAppUsed }else { 'N/A' } } }, Status | Group-Object ClientAppUsed, Status | Select-Object @{Name = 'ClientAppUsed'; Expression = { $_.Values[0] } }, @{Name = 'Status'; Expression = { $_.Values[1] } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ClientAppUsed-Status.csv" -NoTypeInformation -Encoding UTF8
        }
    }

    # XLSX (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ClientAppUsed-Status.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ClientAppUsed-Status.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ClientAppUsed-Status.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\ClientAppUsed-Status.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "ClientAppUsed" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:D1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns A-D
                $WorkSheet.Cells["A:D"].Style.HorizontalAlignment = "Center"
                # ConditionalFormatting - Modern Authentication Clients
                $Green = [System.Drawing.Color]::FromArgb(0, 176, 80)
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:A"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Browser",$A1)))' -BackgroundColor $Green
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:A"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Mobile Apps and Desktop clients",$A1)))' -BackgroundColor $Green
                # ConditionalFormatting - Legacy Authentication Clients
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:A"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Authenticated SMTP",$A1)))' -BackgroundColor Red
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:D"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue '=AND($A1="Authenticated SMTP",$B1="Failure")' -BackGroundColor "Red"
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:A"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Other clients",$A1)))' -BackgroundColor Red
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:D"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue '=AND($A1="Other clients",$B1="Failure")' -BackGroundColor "Red"
            }
        }
    }
            

    # ConditionalAccessStatus (Stats)
    $Total = ($Results | Select-Object ConditionalAccessStatus | Measure-Object).Count
    $ConditionalAccessStatus = $Results | Group-Object ConditionalAccessStatus | Select-Object @{Name = 'ConditionalAccessStatus'; Expression = { $_.Name } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending

    # XLSX
    if ($ConditionalAccessStatus.Count -gt 0) {
        $ConditionalAccessStatus | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\ConditionalAccessStatus.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "ConditionalAccessStatus" -CellStyleSB {
            param($WorkSheet)
            # BackgroundColor and FontColor for specific cells of TopRow
            $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
            Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
            # HorizontalAlignment "Center" of columns A-C
            $WorkSheet.Cells["A:C"].Style.HorizontalAlignment = "Center"
        }
    }
            

    # Conditional Access Status (Investigating Sign-Ins with CA applied)
    # notApplied: No policy applied to the user and application during sign-in.
    # success:    One or more conditional access policies applied to the user and application (but not necessarily the other conditions) during sign-in.
    # failure:    The sign-in satisfied the user and application condition of at least one Conditional Access policy and grant controls are either not satisfied or set to block access.

    # Note: Conditional Access policies are enforced after first-factor authentication is completed. Conditional Access isn't intended to be an organization's first line of defense for scenarios like denial-of-service (DoS) attacks, but it can use signals from these events to determine access.

    # Country / Country Name
                        
    # CSV (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") {
        if ([int](& $xsv count "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") -gt 0) {
            $Total = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object Country | Measure-Object).Count
            Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object Country, "Country Name" | Where-Object { $_.Country -ne '' } | Where-Object { $null -ne ($_.PSObject.Properties | ForEach-Object { $_.Value }) } | Group-Object Country, "Country Name" | Select-Object @{Name = 'Country'; Expression = { $_.Values[0] } }, @{Name = 'Country Name'; Expression = { $_.Values[1] } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Country.csv" -NoTypeInformation -Encoding UTF8
                                
            # Countries
            $Countries = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object Country -Unique | Where-Object { $_.Country -ne '' } | Measure-Object).Count

            # Cities
            $Cities = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object City -Unique | Where-Object { $_.City -ne '' } | Measure-Object).Count

            Write-Output "[Info]  $Countries Countries and $Cities Cities found"
        }
    }

    # XLSX (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Country.csv") {
        $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Country.csv" -Delimiter ","
        if ($IMPORT.Count -gt 0) {
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\Country.xlsx" -NoNumberConversion * -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "Countries" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:D1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns A-D
                $WorkSheet.Cells["A:D"].Style.HorizontalAlignment = "Center"

                # Iterating over the Country-Blacklist HashTable
                foreach ($Country in $CountryBlacklist_HashTable.Keys) {
                    $ConditionValue = 'NOT(ISERROR(FIND("{0}",$B1)))' -f $Country
                    Add-ConditionalFormatting -Address $WorkSheet.Cells["A:D"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                }

            }
        }
    }
            

    # ErrorCode / Status

    # CSV (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") {
        $ErrorCode = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object ErrorCode
        if ($ErrorCode.Count -gt 0) {
            Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," -Encoding UTF8 | Select-Object Status, ErrorCode, FailureReason, AdditionalDetails | Group-Object Status, ErrorCode, FailureReason, AdditionalDetails | Select-Object Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $($ErrorCode.Count)) } }, @{Name = 'Status'; Expression = { $_.Values[0] } }, @{Name = 'ErrorCode'; Expression = { $_.Values[1] } }, @{Name = 'FailureReason'; Expression = { $_.Values[2] } }, @{Name = 'AdditionalDetails'; Expression = { $_.Values[3] } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ErrorCode.csv" -NoTypeInformation -Encoding UTF8
        }
    }

    # XLSX (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ErrorCode.csv") {
        $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ErrorCode.csv" -Delimiter ","
        if ($IMPORT.Count -gt 0) {
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\ErrorCode.xlsx" -NoNumberConversion * -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "ErrorCode" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:F1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns A-D
                $WorkSheet.Cells["A:D"].Style.HorizontalAlignment = "Center"
                # ConditionalFormatting - Suspicious Error Codes
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("50034",$D1)))' -BackgroundColor Red # "The user account does not exist in the tenant directory." --> involving non-existent user accounts
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("50053",$D1)))' -BackgroundColor Red # Sign-in was blocked because it came from an IP address with malicious activity or The account is locked, you've tried to sign in too many times with an incorrect user ID or password.
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("50126",$D1)))' -BackgroundColor Red # "Error validating credentials due to invalid username or password." --> Failed authentication attempts (Password Spraying Attack): Identify a traditional password spraying attack where a high number of users fail to authenticate from one single source IP in a short period of time.
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("90094",$D1)))' -BackgroundColor Red # Admin consent is required for the permissions requested by this application.
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("90095",$D1)))' -BackgroundColor Red # Admin consent is required for the permissions requested by this application. An admin consent request may be sent to the admin.
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("500121",$D1)))' -BackgroundColor Red # "Authentication failed during strong authentication request." --> MFA Fatigue aka MFA Prompt Bombing
                Add-ConditionalFormatting -Address $WorkSheet.Cells["A:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("530032",$D1)))' -BackgroundColor Red # User blocked due to risk on home tenant.
            }
        }
    }
            
    # IpAddress / Country Name
                        
    # CSV (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") {
        $IpAddress = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object IpAddress
        if ($IpAddress.Count -gt 0) {
            Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," -Encoding UTF8 | Select-Object IpAddress, Country, "Country Name", ASN, OrgName | Where-Object { $_.IpAddress -ne '' } | Where-Object { $_."Country Name" -ne '' } | Where-Object { $null -ne ($_.PSObject.Properties | ForEach-Object { $_.Value }) } | Group-Object IpAddress, Country, "Country Name", ASN, OrgName | Select-Object @{Name = 'IpAddress'; Expression = { $_.Values[0] } }, @{Name = 'Country'; Expression = { $_.Values[1] } }, @{Name = 'Country Name'; Expression = { $_.Values[2] } }, @{Name = 'ASN'; Expression = { $_.Values[3] } }, @{Name = 'OrgName'; Expression = { $_.Values[4] } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $($IpAddress.Count)) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\IpAddress.csv" -NoTypeInformation -Encoding UTF8
        }
    }

    # XLSX (Stats)
    if ($IpAddress.Count -gt 0) {
        $IpAddress | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\IpAddress.xlsx" -NoNumberConversion * -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "IpAddress" -CellStyleSB {
            param($WorkSheet)
            # BackgroundColor and FontColor for specific cells of TopRow
            $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
            Set-Format -Address $WorkSheet.Cells["A1:G1"] -BackgroundColor $BackgroundColor -FontColor White
            # HorizontalAlignment "Center" of columns A-G
            $WorkSheet.Cells["A:G"].Style.HorizontalAlignment = "Center"

            # Iterating over the ASN-Blacklist HashTable
            foreach ($ASN in $AsnBlacklist_HashTable.Keys) {
                $ConditionValue = 'NOT(ISERROR(FIND("{0}",$D1)))' -f $ASN
                Add-ConditionalFormatting -Address $WorkSheet.Cells["D:E"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
            }

            # Iterating over the Country-Blacklist HashTable
            foreach ($Country in $CountryBlacklist_HashTable.Keys) {
                $ConditionValue = 'NOT(ISERROR(FIND("{0}",$C1)))' -f $Country
                Add-ConditionalFormatting -Address $WorkSheet.Cells["B:C"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
            }

        }
    }

    # NetworkNames (Stats)
    $NetworkNames = $Results | Select-Object NetworkNames
    $NetworkNames | Group-Object NetworkNames | Select-Object @{Name = 'NetworkNames'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $($NetworkNames.Count)) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\NetworkNames.csv" -NoTypeInformation

    # XLSX
    if ($NetworkNames.Count -gt 0) {
        $NetworkNames | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\NetworkNames.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "NetworkNames" -CellStyleSB {
            param($WorkSheet)
            # BackgroundColor and FontColor for specific cells of TopRow
            $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
            Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
            # HorizontalAlignment "Center" of columns B-C
            $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
        }
    }

    # OperatingSystem (Stats)
    $Total = ($Results | Select-Object OperatingSystem | Measure-Object).Count
    $Results | Group-Object OperatingSystem | Select-Object @{Name = 'OperatingSystem'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\OperatingSystem.csv" -NoTypeInformation

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\OperatingSystem.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\OperatingSystem.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\OperatingSystem.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\OperatingSystem.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "OperatingSystem" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns A-C
                $WorkSheet.Cells["A:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    # ResourceDisplayName (Stats)
    $Total = ($Results | Select-Object ResourceDisplayName | Measure-Object).Count
    $Count = ($Results | Select-Object ResourceDisplayName | Sort-Object ResourceDisplayName -Unique | Measure-Object).Count
    $ResourceDisplayName = '{0:N0}' -f $Count
    $Results | Group-Object ResourceDisplayName | Select-Object @{Name = 'ResourceDisplayName'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ResourceDisplayName.csv" -NoTypeInformation
    Write-Output "[Info]  $ResourceDisplayName Resources found"

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ResourceDisplayName.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ResourceDisplayName.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\ResourceDisplayName.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\ResourceDisplayName.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "ResourceDisplayName" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    # RiskDetail (Stats)
    $Total = ($Results | Select-Object RiskDetail | Measure-Object).Count
    $Results | Group-Object RiskDetail | Select-Object @{Name = 'RiskDetail'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskDetail.csv" -NoTypeInformation

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskDetail.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskDetail.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskDetail.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\RiskDetail.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "RiskDetail" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    # RiskEventTypes_v2 (Stats)
    $Total = ($Results | Select-Object RiskEventTypes_v2 | Measure-Object).Count
    $Results | Group-Object RiskEventTypes_v2 | Select-Object @{Name = 'RiskEventTypes_v2'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskEventTypes_v2.csv" -NoTypeInformation

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskEventTypes_v2.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskEventTypes_v2.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskEventTypes_v2.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\RiskEventTypes_v2.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "RiskEventTypes_v2" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    # RiskLevelDuringSignIn (Stats)
    $Total = ($Results | Select-Object RiskLevelDuringSignIn | Measure-Object).Count
    $Results | Group-Object RiskLevelDuringSignIn | Select-Object @{Name = 'RiskLevelDuringSignIn'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskLevelDuringSignIn.csv" -NoTypeInformation

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskLevelDuringSignIn.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskLevelDuringSignIn.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\RiskLevelDuringSignIn.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\RiskLevelDuringSignIn.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "RiskLevelDuringSignIn" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns A-C
                $WorkSheet.Cells["A:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    # Status (Stats)
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") {
        if ([int](& $xsv count "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv") -gt 0) {
            $Total = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Select-Object Status | Measure-Object).Count
            Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Group-Object Status | Select-Object @{Name = 'Status'; Expression = { $_.Name } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Status.csv" -NoTypeInformation
        }
    }

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Status.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Status.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\Status.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\Status.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "Status" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns A-C
                $WorkSheet.Cells["A:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    # UserAgent (Stats)
    $Total = ($Results | Select-Object UserAgent | Measure-Object).Count
    $Results | Group-Object UserAgent | Select-Object @{Name = 'UserAgent'; Expression = { if ($_.Name) { $_.Name }else { 'N/A' } } }, Count, @{Name = 'PercentUse'; Expression = { "{0:p2}" -f ($_.Count / $Total) } } | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\UserAgent.csv" -NoTypeInformation

    # XLSX
    if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\UserAgent.csv") {
        if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\UserAgent.csv") -gt 0) {
            $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Stats\CSV\UserAgent.csv" -Delimiter ","
            $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\UserAgent.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "UserAgent" -CellStyleSB {
                param($WorkSheet)
                # BackgroundColor and FontColor for specific cells of TopRow
                $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                Set-Format -Address $WorkSheet.Cells["A1:C1"] -BackgroundColor $BackgroundColor -FontColor White
                # HorizontalAlignment "Center" of columns B-C
                $WorkSheet.Cells["B:C"].Style.HorizontalAlignment = "Center"
            }
        }
    }

    $EndTime_Stats = (Get-Date)
    $Time_Stats = ($EndTime_Stats - $StartTime_Stats)
('ADSignInLogsGraph Stats duration:           {0} h {1} min {2} sec' -f $Time_Stats.Hours, $Time_Stats.Minutes, $Time_Stats.Seconds) >> "$OUTPUT_FOLDER\Stats.txt"

}
}

#endregion Stats

#############################################################################################################################################################################################
#############################################################################################################################################################################################

#region Analytics

Function Get-Analytics {

    $StartTime_Analytics = (Get-Date)

    # Brute-Force Detection
    $Import = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," -Encoding UTF8 | Where-Object { $_.Status -eq 'Failure' }
    $Count = ($Import | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Where-Object Count -ge 1000 | Measure-Object).Count

    if ($Count -ge 1) {
        Write-Host "[Alert] Brute-Force Attack detected: 1000+ failed Sign-In events on a single day ($Count)" -ForegroundColor Red
        New-Item "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV" -ItemType Directory -Force | Out-Null
        New-Item "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\XLSX" -ItemType Directory -Force | Out-Null

        # CSV
        $Import | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Where-Object Count -ge 1000 | Select-Object Name, Count | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack-Overview.csv" -NoTypeInformation -Encoding UTF8
        $Import | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Where-Object Count -ge 1000 | Select-Object -Expand Group | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack.csv" -NoTypeInformation -Encoding UTF8

        # Brute-Force-Attack-Overview.xlsx
        if (Get-Module -ListAvailable -Name ImportExcel) {
            if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack-Overview.csv") {
                if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack-Overview.csv") -gt 0) {
                    $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack-Overview.csv" -Delimiter ","
                    $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\XLSX\Brute-Force-Attack-Overview.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "Brute-Force Attack" -CellStyleSB {
                        param($WorkSheet)
                        # BackgroundColor and FontColor for specific cells of TopRow
                        $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                        Set-Format -Address $WorkSheet.Cells["A1:B1"] -BackgroundColor $BackgroundColor -FontColor White
                        # HorizontalAlignment "Center" of columns A-B
                        $WorkSheet.Cells["A:B"].Style.HorizontalAlignment = "Center"
                    }
                }
            }
        }

        # Brute-Force-Attack.xlsx
        if (Get-Module -ListAvailable -Name ImportExcel) {
            if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack.csv") {
                if ([int](& $xsv count -d "," "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack.csv") -gt 0) {
                    $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\CSV\Brute-Force-Attack.csv" -Delimiter ","
                    $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Brute-Force-Attack\XLSX\Brute-Force-Attack.xlsx" -NoNumberConversion * -FreezePane 2, 4 -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "Brute-Force Attack" -CellStyleSB {
                        param($WorkSheet)
                        # BackgroundColor and FontColor for specific cells of TopRow
                        $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                        Set-Format -Address $WorkSheet.Cells["A1:AU1"] -BackgroundColor $BackgroundColor -FontColor White
                        # HorizontalAlignment "Center" of columns A-X and AA-AU
                        $WorkSheet.Cells["A:X"].Style.HorizontalAlignment = "Center"
                        $WorkSheet.Cells["AA:AU"].Style.HorizontalAlignment = "Center"
                        # ConditionalFormatting
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["AA:AA"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Failure",$AA1)))' -BackgroundColor Red
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["H:H"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Authenticated SMTP",$H1)))' -BackgroundColor Red
                    }
                }
            }
        }
    }

    # Basic Authentication (Legacy Authentication Client) detected: Authenticated SMTP
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ClientAppUsed -eq 'Authenticated SMTP' } | Measure-Object).Count

    if ($Count -ge 1) {
        $Failure = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ClientAppUsed -eq 'Authenticated SMTP' } | Where-Object { $_.Status -eq 'Failure' } | Measure-Object).Count
        $Success = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ClientAppUsed -eq 'Authenticated SMTP' } | Where-Object { $_.Status -eq 'Success' } | Measure-Object).Count
        $FailureCount = '{0:N0}' -f $Failure
        $SuccessCount = '{0:N0}' -f $Success
        Write-Host "[Alert] Basic Authentication (Legacy Authentication Client) detected: Authenticated SMTP ($Count)" -ForegroundColor Red
        Write-Host "[Alert] $FailureCount failed Sign-Ins via Legacy Authentication Client detected: Authenticated SMTP" -ForegroundColor Red
        Write-Host "[Alert] $SuccessCount successful Sign-Ins via Legacy Authentication Client detected: Authenticated SMTP" -ForegroundColor Red
    }

    # Suspicious Error Codes

    # ErrorCode: 90095 - Admin consent is required for the permissions requested by this application. An admin consent request may be sent to the admin.
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ErrorCode -eq '90095' } | Measure-Object).Count
    if ($Count -ge 1) {
        Write-Host "[Alert] Suspicious Error Code detected: 90095 - Admin consent is required for the permissions requested by an application ($Count)" -ForegroundColor Red
    }

    #############################################################################################################################################################################################

    # Line Charts
    New-Item "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts" -ItemType Directory -Force | Out-Null

    # Failure (Sign-Ins)
    $Total = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Failure' } | Select-Object IpAddress | Measure-Object).Count
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Failure' } | Select-Object IpAddress -Unique | Measure-Object).Count
    $UniqueFailures = '{0:N0}' -f $Count
    Write-Output "IPAddress" | Out-File "$OUTPUT_FOLDER\IpAddress\Failure.txt" -Encoding UTF8 # Header
    Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Failure' } | Select-Object -ExpandProperty IpAddress -Unique | Out-File "$OUTPUT_FOLDER\IpAddress\Failure.txt" -Append
    Write-Output "[Info]  $UniqueFailures failed Sign-Ins found ($Total)"

    # Authentication: Failure (Line Chart) --> Failed Sign-Ins per day
    $Import = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Failure' } | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Select-Object Count, @{Name = 'CreatedDateTime'; Expression = { $_.Values[0] } } | Sort-Object { $_.CreatedDateTime -as [datetime] }
    $ChartDefinition = New-ExcelChartDefinition -XRange CreatedDateTime -YRange Count -Title "Failed Sign-Ins" -ChartType Line -NoLegend -Width 1200
    $Import | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts\Failure.xlsx" -Append -WorksheetName "Line Chart" -AutoNameRange -ExcelChartDefinition $ChartDefinition

    # Failure (Map)
    if (Test-Path $IPinfoExe) {
        if (Test-Path "$OUTPUT_FOLDER\IpAddress\Failure.txt") {
            if ((Get-Item "$OUTPUT_FOLDER\IpAddress\Failure.txt").Length -gt 0kb) {
                # Internet Connectivity Check (Vista+)
                $NetworkListManager = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet

                if (!($NetworkListManager -eq "True")) {
                    Write-Host "[Error] Your computer is NOT connected to the Internet. IP addresses cannot be checked via IPinfo API." -ForegroundColor Red
                } else {
                    # Check if IPinfo.io is reachable
                    if (!(Test-Connection -ComputerName ipinfo.io -Count 1 -Quiet)) {
                        Write-Host "[Error] ipinfo.io is NOT reachable. IP addresses cannot be checked via IPinfo API." -ForegroundColor Red
                    } else {
                        # Map IPs
                        Get-Content "$OUTPUT_FOLDER\IpAddress\Failure.txt" | & $IPinfoExe map | Out-File "$OUTPUT_FOLDER\IpAddress\IPinfo\Map_Failure.txt"
                    }
                }
            }
        }
    }

    # Success (Sign-Ins)
    Write-Output "IPAddress" | Out-File "$OUTPUT_FOLDER\IpAddress\Success.txt" -Encoding UTF8 # Header
    Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Success' } | Select-Object -ExpandProperty IpAddress -Unique | Out-File "$OUTPUT_FOLDER\IpAddress\Success.txt" -Append

    # Authentication: Success (Line Chart) --> Successful Sign-Ins per day
    $Import = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Success' } | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Select-Object Count, @{Name = 'CreatedDateTime'; Expression = { $_.Values[0] } } | Sort-Object { $_.CreatedDateTime -as [datetime] }
    $ChartDefinition = New-ExcelChartDefinition -XRange CreatedDateTime -YRange Count -Title "Successful Sign-Ins" -ChartType Line -NoLegend -Width 1200
    $Import | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts\Success.xlsx" -Append -WorksheetName "Line Chart" -AutoNameRange -ExcelChartDefinition $ChartDefinition

    # Success (Map)
    if (Test-Path "$OUTPUT_FOLDER\IpAddress\Success.txt") {
        if ((Get-Item "$OUTPUT_FOLDER\IpAddress\Success.txt").Length -gt 0kb) {
            # Internet Connectivity Check (Vista+)
            $NetworkListManager = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet

            if (!($NetworkListManager -eq "True")) {
                Write-Host "[Error] Your computer is NOT connected to the Internet. IP addresses cannot be checked via IPinfo API." -ForegroundColor Red
            } else {
                # Check if IPinfo.io is reachable
                if (!(Test-Connection -ComputerName ipinfo.io -Count 1 -Quiet)) {
                    Write-Host "[Error] ipinfo.io is NOT reachable. IP addresses cannot be checked via IPinfo API." -ForegroundColor Red
                } else {
                    # Map IPs
                    Get-Content "$OUTPUT_FOLDER\IpAddress\Success.txt" | & $IPinfoExe map | Out-File "$OUTPUT_FOLDER\IpAddress\IPinfo\Map_Success.txt"
                }
            }
        }
    }

    # Interrupted (Sign-Ins)
    Write-Output "IPAddress" | Out-File "$OUTPUT_FOLDER\IpAddress\Interrupted.txt" -Encoding UTF8 # Header
    Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Interrupted' } | Select-Object -ExpandProperty IpAddress -Unique | Out-File "$OUTPUT_FOLDER\IpAddress\Interrupted.txt" -Append

    # Authentication: Interrupted (Line Chart) --> Interrupted Sign-Ins per day
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Interrupted' } | Measure-Object).Count

    if ($Count -ge 1) {
        $Import = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Interrupted' } | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Select-Object Count, @{Name = 'CreatedDateTime'; Expression = { $_.Values[0] } } | Sort-Object { $_.CreatedDateTime -as [datetime] }
        $ChartDefinition = New-ExcelChartDefinition -XRange CreatedDateTime -YRange Count -Title "Interrupted Sign-Ins" -ChartType Line -NoLegend -Width 1200
        $Import | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts\Interrupted.xlsx" -Append -WorksheetName "Line Chart" -AutoNameRange -ExcelChartDefinition $ChartDefinition
    }

    #############################################################################################################################################################################################

    # Conditional Access

    # Conditional Access Result: Success (Line Chart)
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ConditionalAccessStatus -eq 'Success' } | Measure-Object).Count

    if ($Count -ge 10) {
        $Import = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ConditionalAccessStatus -eq 'Success' } | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Select-Object Count, @{Name = 'CreatedDateTime'; Expression = { $_.Values[0] } } | Sort-Object { $_.CreatedDateTime -as [datetime] }
        $ChartDefinition = New-ExcelChartDefinition -XRange CreatedDateTime -YRange Count -Title "Conditional Access Result: Success" -ChartType Line -NoLegend -Width 1200
        $Import | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts\ConditionalAccessResult-Success.xlsx" -Append -WorksheetName "Line Chart" -AutoNameRange -ExcelChartDefinition $ChartDefinition
    }

    # Conditional Access Result: Failure (Line Chart)
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ConditionalAccessStatus -eq 'Failure' } | Measure-Object).Count

    if ($Count -ge 10) {
        $Import = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ConditionalAccessStatus -eq 'Failure' } | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Select-Object Count, @{Name = 'CreatedDateTime'; Expression = { $_.Values[0] } } | Sort-Object { $_.CreatedDateTime -as [datetime] }
        $ChartDefinition = New-ExcelChartDefinition -XRange CreatedDateTime -YRange Count -Title "Conditional Access Result: Failure" -ChartType Line -NoLegend -Width 1200
        $Import | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts\ConditionalAccessResult-Failure.xlsx" -Append -WorksheetName "Line Chart" -AutoNameRange -ExcelChartDefinition $ChartDefinition
    }

    # Conditional Access Result: Not applied (Line Chart)
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ConditionalAccessStatus -eq 'notApplied' } | Measure-Object).Count

    if ($Count -ge 10) {
        $Import = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.ConditionalAccessStatus -eq 'notApplied' } | Group-Object { ($_.CreatedDateTime -split "\s+")[0] } | Select-Object Count, @{Name = 'CreatedDateTime'; Expression = { $_.Values[0] } } | Sort-Object { $_.CreatedDateTime -as [datetime] }
        $ChartDefinition = New-ExcelChartDefinition -XRange CreatedDateTime -YRange Count -Title "Conditional Access Result: Not applied" -ChartType Line -NoLegend -Width 1200
        $Import | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\Stats\XLSX\LineCharts\ConditionalAccessResult-NotApplied.xlsx" -Append -WorksheetName "Line Chart" -AutoNameRange -ExcelChartDefinition $ChartDefinition
    }

    # Conditional Access (NOT Blocked)
    Write-Output "IPAddress" | Out-File "$OUTPUT_FOLDER\IpAddress\ConditionalAccess.txt" -Encoding UTF8 # Header
    Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.Status -eq 'Success' } | Where-Object { $_.ConditionalAccessStatus -eq "notApplied" -or $_.ConditionalAccessStatus -eq "success" } | Select-Object -ExpandProperty IpAddress -Unique | & $IPinfoExe grepip -o | Out-File "$OUTPUT_FOLDER\IpAddress\ConditionalAccess.txt" -Append

    # Conditional Access (Map)
    if (Test-Path "$OUTPUT_FOLDER\IpAddress\ConditionalAccess.txt") {
        if ((Get-Item "$OUTPUT_FOLDER\IpAddress\ConditionalAccess.txt").Length -gt 0kb) {
            # Internet Connectivity Check (Vista+)
            $NetworkListManager = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet

            if (!($NetworkListManager -eq "True")) {
                Write-Host "[Error] Your computer is NOT connected to the Internet. IP addresses cannot be checked via IPinfo API." -ForegroundColor Red
            } else {
                # Check if IPinfo.io is reachable
                if (!(Test-Connection -ComputerName ipinfo.io -Count 1 -Quiet)) {
                    Write-Host "[Error] ipinfo.io is NOT reachable. IP addresses cannot be checked via IPinfo API." -ForegroundColor Red
                } else {
                    # Map IPs
                    Get-Content "$OUTPUT_FOLDER\IpAddress\ConditionalAccess.txt" | & $IPinfoExe map | Out-File "$OUTPUT_FOLDER\IpAddress\IPinfo\Map_ConditionalAccess.txt"
                }
            }
        }
    }

    # Conditional Access Status (Investigating Sign-Ins with CA applied)
    # notApplied: No policy applied to the user and application during sign-in.
    # success:    One or more conditional access policies applied to the user and application (but not necessarily the other conditions) during sign-in.
    # failure:    The sign-in satisfied the user and application condition of at least one Conditional Access policy and grant controls are either not satisfied or set to block access.

    # Note: Conditional Access policies are enforced after first-factor authentication is completed. Conditional Access isn't intended to be an organization's first line of defense for scenarios like denial-of-service (DoS) attacks, but it can use signals from these events to determine access.

    # Impact Summary
    # Total: The number of users or sign-ins during the time period where at least one of the selected policies was evaluated.
    # Success: The number of users or sign-ins during the time period where the combined result of the selected policies was “Success” or “Report-only: Success”.
    # Failure: The number of users or sign-ins during the time period where the result of at least one of the selected policies was “Failure” or “Report-only: Failure”.
    # Not applied: The number of users or sign-ins during the time period where none of the selected policies applied.

    #############################################################################################################################################################################################

    # Very Risky Authentication (Microsoft Entra ID Premium P2 required)
    $Count = (Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.RiskLevelDuringSignIn -eq "high" } | Where-Object { $_.RiskState -eq "atRisk" } | Where-Object { ($_.RiskLevelAggregated -eq "medium" -Or $_.RiskLevelAggregated -eq "high") } | Measure-Object).Count

    if ($Count -ge 1) {
        Write-Host "[Alert] Very Risky Authentication(s) detected ($Count)" -ForegroundColor Red
    }

    # Adversary-in-the-Middle (AiTM) Phishing / MFA Attack [T1557]
    # Note: "OfficeHome" is a pretty reliable application for detecting threat actors, in particular when the DeviceId is empty. --> Check for unusual IP address (outside the country, not typical for that user, etc.)
    $Import = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter "," | Where-Object { $_.AppDisplayName -eq "OfficeHome" } | Where-Object { $_.DeviceId -eq "" } | Where-Object { ($_.ErrorCode -eq "0" -or $_.ErrorCode -eq "50074" -or $_.ErrorCode -eq "50140" -or $_.ErrorCode -eq "53000") }
    $Count = ($Import | Measure-Object).Count
    $Users = ($Import | Select-Object UserId -Unique | Measure-Object).Count

    # ApplicationId = 4765445b-32c6-49b0-83e6-1d93765276ca
    # ClientAppUsed = Browser
    # IsInteractive = True

    if ($Count -ge 1) {
        Write-Host "[Alert] Potential Adversary-in-the-Middle (AitM) Phishing Attack(s) detected ($Users account credentials, $Count events)" -ForegroundColor Red
        New-Item "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV" -ItemType Directory -Force | Out-Null
        New-Item "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\XLSX" -ItemType Directory -Force | Out-Null

        # CSV
        $Import | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM.csv" -NoTypeInformation -Encoding UTF8

        # XLSX
        if (Get-Module -ListAvailable -Name ImportExcel) {
            if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM.csv") {
                if ([int](& $xsv count "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM.csv") -gt 0) {
                    $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM.csv" -Delimiter "," | Sort-Object { $_.CreatedDateTime -as [datetime] } -Descending
                    $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\XLSX\AiTM.xlsx" -NoNumberConversion * -FreezePane 2, 4 -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "AiTM" -CellStyleSB {
                        param($WorkSheet)
                        # BackgroundColor and FontColor for specific cells of TopRow
                        $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                        Set-Format -Address $WorkSheet.Cells["A1:AW1"] -BackgroundColor $BackgroundColor -FontColor White
                        # HorizontalAlignment "Center" of columns A-X and AA-AW
                        $WorkSheet.Cells["A:X"].Style.HorizontalAlignment = "Center"
                        $WorkSheet.Cells["AA:AW"].Style.HorizontalAlignment = "Center"
                        # ConditionalFormatting
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["F:F"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("4765445b-32c6-49b0-83e6-1d93765276ca",$F1)))' -BackgroundColor Red # ApplicationId
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["G:G"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("OfficeHome",$G1)))' -BackgroundColor Red # AppDisplayName
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["H:H"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("Browser",$H1)))' -BackgroundColor Red # ClientAppUsed
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["J:J"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("notApplied",$J1)))' -BackgroundColor Red # ConditionalAccessStatus
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["L:L"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("True",$L1)))' -BackgroundColor Red # IsInteractive
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["S:S"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("atRisk",$S1)))' -BackgroundColor Red # RiskState
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["X:X"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("0",$X1)))' -BackgroundColor Red # ErrorCode
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["X:X"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("50074",$X1)))' -BackgroundColor Red # ErrorCode
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["X:X"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("50140",$X1)))' -BackgroundColor Red # ErrorCode
                        Add-ConditionalFormatting -Address $WorkSheet.Cells["X:X"] -WorkSheet $WorkSheet -RuleType 'Expression' 'NOT(ISERROR(FIND("53000",$X1)))' -BackgroundColor Red # ErrorCode

                        # Iterating over the ASN-Blacklist HashTable
                        foreach ($ASN in $AsnBlacklist_HashTable.Keys) {
                            $ConditionValue = 'NOT(ISERROR(FIND("{0}",$AS1)))' -f $ASN
                            Add-ConditionalFormatting -Address $WorkSheet.Cells["AS:AT"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                        }

                        # Iterating over the Country-Blacklist HashTable
                        foreach ($Country in $CountryBlacklist_HashTable.Keys) {
                            $ConditionValue = 'NOT(ISERROR(FIND("{0}",$AP1)))' -f $Country
                            Add-ConditionalFormatting -Address $WorkSheet.Cells["AO:AP"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                        }

                    }
                }
            }
        }

        # Hunt

        # CSV
        $Import | Group-Object UserId, UserPrincipalName, Country, "Country Name", ASN, OrgName, Region, City | Select-Object @{Name = 'UserId'; Expression = { $_.Values[0] } }, @{Name = 'UserPrincipalName'; Expression = { $_.Values[1] } }, @{Name = 'Country'; Expression = { $_.Values[2] } }, @{Name = 'Country Name'; Expression = { $_.Values[3] } }, @{Name = 'ASN'; Expression = { $_.Values[4] } }, @{Name = 'OrgName'; Expression = { $_.Values[5] } }, @{Name = 'Region'; Expression = { $_.Values[6] } }, @{Name = 'City'; Expression = { $_.Values[7] } }, Count | Sort-Object Count -Descending | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM_Hunt.csv" -NoTypeInformation -Encoding UTF8

        # XLSX
        if (Get-Module -ListAvailable -Name ImportExcel) {
            if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM_Hunt.csv") {
                if ([int](& $xsv count "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM_Hunt.csv") -gt 0) {
                    $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\CSV\AiTM_Hunt.csv" -Delimiter ","
                    $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\AiTM\XLSX\AiTM_Hunt.xlsx" -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "AiTM_Hunt" -CellStyleSB {
                        param($WorkSheet)
                        # BackgroundColor and FontColor for specific cells of TopRow
                        $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                        Set-Format -Address $WorkSheet.Cells["A1:I1"] -BackgroundColor $BackgroundColor -FontColor White
                        # HorizontalAlignment "Center" of columns A-I
                        $WorkSheet.Cells["A:I"].Style.HorizontalAlignment = "Center"

                        # Iterating over the ASN-Blacklist HashTable
                        foreach ($ASN in $AsnBlacklist_HashTable.Keys) {
                            $ConditionValue = 'NOT(ISERROR(FIND("{0}",$E1)))' -f $ASN
                            Add-ConditionalFormatting -Address $WorkSheet.Cells["E:F"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                        }

                        # Iterating over the Country-Blacklist HashTable
                        foreach ($Country in $CountryBlacklist_HashTable.Keys) {
                            $ConditionValue = 'NOT(ISERROR(FIND("{0}",$D1)))' -f $Country
                            Add-ConditionalFormatting -Address $WorkSheet.Cells["C:D"] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                        }

                    }
                }
            }
        }
    }

    # Device Code Phishing --> Detect Malicious OAuth Device Code Phishing --> not seen yet
    # https://login.microsoftonline.com/common/oauth2/devicecode?api-version=1.0
    $Import = $Results | Where-Object { $_.ClientAppUsed -eq "Mobile Apps and Desktop clients" } | Where-Object { $_.AuthenticationProtocol -eq "deviceCode" } | Where-Object { $_.AuthenticationRequirement -eq "singleFactorAuthentication" }
    $Count = ($Import | Measure-Object).Count
    $Users = ($Import | Select-Object UserId -Unique | Measure-Object).Count

    if ($Count -ge 1) {
        Write-Host "[Alert] Potential Device Code Authentication (PRT Phishing) detected ($Users account credentials, $Count events)" -ForegroundColor Red
        New-Item "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\CSV" -ItemType Directory -Force | Out-Null
        New-Item "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\XLSX" -ItemType Directory -Force | Out-Null

        # CSV
        $Import | Export-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\CSV\DeviceCode.csv" -NoTypeInformation -Encoding UTF8

        # XLSX
        if (Get-Module -ListAvailable -Name ImportExcel) {
            if (Test-Path "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\CSV\DeviceCode.csv") {
                if ([int](& $xsv count "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\CSV\DeviceCode.csv") -gt 0) {
                    $IMPORT = Import-Csv "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\CSV\DeviceCode.csv" -Delimiter ","
                    $IMPORT | Export-Excel -Path "$OUTPUT_FOLDER\SignInLogsGraph\DeviceCode\XLSX\DeviceCode.xlsx" -NoNumberConversion * -FreezePane 2, 5 -BoldTopRow -AutoSize -AutoFilter -WorkSheetname "Device Code Auth" -CellStyleSB {
                        param($WorkSheet)
                        # BackgroundColor and FontColor for specific cells of TopRow
                        $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
                        Set-Format -Address $WorkSheet.Cells["A1:BO"] -BackgroundColor $BackgroundColor -FontColor White
                        # HorizontalAlignment "Center" of columns A-BO
                        $WorkSheet.Cells["A:BO"].Style.HorizontalAlignment = "Center"
                    }
                }
            }
        }
    }

    # PRT = Primary Refresh Token

    # Detection Methodology
    # ClientAppUsed: Mobile Apps and Desktop clients
    # AuthenticationProtocol: deviceCode
    # AuthenticationRequirement: singleFactorAuthentication
    # AdditionalDetails: MFA requirement satisfied by claim in the token
    # OriginalTransferMethod: deviceCodeFlow
    # AppId: 29d9ed98-a469-4536-ade2-f981bc1d605e // Microsoft Authentication Broker

    # https://twitter.com/ITguySoCal/status/1761184877406572834
    # https://dirkjanm.io/phishing-for-microsoft-entra-primary-refresh-tokens/
    # https://www.invictus-ir.com/news/do-not-use-the-get-mgauditlogsignin-for-your-investigations
    # https://www.inversecos.com/2022/12/how-to-detect-malicious-oauth-device.html
    # https://github.com/pushsecurity/saas-attacks/blob/main/techniques/device_code_phishing/examples/microsoft.md

    #############################################################################################################################################################################################

    # Blacklisting

    # Applications

    # Create HashTable and import 'Application-Blacklist.csv'
    $ApplicationBlacklist_HashTable = [ordered]@{}
    if (Test-Path "$PSScriptRoot\Blacklists\Application-Blacklist.csv") {
        if ([int](& $xsv count "$PSScriptRoot\Blacklists\Application-Blacklist.csv") -gt 0) {
            Import-Csv "$PSScriptRoot\Blacklists\Application-Blacklist.csv" -Delimiter "," | ForEach-Object { $ApplicationBlacklist_HashTable[$_.AppId] = $_.AppDisplayName, $_.Severity }

            # Count Ingested Properties
            $Count = $ApplicationBlacklist_HashTable.Count
            Write-Output "[Info]  Initializing 'Application-Blacklist.csv' Lookup Table ($Count) ..."

            $Data = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter ","

            # Iterating over the HashTable
            foreach ($AppId in $ApplicationBlacklist_HashTable.Keys) {
                $Import = $Data | Where-Object { $_.AppId -eq "$AppId" }
                $Count = [string]::Format('{0:N0}', ($Import | Measure-Object).Count)
                if ($Count -gt 0) {
                    $AppDisplayName = $ApplicationBlacklist_HashTable["$AppId"][0]
                    $Severity = $ApplicationBlacklist_HashTable["$AppId"][1]
                    Write-Host "[Alert] Suspicious OAuth Application detected: $AppDisplayName ($Count)" -ForegroundColor $Severity
                }
            }
        }
    }

    # ASN

    # Create HashTable and import 'ASN-Blacklist.csv'
    $AsnBlacklist_HashTable = [ordered]@{}
    if (Test-Path "$PSScriptRoot\Blacklists\ASN-Blacklist.csv") {
        if ([int](& $xsv count "$PSScriptRoot\Blacklists\ASN-Blacklist.csv") -gt 0) {
            Import-Csv "$PSScriptRoot\Blacklists\ASN-Blacklist.csv" -Delimiter "," | ForEach-Object { $AsnBlacklist_HashTable[$_.ASN] = $_.OrgName, $_.Info }

            # Count Ingested Properties
            $Count = $AsnBlacklist_HashTable.Count
            Write-Output "[Info]  Initializing 'ASN-Blacklist.csv' Lookup Table ($Count) ..."

            $Data = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter ","

            # Iterating over the HashTable
            foreach ($ASN in $AsnBlacklist_HashTable.Keys) {
                $Import = $Data | Where-Object { $_.ASN -eq "$ASN" }
                $Count = [string]::Format('{0:N0}', ($Import | Measure-Object).Count)
                if ($Count -gt 0) {
                    $OrgName = $AsnBlacklist_HashTable["$ASN"][0]
                    Write-Host "[Alert] Suspicious ASN detected: AS$ASN - $OrgName ($Count)" -ForegroundColor Red
                }
            }
        }
    }

    # Country

    # Create HashTable and import 'Country-Blacklist.csv'
    $CountryBlacklist_HashTable = [ordered]@{}
    if (Test-Path "$PSScriptRoot\Blacklists\Country-Blacklist.csv") {
        if ([int](& $xsv count "$PSScriptRoot\Blacklists\Country-Blacklist.csv") -gt 0) {
            Import-Csv "$PSScriptRoot\Blacklists\Country-Blacklist.csv" -Delimiter "," | ForEach-Object { $CountryBlacklist_HashTable[$_."Country Name"] = $_.Country }

            # Count Ingested Properties
            $Count = $CountryBlacklist_HashTable.Count
            Write-Output "[Info]  Initializing 'Country-Blacklist.csv' Lookup Table ($Count) ..."

            $Data = Import-Csv -Path "$OUTPUT_FOLDER\SignInLogsGraph\CSV\Hunt.csv" -Delimiter ","

            # Iterating over the HashTable
            foreach ($CountryName in $CountryBlacklist_HashTable.Keys) {
                $Import = $Data | Where-Object { $_."Country Name" -eq "$CountryName" }
                $Count = [string]::Format('{0:N0}', ($Import | Measure-Object).Count)
                if ($Count -gt 0) {
                    Write-Host "[Alert] Suspicious Country detected: $CountryName ($Count)" -ForegroundColor Red
                }
            }
        }
    }

    $EndTime_Analytics = (Get-Date)
    $Time_Analytics = ($EndTime_Analytics - $StartTime_Analytics)
('ADSignInLogsGraph Analytics duration:       {0} h {1} min {2} sec' -f $Time_Analytics.Hours, $Time_Analytics.Minutes, $Time_Analytics.Seconds) >> "$OUTPUT_FOLDER\Stats.txt"

}

#endregion Analytics

#############################################################################################################################################################################################
#############################################################################################################################################################################################

# Main
Start-Processing
Get-IPLocation
Get-Stats
Get-Analytics

#endregion
#region end
end {
    #region Footer

    # Get End Time
    $endTime = (Get-Date)

    # Echo Time elapsed
    Write-Output ""
    Write-Output "FINISHED!"

    $Time = ($endTime - $startTime)
    $ElapsedTime = ('Overall analysis duration: {0} h {1} min {2} sec' -f $Time.Hours, $Time.Minutes, $Time.Seconds)
    Write-Output "$ElapsedTime"

    # Stop logging
    if ($Transcript) {
        Stop-Transcript
    }
    # MessageBox UI
    $MessageBody = "Status: Sign-In Logs Analysis completed."
    $MessageTitle = "ADSignInLogsGraph-Analyzer.ps1 (https://lethal-forensics.com/)"
    $ButtonType = "OK"
    $MessageIcon = "Information"
    $Result = [System.Windows.Forms.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)

    if ($Result -eq "OK" ) {
        # Reset Progress Preference
        $Global:ProgressPreference = $OriginalProgressPreference

        # Reset Windows Title
        $Host.UI.RawUI.WindowTitle = "$DefaultWindowsTitle"
        Exit
    }

    #endregion end
}
}