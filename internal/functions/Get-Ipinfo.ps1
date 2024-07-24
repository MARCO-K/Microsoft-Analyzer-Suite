
function Get-Ipinfo {   
    <#
    .SYNOPSIS
    Retrieves IP information from a list of IP addresses using the IPInfo API.

    .DESCRIPTION
    The get-Ipinfo function retrieves IP information from a list of IP addresses using the IPInfo API. It reads the IP addresses from a text file and sends requests to the IPInfo API to get the information for each IP address. The function then processes the API response and returns the IP information as objects.

    .PARAMETER Token
    The API token to authenticate the requests to the IPInfo API. By default, it is set to '1210be81001a80'.

    .PARAMETER AddressList
    The path to the text file that contains the list of IP addresses. By default, it is set to 'C:\temp\IpAddress\IP.txt'.

    .PARAMETER Fields
    The fields to include in the API response. By default, it is set to 'ip,hostname,anycast,city,region,country,country_name,loc,org,postal,timezone'.

    .EXAMPLE
    get-Ipinfo -token 'your-api-token' -list 'C:\path\to\ip-list.txt' -fields 'ip,city,country'

    This example retrieves IP information for the IP addresses listed in the 'ip-list.txt' file using the specified API token and includes only the 'ip', 'city', and 'country' fields in the API response.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Token,
    
        [Parameter(Mandatory = $true)]
        [string]$AddressList,
    
        [string]$Fields = 'ip,hostname,anycast,city,region,country,country_name,loc,org,postal,timezone'
    )
    
    begin {
    
        # Check if the IP list exists and is not empty
        if (-not (Test-Path -Path $AddressList -PathType Leaf)) {
            Write-PSFMessage -Level Error -Message "IP list not found: $AddressList"
            throw "IP list not found: $AddressList"
        }
    
        $ipList = Get-Content $AddressList
        if ($ipList.Count -eq 0) {
            Write-PSFMessage -Level Error -Message "IP list is empty: $AddressList"
            throw "IP list is empty: $AddressList"
        }
    
        $ipinfoPath = "$env:LOCALAPPDATA\ipinfo\ipinfo.exe"
        if (-not (Test-Path $ipinfoPath)) {
            Write-PSFMessage -Level Error -Message "IPInfo executable not found: $ipinfoPath"
            throw "IPInfo executable not found: $ipinfoPath"
        }
    
        Write-PSFMessage -Level Verbose -Message "Starting IP info retrieval for $(($ipList | Select-Object -Unique).Count) unique IP addresses"
    }
    
    process {
        try {
            Write-PSFMessage -Level Verbose -Message "Executing IPInfo command"
            $result = $ipList | Select-Object -Unique -Skip 1 | & $ipinfoPath -t $Token -f $Fields
            $rawData = $result | ConvertFrom-Csv -Delimiter ','
    
            Write-PSFMessage -Level Verbose -Message "Processing $(($rawData | Measure-Object).Count) IP records"
            $ipData = foreach ($data in $rawData) {
                [PSCustomObject]@{
                    ASN          = ($data.org -split ' ', 2)[0]
                    Organization = ($data.org -split ' ', 2)[1]
                    IP           = [System.Net.IPAddress]$data.ip
                    Hostname     = $data.hostname
                    Anycast      = [System.Convert]::ToBoolean($data.anycast)
                    City         = $data.city
                    Region       = $data.region
                    Country      = $data.country
                    CountryName  = $data.country_name
                    Location     = $data.loc
                    PostalCode   = $data.postal
                    Timezone     = $data.timezone
                }
            }
        } catch {
            Write-PSFMessage -Level Error -Message "Error occurred while processing IP data" -ErrorRecord $_
            throw
        }
    }
    
    end {
        Write-PSFMessage -Level Verbose -Message "IP info retrieval completed successfully"
        $ipData
    }
}