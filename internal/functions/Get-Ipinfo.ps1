
function Get-Ipinfo{   
    <#
    .SYNOPSIS
    Retrieves IP information from a list of IP addresses using the IPInfo API.

    .DESCRIPTION
    The get-Ipinfo function retrieves IP information from a list of IP addresses using the IPInfo API. It reads the IP addresses from a text file and sends requests to the IPInfo API to get the information for each IP address. The function then processes the API response and returns the IP information as objects.

    .PARAMETER token
    The API token to authenticate the requests to the IPInfo API. By default, it is set to '1210be81001a80'.

    .PARAMETER list
    The path to the text file that contains the list of IP addresses. By default, it is set to 'C:\temp\IpAddress\IP.txt'.

    .PARAMETER fields
    The fields to include in the API response. By default, it is set to 'ip,hostname,anycast,city,region,country,country_name,loc,org,postal,timezone'.

    .EXAMPLE
    get-Ipinfo -token 'your-api-token' -list 'C:\path\to\ip-list.txt' -fields 'ip,city,country'

    This example retrieves IP information for the IP addresses listed in the 'ip-list.txt' file using the specified API token and includes only the 'ip', 'city', and 'country' fields in the API response.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$token,
        [Parameter(Mandatory=$true)]
        [string]$addresslist,
        [string]$fields = 'ip,hostname,anycast,city,region,country,country_name,loc,org,postal,timezone'
    )
    
    begin {
            # Check if the IP list exists and is not empty
            if (Test-Path -Path $list -PathType Leaf)
                {$ip_list = Get-Content $list}
            else {
                'IP list not found.'
                break
            }
            if(-not($ip_list.Count -gt 1)) {
                'IP list is empty.'
                break
            }
            
            if (Test-Path "$env:LOCALAPPDATA\ipinfo\ipinfo.exe") {
                $ipinfo = "$env:LOCALAPPDATA\ipinfo\ipinfo.exe"
               }
            else {
                'IPInfo executable not found.'
                break
            }
    }
    
    process {

        try {
            $result = $ip_list | Select-Object -Unique -Skip 1 | & $ipinfo -t $token -f $fields 
        }
        catch {
            Write-Error $_.Exception.Message
        }
        
        $rawdata = $result | ConvertFrom-csv -Delimiter ','
        $ipdata = foreach ($data in $rawdata) {
            [PSCustomObject]@{
                asn = ($data.org -split ' ',2)[0]
                org = ($data.org -split ' ',2)[1]
                ip = [System.Net.IPAddress]$data.ip
                hostname = $data.hostname
                anycast = $data.anycast
                city = $data.city
                region = $data.region
                country = $data.country
                country_name = $data.country_name
                gps_loc = $data.loc
                postalcode = $data.postal
                timezone = $data.timezone
        }}
    }
    
    end {
        $ipdata
    }
}