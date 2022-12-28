
# Generic function to avoid code duplication
function Set-SC365PropertiesFromConfigJson
{
    [CmdLetBinding()]
    Param
    (
        [psobject] $InputObject,
        [psobject] $Json,
        [SC365.MailRouting] $Routing,
        [SC365.ConfigOption[]] $Option,
        [SC365.GeoRegion] $Region
    )

    # Set all properties that aren't version specific
    $json.psobject.properties | Foreach-Object {
        if ($_.Name -notin @("Name", "Option", "Routing", "Region"))
        { $InputObject.$($_.Name) = $_.Value }
    }

    if($routing -and $json.Routing)
    {
        $json.Routing.$Routing.psobject.properties | Foreach-Object {
            $InputObject.$($_.Name) = $_.Value
        }
    }

    if($Option -and $json.Option)
    {
        $Option | Where-Object {$json.Option.$_} | ForEach-Object{
            $Json.Option.$_.psobject.properties | ForEach-Object{
                $InputObject.$($_.Name) = $_.Value
            }
        }
    }


    if($Region -and $json.Region)
    {
        $json.Region.$Region.psobject.properties | %Foreach-Object {
            $InputObject.$($_.Name) = $_.Value
        }
    }
}

# Essentially a factory function for either an empty
# settings object, filled with necessary attributes to identify
# the O365 object (i.e. the Name), or version specific settings.
function Get-SC365InboundConnectorSettings
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $routing,
        $option
    )

    Write-Verbose "Loading inbound connector settings for routingtype $Routing"
    $inBoundRaw = (Get-Content "$PSScriptRoot\..\ExOConfig\Connectors\InBound.json" -Raw|Convertfrom-Json -AsHashtable)
    $ret = $inBoundRaw.routing.($routing.Tolower())

    return $ret
}

function Get-SC365OutboundConnectorSettings
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $routing,
        $option
    )

    Write-Verbose "Loading outbound connector settings"
    $outBoundRaw = (Get-Content "$PSScriptRoot\..\ExOConfig\Connectors\OutBound.json" -Raw|Convertfrom-Json -AsHashtable)
    $ret= $outBoundRaw.routing.($routing.ToLower())
    return $ret
}

function Get-SC365TransportRuleSettings
{
    [CmdLetBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $routing,
        [Parameter(Mandatory = $true)]
        [string] $file,
        [switch] $IncludeSkipped
    )

    begin {
        $ret = $null
        $raw = $null
    }
    process {
        $raw = (Get-Content $File -Raw|Convertfrom-Json -AsHashtable)
        $ret = $raw.routing.($routing.ToLower())
    }
    end {
        return $ret    
    }
}
function Get-SC365CloudConfig
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [SC365.GeoRegion] $Region
    )

    Write-Verbose "Loading inbound connector settings for region $Region"

    $json = ConvertFrom-Json (Get-Content -Path "$PSScriptRoot\..\ExOConfig\CloudConfig\GeoRegion.json" -Raw)

    $ret = [SC365.PoliciesAntiSpamSettings]::new($json.Name, $Region)

    Set-SC365PropertiesFromConfigJson $ret -Json $json -Option $Option -Region $Region

    return $ret
}

function Resolve-SC365IPv4Address {
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'DNS Name'
        )]
        $fqdn
    )
    try {
        $ret = [System.Net.Dns]::GetHostAddresses($fqdn) |where-object AddressFamily -eq 'Internetwork'|select-object -expandproperty ipaddresstostring
        return = $ret    
    } catch {
        '--- IP4 Address could not be resolved ---'
    }
}

function Resolve-SC365IPv6Address {
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'DNS Name'
        )]
        $fqdn
    )
    try {
        $ret = [System.Net.Dns]::GetHostAddresses($fqdn) |where-object AddressFamily -eq 'InterNetworkV6'|select-object -expandproperty ipaddresstostring
        return = $ret    
    } catch {
        '--- IP6 Address could not be resolved ---'
    }
}

function Resolve-SC365DNSName {
    [CmdLetBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'IP4 or IPv6 IP address'
        )]
        [String]$ipAddress
    )
    $DnsName = $Null

    <#
    Determine Addressfamily with:
    [System.Net.DNs]::GetHostAddresses($ipv6)
    #>
    try {
        Write-Verbose "Resolving $iPAddress to HostName"
        $DNSName = [System.Net.Dns]::GetHostEntry($ipaddress).Hostname
     } 
     catch {
        $DnsName = "---IP could not be resolved---"
    }
    return $DNSName
}

<#
.SYNOPSIS
    Validates if a given strig (domainname) is the tenant default maildomain
.DESCRIPTION
    It takes the string and prooves if the domain is a member of the tenant. If this fails, the function returns an error and stops.
    If the domain is a tenant-member it returns true/false if its the domain default.
.NOTES
    - none -
.LINK
    - none -
.EXAMPLE
    Confirm-TenantDefaultDomain -Domain 'contoso.eu'
    Returns either an error if the domain is NOT in the tenant, of true or false.
#>
function Confirm-SC365TenantDefaultDomain {
    param (
        [CmdLetBinding(Mandatory = $true)]
        [String]$ValidationDomain
    )

    begin {
        $TenantDomains = Get-AcceptedDomain
        $TenantDefaultDomain = $TenantDomains.Where({$_.Default -eq $true}).DomainName     
    }
    process {
        If (!($TenantDomains.DomainName -contains $ValidationDomain)) {
            Write-Error "$ValidationDomain is not member of the connected tenant. Retry using only tenant-domains (Use CmdLet: Get-AcceptedDomain)"
            break
        } else {
            if ($Validationdomain -eq $TenantDefaultDomain) {
                return $true
            } else {
                return $false
            }    
        }
    }
    end {
    }
}


# SIG # Begin signature block
# MIIL/AYJKoZIhvcNAQcCoIIL7TCCC+kCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCTVVaP3Y/2HSMO
# C6kQnkGzHkm3Uj/b6/AIzOO3AP35DaCCCUAwggSZMIIDgaADAgECAhBxoLc2ld2x
# r8I7K5oY7lTLMA0GCSqGSIb3DQEBCwUAMIGpMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMdGhhd3RlLCBJbmMuMSgwJgYDVQQLEx9DZXJ0aWZpY2F0aW9uIFNlcnZpY2Vz
# IERpdmlzaW9uMTgwNgYDVQQLEy8oYykgMjAwNiB0aGF3dGUsIEluYy4gLSBGb3Ig
# YXV0aG9yaXplZCB1c2Ugb25seTEfMB0GA1UEAxMWdGhhd3RlIFByaW1hcnkgUm9v
# dCBDQTAeFw0xMzEyMTAwMDAwMDBaFw0yMzEyMDkyMzU5NTlaMEwxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwx0aGF3dGUsIEluYy4xJjAkBgNVBAMTHXRoYXd0ZSBTSEEy
# NTYgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAm1UCTBcF6dBmw/wordPA/u/g6X7UHvaqG5FG/fUW7ZgHU/q6hxt9nh8BJ6u5
# 0mfKtxAlU/TjvpuQuO0jXELvZCVY5YgiGr71x671voqxERGTGiKpdGnBdLZoh6eD
# MPlk8bHjOD701sH8Ev5zVxc1V4rdUI0D+GbNynaDE8jXDnEd5GPJuhf40bnkiNIs
# KMghIA1BtwviL8KA5oh7U2zDRGOBf2hHjCsqz1v0jElhummF/WsAeAUmaRMwgDhO
# 8VpVycVQ1qo4iUdDXP5Nc6VJxZNp/neWmq/zjA5XujPZDsZC0wN3xLs5rZH58/eW
# XDpkpu0nV8HoQPNT8r4pNP5f+QIDAQABo4IBFzCCARMwLwYIKwYBBQUHAQEEIzAh
# MB8GCCsGAQUFBzABhhNodHRwOi8vdDIuc3ltY2IuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwMgYDVR0fBCswKTAnoCWgI4YhaHR0cDovL3QxLnN5bWNiLmNvbS9UaGF3
# dGVQQ0EuY3JsMB0GA1UdJQQWMBQGCCsGAQUFBwMCBggrBgEFBQcDAzAOBgNVHQ8B
# Af8EBAMCAQYwKQYDVR0RBCIwIKQeMBwxGjAYBgNVBAMTEVN5bWFudGVjUEtJLTEt
# NTY4MB0GA1UdDgQWBBRXhptUuL6mKYrk9sLiExiJhc3ctzAfBgNVHSMEGDAWgBR7
# W0XPr87Lev0xkhpqtvNG61dIUDANBgkqhkiG9w0BAQsFAAOCAQEAJDv116A2E8dD
# /vAJh2jRmDFuEuQ/Hh+We2tMHoeei8Vso7EMe1CS1YGcsY8sKbfu+ZEFuY5B8Sz2
# 0FktmOC56oABR0CVuD2dA715uzW2rZxMJ/ZnRRDJxbyHTlV70oe73dww78bUbMyZ
# NW0c4GDTzWiPKVlLiZYIRsmO/HVPxdwJzE4ni0TNB7ysBOC1M6WHn/TdcwyR6hKB
# b+N18B61k2xEF9U+l8m9ByxWdx+F3Ubov94sgZSj9+W3p8E3n3XKVXdNXjYpyoXY
# RUFyV3XAeVv6NBAGbWQgQrc6yB8dRmQCX8ZHvvDEOihU2vYeT5qiGUOkb0n4/F5C
# ICiEi0cgbjCCBJ8wggOHoAMCAQICEF0xOuf5lHR9Mf0X/F6tAjYwDQYJKoZIhvcN
# AQELBQAwTDELMAkGA1UEBhMCVVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEmMCQG
# A1UEAxMddGhhd3RlIFNIQTI1NiBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwMzE2MDAw
# MDAwWhcNMjMwMzE2MjM1OTU5WjBdMQswCQYDVQQGEwJDSDEPMA0GA1UECAwGQWFy
# Z2F1MREwDwYDVQQHDAhOZXVlbmhvZjEUMBIGA1UECgwLU0VQUG1haWwgQUcxFDAS
# BgNVBAMMC1NFUFBtYWlsIEFHMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAoTng2flWvxhwRxO/mTS8XJbbqTMG3097ZRxrB+OfKq9ilZ4K45fHL1deJ/bh
# E24FE4Rzd6pXgQdsD2/5w6sTWZ7gdNy3tvJEzQBALOht4UDJ6Ld7J4Nu9v3relQP
# 5euW4JBL4vzSfzIKZw7AH7B5f5iPYJjuzcdl9sCvNPEEejkAoQOn1fHKgJz+1RA6
# rq/L7B1s458uGs0yQMMKwFZIuqKPHe3UT5+sBgXu1do7n8UQZN5v86Neo6miig0v
# fwWYImHQIpvLT3SwEk/nBKJ6efU12GYE4pEv7QrduiajOhiZCxylPkB9ajikVCb7
# xsbhB01pPwBXgAc9vjN/bdvFZQIDAQABo4IBajCCAWYwCQYDVR0TBAIwADAfBgNV
# HSMEGDAWgBRXhptUuL6mKYrk9sLiExiJhc3ctzAdBgNVHQ4EFgQU8p/VNr/y8AxO
# eDZl4HNdfm6GnyMwKwYDVR0fBCQwIjAgoB6gHIYaaHR0cDovL3RsLnN5bWNiLmNv
# bS90bC5jcmwwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMG4G
# A1UdIARnMGUwYwYGZ4EMAQQBMFkwJgYIKwYBBQUHAgEWGmh0dHBzOi8vd3d3LnRo
# YXd0ZS5jb20vY3BzMC8GCCsGAQUFBwICMCMMIWh0dHBzOi8vd3d3LnRoYXd0ZS5j
# b20vcmVwb3NpdG9yeTBXBggrBgEFBQcBAQRLMEkwHwYIKwYBBQUHMAGGE2h0dHA6
# Ly90bC5zeW1jZC5jb20wJgYIKwYBBQUHMAKGGmh0dHA6Ly90bC5zeW1jYi5jb20v
# dGwuY3J0MA0GCSqGSIb3DQEBCwUAA4IBAQB2zM1TxEwHrD3KWrJcbcSNae+KLtqK
# A3T/wE3Bkh6os/1pWseEMQirPRG74okbvb1KCu0uBzs56eq5toMuIuntggICxkFR
# p7dr4xLuJ+BUu4MoFJ2wwQf6hG43WIYx+cCHKSDQ/mW11GU18O9CUgWKoYJxFpoZ
# ULd+qj+ls+21tbfuizIjFOCYHiERazWOjy4MQfHDmrIHD/cIaQGF/kY1z2ZQ4d4Q
# EyfM56yaWruH1HZDkVo8fvgb0VfEAbbP+sIDsLBzDDVCBvWT7d34Ua9SZUbEkTjl
# 3atWPX5qU1z5lwH1hr37yOEyywwuOEnUF/NRZ5Z1mpEY2x8WP5pmyTDGMYICEjCC
# Ag4CAQEwYDBMMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhhd3RlLCBJbmMuMSYw
# JAYDVQQDEx10aGF3dGUgU0hBMjU2IENvZGUgU2lnbmluZyBDQQIQXTE65/mUdH0x
# /Rf8Xq0CNjANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAM
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCL84uKQqk6Aryo1R0WqmyqR2Fu
# h8F2ICdpDV2bQFp5bzANBgkqhkiG9w0BAQEFAASCAQBk8xh9VEV0jrhf2e1x/Uoo
# w9VGfWZXGs1WipINq/zreQhcIvLOXWo257n/k5VNSFAd0WJpxbJcd6XAvBl56CeH
# //nr9MlXFN5FWyn4I9f3amdttkyXW2d6RPbaZf6EgTdiWfzWP6A9Ax55LhUrp5SR
# 6kBHrPbMUN9+o9xtdfVWTELERUGh1ksrYCyXulLZA/MmU3TWC3PZiXj+Mh7zsK5N
# PxH2ZVAxjNEkG2B26KsBWn2umg7rLAYqizdVw8Wbjzr0RRuzPI9Yp/PSlBPstAju
# PxATqOK1C9Zuv54hGz2Zo73Sw5NJVapAaon24RvEAw1dZ00IyxOtxtGANhF/S8R7
# SIG # End signature block
