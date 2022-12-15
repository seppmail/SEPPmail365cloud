
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

function Resolve-IPv4Address {
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'DNS Name'
        )]
        $fqdn
    )

    $ret = [System.Net.Dns]::GetHostAddresses($fqdn) |where-object AddressFamily -eq 'Internetwork'|select-object -expandproperty ipaddresstostring
    return = $ret
}

function Resolve-IPv4Address {
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'DNS Name'
        )]
        $fqdn
    )

    $ret = [System.Net.Dns]::GetHostAddresses($fqdn) |where-object AddressFamily -eq 'Internetwork'|select-object -expandproperty ipaddresstostring
    return = $ret
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


# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUORvc5QCBHI6Q76eRGhUJXVl5
# lvKggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
# AQsFADCBqTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEoMCYG
# A1UECxMfQ2VydGlmaWNhdGlvbiBTZXJ2aWNlcyBEaXZpc2lvbjE4MDYGA1UECxMv
# KGMpIDIwMDYgdGhhd3RlLCBJbmMuIC0gRm9yIGF1dGhvcml6ZWQgdXNlIG9ubHkx
# HzAdBgNVBAMTFnRoYXd0ZSBQcmltYXJ5IFJvb3QgQ0EwHhcNMTMxMjEwMDAwMDAw
# WhcNMjMxMjA5MjM1OTU5WjBMMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhhd3Rl
# LCBJbmMuMSYwJAYDVQQDEx10aGF3dGUgU0hBMjU2IENvZGUgU2lnbmluZyBDQTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJtVAkwXBenQZsP8KK3TwP7v
# 4Ol+1B72qhuRRv31Fu2YB1P6uocbfZ4fASerudJnyrcQJVP0476bkLjtI1xC72Ql
# WOWIIhq+9ceu9b6KsRERkxoiqXRpwXS2aIengzD5ZPGx4zg+9NbB/BL+c1cXNVeK
# 3VCNA/hmzcp2gxPI1w5xHeRjyboX+NG55IjSLCjIISANQbcL4i/CgOaIe1Nsw0Rj
# gX9oR4wrKs9b9IxJYbpphf1rAHgFJmkTMIA4TvFaVcnFUNaqOIlHQ1z+TXOlScWT
# af53lpqv84wOV7oz2Q7GQtMDd8S7Oa2R+fP3llw6ZKbtJ1fB6EDzU/K+KTT+X/kC
# AwEAAaOCARcwggETMC8GCCsGAQUFBwEBBCMwITAfBggrBgEFBQcwAYYTaHR0cDov
# L3QyLnN5bWNiLmNvbTASBgNVHRMBAf8ECDAGAQH/AgEAMDIGA1UdHwQrMCkwJ6Al
# oCOGIWh0dHA6Ly90MS5zeW1jYi5jb20vVGhhd3RlUENBLmNybDAdBgNVHSUEFjAU
# BggrBgEFBQcDAgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgEGMCkGA1UdEQQiMCCk
# HjAcMRowGAYDVQQDExFTeW1hbnRlY1BLSS0xLTU2ODAdBgNVHQ4EFgQUV4abVLi+
# pimK5PbC4hMYiYXN3LcwHwYDVR0jBBgwFoAUe1tFz6/Oy3r9MZIaarbzRutXSFAw
# DQYJKoZIhvcNAQELBQADggEBACQ79degNhPHQ/7wCYdo0ZgxbhLkPx4flntrTB6H
# novFbKOxDHtQktWBnLGPLCm37vmRBbmOQfEs9tBZLZjgueqAAUdAlbg9nQO9ebs1
# tq2cTCf2Z0UQycW8h05Ve9KHu93cMO/G1GzMmTVtHOBg081ojylZS4mWCEbJjvx1
# T8XcCcxOJ4tEzQe8rATgtTOlh5/03XMMkeoSgW/jdfAetZNsRBfVPpfJvQcsVncf
# hd1G6L/eLIGUo/flt6fBN591ylV3TV42KcqF2EVBcld1wHlb+jQQBm1kIEK3Osgf
# HUZkAl/GR77wxDooVNr2Hk+aohlDpG9J+PxeQiAohItHIG4wggSfMIIDh6ADAgEC
# AhBdMTrn+ZR0fTH9F/xerQI2MA0GCSqGSIb3DQEBCwUAMEwxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwx0aGF3dGUsIEluYy4xJjAkBgNVBAMTHXRoYXd0ZSBTSEEyNTYg
# Q29kZSBTaWduaW5nIENBMB4XDTIwMDMxNjAwMDAwMFoXDTIzMDMxNjIzNTk1OVow
# XTELMAkGA1UEBhMCQ0gxDzANBgNVBAgMBkFhcmdhdTERMA8GA1UEBwwITmV1ZW5o
# b2YxFDASBgNVBAoMC1NFUFBtYWlsIEFHMRQwEgYDVQQDDAtTRVBQbWFpbCBBRzCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKE54Nn5Vr8YcEcTv5k0vFyW
# 26kzBt9Pe2UcawfjnyqvYpWeCuOXxy9XXif24RNuBROEc3eqV4EHbA9v+cOrE1me
# 4HTct7byRM0AQCzobeFAyei3eyeDbvb963pUD+XrluCQS+L80n8yCmcOwB+weX+Y
# j2CY7s3HZfbArzTxBHo5AKEDp9XxyoCc/tUQOq6vy+wdbOOfLhrNMkDDCsBWSLqi
# jx3t1E+frAYF7tXaO5/FEGTeb/OjXqOpoooNL38FmCJh0CKby090sBJP5wSienn1
# NdhmBOKRL+0K3bomozoYmQscpT5AfWo4pFQm+8bG4QdNaT8AV4AHPb4zf23bxWUC
# AwEAAaOCAWowggFmMAkGA1UdEwQCMAAwHwYDVR0jBBgwFoAUV4abVLi+pimK5PbC
# 4hMYiYXN3LcwHQYDVR0OBBYEFPKf1Ta/8vAMTng2ZeBzXX5uhp8jMCsGA1UdHwQk
# MCIwIKAeoByGGmh0dHA6Ly90bC5zeW1jYi5jb20vdGwuY3JsMA4GA1UdDwEB/wQE
# AwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBuBgNVHSAEZzBlMGMGBmeBDAEEATBZ
# MCYGCCsGAQUFBwIBFhpodHRwczovL3d3dy50aGF3dGUuY29tL2NwczAvBggrBgEF
# BQcCAjAjDCFodHRwczovL3d3dy50aGF3dGUuY29tL3JlcG9zaXRvcnkwVwYIKwYB
# BQUHAQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vdGwuc3ltY2QuY29tMCYGCCsG
# AQUFBzAChhpodHRwOi8vdGwuc3ltY2IuY29tL3RsLmNydDANBgkqhkiG9w0BAQsF
# AAOCAQEAdszNU8RMB6w9ylqyXG3EjWnvii7aigN0/8BNwZIeqLP9aVrHhDEIqz0R
# u+KJG729SgrtLgc7OenqubaDLiLp7YICAsZBUae3a+MS7ifgVLuDKBSdsMEH+oRu
# N1iGMfnAhykg0P5ltdRlNfDvQlIFiqGCcRaaGVC3fqo/pbPttbW37osyIxTgmB4h
# EWs1jo8uDEHxw5qyBw/3CGkBhf5GNc9mUOHeEBMnzOesmlq7h9R2Q5FaPH74G9FX
# xAG2z/rCA7Cwcww1Qgb1k+3d+FGvUmVGxJE45d2rVj1+alNc+ZcB9Ya9+8jhMssM
# LjhJ1BfzUWeWdZqRGNsfFj+aZskwxjGCAgEwggH9AgEBMGAwTDELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEmMCQGA1UEAxMddGhhd3RlIFNIQTI1
# NiBDb2RlIFNpZ25pbmcgQ0ECEF0xOuf5lHR9Mf0X/F6tAjYwCQYFKw4DAhoFAKB4
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFHKswGZ030V5h9QLgVFZFEM3GqxvMA0GCSqGSIb3DQEBAQUABIIBACY/lN1Q
# dOv1j2at6wNoqg5viLW7FORjd3OrAHbkegrxToxiL+SP7XBm7E1pSI/yEjpIY/E/
# KgpnldetW2ZKwK08bgQJgDz575icygygWbxVM/evEYvFUeiCIeCH70i1FtQfTsyS
# LxjUsfcJZ2co67AjRBHWAExGYHuNA7HRgvTnalb+XspLqmzpRSQX6Qv6ee9hDi9i
# tA+jYDzLKiX0uiQQfqkHcbcA40l54RQ2Hd0TMDstk6RSCDjKUZRG8ucpZcU5i9vo
# 4UCKaD+PDxMQljAO3i0xA8RlYD3iqW1EvFRl+UP8f+9odcvRqGpJzNLT49XqFN/o
# Ehpxf+tt0rS0oXw=
# SIG # End signature block
