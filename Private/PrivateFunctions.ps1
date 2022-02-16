
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
        <#[Parameter(Mandatory=$true)]
        [SC365.MailRouting] $Route,
        [Parameter(Mandatory=$true)]
        [SC365.Region] $Region,
        [SC365.ConfigOption[]] $Option#>
        [Parameter(Mandatory=$true)]
        $routing,
        $option
    )

    Write-Verbose "Loading inbound connector settings for routingtype $Routing"
    <#
    $json = ConvertFrom-Json (Get-Content -Path "$PSScriptRoot\..\ExOConfig\Connectors\Inbound.json" -Raw)

    $ret = [SC365.InboundConnectorSettings]::new($json.Name, $Route)

    Set-SC365PropertiesFromConfigJson $ret -Json $json -Route $Route -Region $Region -Option $Option
    #>
    $inBoundRaw = (Get-Content "$PSScriptRoot\..\ExOConfig\Connectors\Inbound.json" -Raw|Convertfrom-Json -AsHashtable)
    $ret = $inBoundRaw.routing.($routing.Tolower())

    return $ret
}

function Get-SC365OutboundConnectorSettings
{
    [CmdletBinding()]
    Param
    (
        <#[Parameter(Mandatory=$true)]
        [SC365.MailRouting] $Routing,
        [SC365.ConfigOption[]] $Option#>
        [Parameter(Mandatory=$true)]
        $routing,
        $option
    )

    Write-Verbose "Loading outbound connector settings"
    <#
    $json = ConvertFrom-Json (Get-Content -Path "$PSScriptRoot\..\ExOConfig\Connectors\Outbound.json" -Raw)

    $ret = [SC365.OutboundConnectorSettings]::new($json.Name, $Routing)

    Set-SC365PropertiesFromConfigJson $ret -Json $json -Routing $Routing -Option $Option
    #>
    $outBoundRaw = (Get-Content "$PSScriptRoot\..\ExOConfig\Connectors\Outbound.json" -Raw|Convertfrom-Json -AsHashtable)
    $ret= $outBoundRaw.routing.($routing.ToLower())
    return $ret
}

function Get-SC365TransportRuleSettings
{
    [CmdLetBinding()]
    Param
    (
        <#[Parameter(Mandatory=$true)]
        [SC365.MailRouting] $Route,
        [Parameter(Mandatory=$true)]
        [SC365.Region] $Region,
        [SC365.ConfigOption[]] $Option,
        [SC365.AvailableTransportRuleSettings[]] $Settings =[SC365.AvailableTransportRuleSettings]::All,
        [switch] $IncludeSkipped
        #>
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
    <#Write-Verbose "Loading transport rule settings for routingtype $Region"

    $settingsToFetch = 0
    foreach($set in $Settings)
    {$settingsToFetch = $settingsToFetch -bor $set}

    $configs = array string
    $ret = array SC365.TransportRuleSettings -Capacity $configs.Count
    $addSetting = {
        Param
        (
            [string] $FileName,
            [SC365.AvailableTransportRuleSettings] $Type
        )
        $json = ConvertFrom-Json (Get-Content "$PSScriptRoot\..\ExOConfig\Rules\$FileName" -Raw)

        $settings = [SC365.TransportRuleSettings]::new($json.Name, $Region, $Route, $Type)

        Set-SC365PropertiesFromConfigJson $settings -Json $json -Region $Region -Route $Route -Option $Option

        if(!$settings.Skip -or ($settings.Skip -and $IncludeSkipped))
        {$ret.Add($settings)}
    }

    if([SC365.AvailableTransportRuleSettings]::OutgoingHeaderCleaning -band $settingsToFetch)
    {& $addSetting "X-SM-outgoing-header-cleaning.json" "OutgoingHeaderCleaning"}

    if([SC365.AvailableTransportRuleSettings]::DecryptedHeaderCleaning -band $settingsToFetch)
    {& $addSetting "X-SM-decrypted-header-cleaning.json" "DecryptedHeaderCleaning"}

    if([SC365.AvailableTransportRuleSettings]::EncryptedHeaderCleaning -band $settingsToFetch)
    {& $addSetting "X-SM-encrypted-header-cleaning.json" "EncryptedHeaderCleaning"}

    if([SC365.AvailableTransportRuleSettings]::SkipSpfIncoming -band $settingsToFetch)
    {& $addSetting "Skip-SPF-incoming.json" "SkipSpfIncoming"}

    if([SC365.AvailableTransportRuleSettings]::SkipSpfInternal -band $settingsToFetch)
    {& $addSetting "Skip-SPF-internal.json" "SkipSpfInternal"}

    if([SC365.AvailableTransportRuleSettings]::Inbound -band $settingsToFetch)
    {& $addSetting "Inbound.json" "Inbound"}

    if([SC365.AvailableTransportRuleSettings]::Outbound -band $settingsToFetch)
    {& $addSetting "Outbound.json" "Outbound"}

    # Deactivated, because it seems unnecessary
    # if([SC365.AvailableTransportRuleSettings]::Internal -band $settingsToFetch)
    # {& $addSetting "Internal.json" "Internal"}

    # Return the array in reverse SMPriority order, so that they can be created with the
    # same priority, i.e.:
    # New-TransportRule @param -Priority 3
    # But via this sorting, an SMPriority 0 rule will actually be at the top (but at priority 3).
    $ret | Sort-Object -Property SMPriority -Descending
    #>

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

# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUD3pTm+nN1r2FavWS1RDo9KQK
# rUCggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# MRYEFMtQAGWmn5sHPQUg7Kqgyi4bv1t5MA0GCSqGSIb3DQEBAQUABIIBAH7X0cWs
# MX+HlxH/5Ep496WMTASdmQXsKboDxjCxRFgLH1wWw/WUe2jW/ClTfuRQQZlyqfX4
# aWboGxuBprClBm2dg2xgrTrFaJIUbgsSpWHgl1fkyaxeQBLfgF+YOu1DUfY9Dxu/
# FAuTx0+uEliUMzrVkyGpgU8DTEnVhlTQErqZVUZnBYBBjMGsa7CY5FWC1RFMX++h
# b0D1YcpWwmpu0wa0M1Qrmn7D+apUJRrDVCSPCTBR4KOqkBWODDVdOksgzzjHU4vt
# sWt5SHwSLc764gb9fucdXIItoha9VXvSv19QFji6X1bgz1XN73+JEcHI/7tbPhQE
# CYo3oU8+91uMUr4=
# SIG # End signature block
