[CmdLetBinding()]

$InformationPreference = $true
$ModulePath = $PSScriptRoot
$InteractiveSession = [System.Environment]::UserInteractive

Write-Verbose 'Request terminating errors by default'
$PSDefaultParameterValues['*:ErrorAction'] = [System.Management.Automation.ActionPreference]::Stop

Write-Verbose 'Loading Module Files'
. $ModulePath\Private\PrivateFunctions.ps1
. $ModulePath\Private\ConfigBundle.ps1
. $ModulePath\Public\Common.ps1
. $ModulePath\Public\Rules.ps1
. $ModulePath\Public\Connectors.ps1


if ($sc365notests -ne $true) {
    #Check Environment
    If ($psversiontable.PsVersion.ToString() -notlike '7.*') {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+           ! WRONG POWERSHELL VERSION !               +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+          PLEASE Install PowerShell CORE 7.2+         +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+          The module will not load on                 +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+           Windows Powershell 5.1  :-( :-(            +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
        Break
    }
    # Check Exo Module Version 
    if ((Get-Module ExchangeOnlineManagement).Version -notlike '3.*') {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+   WRONG Version of  ExchangeOnlineManagement Module  +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+          Install Version 3.0.0+ of the               +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+         ExchangeOnlineManagement Module with:        +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+   Install-Module ExchangeOnlineManagement -Force     +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+          # RESTART THE POWERSHELL SESSION #          +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+        Import-Module ExchangeOnlineManagement        +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+                                                      +" -ForegroundColor Green -BackgroundColor DarkGray
        Write-Host "+------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
    }
    Write-Verbose "Testing Exchange Online connectivity"
    if (!(Test-SC365ConnectionStatus)) {
        Write-Warning "You are not connected to Exchange Online. Use Connect-ExchangeOnline to connect to your tenant"
    }
    Write-Verbose 'Test new version available'
    try {
        $onLineVersion = Find-Module -Name 'SEPPmail365cloud'|Select-Object -expandproperty Version
        $offLineVersion = Test-ModuleManifest (Join-Path $ModulePath -ChildPath SEPPmail365cloud.psd1) |Select-Object -ExpandProperty Version 
        if ($onLineVersion -gt $offLineVersion) {
            Write-Warning "You have version $offlineVersion, but there is the new version $onLineVersion of the SEPPmail365cloud module available on the PowerShell Gallery. Update the module as soon as possible. More info here https://www.powershellgallery.com/packages/SEPPMail365cloud"
        }   
    }
    catch {
        Write-Error "Could not determine newest module version due to exception $($_.Exception.Message)"
    }
}

Write-Verbose 'Initialize argument completer scriptblocks'
$paramDomSB = {
    # Read Accepted Domains for domain selection
    Get-AcceptedDomain -Erroraction silentlycontinue|select-Object -ExpandProperty DomainName
}


Export-ModuleMember -Alias * -Function *

# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUgy4ISHeelKbBFuR2CR8nGugy
# A+2ggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# MRYEFCEdQzx0P8V/bcPLJWHMBAH2/vMhMA0GCSqGSIb3DQEBAQUABIIBAGuAhSMo
# xRewubLQQQIvgDY2rRAbgaAgu2PAuPWQTKYnFZFRcBrsHXWRc9ZHLc3AlzTzArbP
# VVxRjZKfYWczhWCPezJ3OxGA1Mypj1EUWGwNu/e4552nHlACWCu3vugEb36Ne+1C
# 9UUOf5XAoh3M3taS9Yl0uaJIj62XuvHAmjc8nM9hLkTfEez8pKxNXZVrEtqGXS5K
# s52IpYfEV3utjZXEngXijSaSLx6jlYCqk+Xu7NAggCT5a2TR660bGEtcHbFLXSaV
# gMR+L01QcXigKfNRV0GTjx2N5vdz455MZT21K+r3FSfjT3LNQbKR4HfbwdA6qaRq
# gbMVGfIQBlpLgB0=
# SIG # End signature block
