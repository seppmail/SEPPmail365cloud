[CmdLetBinding()]

$InformationPreference = $true
$ModulePath = $PSScriptRoot
$InteractiveSession = [System.Environment]::UserInteractive

Write-Verbose 'Request terminating errors by default'
$PSDefaultParameterValues['*:ErrorAction'] = [System.Management.Automation.ActionPreference]::Stop

$global:tenantAcceptedDomains = $null

Write-Verbose 'Loading Module Files'
. $ModulePath\Private\PrivateFunctions.ps1
. $ModulePath\Public\Common.ps1
. $ModulePath\Public\Rules.ps1
. $ModulePath\Public\Connectors.ps1

Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Welcome to the SEPPmail.cloud PowerShell setup module               |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Please read the documentation on GitHub if you are unfamiliar       |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| with the module and its CmdLets before continuing !                 |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md    |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Press <CTRL><Klick> to open the Link                                |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray

if ($sc365notests -ne $true) {
    # Check Module availability
    if (!(Get-Module DNSClient-PS -ListAvailable)) {
        try {
            Write-Information "Installing required module DNSClient-PS" -InformationAction Continue
            Install-Module DNSCLient-PS -WarningAction SilentlyContinue
            Import-Module DNSClient-PS -Force
        } 
        catch {
            Write-Error "Could not install requirem Module 'DNSClient'. Please install manually from the PowerShell Gallery"
        }
    }
    if (!(Get-Module ExchangeOnlineManagement -ListAvailable|Where-Object Version -like '3.*')) {
        try {
            Write-Information "Installing required module ExchangeOnlineManagement" -InformationAction Continue
            Install-Module ExchangeOnlineManagement -WarningAction SilentlyContinue
            Import-Module ExchangeOnlineManagement
        } 
        catch {
            Write-Error "Could not install required Module 'ExchangeOnlineManagement'. Please install manually from the PowerShell Gallery"
            break
        }
    }
    
    #Check Environment
    If ($psversiontable.PsVersion.ToString() -notlike '7.*') {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           ! WRONG POWERSHELL VERSION !               |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           PLEASE install PowerShell CORE 7.2+        |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           The module will not work on                |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           Windows Powershell 5.1  :-( :-(            |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Break
    }
    # Check Exo Module Version 
    if (!((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Where({$_.Version -ge [version]'3.0.0'}))) {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|   WRONG Version of ExchangeOnlineManagement Module   |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|          Install version 3.0.0 ++ of the             |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|         ExchangeOnlineManagement Module with:        |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|  `"Install-Module ExchangeOnlineManagement -Force`"    |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|     # EXIT and RESTART THE POWERSHELL SESSION #      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|       `"Import-Module ExchangeOnlineManagement`"       |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
    }
    Write-Verbose "Testing Exchange Online connectivity"
    if (!(Test-SC365ConnectionStatus)) {
        Write-Warning "You are not connected to Exchange Online. Use `"Connect-ExchangeOnline`" to connect to your tenant"
    } else {
        try {
            if ((Get-OrganizationConfig).IsDehydrated) {
                Write-Verbose "Organisation is not enabled for customizations -- is 'Dehyrated'. Turning this on now"
                Enable-OrganizationCustomization  #-confirm:$false
            }        
        }
        catch {
            Write-Warning "Cannot detect Tenant hydration - maybe disconnected"
        }
        
        try {
            Write-verbose "Creating Test OnPrem Connector to check if tenant allows connector creation"
            New-InboundConnector -Name '[SEPPmail.cloud] TempConnector EX505293' -ConnectorType OnPrem -TlsSenderCertificateName 'test.nowhere.org' -SenderDomains 'test.nowhere.org' -RequireTls $true -enabled $false
            Remove-InboundConnector -Identity '[SEPPmail.cloud] TempConnector EX505293' -Confirm:$false
        }
        catch {
            Write-Error "This Tenant is not yet allowed to create OnPrem-Connectors (Exchange Error EX505293).If this tenant shall be integrated in PARALLEL mode, contact Microsoft Support and request connector creation"
        }            
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
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    $tenantAccetedDomains.Domain | Where-Object {
        $_ -like "$wordToComplete*"
            } | ForEach-Object {
                "'$_'"
                }
}

Export-ModuleMember -Alias * -Function *

# SIG # Begin signature block
# MIIL/AYJKoZIhvcNAQcCoIIL7TCCC+kCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDyfPnY6u+k8vqH
# yx9GDH90q4meXRsdibsPysZGmUShiaCCCUAwggSZMIIDgaADAgECAhBxoLc2ld2x
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCRIuI4cu+dZP7p9Qc72OhrAawn
# 8lclrQyp+Ao562J7+TANBgkqhkiG9w0BAQEFAASCAQADhgF3brKS8uE0CAnmprut
# 3C/zPwbzBxxVpXzzWbsrIpVUsKyHuKdpwFYzOsCtLSGm+YBeFLYnP25elAopB/fC
# nOcX+pcYpa9EHa62tPWsN7Wp9YUlkPEDwU3ThWNnSEhIbR24/cJIEuZuPEzsg74v
# RkJVJ060AO1uLS9V3j//OPOYUt2Mp+3IIJPbBeaT15rmk0SqVZVNzgASo7dCyYKf
# JkzuJoy5H0lYjc36Mvho2MmqA7mHOb9k9ZFTQwML2FsjOOQxc04RfeOsUbZfdf0F
# 9kigcQCMskyobxW+5VEBi4mRXfquDeUlr9R7rpC4HDyd/QE/iGlErwchmXYiLP4K
# SIG # End signature block
