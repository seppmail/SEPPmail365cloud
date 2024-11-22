# Detect installed module version
$ModuleRootPath = Split-Path -Path $PSScriptRoot -Parent
$ManiFestFile = Import-PowerShellDataFile -Path $ModuleRootPath/SEPPmail365cloud.psd1
$ModuleVersion = $ManiFestFile.Moduleversion.ToString()

Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Welcome to the SEPPmail.cloud PowerShell setup module version $ModuleVersion |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Please read the documentation if you are unfamiliar                 |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| with the module and its CmdLets before continuing !                 |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| https://docs.seppmail.com/en/cloud/c07_cloud_powershell.html        |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Press <CTRL><Klick> to open the Link                                |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray

Write-Verbose "Running InitScript.ps1 and doing requirement-checks" 
if ($sc365notests -ne $true) {
    if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
        Write-Warning "You do not Trust the PowerShellGallery as a module installation source."
        Write-Output "Run 'Set-PSRepository -Name PSGallery -Trusted' to avoid confirmation on module installs."
    }
    # Check Module availability
    if (!(Get-Module DNSClient-PS -ListAvailable)) {
        try {
            Write-Information "Installing required module DNSClient-PS" -InformationAction Continue
            if (Get-Command -Name Install-PSResource) {
                Install-PSResource 'DNSClient-PS' -Reinstall -WarningAction SilentlyContinue
            } else {
                Install-Module 'DNSClient-PS' -Force -WarningAction SilentlyContinue
            }
            Import-Module DNSClient-PS -Force
        } 
        catch {
            Write-Error "Could not install required Module 'DNSClient'. Please install manually from the PowerShell Gallery"
        }
    }
    if (!(Get-Module PSWriteHtml -ListAvailable)) {
        try {
            Write-Information "Installing required module PSWriteHtml" -InformationAction Continue
            if (Get-Command -Name Install-PSResource) {
                Install-PSResource 'PSWriteHtml' -reinstall -WarningAction SilentlyContinue
            } else {
                Install-Module 'PSWriteHtml' -force -WarningAction SilentlyContinue         
            }
            Import-Module PSWriteHtml -Force
        } 
        catch {
            Write-Error "Could not install required Module 'PSWriteHtml'. Please install manually from the PowerShell Gallery"
        }
    }
    if (!(Get-Module ExchangeOnlineManagement -ListAvailable|Where-Object Version -like '3.6.0')) {
        try {
            Write-Information "Installing required module ExchangeOnlineManagement" -InformationAction Continue
            if (Get-Command -Name Install-PSResource) {
                Install-PSResource ExchangeOnlineManagement -Reinstall -WarningAction SilentlyContinue
            } else {
                Install-Module ExchangeOnlineManagement -force -WarningAction SilentlyContinue
            }
            Import-Module ExchangeOnlineManagement
        } 
        catch {
            Write-Error "Could not install required Module 'ExchangeOnlineManagement'. Please install manually from the PowerShell Gallery"
            break
        }
    }
    #Check PowerShell Version
    [string]$requiredPSVersion = '7.4.6'
    [String]$InstPSVersion = ((($PSVersionTable.PSVersion.ToString())) -Split '\.')[0..2] -join '.')
    if (($PSVersiontable.PSEdition -eq "Desktop") -and (!($PSVersionTable.Platform))) {
        $minVersion = ConvertTo-SemanticVersion -VersionString $requiredPSVersion
        $instVersion = ConvertTo-SemanticVersion -VersionString $instPSVersion
    } else {
        $minVersion = [System.Management.Automation.SemanticVersion]::Parse($requiredPSVersion)
        $instVersion = [System.Management.Automation.SemanticVersion]::Parse($InstPSVersion)
    }
    if ($instVersion -ge $minVersion) {
        Write-Verbose "PowerShell version is $instPSVersion and equal or newer than required version $minPSVersion"
    } else {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           ! WRONG POWERSHELL VERSION !               |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           PLEASE install PowerShell CORE $minPSVersion+      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           The module will not work on                |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           Windows Powershell 5.1 and earlier :-( :-( |" -ForegroundColor Red -BackgroundColor Black
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
        Write-Host "|          Install version 3.6.0 ++ of the             |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|         ExchangeOnlineManagement Module with:        |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|  `"Install-Module ExchangeOnlineManagement -Force`"    |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                         or                           |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "| `"Install-PSResource ExchangeOnlineManagement -Force`" |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|     # EXIT and RESTART THE POWERSHELL SESSION #      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|       `"Import-Module ExchangeOnlineManagement`"       |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
    }
}

Write-Verbose 'Initialize argument completer scriptblocks'
$script:paramDomSB = {
    # Read Accepted Domains for domain selection
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    $tenantAcceptedDomains.Domain | Where-Object {
        $_ -like "$wordToComplete*"
            } | ForEach-Object {
                "'$_'"
                }
}
# SIG # Begin signature block
# MIIVzAYJKoZIhvcNAQcCoIIVvTCCFbkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDVI9NpxT0Nb23n
# AgY7aEUCk6o/fHcpktXKxxew+dDbiKCCEggwggVvMIIEV6ADAgECAhBI/JO0YFWU
# jTanyYqJ1pQWMA0GCSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# DBJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoM
# EUNvbW9kbyBDQSBMaW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2Vy
# dmljZXMwHhcNMjEwNTI1MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjBWMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYDVQQDEyRTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCN55QSIgQkdC7/FiMCkoq2rjaFrEfUI5ErPtx94jGgUW+s
# hJHjUoq14pbe0IdjJImK/+8Skzt9u7aKvb0Ffyeba2XTpQxpsbxJOZrxbW6q5KCD
# J9qaDStQ6Utbs7hkNqR+Sj2pcaths3OzPAsM79szV+W+NDfjlxtd/R8SPYIDdub7
# P2bSlDFp+m2zNKzBenjcklDyZMeqLQSrw2rq4C+np9xu1+j/2iGrQL+57g2extme
# me/G3h+pDHazJyCh1rr9gOcB0u/rgimVcI3/uxXP/tEPNqIuTzKQdEZrRzUTdwUz
# T2MuuC3hv2WnBGsY2HH6zAjybYmZELGt2z4s5KoYsMYHAXVn3m3pY2MeNn9pib6q
# RT5uWl+PoVvLnTCGMOgDs0DGDQ84zWeoU4j6uDBl+m/H5x2xg3RpPqzEaDux5mcz
# mrYI4IAFSEDu9oJkRqj1c7AGlfJsZZ+/VVscnFcax3hGfHCqlBuCF6yH6bbJDoEc
# QNYWFyn8XJwYK+pF9e+91WdPKF4F7pBMeufG9ND8+s0+MkYTIDaKBOq3qgdGnA2T
# OglmmVhcKaO5DKYwODzQRjY1fJy67sPV+Qp2+n4FG0DKkjXp1XrRtX8ArqmQqsV/
# AZwQsRb8zG4Y3G9i/qZQp7h7uJ0VP/4gDHXIIloTlRmQAOka1cKG8eOO7F/05QID
# AQABo4IBEjCCAQ4wHwYDVR0jBBgwFoAUoBEKIz6W8Qfs4q8p74Klf9AwpLQwHQYD
# VR0OBBYEFDLrkpr/NZZILyhAQnAgNpFcF4XmMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMDMBsGA1UdIAQUMBIwBgYE
# VR0gADAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21v
# ZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEMBQADggEBABK/oe+LdJqYRLhpRrWrJAoMpIpnuDqBv0WKfVIHqI0fTiGF
# OaNrXi0ghr8QuK55O1PNtPvYRL4G2VxjZ9RAFodEhnIq1jIV9RKDwvnhXRFAZ/ZC
# J3LFI+ICOBpMIOLbAffNRk8monxmwFE2tokCVMf8WPtsAO7+mKYulaEMUykfb9gZ
# pk+e96wJ6l2CxouvgKe9gUhShDHaMuwV5KZMPWw5c9QLhTkg4IUaaOGnSDip0TYl
# d8GNGRbFiExmfS9jzpjoad+sPKhdnckcW67Y8y90z7h+9teDnRGWYpquRRPaf9xH
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggYaMIIEAqADAgECAhBiHW0M
# UgGeO5B5FSCJIRwKMA0GCSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxLTArBgNVBAMTJFNlY3RpZ28gUHVibGljIENv
# ZGUgU2lnbmluZyBSb290IFI0NjAeFw0yMTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5
# NTlaMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzAp
# BgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwggGiMA0G
# CSqGSIb3DQEBAQUAA4IBjwAwggGKAoIBgQCbK51T+jU/jmAGQ2rAz/V/9shTUxjI
# ztNsfvxYB5UXeWUzCxEeAEZGbEN4QMgCsJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NV
# DgFigOMYzB2OKhdqfWGVoYW3haT29PSTahYkwmMv0b/83nbeECbiMXhSOtbam+/3
# 6F09fy1tsB8je/RV0mIk8XL/tfCK6cPuYHE215wzrK0h1SWHTxPbPuYkRdkP05Zw
# mRmTnAO5/arnY83jeNzhP06ShdnRqtZlV59+8yv+KIhE5ILMqgOZYAENHNX9SJDm
# +qxp4VqpB3MV/h53yl41aHU5pledi9lCBbH9JeIkNFICiVHNkRmq4TpxtwfvjsUe
# dyz8rNyfQJy/aOs5b4s+ac7IH60B+Ja7TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz4
# 4MPZ1f9+YEQIQty/NQd/2yGgW+ufflcZ/ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBM
# dlyh2n5HirY4jKnFH/9gRvd+QOfdRrJZb1sCAwEAAaOCAWQwggFgMB8GA1UdIwQY
# MBaAFDLrkpr/NZZILyhAQnAgNpFcF4XmMB0GA1UdDgQWBBQPKssghyi47G9IritU
# pimqF6TNDDAOBgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNV
# HSUEDDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsG
# A1UdHwREMEIwQKA+oDyGOmh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1
# YmxpY0NvZGVTaWduaW5nUm9vdFI0Ni5jcmwwewYIKwYBBQUHAQEEbzBtMEYGCCsG
# AQUFBzAChjpodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2Rl
# U2lnbmluZ1Jvb3RSNDYucDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0
# aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEABv+C4XdjNm57oRUgmxP/BP6YdURh
# w1aVcdGRP4Wh60BAscjW4HL9hcpkOTz5jUug2oeunbYAowbFC2AKK+cMcXIBD0Zd
# OaWTsyNyBBsMLHqafvIhrCymlaS98+QpoBCyKppP0OcxYEdU0hpsaqBBIZOtBajj
# cw5+w/KeFvPYfLF/ldYpmlG+vd0xqlqd099iChnyIMvY5HexjO2AmtsbpVn0OhNc
# WbWDRF/3sBp6fWXhz7DcML4iTAWS+MVXeNLj1lJziVKEoroGs9Mlizg0bUMbOalO
# hOfCipnx8CaLZeVme5yELg09Jlo8BMe80jO37PU8ejfkP9/uPak7VLwELKxAMcJs
# zkyeiaerlphwoKx1uHRzNyE6bxuSKcutisqmKL5OTunAvtONEoteSiabkPVSZ2z7
# 6mKnzAfZxCl/3dq3dUNw4rg3sTCggkHSRqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5J
# KdGvspbOrTfOXyXvmPL6E52z1NZJ6ctuMFBQZH3pwWvqURR8AgQdULUvrxjUYbHH
# j95Ejza63zdrEcxWLDX6xWls/GDnVNueKjWUH3fTv1Y8Wdho698YADR7TNx8X8z2
# Bev6SivBBOHY+uqiirZtg0y9ShQoPzmCcn63Syatatvx157YK9hlcPmVoa1oDE5/
# L9Uo2bC5a4CH2RwwggZzMIIE26ADAgECAhAMcJlHeeRMvJV4PjhvyrrbMA0GCSqG
# SIb3DQEBDAUAMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYw
# HhcNMjMwMzIwMDAwMDAwWhcNMjYwMzE5MjM1OTU5WjBqMQswCQYDVQQGEwJERTEP
# MA0GA1UECAwGQmF5ZXJuMSQwIgYDVQQKDBtTRVBQbWFpbCAtIERldXRzY2hsYW5k
# IEdtYkgxJDAiBgNVBAMMG1NFUFBtYWlsIC0gRGV1dHNjaGxhbmQgR21iSDCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOapobQkNYCMP+Y33JcGo90Soe9Y
# /WWojr4bKHbLNBzKqZ6cku2uCxhMF1Ln6xuI4ATdZvm4O7GqvplG9nF1ad5t2Lus
# 5SLs45AYnODP4aqPbPU/2NGDRpfnceF+XhKeiYBwoIwrPZ04b8bfTpckj/tvenB9
# P8/9hAjWK97xv7+qsIz4lMMaCuWZgi8RlP6XVxsb+jYrHGA1UdHZEpunEFLaO9Ss
# OPqatPAL2LNGs/JVuGdq9p47GKzn+vl+ANd5zZ/TIP1ifX76vorqZ9l9a5mzi/HG
# vq43v2Cj3jrzIQ7uTbxtiLlPQUqkRzPRtiwTV80JdtRE+M+gTf7bT1CTvG2L3scf
# YKFk7S80M7NydxV/qL+l8blGGageCzJ8svju2Mo4BB+ALWr+gBmCGqrM8YKy/wXR
# tbvdEvBOLsATcHX0maw9xRCDRle2jO+ndYkTKZ92AMH6a/WdDfL0HrAWloWWSg62
# TxmJ/QiX54ILQv2Tlh1Al+pjGHN2evxS8i+XoWcUdHPIOoQd37yjnMjCN593wDzj
# XCEuDABYw9BbvfSp29G/uiDGtjttDXzeMRdVCJFgULV9suBVP7yFh9pK/mVpz+aC
# L2PvqiGYR41xRBKqwrfJEdoluRsqDy6KD985EdXkTvdIFKv0B7MfbcBCiGUBcm1r
# fLAbs8Q2lqvqM4bxAgMBAAGjggGpMIIBpTAfBgNVHSMEGDAWgBQPKssghyi47G9I
# ritUpimqF6TNDDAdBgNVHQ4EFgQUL96+KAGrvUgJnXwdVnA/uy+RlEcwDgYDVR0P
# AQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwSgYD
# VR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAjBggrBgEFBQcCARYXaHR0cHM6Ly9z
# ZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkGA1UdHwRCMEAwPqA8oDqGOGh0dHA6
# Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5nQ0FSMzYu
# Y3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEFBQcwAoY4aHR0cDovL2NydC5zZWN0
# aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcnQwIwYIKwYB
# BQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMB4GA1UdEQQXMBWBE3N1cHBv
# cnRAc2VwcG1haWwuY2gwDQYJKoZIhvcNAQEMBQADggGBAHnWpS4Jw/QiiLQi2EYv
# THCtwKsj7O3G7wAN7wijSJcWF7iCx6AoCuCIgGdWiQuEZcv9pIUrXQ6jOSRHsDNX
# SvIhCK9JakZJSseW/SCb1rvxZ4d0n2jm2SdkWf5j7+W+X4JHeCF9ZOw0ULpe5pFs
# IGTh8bmTtUr3yA11yw4vHfXFwin7WbEoTLVKiL0ZUN0Qk+yBniPPSRRlUZIX8P4e
# iXuw7lh9CMaS3HWRKkK89w//18PjUMxhTZJ6dszN2TAfwu1zxdG/RQqvxXUTTAxU
# JrrCuvowtnDQ55yXMxkkSxWUwLxk76WvXwmohRdsavsGJJ9+yxj5JKOd+HIZ1fZ7
# oi0VhyOqFQAnjNbwR/TqPjRxZKjCNLXSM5YSMZKAhqrJssGLINZ2qDK/CEcVDkBS
# 6Hke4jWMczny8nB8+ATJ84MB7tfSoXE7R0FMs1dinuvjVWIyg6klHigpeEiAaSaG
# 5KF7vk+OlquA+x4ohPuWdtFxobOT2OgHQnK4bJitb9aDazGCAxowggMWAgEBMGgw
# VDELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UE
# AxMiU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNgIQDHCZR3nkTLyV
# eD44b8q62zANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAM
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBlQ1KRul6hyglOELYdwghTOPYN
# +fASUNaaanXmY2lyyjANBgkqhkiG9w0BAQEFAASCAgApXjQVoLMtg1Q6Z2PPohEQ
# hVZNj/0sDNLC9YrPMlC4Obr3M+NnOpwbot4jbTB8j19XoSpdg/3jPupOi4yzCPvd
# Va96NdG5Y4QWusvi/Lbph/2US1mHe6+IHZW6ahkHCanKNeMgic1BBaXo2RbUgkP3
# 8saSwN/jx1T3DBiuTfrJzNrQPxffT7jUfCAOGvQl5UItb9Nv68Z6PCP3z/6HiQO7
# TSUOuCBIktX9RbSz+HfNt7nRriFIgovOzEdDi6jlrB3Mg8l+yo89IkExsv/IJUVH
# Gq+npTNBGUIFybKimPU3isfVTU9E+GiDS+mKliRGZDmYAjBoa6RMMWsffXhzXTaW
# nCOJ2X+VI0ocLvlPF+vIzrmP1eCsb4LGSM9wuodg2ceNNoE1WqR+uozZSC1gBWlm
# YTSuG/IaVYR7YoQrOqp/slRCtVsjlOgJ3XyxSzx5JcFcypTBLg36QlkuKq4W2MPl
# dTWoTV+5Jf/e9jTRPW7bPwtN72NvAStZF1MSn948Jb2vHiHXR+66tnYTGAKHp2IS
# GOuTOzW+TC8EOhMFhdw6YRc9VZQ7JPbNsgm8LYNA0cRMZrFQaI7lrufgV7yya8U8
# BGNyFyseMCuDkpa6ztwPiqgkodIoHhWeDa5F6thns29wRSdYu7kifQOheZAivRYV
# VZ9eBB106KLHHwcT3lZbXw==
# SIG # End signature block
