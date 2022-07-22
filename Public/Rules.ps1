<#
.SYNOPSIS
    Read existing SEPPmail.cloud transport rules in the exchange online environment.
.DESCRIPTION
    Use this tofigure out if there are already SEPPmail.cloud rules implemented in Exchange oinline.
    It is only emitting installed rules which come with the seppmail365cloud PowerShell Module.
    If you want to get all installed transport rules, usw New-SC365ExoReport-

.EXAMPLE
    Get-SC365Rules -Routing 'microsoft'
#>
function Get-SC365Rules {
    [CmdletBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    param
    (
        [Parameter(Mandatory = $false)]
        $routing = 'microsoft'
    )

    if (!(Test-SC365ConnectionStatus))
    { 
        throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
    }
    else 
    {
        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

        if ($routing -eq 'microsoft') {
            #Wait-Debugger

            $transportRuleFiles = Get-Childitem "$psscriptroot\..\ExoConfig\Rules\"

            foreach($file in $transportRuleFiles) {
            
                $setting = Get-SC365TransportRuleSettings -File $file -Routing $routing
                $rule = Get-TransportRule $setting.Name -ErrorAction SilentlyContinue
                if ($rule) {
                    $rule
                }
                else
                {
                    Write-Warning "Rule $($setting.Name) does not exist"
                }
            }    
        }
        else {
            Write-Warning "No transport rules needed for routingtype $routing"
        }
    }
}

<#
.SYNOPSIS
    Create transport rules for routingmode "parallel"
.DESCRIPTION
    Creates all necessary transport rules in Exchange Online to send E-Mails through seppmail.cloud for cryptographic processing.
.EXAMPLE
    PS C:\> New-SC365Rules -routing 'microsoft'
    Thats the one-show-solves-all-problem CmdLet with interactive questioning on rule generation. It will search for existing rules, create new rules and ask if rules are placed on top (before all other) or bottom (after all other).
.EXAMPLE
    PS C:\> New-SC365Rules -routing 'microsoft' -PlacementPriority Bottom
    Places the transport rules AFTER all other rules. If you want to place them before, use "TOP" as parameter value.
.EXAMPLE
    PS C:\> New-SC365Rules -routing 'microsoft' -disabled
    Sets the transport rules up, but keeps them inactive. For a smoother integration.
.EXAMPLE
    PS C:\> New-SC365Rules -routing 'seppmail'
    Does literally nothing, except a message to the user that there are no rules needed.
.INPUTS
    none
.OUTPUTS
    transport rules
.NOTES
    
#>
function New-SC365Rules
{
    [CmdletBinding(SupportsShouldProcess = $true,
                   ConfirmImpact = 'Medium',
                   HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
                  )]
    param
    (
        [Parameter(Mandatory=$false,
                   HelpMessage='Should the new rules be placed before or after existing ones (if any)')]
        [ValidateSet('Top','Bottom')]
        [String] $PlacementPriority = 'Bottom',

        [Parameter(
            Mandatory = $true,
            HelpMessage = 'MX record->SEPPmail means routingtype seppmail, MX->Microsoft means routingtype microsoft'
        )]
        [ValidateSet('microsoft')]
        [String]$routing = 'microsoft',

        [Parameter(Mandatory=$false,
                   HelpMessage='E-Mail domains you want to exclude from beeing routed through the SEPPmail.cloud')]
        [String[]]$ExcludeEmailDomain,

        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Add rules if you provisioned internal e-mail signature in the SEPPmail.cloud Service'
        )]
        
        [switch]$Disabled
    )

    begin
    {
        if (!(Test-SC365ConnectionStatus))
        { throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" }

        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

        $outboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors $true | Where-Object { $_.Name -match "^\[SEPPmail.cloud\]" }
        if(!($outboundConnectors))
        {
            throw [System.Exception] "No SEPPmail.cloud outbound connector found. Run `"New-SC365Connectors`" to add the proper SEPPmail.cloud connectors"
        }
        if ($($outboundConnectors.Enabled) -ne $true) {
            throw [System.Exception] "SEPPmail.cloud outbound-connector is disabled, cannot create rules. Create connectors without -Disable switch, or enable them in the admin portal."
        }
    }

    process
    {
        try
        {
            Write-Verbose "Read all `"non-[SEPPmail`" transport rules"
            $existingTransportRules = Get-TransportRule | Where-Object Name -NotMatch '\[SEPPmail*'
            [int] $placementPrio = @(0, $existingTransportRules.Count)[!($PlacementPriority -eq "Top")] <# Poor man's ternary operator #>
            if ($existingTransportRules)
            {
                if($InteractiveSession -and !$PSBoundParameters.ContainsKey("PlacementPriority") <# Prio already set, so no need to ask #>)
                {
                    Write-Warning 'Found existing custom transport rules.'
                    Write-Warning '--------------------------------------------'
                    foreach ($etpr in $existingTransportRules) {
                        Write-Warning "Rule name `"$($etpr.Name)`" with state `"$($etpr.State)`" has priority `"$($etpr.Priority)`""
                    }
                    Write-Warning '--------------------------------------------'
                    Do {
                        try {
                            [ValidateSet('Top', 'Bottom', 'Cancel', 't', 'T', 'b', 'B', 'c', 'C', $null)]$existingRulesAction = Read-Host -Prompt "Where shall we place the SEPPmail.cloud rules ? (Top(Default)/Bottom/Cancel)"
                        }
                        catch {}
                    }
                    until ($?)

                    switch ($existingRulesAction) {
                        'Top' { $placementPrio = '0' }
                        't' { $placementPrio = '0' }
                        'Bottom' { $placementPrio = ($existingTransportRules).count }
                        'b' { $placementPrio = ($existingTransportRules).count }
                        'Cancel' { return }
                        'c' { return }
                        default { $placementPrio = '0' }
                    }
                }
            }
            else
            {
                Write-Verbose 'No existing custom rules found'
            }
            Write-Verbose "Placement priority is $placementPrio"

            Write-Verbose "Read existing [SEPPmail.cloud] transport rules"
            $existingSMCTransportRules = Get-TransportRule | Where-Object Name -Match '\[SEPPmail*'
            [bool] $createRules = $true
            if ($existingSMCTransportRules)
            {
                if($InteractiveSession)
                {
                    Write-Warning 'Found existing [SEPPmail* transport rules.'
                    Write-Warning '--------------------------------------------'
                    foreach ($eSMCtpr in $existingSMCTransportRules) {
                        Write-Warning "Rule name `"$($eSMCtpr.Name)`" with state `"$($eSMCtpr.State)`" has priority `"$($eSMCtpr.Priority)`""
                    }
                    Write-Warning '--------------------------------------------'
                    Do {
                        try {
                            [ValidateSet('y', 'Y', 'n', 'N')]$recreateSMRules = Read-Host -Prompt "Shall we delete and recreate them ? (Y/N)"
                        }
                        catch {}
                    }
                    until ($?)
                    if ($recreateSMRules -like 'y') {
                        $existingSMCTransportRules|ForEach-Object {Remove-Transportrule -Identity $_.Identity -Confirm:$false}
                    }
                    else {
                        $createRules = $false
                    }
                }
                else
                {
                    throw [System.Exception] "SEPPmail* transport rules already exist"
                }
            }

            if($createRules){
               
                $transportRuleFiles = Get-Childitem -Path "$psscriptroot\..\ExoConfig\Rules\"

                foreach($file in $transportRuleFiles) {
                
                    $setting = Get-SC365TransportRuleSettings -File $file -Routing $routing
                    # $setting = $_

                    $setting.Priority = $placementPrio + $setting.SMPriority
                    $setting.Remove('SMPriority')
                    if ($Disabled -eq $true) {$setting.Enabled = $false}

                    if (($ExcludeEmailDomain.count -ne 0) -and ($Setting.Name -eq '[SEPPmail.cloud] - Route incoming e-mails to SEPPmail')) {
                        Write-Verbose "Excluding Inbound E-Mails domains $ExcludeEmailDomain"
                        $Setting.ExceptIfRecipientDomainIs = $ExcludeEmailDomain
                    }

                    if (($ExcludeEmailDomain.count -ne 0) -and ($Setting.Name -eq '[SEPPmail.cloud] - Route outgoing e-mails to SEPPmail')) {
                        Write-Verbose "Excluding Outbound E-Mail domains $ExcludeEmailDomain"
                        $Setting.ExceptIfSenderDomainIs = $ExcludeEmailDomain
                    }

                    if ($PSCmdlet.ShouldProcess($setting.Name, "Create transport rule"))
                    {
                        $Now = Get-Date
                        Write-Verbose "Adding Timestamp $now to Comment"
                        $setting.Comments += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                        New-TransportRule @setting
                    }
                }
            }
        }
        catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
        }
    }
    end
    {

    }
}

<#
.SYNOPSIS
    Removes the SEPPmail.cloud transport rules
.DESCRIPTION
    Convenience function to remove the SEPPmail.cloud rules in one CmdLet.
.EXAMPLE
    Remove-SC365Rules -routing 'microsoft'
#>
function Remove-SC365Rules {
    [CmdletBinding(SupportsShouldProcess = $true,
                   ConfirmImpact = 'Medium',
                   HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
                  )]
    param
    (
        [Parameter(
            Mandatory = $false,
            HelpMessage = 'Use seppmail if the MX record points to SEPPmail and microsoft if the MX record points to the Microsoft Inrastructure'
        )]
        [ValidateSet('microsoft')]
        [String]$routing = 'microsoft'
    )

    begin {
        $transportRuleFiles = Get-Childitem "$psscriptroot\..\ExoConfig\Rules\"
        if (!(Test-SC365ConnectionStatus))
        { 
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
        }
        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
    }

    process {
        Write-Verbose "Removing current version module rules"
        
        if ($routing -eq 'microsoft') {
            foreach ($file in $transportRuleFiles) {
                $setting = Get-SC365TransportRuleSettings -routing $routing -file $file
                
                if($PSCmdlet.ShouldProcess($setting.Name, "Remove transport rule")) {
                    $rule = Get-TransportRule $setting.Name -ErrorAction SilentlyContinue
                    if($rule -ne $null)
                        {$rule | Remove-TransportRule -Confirm:$false}
                    else
                        {Write-Warning "Rule $($setting.Name) does not exist"}
                }
            }    
        }
        else {
            Write-Warning "Routingtype 'seppmail' doesnt require any Mail Transport Rules. Use '-routing 'microsoft'' to remove existing rules."
        }

    }
    end {

    }
}

<#
.SYNOPSIS
    Backs up all existing transport rules to individual json files
.DESCRIPTION
    Convenience function to perform a backup of all existing transport rules
.EXAMPLE
    Backup-SC365Rules -OutFolder "C:\temp"
#>
function Backup-SC365Rules
{
    [CmdletBinding()]
    param
    (
        [Parameter(
             Mandatory = $true,
             HelpMessage = 'Folder in which the backed up configuration will be stored'
         )]
        [Alias("Folder")]
        [String] $OutFolder
    )

    begin
    {
        if (!(Test-SC365ConnectionStatus))
        { throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" }

        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
    }

    process
    {
        if(!(Test-Path $OutFolder))
        {New-Item $OutFolder -ItemType Directory}

        Get-TransportRule | Foreach-Object{
            $n = $_.Name
            $n = $n -replace "[\[\]*\\/?:><`"]"

            $p = "$OutFolder\rule_$n.json"
            Write-Verbose "Backing up $($_.Name) to $p"
            ConvertTo-Json -InputObject $_ | Out-File $p
        }
    }
}

if (!(Get-Alias 'Set-SC365rules' -ErrorAction SilentlyContinue)) {
    New-Alias -Name Set-SC365Rules -Value New-SC365Rules
}

Register-ArgumentCompleter -CommandName New-SC365Rules -ParameterName ExcludeEmailDomain -ScriptBlock $paramDomSB

# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfe6aAkZTU3y/UgqfshQRW7iJ
# 0OGggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# MRYEFCCfdbM5T/3Iv8FAAuOVSOvfvAVOMA0GCSqGSIb3DQEBAQUABIIBADSkb5Ll
# dY+rgU+9jk2AGyVT1f/7LLhJkuIQbXOz808kguCbdyT31QxAjLdTVYEygi2Kb95F
# dHUs6nVbzayi8bMP7hZ5YWCzclEHxPHlb0zvIJ3GWdPhNA/sg9o2j07LgRyQoPbK
# leGh2c/NuLXIFW886kOkXZJF+11cbWqt70+94mLUFaF/y93jfYlWTCrTthe3ZHqw
# hmAOQ5LoetzjBASKeE6WL1oPhwsiCay09kvUoyG+qvJFm9iZPQXe67rJa1ghmXhd
# viYHXdZdcLJk/RwdGIkW4ymuqjJjHGftRWhNYr1CBdVUtbESkBQKcKr5wDeP06yW
# u+43VwxKw7DT1p4=
# SIG # End signature block
