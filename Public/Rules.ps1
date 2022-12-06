<#
.SYNOPSIS
	Read existing SEPPmail.cloud transport rules in the exchange online environment.
.DESCRIPTION
	Use this tofigure out if there are already SEPPmail.cloud rules implemented in Exchange online.
	It is only emitting installed rules which come with the seppmail365cloud PowerShell Module.
	If you want to be informed about all installed transport rules, use New-SC365ExoReport.

.EXAMPLE
	Get-SC365Rules
#>
function Get-SC365Rules {
	[CmdletBinding(
		HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
	)]
	param
	(
		[Parameter(Mandatory = $false)]
		$routing = 'parallel'
	)

	if (!(Test-SC365ConnectionStatus))
	{ 
		throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
	}
	else 
	{
		Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

		if ($routing -eq 'parallel') {

			$transportRuleFiles = Get-Childitem "$psscriptroot\..\ExOConfig\Rules\"

			foreach ($file in $transportRuleFiles) {
			
				$setting = Get-SC365TransportRuleSettings -File $file -Routing $routing
				$rule = Get-TransportRule $setting.Name -ErrorAction SilentlyContinue
				if ($rule) {
					
					if ($rule.Identity -like '*100*') {
						$rule|Select-Object Identity,Priority,State,@{Name = 'ExcludedDomains'; Expression={$_.ExceptIfRecipientDomainIs}}
					}
					elseif ($rule.Identity -like '*200*') {
						$rule|Select-Object Identity,Priority,State,@{Name = 'ExcludedDomains'; Expression={$_.ExceptIfSenderDomainIs}}
					}
					else {
						$rule|Select-Object Identity,Priority,State,ExcludedDomains
					}
				}
				else
				{
					Write-Warning "No transport rule '$($setting.Name)'"
				}
			}    
		}
		else {
			Write-Information "No transport rules needed for routingtype $routing"
		}
	}
}

<#
.SYNOPSIS
	Create transport rules for routingmode "parallel"
.DESCRIPTION
	Creates all necessary transport rules in Exchange Online to send E-Mails through seppmail.cloud for cryptographic processing.
.EXAMPLE
		PS C:\> New-SC365Rules -SEPPmailCloudDomain 'contoso.eu'
		Creates the rules for specific domains. Excludes all other e-mail domains from processing by SEPPmail.cloud
.EXAMPLE
	PS C:\> New-SC365Rules -SEPPmailCloudDomain 'contoso.eu' -PlacementPriority Top
	Places the transport rules BEFORE all other rules. This is unusual and against the default. It may make sense in some situations.
.EXAMPLE
	PS C:\> New-SC365Rules -SEPPmailCloudDomain 'contoso.eu' -disabled
	Sets the transport rules up, but keeps them inactive. Useful for a smoother integration.
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
			Mandatory = $false,
			HelpMessage = 'MX record->SEPPmail means routingtype inline, MX->Microsoft means routingtype parallel'
		)]
		[ValidateSet('parallel')]
		[String]$routing = 'parallel',

		[Parameter(Mandatory=$true,
				   HelpMessage='E-Mail domains you have registered in the SEPmail.Cloud')]
	   [String[]]$SEPPmailCloudDomain,

	   [Parameter(Mandatory=$false,
	   				HelpMessage='SCL Value for inbound Mails which should NOT be processed by SEPPmail.Cloud. Default is 5')]
	   [ValidateSet('-1','0','5','6','8','9')]
	   [int]$SCLInboundValue=5,

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

	 	Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

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
			   
				$transportRuleFiles = Get-Childitem -Path "$psscriptroot\..\ExOConfig\Rules\"

				[System.Collections.ArrayList]$ExcludeEmailDomain = (Get-Accepteddomain).DomainName
				$SEPPmailCloudDomain|foreach-object {$ExcludeEmailDomain.Remove($_)}

				$moduleVersion = $myInvocation.MyCommand.Version

				foreach($file in $transportRuleFiles) {
				
					$setting = Get-SC365TransportRuleSettings -File $file -Routing $routing
					# $setting = $_

					$setting.Priority = $placementPrio + $setting.SMPriority
					$setting.Remove('SMPriority')
					if ($Disabled -eq $true) {$setting.Enabled = $false}

					switch ($setting.Name)
					{
						"[SEPPmail.cloud] - 100 Route incoming e-mails to SEPPmail" {
							Write-Verbose "Excluding all other domains than $SEPPmailCloudDomain"
							$Setting.ExceptIfRecipientDomainIs = $ExcludeEmailDomain
							if ($SCLInboundValue -ne 5) {
								Write-Verbose "Setting Value $SCLInboundValue to Inbound flowing to SEPPmail.cloud"
							$Setting.ExceptIfSCLOver = $SCLInboundValue
							}
						}
						"[SEPPmail.cloud] - 200 Route outgoing e-mails to SEPPmail" {
							Write-Verbose "Excluding Outbound E-Mail domains $SEPPmailCloudDomain"
							$Setting.ExceptIfSenderDomainIs = $ExcludeEmailDomain	
						}
						"[SEPPmail.cloud] - 800 Add outbound header X-SM-ruleversion" {
							"MODULEVERSION is" + $ModuleVesion
							Write-Verbose "Add rule version $Moduleversion"
							$Setting.SetHeaderValue = $Moduleversion.ToString()	
						}
					}

					if ($PSCmdlet.ShouldProcess($setting.Name, "Create transport rule"))
					{
						$Now = Get-Date
						Write-Verbose "Adding Timestamp $now to Comment"
						$setting.Comments += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
						New-TransportRule @setting |Select-Object Identity,Priority,State
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
	Remove-SC365Rules
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
		[ValidateSet('parallel','inline')]
		[String]$routing = 'parallel'
	)

	begin {
		$transportRuleFiles = Get-Childitem "$psscriptroot\..\ExOConfig\Rules\"
		if (!(Test-SC365ConnectionStatus))
		{ 
			throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
		}
		Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
	}

	process {
		Write-Verbose "Removing current version module rules"
		
		if ($routing -eq 'parallel') {
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
			Write-Warning "Routingtype 'inline' doesnt require any Mail Transport Rules. Use '-routing 'parallel'' to remove existing rules."
		}

	}
	end {

	}
}

if (!(Get-Alias 'Set-SC365rules' -ErrorAction SilentlyContinue)) {
	New-Alias -Name Set-SC365Rules -Value New-SC365Rules
}

Register-ArgumentCompleter -CommandName New-SC365Rules -ParameterName SEPPmailCloudDomain -ScriptBlock $paramDomSB

# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUrl9e1wZBE6mPrqmLqBuYCm8f
# GXaggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# MRYEFDTyQEQaYCIQP6AihYwzQv3IRY/tMA0GCSqGSIb3DQEBAQUABIIBAHFQpKaV
# 4goUntoqIySmpaqjriVmlCZEBHczaMGf8qM78lMx7v7ipdFsvWEQ0iRA2kBWjtcD
# OCVq4s7fQ8LBzKQ+jhHPIiLi7bOoJ1pV4T1pODifztLmNIR1N7r54UlctrycfIve
# lkNJ0GuHzrJUditkXhAiCKyvpA4YojPFrx6GrX+vG0pJefz2W/tT8g0ZKv8MIrOG
# Bf6ztSxCIB8hBXKbhwqOgyqSe1wKkJrE/0IxHkTpYLRhz6LIuPiSfTWhEGeR3OOk
# G1ZZzu2IGQqKRKUAFl/wpDV+i621cATiy74CdcObkcV64ht9oKQD+dl326dN81fo
# lo/qkAiUVKZ+99g=
# SIG # End signature block
