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
		[Parameter(Mandatory = $true)]
		[ValidateSet('inline','parallel')]
		$routing
	)

	begin {
		if (!(Test-SC365ConnectionStatus))
		{ 
			throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
		}
		else 
		{
			Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
			$transportRuleFiles = Get-Childitem "$psscriptroot\..\ExOConfig\Rules\"
		}
	}
	process {
		foreach ($file in $transportRuleFiles) {
			$setting = Get-SC365TransportRuleSettings -File $file -Routing $routing
			if ($setting.values) {
				$rule = Get-TransportRule $setting.Name -ErrorAction SilentlyContinue
				if ($rule) {
					if ($rule.Identity -like '*100*') {
						$rule|Select-Object Identity,Priority,State,@{Name = 'IncludedDomains'; Expression={$_.RecipientDomainIs}}
					}
					elseif ($rule.Identity -like '*200*') {
						$rule|Select-Object Identity,Priority,State,@{Name = 'IncludedDomains'; Expression={$_.SenderDomainIs}}
					}
					else {
						$rule|Select-Object Identity,Priority,State,IncludedDomains
					}
				}
				else
				{
					Write-Warning "No transport rule '$($setting.Name)'"
				}
			}
		}
	}
	end {

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
			Mandatory = $true,
			HelpMessage = 'MX record->SEPPmail means routingtype inline, MX->Microsoft means routingtype parallel'
		)]
		[ValidateSet('parallel','inline')]
		[String]$routing,

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

		 $TenantDomains = Get-AcceptedDomain
		 foreach ($namedDomain in $SEPPmailCloudDomaindomain) {
			If (!($TenantDomains.DomainName -contains $namedDomain)) {
				$PrimaryDomain = $TenantDomain|Where-Object 'Default' -eq $true|Select-Object -ExpandProperty DomainName
				Write-Information "Typo ? Domain should be $PrimaryDomain" 
				Write-Error "$SEPPmailCloudDomain is not member of the connected tenant. Retry using only tenant-domains"
				break
			 }	
		 }
 
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

				$moduleVersion = $myInvocation.MyCommand.Version
				foreach($file in $transportRuleFiles) {
				
					$setting = Get-SC365TransportRuleSettings -File $file -Routing $routing
					if ($setting.Values) {
						$setting.Priority = $placementPrio + $setting.SMPriority
						$setting.Remove('SMPriority')
						if ($Disabled -eq $true) {$setting.Enabled = $false}
	
						switch ($setting.Name)
						{
							"[SEPPmail.cloud] - 060 Add header X-SM-ruleversion" {
								Write-Verbose "Add rule version $Moduleversion"
								$Setting.SetHeaderValue = $Moduleversion.ToString()	
							}
							"[SEPPmail.cloud] - 100 Route incoming e-mails to SEPPmail" {
								Write-Verbose "Including all managed domains $SEPPmailCloudDomain"
								$Setting.RecipientDomainIs = $SEPPmailCloudDomain
								if ($SCLInboundValue -ne 5) {
									Write-Verbose "Setting Value $SCLInboundValue to Inbound flowing to SEPPmail.cloud"
								$Setting.ExceptIfSCLOver = $SCLInboundValue
								}
							}
							"[SEPPmail.cloud] - 200 Route outgoing e-mails to SEPPmail" {
								Write-Verbose "Including only Outbound E-Mails from domains $SEPPmailCloudDomain"
								$Setting.SenderDomainIs = $SEPPmailCloudDomain	
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
			Mandatory = $true,
			HelpMessage = 'Use seppmail if the MX record points to SEPPmail and microsoft if the MX record points to the Microsoft Inrastructure'
		)]
		[ValidateSet('parallel','inline')]
		[String]$routing
	)

	begin {
		if (!(Test-SC365ConnectionStatus))
		{ 
			throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
		}
		Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
		$transportRuleFiles = Get-Childitem "$psscriptroot\..\ExOConfig\Rules\"
	}
	process {
		Write-Verbose "Removing current version module rules"
		foreach ($file in $transportRuleFiles) {
			$setting = Get-SC365TransportRuleSettings -routing $routing -file $file
			if ($setting.values) {
				if($PSCmdlet.ShouldProcess($setting.Name, "Remove transport rule")) {
				$rule = Get-TransportRule $setting.Name -ErrorAction SilentlyContinue
				if($rule -ne $null)
					{$rule | Remove-TransportRule -Confirm:$false}
				}
			}
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
# MIIL/AYJKoZIhvcNAQcCoIIL7TCCC+kCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDbWmKksjlPG3m0
# z9UKSepXcWaprEjvAcDUI07RjsXUkKCCCUAwggSZMIIDgaADAgECAhBxoLc2ld2x
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDHk6rrjbCxXZ9zfbk9Ev4hUSwf
# JA9+EaBN7Jf1Z3dz5DANBgkqhkiG9w0BAQEFAASCAQCPOxut7iYXh1BN9MHh8XWR
# 4C8zsIsFOAuSbgm5LF0brxiZXOP3/WjHZxLpuWjVgA8/U/bUJyTbjfvIf97qQOul
# GvpygmISQRAk6aXP310ciXTxwMO9lmOE15g9JysV4U40yjKjPu6S8mSgZZtHXAzG
# l4nytTTt9N9AaoScu+yrtMFszd21rf+pG4+l0MsqCZSw8upNjUpErCm5mEiA3cUt
# lMoClijgO1dnESkQL2fQz1kGUccXOd0DCRFY8N+37Jv1D/+IAXJs0txNIjNbtC0o
# ZxUNVEJ4JTVmsB9jon+MFb3Srgb+VKfQ20XnF7wzjXP5rKayXQoBp5WU4X8wCDpK
# SIG # End signature block
