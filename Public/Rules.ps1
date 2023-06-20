<#
.SYNOPSIS
	Read existing SEPPmail.cloud transport rules in the exchange online environment.
.DESCRIPTION
	Use this to figure out if there are already SEPPmail.cloud rules implemented in Exchange online.
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

	)

	begin {
		if (!(Test-SC365ConnectionStatus))
		{ 
			throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
		}
		else 
		{
			Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
		}
		#if ($routing -eq 'p') {$routing = 'parallel'}
		#if ($routing -eq 'i') {$routing = 'inline'}
		#$transportRuleFiles = Get-Childitem "$psscriptroot\..\ExOConfig\Rules\"
	}
	process {
		#foreach ($file in $transportRuleFiles) {
		$allSEPPmailCloudRules = Get-TransportRule -Identity '[SEPPmail*'
		 	if ($allSEPPmailCloudRules) {
				Foreach ($rule in $allSEPPmailCloudRules) {
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
				}
			}
			else {
				Write-Warning "No transport rules found matching [SEPPmail.Cloud]* in your tenant."
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
	PS C:\> New-SC365Rules -SEPPmailCloudDomain 'contoso.eu','contoso.com' -routing inline
	Creates the rules for specific domains. Includes only defined e-mail domains from processing by SEPPmail.cloud
.EXAMPLE
	PS C:\> New-SC365Rules -SEPPmailCloudDomain 'contoso.eu' -PlacementPriority Top -routing parallel
	Places the transport rules BEFORE all other rules. This is unusual and against the default. It may make sense in some situations.
.EXAMPLE
	PS C:\> New-SC365Rules -SEPPmailCloudDomain 'contoso.eu' -routing parallel -disabled
	Sets the transport rules up, but keeps them inactive. Useful for a smoother integration.
.INPUTS
	none
.OUTPUTS
	transport rules
.NOTES
	
#>
function New-SC365Rules
{
	[CmdletBinding(
		SupportsShouldProcess = $true,
		ConfirmImpact = 'Medium',
		HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
		)
	]
	param
	(
		[Parameter(Mandatory = $false,
			HelpMessage = 'Should the new rules be placed before or after existing ones (if any)')]
		[ValidateSet('Top', 'Bottom')]
		[String] $PlacementPriority = 'Bottom',

		[Parameter(
			Mandatory = $true,
			HelpMessage = 'MX record->SEPPmail means routingtype inline, MX->Microsoft means routingtype parallel'
		)]
		[ValidateSet('parallel', 'inline', 'p', 'i')]
		[String]$routing,

		[Parameter(Mandatory = $true,
			HelpMessage = 'E-Mail domains you have registered in the SEPPmail.Cloud')]
		[String[]]$SEPPmailCloudDomain,

		[Parameter(Mandatory = $false,
			HelpMessage = 'SCL Value for inbound Mails which should NOT be processed by SEPPmail.Cloud. Default is 5')]
		[ValidateSet('-1', '0', '5', '6', '8', '9')]
		[int]$SCLInboundValue = 5,

		[Parameter(
			  Mandatory = $false,
			HelpMessage = 'Rule 100 will only send E-Mails to SEPPmail.cloud which requires cryptographic processing'
		)]
		[bool]$CryptoContentOnly = $true,

		[Parameter(
			Mandatory = $false,
			HelpMessage = 'Add rules if you provisioned internal e-mail signature in the SEPPmail.cloud Service'
		)]
		[switch]$Disabled
	)

	begin
	{
		if (!(Test-SC365ConnectionStatus)) { 
			throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
		} else {
			Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`" " 
		}

		if ($routing -eq 'p') {$routing = 'parallel'}
		if ($routing -eq 'i') {$routing = 'inline'}

		foreach ($validationDomain in $SEPPmailCloudDomain) {
			if ((Confirm-SC365TenantDefaultDomain -ValidationDomain $validationDomain) -eq $true) {
				Write-verbose "Domain is part of the tenant and the Default Domain"
			} else {
				if ((Confirm-SC365TenantDefaultDomain -ValidationDomain $validationDomain) -eq $false) {
					Write-verbose "Domain is part of the tenant"
				} else {
					Write-Error "Domain is NOT Part of the tenant"
					break
				}
			}
		}
 
		# Filter onmicrosoft domains
		try {
			$FilteredCloudDomain = Remove-SC365OnMicrosoftDomain -DomainList $SEPPmailCloudDomain
		} catch {
			Write-Warning "Could not remove onMicrosoft.com domains"
			$FilteredCloudDomain = $SEPPmailCloudDomain
		}

		$outboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors $true | Where-Object { $_.Name -match "^\[SEPPmail.cloud\]" }
		if ($PSCmdlet.ShouldProcess('Exchange Online', "Check existence of Outbound Connector")) {
			if(!($outboundConnectors))
			{
				throw [System.Exception] "No SEPPmail.cloud outbound connector found. InBoundOnly Mode ? Run `"New-SC365Connectors`" to add the proper SEPPmail.cloud connectors"
			}
			if ($($outboundConnectors.Enabled) -ne $true) {
				throw [System.Exception] "SEPPmail.cloud outbound-connector is disabled, cannot create rules. Create connectors without -Disable switch, or enable them in the admin portal."
			}	
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

						$Now = Get-Date
						Write-Verbose "Adding Timestamp $now to Comment"
						$setting.Comments += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"

						if ($PSCmdlet.ShouldProcess($setting.Name, "Create transport rule"))
						{
							switch ($setting.Name)
							{
								"[SEPPmail.cloud] - 060 Add header X-SM-ruleversion" {
									Write-Verbose "Add rule version $Moduleversion"
									$Setting.SetHeaderValue = $Moduleversion.ToString()
									New-TransportRule @setting #|Out-Null
								}
								"[SEPPmail.cloud] - 100 Route incoming e-mails to SEPPmail" {
									Write-Verbose "Including all managed domains $FilteredCloudDomain"
									$Setting.RecipientDomainIs = $FilteredCloudDomain
									if ($SCLInboundValue -ne 5) {
										Write-Verbose "Setting Value $SCLInboundValue to Inbound flowing to SEPPmail.cloud"
										$Setting.ExceptIfSCLOver = $SCLInboundValue
									}
									if ($cryptoContentOnly) {
										Write-Verbose 'Adding Setting to send only crptographic needed E-Mails to SEPPmail.cloud'
										$Setting.HeaderContainsMessageHeader = 'content-type'
										$Setting.HeaderContainsWords = "application/x-pkcs7-mime","application/pkcs7-mime","application/x-pkcs7-signature","application/pkcs7-signature","multipart/signed","application/pgp-signature","multipart/encrypted","application/pgp-encrypted","application/octet-stream"
									}
									New-TransportRule @setting #|Out-Null
								}
								"[SEPPmail.cloud] - 200 Route outgoing e-mails to SEPPmail" {
									Write-Verbose "Including only Outbound E-Mails from domains $FilteredCloudDomain"
									$Setting.SenderDomainIs = $FilteredCloudDomain
									New-TransportRule @setting #|Out-Null
								}
								Default {
									New-TransportRule @setting #|Out-Null
								}
							}						
						}	
					}
				# Get-TransportRule -Identity $Setting.Name|Select-Object Identity,Priority,State ## Unfinished
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
				   ConfirmImpact = 'High',
				   HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
				  )]
	param
	(
	)

	begin {
		if (!(Test-SC365ConnectionStatus))
		{ 
			throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
		} else {
			Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`" " 
		}
	}
	process {
		Write-Verbose "Removing current version module rules"
		$allSEPPmailCloudRules = Get-TransportRule -Identity '[SEPPmail*'
		foreach ($rule in $allSEPPmailCloudRules) {
			if($PSCmdlet.ShouldProcess($rule.Name, "Remove transport rule")) {
					#$rule = Get-TransportRule $setting.Name -ErrorAction SilentlyContinue
				Remove-TransportRule -Identity $rule -confirm:$false
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
# MIIVzAYJKoZIhvcNAQcCoIIVvTCCFbkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqMs55/b8EllNm
# pFeoEqqF7rbkKstyTZjJpP+DDJCD0qCCEggwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCRMjJfxvfAGY+R6EelaIz2lx4z
# KFvBONOCDoq7n+O5DzANBgkqhkiG9w0BAQEFAASCAgC4Spis6BJi9pvDjgF/dVeB
# qH35czcAm/SxWiul8dZIcNqkCFaCy0EneUXK3oKIwcFraCRcFemwq7XNuqASvEwz
# aaqpogD1FyS+CVUNHG140fmoFQFmbruUSLHGL/rAxuRYBDE3napaDZq0pE+QI6zy
# vx9Uq7Zeq7wl19rnra9AvK7UP7XgHUG0VPEkja5+HBUnZ5JbxEPDjcLGq9ABdktc
# hEqPEihZp91FSINyboe4ofVlp7zxUADeAnI16SxLrRC1DjXG9gfapHc+GI1UuEFf
# AeF/dGJz9VkQA5sswvuuIm2fZWPnbecNNxk7SRKrwdnQnCQ5/y/1lhycDj4wi80d
# TAb8ZVS8Gcx/0w+WDpNNB/yttORVI/Z5NhmolwfwAS/yyYcWBx+K65kSWA4RMWWs
# VuGo14V0HRyy7nwkHnlFLTYJM+rnsg93iT4a9/80lHTDgcfIr5lQ7TKRYB/5/4ar
# ltXgRk+FNhpQtKaHhhLq9EMY/etjFQWxpGciRL4wcRfJAQnU2pE3JtIQ1CV73fP4
# bo4db9S9GEqvBGh5PH37XZAAi9BTN+aqzGQmAtFsfOSLlILMISzB+EpT1ISW3BxZ
# hUKCOLQiqURvvdJ+0GiUk1iemDnGEYaGA2ehQrsDAGLO0FVDFeV9omjBPu2/qxZV
# 6Ux+uZtKbY0dj5OP/pcrXw==
# SIG # End signature block
