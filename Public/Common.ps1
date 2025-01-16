<#
.SYNOPSIS
    Detects the SEPPmail.Cloud Deploymemnt status, based in M365 Tenant information
.DESCRIPTION
    Checks Deplyoment Status of the SEPPmail.Cloud based on M365 Tenantinformation from outside, based on DNS information
    
    - Queries routingmode based on available MX records in the SEPPmail.cloud
    - Queries region based on IP adresses and mailhosts
    - Tests if MX record is set correctly in Inline Mode
    - Checks if the SEPPmail.cloud is prepared for Certificate-based-Connectors
    
    Creates a PSObject with the following values:
    
    Routing = inline/parallel                           # Routing Mode
    Region = ch/de/prv                                  # Cloud Region (Datacenter location)
    SEPPmailCloudDomain = 'contoso.de','contoso.ch'     # Maildomains which will be routet via SEPPmail. Is basis for naming the mailrouting hosts (gate/relay/mail) 
    CBCenabled = $true/$false                           # Certificate Based Connectors setup available
    CBCConnectorHost = '<tenantid>.<rg>.seppmail.cloud' # Hostname of TLS host for CBC
    InlineMXMatch = $true/$false                        # (Inline Mod only) MX record points to the correct (SEPPmail) host
    RelayHost = domain-tls.relay.seppmail.cloud         # Name of relay host
    GateHost = domain-tls.gate.seppmail.cloud           # Name of gate host
    MailHost = domain-tls.mail.seppmail.cloud           # Name of mail host

.NOTES
    Emits parameters for New-SC365Setup
.LINK
    https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md
.EXAMPLE
    Get-SC365DeploymentInfo

    DeployMentStatus    : True
    SEPPmailCloudDomain : contoso.com
    Region              : ch
    Routing             : inline
    InBoundOnly         : False
    CBCDeployed         : True
    CBCConnectorHost    : 271dd771-832d-4913-80d7-9c21616accd4.ch.seppmail.cloud
    CBCDnsEntry         : c60abc9d247a2bf21cbc3344eef199eb738876b2.cbc.seppmail.cloud
    InlineMXMatch       : True
    MailHost            : 
    RelayHost           : contoso-com.relay.seppmail.cloud
    GateHost            : contoso-com.gate.seppmail.cloud
.EXAMPLE
    Get-SC365DeploymentInfo -SEPPmailCLoudDomain contoso.eu

    DeployMentStatus    : True
    SEPPmailCloudDomain : contoso.eu
    Region              : de
    Routing             : inline
    InBoundOnly         : False
    CBCDeployed         : True
    CBCConnectorHost    : 271dd771-832d-4913-80d7-9c21616accd4.de.seppmail.cloud
    CBCDnsEntry         : c60abc9d247a2bf21cbc3344eef199eb738876b2.cbc.seppmail.cloud
    InlineMXMatch       : True
    MailHost            : 
    RelayHost           : contoso-eu.relay.seppmail.cloud
    GateHost            : contoso-eu.gate.seppmail.cloud

#>
function Get-SC365DeploymentInfo {
    [CmdletBinding()]
    param (
        [Parameter(   
            Mandatory   = $false,
            HelpMessage = "Domain name you selected for Tenant-onboarding"
         )]
         [string[]]$SEPPmailCloudDomain
    )
    
    begin {
        if (!(Test-SC365ConnectionStatus)){
            throw [System.Exception] "You're not connected to Exchange Online - please connect to the designated tenant prior to using this CmdLet" }
        else {
            Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        }
        # Reset the Stage
        $DeployMentStatus = $null
        $region = $null
        $routing = $null
        $inBoundOnly = $null

        $DeploymentInfo = [PSCustomObject]@{
            DeploymentStatus    = $null
            SEPPmailCloudDomain = $Null
            Region              = $null
            Routing             = $null
            InboundOnly         = $null
            CBCDeployed         = $Null
            CBCConnectorHost    = $null
            CBCDnsEntry         = $null
            InlineMXMatch       = $null
            MailHost            = $null
            GateHost            = $null
            RelayHost           = $null
            SwissSignCheckTXT   = $null
            spfTXT              = $null
            DnsSecurEmailCNAME  = $null
            DnsLetsEncryptCNAME = $null
            DnsDKIMTXT          = $null
            DnsWildCardActive   = $null
        }
    }
    
    process {
        #region Select DefaultDomain
            if (!($SEPPmailCloudDomain)) {
                [String]$DNSHostDomain = $tenantAcceptedDomains |Where-Object 'Default' -eq $true |select-Object -ExpandProperty DomainName
                Write-Verbose "Extracted Default-Domain with name $DNSHostDomain from TenantDefaultDomains"
                }
            else {
                foreach ($dom in $SEPPmailCloudDomain) {
                    if (!($tenantAcceptedDomains.DomainName.Contains($Dom))) {
                        throw [System.Exception] "Domain $Dom is not member of this tenant! Check for typos or connect to different tenant"
                    }
                    else {
                        Write-Verbose "Domain $Dom is member of the tenant domains"
                        if ($dom -eq ($tenantAcceptedDomains |Where-Object 'Default' -eq $true |select-Object -ExpandProperty DomainName)) {
                            $DnsHostDomain = $dom
                        }
                        else {
                            Write-Verbose "TenantDefaultDomain not selected, using $dom ans DNSHostDoman"
                            $DnsHostDomain = $dom
                        }
                    }
                }
            }

        #endregion Select DefaultDomain

        #region Query SEPPmail routing-Hosts DNS records and detect routing mode and in/oitbound
            [string]$relayHost = $DnsHostDomain.Replace('.','-') + '.relay.seppmail.cloud'
             [string]$mailHost = $DnsHostDomain.Replace('.','-') + '.mail.seppmail.cloud'
             [string]$gateHost = $DnsHostDomain.Replace('.','-') + '.gate.seppmail.cloud'
        
        $defEAPref = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'

        $DeploymentStatus = 'unknown'
        if (((Resolve-Dns -Query $GateHost).Answers)) {
            $routing = 'inline'
            if ((Resolve-Dns -Query $RelayHost).Answers) {                    
                $inBoundOnly = $false
                Write-Verbose "$GateHost and $relayHost alive ==> InLine-bidirectional"
                $deploymentStatus = $true                
            } else {
                if (!((Resolve-Dns -Query $RelayHost).Answers)) {
                    $inBoundOnly = $true
                    $relayHost = $null
                    Write-Verbose "$GateHost alive,$relayHost missing ==> InLine-InBound only"
                    $deploymentStatus = $true                
                }
            }
        }
        else {
            if ((Resolve-Dns -Query $MailHost ).Answers) {
                Write-verbose "$mailHost alive ==> parallel"
                $routing = 'parallel'
                $deploymentStatus = $true                
            } else {
                $mailHost = $null
                $deploymentStatus = $false
            }            
        }
        #endregion Mailhost queries

        #region DoubleCheck if MX Record is set correctly
            $mxFull = get-mxrecordreport -Domain $DnsHostDomain
            
        if ($mxFull.Count -eq 0) {
            $DeployMentStatus = $false
        }
        else {
            if ($mxFull.Count -eq 1) {
                $mx = $mxFull | Select-Object -ExpandProperty highestpriorityMailHost -Unique
            } 
            if ($mxFull.Count -gt 1) {
                $mx = $mxFull[0] | Select-Object -ExpandProperty highestpriorityMailHost -Unique
            }
                
            if (($mx.Split($DnsHostDomain.Replace('.', '-'))) -eq '.gate.seppmail.cloud') {
                Write-Verbose "MX = SEPPmail"
            }
            if (($mx.Split($DnsHostDomain.Replace('.', '-'))) -eq '.mail.protection.outlook.com') {
                Write-Verbose "MX = Microsoft"
            }
            if ($routing -eq 'inline') {
                if ($mx -eq $gateHost) {
                    Write-Verbose "MX record $mx in M365 Config matches $gateHost"
                    $mxMatch = $true
                }
                else {
                    Write-Warning "MX Record $mx configured in M365 does not fit to SEPPmail GateHost $gateHost in DNS - Check your provisioning Status in SEPPmail.cloud Portal"
                    $mxMatch = $false
                    $DeployMentStatus = $false
                }
            }   
        }
        #endRegion MX Check

        #region Identify region based on Cloud-IPAddresses
            $region = $null
            $ch = Get-SC365CloudConfig -region 'ch'
            $de = Get-SC365CloudConfig -region 'de'
            $prv = Get-SC365CloudConfig -region 'prv'
            $dev = Get-SC365CloudConfig -region 'dev'
            if ($routing -eq 'inline') {
                [String[]]$GateIP = ((Resolve-Dns -Query $GateHost).Answers)|Select-Object -expand Address| Select-Object -expand IPAddressToString
                Foreach ($IP in $GateIP) {if ($ch.IPv4GateIPs.Contains($Ip)) {$region = 'ch';break}}
                Foreach ($IP in $GateIP) {if ($de.IPv4GateIPs.Contains($Ip)) {$region = 'de';break}}
                Foreach ($IP in $GateIP) {if ($prv.IPv4GateIPs.Contains($Ip)) {$region = 'prv';break}}
                Foreach ($IP in $GateIP) {if ($dev.IPv4GateIPs.Contains($Ip)) {$region = 'dev';break}}
            }
            if ($routing -eq 'parallel') {
                [string[]]$MailIP = ((Resolve-Dns -Query $MailHost).Answers)|Select-Object -expand Address| Select-Object -expand IPAddressToString
                Foreach ($ip in $mailIp) {if ($ch.IPv4MailIPs.Contains($Ip)) { $region = 'ch';break}}
                Foreach ($ip in $mailIp) {if ($de.IPv4MailIPs.Contains($Ip)) { $region = 'de';break}}
                Foreach ($ip in $mailIp) {if ($prv.IPv4MailIPs.Contains($IP)) { $region = 'prv';break}}
                Foreach ($ip in $mailIp) {if ($dev.IPv4MailIPs.Contains($IP)) { $region = 'dev';break}}
            }
        #endregion Cloud-IP-Addresses

        #region Check CBC Availability
            [String]$TenantID = Get-SC365TenantID -MailDomain $DnsHostDomain -OutVariable "TenantID"
            $TenantIDHash = Get-SC365StringHash -String $TenantID
            [string]$hashedDomain =  $TenantIDHash + '.cbc.seppmail.cloud'
            if (((resolve-dns -query $hashedDomain -QueryType TXT).Answers)) {
               $CBCDeployed = $true
               Write-Verbose "$HashedDomain of TenantID $tenantId has a CBC entry"
            } else {
               $CBCDeployed = $false
               Write-Warning "Could not find TXT Entry for TenantID $TenantID of domain $DNSHostCloudDomain. Setup will most likely fail! Go to the SEPPmail.cloud-portal and check the deployment status."
            }
        #endregion CBC 
        
        #region Advanced DNS Queries
        
        # TXT records (Swissign check and SPF)
        $txtRecords = (Resolve-dns -querytype TXT $DNSHostdomain).Answers
        if ($txtRecords) { 
            $swisssignTXTRecord = $txtRecords|Where-Object EscapedText -like 'swisssign-check*'
            if ($swissSignTXTRecord) {
                $swissSignCheck = $swissSignTXTRecord.EscapedText.Trim('{','}')
            }
            else {
                Write-Warning "Swisssign TXT (swissSign-check) record is missing - SC-CERT deployment will fail!"
            }
            $spfTXTrecord = $txtRecords|Where-Object EscapedText -like 'v=spf*'
            if ($spfTXTrecord) {
                $spf = $spfTXTrecord.EscapedText.Trim('{','}')
            }
            else {
                Write-Warning "SPF TXT (v=spf*) record is missing"
            }
        }

        ## WebService hosts
        [String]$SecurEmailCNAME = (Resolve-dns -querytype CNAME -Query ('securemail.' + $DNSHostdomain)).Answers.CanonicalName.Value

        ## Letsencrypt
        [String]$LetsEncryptCNAME = (Resolve-dns -querytype CNAME -Query ('_acme-challenge.securemail.' + $DNSHostdomain)).Answers.CanonicalName.Value

        ## DKIM (default._domainkey)
        [String]$dkimRecord = (resolve-dns -QueryType TXT -query ('default._domainkey.' + $DNSHostDomain)).Answers.EscapedText

        ## Wildcard records
        if ((resolve-dns -query ('jioak84-nlkjec.' + $DNSHostDomain)).Answers.EscapedText) {
            $WildcardRecord = $true
        }
        else {
            $WildcardRecord = $false
        }
        $ErrorActionPreference = $defEAPref
        #endregion
    }
    end {
        $DeploymentInfo.DeploymentStatus = $DeploymentStatus
        $DeploymentInfo.Region = $region
        $DeploymentInfo.Routing = $routing
        $DeploymentInfo.InboundOnly  = $inBoundOnly
        $DeploymentInfo.SEPPmailCloudDomain = $DNSHostDomain
        $DeploymentInfo.CBCDeployed = $CBCDeployed
        if ($region) {$DeploymentInfo.CBCConnectorHost = ($tenantId + ((Get-Variable $region).Value.TlsCertificate).Replace('*',''))}
        if ($CBCDeployed -eq $true) {$DeploymentInfo.CBCDnsEntry = ($TenantIDHash + '.cbc.seppmail.cloud')}
        if ($routing -eq 'inline') {$DeploymentInfo.InlineMXMatch = $MxMatch}
        if (($routing -eq 'inline') -and (!($inBoundOnly))) {$DeploymentInfo.RelayHost = $relayHost}
        if ($routing -eq 'inline') {$DeploymentInfo.GateHost = $gateHost}
        if ($routing -eq 'parallel') {$DeploymentInfo.MailHost = $MailHost}
        # DNS Records
        $DeploymentInfo.swisssignCheckTXT = $swisssignCheck
        $DeploymentInfo.spfTXT = $spf
        $DeploymentInfo.DnsSecurEmailCNAME = $SecurEmailCNAME
        $DeploymentInfo.DnsLetsEncryptCNAME = $LetsEncryptCNAME
        $DeploymentInfo.DnsDKIMTXT = $DKIMRecord
        $DeploymentInfo.DnsWildCardActive = $wildcardRecord
        return $DeploymentInfo
    }
}

<#
.SYNOPSIS
    Removes all Rules and Connectors
.DESCRIPTION
    Based on autodiscovery, or forced values through parameters, Remove-SC365Setup removes all connectors and rules from an Exo-Tenant
.NOTES
    - none -
.LINK
    https://github.com/seppmail/seppmail365cloud
.EXAMPLE
    Remove-SC365Setup
    # Without any parameters, it runs discovery mode and removes rules and connectors
.EXAMPLE
    Remove-SC365Setup -parallel
    # Forces to remove parallel setup config
.EXAMPLE
    Remove-SC365Setup -inline
    # Forces to remove inline setup config
.EXAMPLE
    Remove-SC365Setup -inline -inBoundOnly
    # Forces to remove inline in "InbohndOnly" mode setup config
#>
function Remove-SC365Setup {
    [CmdletBinding(
        SupportsShouldProcess = $true,
        ConfirmImpact = 'Medium',
        DefaultParameterSetName='parallel'
        )]
    
    param(
        [Parameter(
            ParameterSetName = 'parallel',
            Mandatory=$false,
            HelpMessage="Inline routing via SEPPmail (MX ==> SEPPmail), or routing via Microsoft (MX ==> Microsoft)"
            )]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('parallel','inline','p','i')]
        [Parameter(
            ParameterSetName = 'inline',
            Mandatory=$false,
            HelpMessage="Inline routing via SEPPmail (MX ==> SEPPmail), or routing via Microsoft (MX ==> Microsoft)"
            )]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('parallel','inline','p','i')]
        [String]$routing,

        [Parameter(
            ParameterSetName = 'inline',
            Mandatory=$false,
            HelpMessage="No routing of outbound traffic via SEPPmail.cloud"
            )]
        [switch]$InBoundOnly
    )
    Begin {
        if(!(Test-SC365ConnectionStatus)) {
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"        
        }
        else {
            if ((!($InboundOnly)) -or (!($routing)) ) {
                try {
                    $deploymentInfo = Get-SC365DeploymentInfo
                } catch {
                    Throw [System.Exception] "Could not autodetect SEPPmail.cloud Deployment Status, use manual parameters"
                }
                if ($DeploymentInfo.DeployMentStatus -eq $false) {
                    Write-Error "SEPPmail.cloud setup not (fully) deployed. Use Cloud-Portal and fix deployment."
                    break
                } 
                else {
                    if ($Deploymentinfo) {
                                       if ($deploymentInfo.Routing) {$Routing = $deploymentInfo.Routing} else {Write-Error "Cloud not autodetect routig info, use manual parameters"; break}
                         if ($DeploymentInfo.inBoundOnly -eq $true) {$inboundOnly = $true}
                        if ($DeploymentInfo.inBoundOnly -eq $false) {$inboundOnly = $false}
                         if ($null -eq $DeploymentInfo.inBoundOnly) {$inboundOnly = $false}
                    }
                }
            } 
            else {
                if ($deploymentInfo.routing -eq 'p') {$routing = 'parallel'}
                if ($deploymentInfo.routing -eq 'i') {$routing = 'inline'}
            }
        }
    }
    Process {
        Write-Verbose "Creating Progress Bar"
        $objectCount = $null
        # Count Rules
        foreach ($file in (Get-Childitem -Path "$psscriptroot\..\ExOConfig\Rules\")) {
            $objectCount += if ((Get-SC365TransportRuleSettings -routing $routing -file $file).count -gt 0) {1}
        }

        # Count Connectors
        #if ((Get-SC365InboundConnectorSettings -routing $routing -file $file).count -gt 0) {1}
        $objectCount += 2

        try {
            if ($InBoundOnly) {
                #Write-Progress -Activity "Removing SEPPmail.Cloud Setup" -Status "Removing Rules" -PercentComplete (0)
                Write-Information '--- Remove connector(s) ---' -InformationAction Continue
                Remove-SC365Connectors -routing $routing -Inboundonly:$inboundonly
            }
            else {
                #Write-Progress -Activity "Removing SEPPmail.Cloud Setup" -Status "Removing Rules" -PercentComplete (0)
                Write-Information '--- Removing transport rules ---' -InformationAction Continue
                Remove-SC365Rules
                Write-Information '--- Remove connector(s) ---' -InformationAction Continue
                Remove-SC365Connectors -routing $routing -Inboundonly:$inboundonly
            }
        } catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
            break
        }
    }
    End{
        Write-Information "--- Successfully removed SEPPmail.cloud Setup in $routing mode ---" -InformationAction Continue

    }
}

<#
.SYNOPSIS
    Creates all Rules and Connectors for SEPPmail.cloud
.DESCRIPTION
    Based on autodiscovery, or forced values through parameters, New-SC365Setup creates all connectors and rules for an Exo-Tenant
.NOTES
    - none -
.LINK
    https://github.com/seppmail/seppmail365cloud
.EXAMPLE
    New-SC365Setup
    # Without any parameters, it runs discovery mode and created rules and connectors
.EXAMPLE
    New-SC365Setup -force
    # The force parameter will force the removal of an existig setup and recreate connectors and rules 
.EXAMPLE
    New-SC365Setup -SEPPmailCloudDomain contoso.com -routing parallel -region ch
    # Creates a setup for one domain in parallel mode and in region Switzerland
.EXAMPLE
    New-SC365Setup -SEPPmailCloudDomain contoso.de -routing inline -region de
    # Creates a setup for one domain in inline mode and in region Germany/EU
.EXAMPLE
    New-SC365Setup -SEPPmailCloudDomain contoso.de -routing inline -region de -inboundonly
    # Creates a setup for one domain in inline mode and in region Germany/EU inbound only.
#>
function New-SC365Setup {
    [CmdLetBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact = 'Medium',
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md'
    )]

    # Specifies a path to one or more locations.
    param(
        [Parameter(
            Mandatory=$false,
            HelpMessage="The primary domain, booked in the SEPPmail.cloud"
            )]
            [Alias('domain')]
            [ValidateNotNullOrEmpty()]
            [String[]]$SEPPmailCloudDomain,
    
        [Parameter(
            Mandatory=$false,
            HelpMessage="Inline routing via SEPPmail (MX ==> SEPPmail), or routing via Microsoft (MX ==> Microsoft)"
            )]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('parallel','inline','p','i')]
        [String]$routing,
    
        [Parameter(
            Mandatory=$false,
            HelpMessage="Physical location of your data"
            )]
            [ValidateSet('dev','prv','de','ch')]
        [String]$region,

        [Parameter(
            Mandatory=$false,
            HelpMessage="No routing of outbound traffic via SEPPmail.cloud"
            )]
        [switch]$InBoundOnly,

        [Parameter(
            Mandatory=$false,
            HelpMessage="Removes existing setup"
        )]
        [switch]$force
    )

    Begin {
        if(!(Test-SC365ConnectionStatus)) {
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
        } else {
            Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        }

        # If user tries to use *.onmicrosoft.com domain ==> BREAK
        if ($SEPPmailCloudDomain -like '*.onmicrosoft.com') {
            Write-Error "Domain $SEPPmailcloudDomain is not intended for E-Mail sending and cannot be booked for the SEPPmail-cloud Service. Specify a custom domain of your tenant and retry."
            break
        }
        Write-Verbose "Detecting Deploymentstatus frpm SEPPmail.cloud setup"
        try {
            $deploymentInfo = Get-SC365DeploymentInfo
        } catch {
            Throw [System.Exception] "Could not autodetect SEPPmail.cloud deployment status, check SEPPmail.cloud portal deployment status"
        }

        Write-Verbose "Not enough parameters given, reading from Tenant, otherwise use data from console"
        if ((!($SEPPmailCloudDomain)) -or (!($region)) -or (!($routing)) ) {
            # Customers where TDAD is set to *.onmicrosoft.com  ==> BREAK
            if ($DeploymentInfo.SEPPmailCloudDomain -like '*.onmicrosoft.com') {
                Write-Error "Domain $($DeploymentInfo.SEPPmailCloudDomain) is set as the tenant default accepted domain. $($DeploymentInfo.SEPPmailCloudDomain) is not intended for E-Mail sending and cannot be booked for the SEPPmail-cloud Service. Specify a custom domain of your tenant and retry or change the Default accepted domain in your Exchange Online tenant."
                break
            }
            if ($DeploymentInfo.DeployMentStatus -eq $false) {
                Write-Error "SEPPmail.cloud setup for domain $deploymentinfo.SEPPmailCloudDomain is not (fully) deployed. Use Cloud-Portal and fix deployment."
                break
            } else {
                if ($Deploymentinfo) {
                            if ($deploymentInfo.Routing) {$Routing = $deploymentInfo.Routing} else {Write-Error "Cloud not autodetect routig info, use manual parameters"; break}
               if ($deploymentInfo.Routing -ne $routing) {Write-Error "SEPPmail.cloud is deployed with routing $deploymentInfo.Routing but the routing parameter is set to $routing, this will NOT WORK, exiting ..."; break}
                             if ($deploymentInfo.Region) {$Region = $deploymentInfo.Region} else {Write-Error "Could not autodetect region. Use manual parameters"; break}
                 if ($deploymentInfo.Region -ne $region) {Write-Error "SEPPmail.cloud is deployed in region $deploymentInfo.Region but the region parameter is set to $region, this will NOT WORK, exiting ..."; break}
                if ($DeploymentInfo.SEPPmailCloudDomain) {$SEPPmailCloudDomain = $DeploymentInfo.SEPPmailCloudDomain} else {Write-Error "Could not autodetect SEPPmailCloudDomain. Use manual parameters"; break}          
              if ($DeploymentInfo.inBoundOnly -eq $true) {$inboundOnly = $true}
             if ($DeploymentInfo.inBoundOnly -eq $false) {$inboundOnly = $false}
              if ($null -eq $DeploymentInfo.inBoundOnly) {$inboundOnly = $false}
                }
            }
        } else {
            if ($deploymentInfo.routing -eq 'p') {$routing = 'parallel'}
            if ($deploymentInfo.routing -eq 'i') {$routing = 'inline'}

            Write-Verbose "Checking if console parameter fit to deployment Info"
            if ($SEPPmailCloudDomain -ne $deploymentInfo.SEPPmailCloudDomain) {Write-Warning "Domain `"$SEPPmailCloudDomain`" does not fit to collected deployment info, just detected domain `"$($deploymentInfo.SEPPmailcloudDomain)`""}
                                    if ($routing -ne $deploymentInfo.routing) {Write-Error "Routing mode `"$routing`" does not fit to deployment info, just detected routing mode `"$($DeploymentInfo.routing)`" STOPPING because deployment will FAIL";break}
                                      if ($region -ne $deploymentInfo.region) {Write-Error "Region `"$region`" does not fit to deployment info, just detected region `"$($deploymentInfo.region)`" STOPPING because deployment will FAIL";break}

            Write-Verbose "Confirming if $SEPPmailCloudDomain is part or the tenant"
            $TenantDefaultDomain = $null
            foreach ($validationDomain in $SEPPmailCloudDomain) {
                if ((Confirm-SC365TenantDefaultDomain -ValidationDomain $validationDomain) -eq $true) {
                    Write-verbose "Domain is part of the tenant and the Default Domain"
                    $TenantDefaultDomain = $ValidationDomain
                } else {
                    if ((Confirm-SC365TenantDefaultDomain -ValidationDomain $validationDomain) -eq $false) {
                        Write-verbose "Domain is part of the tenant"
                    } else {
                        Write-Error "Domain is NOT Part of the tenant"
                        break
                    }
                }
             }    
        }

    }
    Process {
        try {
            if ($force) {
                    Remove-SC365Setup
            }    
        } catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
            Write-Error "Setup removal failed. Try removing SEPPmail.cloud Rules and Connectors from the Microsoft portal admin.microsoft.com or with native Exchange Online PowerShell Module CmdLets."
            break
        }

        # For Connectors - use Tenant Default Domain
        # For TransportRules, use all domains in the array
        if ($SEPPmailCloudDomain.count -le 1) {
            $ConnectorDomain = $SEPPmailCloudDomain[0]
        } else {
            $ConnectorDomain = $TenantDefaultDomain
        }

        try {
            if ($InBoundOnly -eq $true) {
                    Write-Information '--- Creating inbound connector ---' -InformationAction Continue
                    New-SC365Connectors -SEPPmailCloudDomain $ConnectorDomain -routing $routing -region $region -inboundonly:$true
            } else {
                    Write-Information '--- Creating in and outbound connectors ---' -InformationAction Continue
                    New-SC365Connectors -SEPPmailCloudDomain $ConnectorDomain -routing $routing -region $region
            }
        } catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
            break
        }
        try {
            if ($inboundonly -eq $false) {
                    Write-Information '--- Creating transport rules ---' -InformationAction Continue
                    New-SC365Rules -SEPPmailCloudDomain $SEPPmailCloudDomain -routing $routing

            }
        } catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
            break
        }
    }
    End{
        Write-Information "--- Successfully created SEPPmail.cloud Setup for $SEPPmailCloudDomain in region $region in $routing mode ---" -InformationAction Continue
        Write-Information "--- Wait a few minutes until changes are applied in the Microsoft cloud ---" -InformationAction Continue
        Write-Information "--- Afterwards, start testing E-Mails in and out ---" -InformationAction Continue
    }
}

<#
.SYNOPSIS
    Reads all Rules and Connectors for SEPPmail.cloud in an Exo-Tenant
.DESCRIPTION
    Based on autodiscovery, or forced values through parameters, Get-SC365Setup reads all connectors and rules from an Exo-Tenant
.NOTES
    - none -
.LINK
    https://github.com/seppmail/seppmail365cloud
.EXAMPLE
    Get-SC365Setup
    # Without any parameters, it runs discovery mode and reads rules and connectors
.EXAMPLE
    Get-SC365Setup -parallel
    # Reads parallel setup config
.EXAMPLE
    Get-SC365Setup -inline
    # Reads inline setup config
.EXAMPLE
    Get-SC365Setup -inline -inBoundOnly
    # Reads inline in "InbohndOnly" mode setup config
#>
function Get-SC365Setup {
    [CmdletBinding(DefaultParameterSetName='parallel')]
    
    param(
        [Parameter(
            ParameterSetName = 'parallel',
            Mandatory=$false,
            HelpMessage="Inline routing via SEPPmail (MX ==> SEPPmail), or routing via Microsoft (MX ==> Microsoft)"
            )]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('parallel','inline','p','i')]
        [Parameter(
            ParameterSetName = 'inline',
            Mandatory=$false,
            HelpMessage="Inline routing via SEPPmail (MX ==> SEPPmail), or routing via Microsoft (MX ==> Microsoft)"
            )]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('parallel','inline','p','i')]
        [String]$routing,

        [Parameter(
            ParameterSetName = 'inline',
            Mandatory=$false,
            HelpMessage="No routing of outbound traffic via SEPPmail.cloud"
            )]
        [switch]$InBoundOnly
    )
    Begin {
        if(!(Test-SC365ConnectionStatus)) {
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
        } else {
            Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        }
        if (!((!($InboundOnly)) -or (!($routing)))) {
            try {
                $deploymentInfo = Get-SC365DeploymentInfo
            } catch {
                Throw [System.Exception] "Could not autodetect SEPPmail.cloud Deployment Status, use manual parameters"
            }
            
            if ($DeploymentInfo.DeployMentStatus -eq $false) {
                Write-Error "SEPPmail.cloud setup not (fully) deployed. Use Cloud-Portal and fix deployment."
                break
            } else {
                if ($Deploymentinfo) {
                    if ($deploymentInfo.Routing) {$Routing = $deploymentInfo.Routing} else {Write-Error "Cloud not autodetect routig info, use manual parameters"; break}
                    if ($deploymentInfo.InBoundOnly -eq $true) {$InBoundOnly = $deploymentInfo.InBoundOnly} else {$InBoundOnly = $false}
                }
            }
        } else {
            if ($deploymentInfo.routing -eq 'p') {$routing = 'parallel'}
            if ($deploymentInfo.routing -eq 'i') {$routing = 'inline'}
        }
    }
    Process {
        try {
            if ($InBoundOnly -eq $true) {
                Write-Verbose "Get SEPPmail.cloud Connectors in inbound-only mode"
                $smcConn = Get-SC365Connectors -Routing $routing -inboundonly:$true
            } else {
                Write-Verbose "Get SEPPmail.cloud Connectors"
                $smcConn = Get-SC365Connectors -Routing $routing -inboundonly:$false
            }
            if ($InBoundOnly -eq $false) {
                Write-Verbose "Get SEPPmail.cloud Transpprt Rules"
                $smcTRules = Get-SC365Rules
            }    
        }
        catch {
            Write-Warning "Found no or incomplete setup, please check manually in EAC."
            break
        }
    }
    End{
        Out-Host -InputObject $smcConn -Paging
        if ($InBoundOnly -eq $false) {
            Out-Host -InputObject $smcTRules -Paging
        }
     }
}

function Update-SC365Setup {
    [CmdletBinding(
        SupportsShouldProcess = $true
        )
    ]
    param(
        # Indicates whether the old setup should be removed during the update process.
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Specify if the old setup should be removed during the update process."
        )]
        [bool]$remove = $false,

        # Specifies the name for the backup to be created during the update process.
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Provide a custom name for the backup mailflow object during the update process."
        )]
        [string]$BackupName = 'SC-BKP',

        # Specifies the name for the backup to be created during the update process.
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Prefix for connectors for temporary swapping rules traffic"
        )]
        [string]$TempPrefix = 'temp'
    )

    begin {
        $existEAValue = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
    }
    process {
        #region Infoblock
        Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "| This script is a helper and provides basic steps to upgrade your    |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "| SEPPmail.cloud/Exchange integration. It covers only STANDARD Setups!|" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "| The Script will:                                                    |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    1.) Check if there are any orphaned rule or connector objects    |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    2.) Rename SEPPmail.cloud Transport rules to `$backupName         |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    3.) Create Connectors with Temp Name                             |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    4.) Set (200) outbound transport rule to New-Connector           |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    5.) Rename SEPPmail.cloud Connectors to `$backupName              |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    6.) Attach old Transport rules to old Connector with BackupNam   |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    ----------------- OLD SETUP STILL RUNNING ------------------     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    7.) Rename NEW Connectors to original names                      |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    8.) Create new transport rules -PlacementPriority TOP            |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    -------------------- NEW SETUP RUNNING ----------------------    |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|    9.) Disable old rules                                            |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|   10.) Disable old connectors                                       |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|   11.) on -remove delete old transport rules and connectors         |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "| If you have any:                                                    |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|   - customizations to SEPPmail.cloud rules                          |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|   - other corporate transport rules                                 |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|   - disclaimer Services integrated via rules                        |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|   - or other special scenarios in your Exo-Tenant                   |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "| you need to adapt/change/post-configure the outcome of this script! |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "| DO NOT JUST FIRE IT UP AND HOPE THINGS ARE GOING TO WORK !!!!!!     |" -ForegroundColor Magenta -BackgroundColor Black
        Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Magenta -BackgroundColor Black
        #endregion    
        $response = Read-Host "I have read and understood the above warning (Type MURPHY if you agree)!"
        if ($response -eq 'MURPHY') {

            Write-Verbose '1 - Checking if there are existing Backup objects in the Exchange Tenant'
            $backupWildCard = '[' + $backupName + '*'
            Write-Verbose "1a - Checking if there are any orphaned backup objects"
            if (!((Get-InboundConnector -Identity $BackupWildcard -ea SilentlyContinue) -or (Get-OutboundConnector -Identity $backupWildCard -ea SilentlyContinue) -or (Get-TransportRule -Identity $backupWildCard -ea SilentlyContinue))) {
                #region 1a - Get-DeploymentInfo
                Write-Verbose "1b - Getting DeploymentInfo"
                $DeplInfo = Get-SC365DeploymentInfo

                #region 2 - rename existing rules to backup name
                Write-Verbose "2 - Rename existing SEPPmail.cloud rules"
                $oldTrpRls = Get-TransportRule -Identity '[SEPPmail.cloud]*'
                foreach ($rule in $oldTrpRls) {
                    Set-TransportRule -Identity $rule.Name -Name ($rule.Name -replace 'SEPPmail.Cloud', $BackupName) @psBoundParameters
                }
                #endregion rename existing rules to backup name

                #region 3 - create new connectors with temp Name
                Write-Verbose "3 - Creating new connectors with temp name" 
                #$newConnectors = New-SC365connectors -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -region $DeplInfo.region -routing $DeplInfo.routing -NamePrefix $tempPrefix # -routing $($DeplInfo.Routing) -region $($DeplInfo.region) #FIXME:

                if ($PSCmdlet.ShouldProcess('New connectors','create')){
                    $newConnectors = New-SC365connectors -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -region $DeplInfo.region -routing $DeplInfo.routing -NamePrefix $tempPrefix @psBoundParameters
                }
                #endregion create new connectors with temp Name
                #region 4 - set outbound rule to new connector (Inline 200, parallel 1xx)
                Write-Verbose "4 - Set outbound rules to new connector" 
                $oldObc = Get-OutboundConnector -Identity '[SEPPmail.cloud] Outbound-*'
                [string]$tempObcName = $TempPrefix + $($OldObc.Identity)
                $rulesToChange = Get-TransportRule -Identity '[*' | Where-Object {($null -ne $_.RouteMessageOutboundConnector) -and ($_.Name -like "*$backupName*")}
                foreach ($rule in $rulesToChange) {
                    Set-TransportRule -Identity $($rule.Identity) -RouteMessageOutboundConnector $tempObcName @psBoundParameters
                }
                #endregion set outbound rule to new connector

                #region 5 - rename old connectors to backup Names
                Write-Verbose "5 - Rename existing SEPPmail.cloud Connectors to $backupName"
                Write-Verbose "5a - Rename existing SEPPmail.cloud Inbound Connector"
                $oldIbc = Get-InboundConnector -Identity '[SEPPmail.cloud] Inbound-*' 

                Set-InboundConnector -Identity $($OldIbc.Identity) -Name ($($OldIbc.Identity) -replace 'SEPPmail.Cloud',$backupName) @PSBoundParameters

                Write-Verbose "5b - Rename existing SEPPmail.cloud Outbound Connector"
                $oldObc = Get-OutBoundConnector -Identity "[SEPPmail.cloud] OutBound-*"
                Set-OutBoundConnector -Identity $($oldObc.Identity) -Name ($oldObc.Identity -replace 'SEPPmail.Cloud',$backupName) @PSBoundParameters

                #endregion

                #region 6 - attach old outbound rules to old connector
                Write-Verbose "6 - Set outbound rule to old backup connector again"
                $bkpConnWildcard = "[" + $backupName + "]*"
                $bkpObc = Get-OutboundConnector -Identity $bkpConnWildcard
                foreach ($rule in $rulesToChange) {
                    Set-TransportRule -Identity $($rule.Identity) -RouteMessageOutboundConnector $bkpObc @PSBoundParameters
                }

                #endregion

                #region 7 - rename new connectors to final Names
                Write-Verbose "7 - Rename existing $tempPrefix connectors to final name"
                $finalObCName = ($newConnectors|Where-Object Identity -like '*OutBound*').Identity -replace "^$([regex]::Escape($tempPrefix))", ""

                Set-OutboundConnector -Identity ($newConnectors|Where-Object Identity -like '*OutBound*').Identity -Name $finalObcName @PSBoundParameters

                $finalIbcName = ($newConnectors|Where-Object Identity -like '*Inbound*').Identity -replace "^$([regex]::Escape($tempPrefix))", ""
                Set-InBoundConnector -Identity ($newConnectors|Where-Object Identity -like '*InBound*').Identity -Name $finalIbcName @PSBoundParameters

                #endregion

                #region 8 - create New Transport rules
                Write-Verbose "8 - Creating new Transport Rules" 
                New-SC365Rules -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -routing $DeplInfo.routing  -PlacementPriority Top @PSBoundParameters

                #endregion

                #region 9 - disable old Transport rules
                $trWildcard = '[' + $BackupName + ']*'
                Write-Verbose "9 - Disable old Transport Rules"
                if ($PSCmdlet.ShouldProcess("Disabling Transport Rules matching $trWildcard")) {
                    Get-TransportRule -Identity $trWildcard | Disable-TransportRule -Confirm:$false @psboundParameters
                } 
                #endregion 9

                #region 10 - Disable old connectors
                Write-Verbose "10 - Disable old connectors" 
                Set-InBoundConnector -Identity $bkpConnWildcard -Enabled:$false @psBoundParameters
                Set-OutBoundConnector -Identity $bkpConnWildcard -Enabled:$false @PSBoundParameters

                #endregion 10

                #region 11 Remove old stuff
                if ($remove) { 
                    Write-Verbose "11a - Deleting old Transport Rules"
                    Get-TransportRule -Identity $trWildcard | Remove-TransportRule -confirm:$false
                    Write-Verbose "11b - Deleting old Inbound Connector"
                    Remove-InBoundConnector -Identity $bkpConnWildcard -confirm:$false
                    Write-Verbose "11c - Deleting old Outbound Connector"
                    Remove-OutBoundConnector -Identity $bkpConnWildcard -confirm:$false
                }
                #endregion
            }
            else {
                Write-Error "STOPPING - Found Existing Backup Objects - clean up the environment from $BackupName objects (rules and connectors) and TRY again"
                break
            }
        } else {
            Write-Host "Wise decision! Analyze your integration with New-SC365ExoReport and come back again if you are more familiar with the environment." -ForegroundColor Green -BackgroundColor DarkGray
            }
        }
        end {
            $ErrorActionPreference = $existEAValue
        }
    }

Register-ArgumentCompleter -CommandName New-SC365Setup -ParameterName SEPPmailCloudDomain -ScriptBlock $paramDomSB


# SIG # Begin signature block
# MIIVzAYJKoZIhvcNAQcCoIIVvTCCFbkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUzHh4lye5vmEK
# idpADD7kV2w1C6NRUHwYOk3Jcz2JX6CCEggwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBjP0RNwieyq1WaI1mvAAXV6MMY
# xzvAp3fgaCRNVVQ3SzANBgkqhkiG9w0BAQEFAASCAgAFvOWrXmZbv6knBuhsYlzo
# IIAHVaOjFw4e8nWS/eQTPLqZfdxdI+E28DQIj5qBoR8HF+xSgBsg7u4DVtdRpNS+
# 7mo8HkoPhwgK3O8xlImk1Yl2RfMqje5EPOnyHjMbx2dai1dNa4ibXfQwPeIRTjt3
# Y4qEwl/+CZlzsiv5IBBPS2uNw+Km7WY3faQW8FHisIKjYGgE+szxH4GqhVjump6y
# VgpYCufiT6ciHS01tzuYWLMIFVlvVShNCHDtT0UoNpjLBr0XZsKZ0k0kRdPzPq5Y
# B8xn52i5y3JS3nGAvX2K/pl6Znb3dfSL/4QPhBCIqIxK8ZRK2TeGFBM85/ErOPZh
# DGkNVWZsNWDNk8TVVs8sgppNRIaIk24a3F+A++L+Asic5FAHh/3MxUUPYwY9y9bp
# 9chA/fOnsb7wo1goXU3BVo+NmnGLSinfkxvNLi8iHNoY/EkAN2URsW8IvtFPyQcR
# 8Gn+aeZl3QyjGX7IALY0UljN72EhJMvTdk3JcgjR4CtYINs24hhoQBj0uzVJS/yA
# vPe5wK3NjcV8LmFF6Lxy6Wcvzc2t7QlIay9a8WbRXwIMMvCYdtdbIfS9qu4moXIf
# KgV26/VBuBf89Qg4DES75vQiAkyKt8S8VkjIhg/UDrdbB544bN5SPRHsQmSmIH4G
# CeVjFZUj0OjeAWMnP1VQjg==
# SIG # End signature block
