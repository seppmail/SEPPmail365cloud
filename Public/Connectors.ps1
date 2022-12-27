<#
.SYNOPSIS
    Read existing SEPPmail.cloud Exchange Online connectors
.DESCRIPTION
    SEPPmail.cloud uses 2 Connectors to transfer messages between SEPPmail.cloud and Exchange Online
    This commandlet will show existing connectors.
.EXAMPLE
    Get-SC365Connectors -routing -parallel
    Shows Connectors in parallel mode
.EXAMPLE
    Get-SC365Connectors -routing -inline
    Shows Connectors in inline mode
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function Get-SC365Connectors
{
    [CmdletBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    Param
    (
        [Parameter(
            Mandatory = $true
        )]
        [ValidateSet('parallel','inline','p','i')]
        $routing
    )

    if (!(Test-SC365ConnectionStatus))
    { 
        throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
    }
    else {
        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

        $inbound = Get-SC365InboundConnectorSettings -Routing $routing
        $outbound = Get-SC365OutboundConnectorSettings -Routing $routing
        $obc = Get-OutboundConnector $outbound.Name -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $ibc = Get-InboundConnector $inbound.Name -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

        if ($obc | Where-Object Identity -eq $($outbound.Name))
        {
            $obc|select-object Name,Enabled,WhenCreated,@{Name = 'Region'; Expression = {($_.TlsDomain.Split('.')[1])}}
        }
        else {
            Write-Warning "No SEPPmail.cloud Outbound Connector with name `"$($outbound.Name)`" found. Wrong routing mode or Inboundonly ?"
        }
        if ($ibc | Where-Object Identity -eq $($inbound.Name))
        {
            $ibc|select-object Name,Enabled,WhenCreated,@{Name = 'Region'; Expression = {($_.TlsSenderCertificateName.Split('.')[1])}}
        }
        else 
        {
            Write-Warning "No SEPPmail.cloud Inbound Connector with Name `"$($inbound.Name)`" found - Wrong routing mode ?"
        }

    }
}

<#
.SYNOPSIS
    Adds/Renews SEPPmail.cloud Exchange Online connectors
.DESCRIPTION
    SEPPmail.cloud uses 2 Connectors to transfer messages between SEPPmail.cloud and Exchange Online
    This commandlet will create the connectors for you, depending on routing mode and the region.  
.EXAMPLE
    PS C:\> New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -region 'ch' -routing 'inline'
    Creates Connectors for the maildomain contoso.eu. seppmail.cloud environment ist Switzerland and customers uses seppmail.cloud mailfilter. MX points to seppmail.cloud
.EXAMPLE
    PS C:\> New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -region 'ch' -routing 'inline' -disabled
    Creates Connectors for the maildomain contoso.eu, the seppmail.cloud environment ist Switzerland and customers uses seppmail.cloud mailfilter. MX points to seppmail.cloud.
    Connectors will be created in "disabled"-mode. You need to enable them manually.
.EXAMPLE
    PS C:\> New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -region 'ch' -routing 'inline' -Confirm:$false -Force
    Creates Connectors for the maildomain contoso.eu, the seppmail.cloud environment ist Switzerland and customers uses seppmail.cloud mailfilter. MX points to seppmail.cloud.
    Connectors will be created and existing connectors will be deleted without any further interaction.
.EXAMPLE
    PS C:\> New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'parallel' -region 'de'
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Germany and customers uses Microsoft mailfilter. MX points to Microsoft.
.EXAMPLE
    PS C:\> New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'parallel' -region 'de' -InboundEFSkipIPs
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Germany and customers uses Microsoft mailfilter. MX points to Microsoft.
    In addition the IP-Addresses of SEPPmail.cloud are listed in the "Enhanced Filter Skip list". This should not be neeed with Version 1.2.0+ as we do ARC-signing!
.EXAMPLE
    PS C:\> New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'parallel' -region 'de' -option AntiSpamAllowListing
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Germany and customers uses Microsoft mailfilter. MX points to Microsoft.
    In addition the IP-addresses of SEPPmail.cloud are listed in the Default Hosted Connection Filter Policy. This will impact SPAM of detection of MS Defender, USE WITH CARE!
.INPUTS
    
.OUTPUTS
    Inbound and OutboundConnectors
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function New-SC365Connectors
{
    [CmdletBinding(
         SupportsShouldProcess = $true,
         ConfirmImpact = 'Medium',
         DefaultparameterSetname = 'BothDirections',
         HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
     )]

    param
    (
        [Parameter(
            Mandatory = $true,
            ParameterSetname = 'Bothdirections',
            Helpmessage = 'Default E-Mail domain of your Exchange Online tenant.',
            Position = 0
            )]
        [Parameter(
            Mandatory = $true,
            ParameterSetname = 'InBoundOnly',
            Helpmessage = 'Default E-Mail domain of your Exchange Online tenant.',
            Position = 0
            )]
        [Alias('domain','maildomain','primaryMailDomain')]
        [String] $SEPPmailCloudDomain,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'BothDirections',
            HelpMessage = 'Geographcal region of the seppmail.cloud service',
            Position = 1
        )]
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Geographcal region of the seppmail.cloud service',
            Position = 1
        )]
        [ValidateSet('ch','prv','de')]
        [String]$region,

        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'BothDirections',
            Helpmessage = '`"seppmailcloud`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft',
            Position = 2
            )]
        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'InBoundOnly',
            Helpmessage = '`"inline`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft',
            Position = 2
            )]
        [ValidateSet('parallel','inline','p','i')]
        [String] $routing,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'For routingtype `"inline`", if only inbound service is used.'
        )]
        [switch]$inBoundOnly = $false,

        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'BothDirections',
            Helpmessage = 'Force set IP-addresses of sending SEPPmail.cloud servers in EFSkipIPs in inbound connector.'
            )]
        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'InBoundOnly',
            Helpmessage = 'Force set IP-addresses of sending SEPPmail.cloud servers in EFSkipIPs in inbound connector'
            )]
        [switch]$InboundEFSkipIPs = $false,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'BothDirections',
            HelpMessage = 'Which configuration option to use'
        )]
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Which configuration option to use'
        )]
        [ValidateSet('AntiSpamAllowListing')]
        [String[]]$option,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'BothDirections',
            HelpMessage = 'Disable the connectors on creation'
        )]
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Disable the connectors on creation'
        )]
        [switch]$disabled,

        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'BothDirections',
            HelpMessage = 'Force overwrite of existing connectors and ignore hybrid setup'
        )]
        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'InBoundOnly',
            HelpMessage = 'Force overwrite of existing connectors and ignore hybrid setup'
        )]
        [switch]$force
    )

    begin
    {
        if(!(Test-SC365ConnectionStatus))
        {throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"}
        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
 
        $TenantDomains = Get-AcceptedDomain
        If (!($TenantDomains.DomainName -contains $SEPPmailCloudDomain)) {
           $PrimaryDomain = $TenantDomain|Where-Object 'Default' -eq $true|Select-Object -ExpandProperty DomainName
           Write-Information "Typo ? Domain should be $PrimaryDomain" 
           Write-Error "$SEPPmailCloudDomain is not member of the connected tenant. Retry using only tenant-domains"
           break
        }

        #region Preparing common setup
        Write-Debug "Preparing values for Cloud configuration"

        Write-debug "Prepare smarthosts for e-Mail domain $SEPPmailCloudDomain"
        if ($routing -eq 'inline') {
            $OutboundSmartHost = ($SEPPmailCloudDomain.Replace('.','-')) + '.relay.seppmail.cloud'
        }
        if ($routing -eq 'parallel') {
            $OutboundSmartHost = ($SEPPmailCloudDomain.Replace('.','-')) + '.mail.seppmail.cloud'
        }
        Write-Verbose "Outbound SmartHost is: $OutboundSmartHost"

        Write-Debug "Prepare GeoRegion configuration for region: $region"
        $CloudConfig = Get-Content "$PSScriptRoot\..\ExOConfig\CloudConfig\GeoRegion.json" -raw|Convertfrom-Json -AsHashtable
        $regionConfig = $cloudConfig.GeoRegion.($region.Tolower())
        $SEPPmailIPv4Range = $regionConfig.IPv4AllowList
        $TenantID = Get-SC365Tenantid -Maildomain $SEPPmailCloudDomain
        $TlsCertificateName = $regionConfig.TlsCertificate
        Write-Verbose "TLS Certificate is $TlsCertificateName"


        $TenantIdCertificateName = $tenantId + ($regionConfig.TlsCertificate).Replace('*','')
        Write-verbose "Tenant certificate Name is $TenantIdCertificateName"

        Write-Debug "Set timestamp and Moduleversion for Comments"
        $Now = Get-Date
        $moduleVersion = $myInvocation.MyCommand.Version
        #endregion commonsetup

        #region collecting existing connectors and test for hybrid Setup
        Write-Debug "Collecting existing connectors"
        $allInboundConnectors = Get-InboundConnector
        $allOutboundConnectors = Get-OutboundConnector -WarningAction SilentlyContinue

        Write-Verbose "Testing for hybrid Setup"
        $HybridInboundConn = $allInboundConnectors |Where-Object {(($_.Name -clike 'Inbound from *') -or ($_.ConnectorSource -clike 'HybridWizard'))}
        $HybridOutBoundConn = $allOutboundConnectors |Where-Object {(($_.Name -clike 'Outbound to *') -or ($_.ConnectorSource -clike 'HybridWizard'))}
        
        if (($HybridInboundConn -or $HybridOutBoundConn) -and (!($force)))
        {
            Write-Warning "!!! - Hybrid Configuration detected - we assume you know what you are doing. Be sure to backup your connector settings before making any change."
            if($InteractiveSession)
            {
                Write-Verbose "Ask user to continue if Hybrid is found."
                Do {
                    try {
                        [ValidateSet('y', 'Y', 'n', 'N')]$hybridContinue = Read-Host -Prompt "Create SEPPmail connectors in hybrid environment ? (Y/N)"
                    }
                    catch {
                    }
                }
                until ($?)
                if ($hybridContinue -eq 'n') {
                    Write-Verbose "Exiting due to user decision."
                    break
                }
            }
            else
            {
                # should we error out here, since connector creation might be dangerous?
            }
        } else {
            Write-Verbose "No Hybrid Connectors detected, seems to be a clean cloud-only environment" -InformationAction Continue
        }

        Write-Debug "Checking for existing SEPPmail.cloud rules"
        $existingSc365Rules = Get-TransportRule -Identity '[SEPPmail.Cloud]*' -WarningAction SilentlyContinue -erroraction SilentlyContinue

        #endregion collecting existing connectors and test for hybrid Setup

        Write-Debug "Look for ARC-Signature for seppmail.cloud and add if missing"
        try {
            [string]$ath = (Get-ArcConfig).ArctrustedSealers
            if ($ath) {
                if ((!($ath.Contains('SEPPmail.cloud')))) {
                    Write-Verbose "ARC header `'$ath`' exists, but does not contain SEPPmail.cloud"
                    $athnew = $ath.Split(',') + 'SEPPmail.cloud'
                    Set-ArcConfig -Identity default -ArcTrustedSealers $athNew|Out-Null
                }
            }
            if (!($ath)) {
                Write-Verbose "ARC header is empty, adding SEPPmail.cloud"
                Set-ArcConfig -Identity default -ArcTrustedSealers 'SEPPmail.cloud'|Out-Null
            }
            if (($ath.Contains('SEPPmail.cloud'))) {
                Write-Verbose "ARC header `'$ath`' exists and contains SEPPmail.cloud, no action required"
            }
        } catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
        }
    }

    process
    {
            #region - OutboundConnector
            Write-Debug "Building Outbound parameters based on smarthost $outboundtlsdomain"
            #$outbound = Get-SC365OutboundConnectorSettings -routing $routing -Option $Option
            $param = Get-SC365OutboundConnectorSettings -Routing $routing -Option $option
            $param.SmartHosts = $OutboundSmartHost
            Write-Verbose "Outbound Connector Smarthosts are $($param.SmartHosts)"
            $param.TlsDomain = $TlsCertificateName
            Write-Verbose "Outbound Connector TLS certificate name is $($param.TlsDomain)"
            
            Write-verbose "if -disabled switch is used, the connector stays deactivated"
            if ($Disabled) {
                $param.Enabled = $false
            }

            Write-Debug "Read existing SEPPmail.cloud outbound connector"
            $existingSMOutboundConn = $allOutboundConnectors | Where-Object Name -eq $param.Name
            # only $false if the user says so interactively

            [bool]$createOutBound = $true #Set Default Value
            #wait-debugger
            if ($existingSMOutboundConn)
            {
                if (!($existingsc365rules)) {

                    Write-Warning "Found existing SEPPmail.cloud outbound connector with name: `"$($existingSMOutboundConn.Name)`" created on `"$($existingSMOutboundConn.WhenCreated)`" pointing to SEPPmail `"$($existingSMOutboundConn.TlsDomain)`" "
                    if (($InteractiveSession) -and (!($force)))
                    {
                        [string] $tmp = $null

                        Do {
                            try {
                                [ValidateSet('y', 'Y', 'n', 'N')]$tmp = Read-Host -Prompt "Shall we delete and recreate the outbound connector (will only work if no rules use it)? (Y/N)"
                                break
                            }
                            catch {}
                        }
                        until ($?)

                        if ($tmp -eq 'y') {
                            $createOutbound = $true

                            Write-Verbose "Removing existing Outbound Connector $($existingSMOutboundConn.Name) !"
                            if ($PSCmdLet.ShouldProcess($($existingSMOutboundConn.Name), 'Removing existing SEPPmail.cloud Outbound Connector')) {
                                Remove-OutboundConnector -Identity $($ExistingSMOutboundConn.Identity) -Confirm:$false # user confirmation action

                                if (!$?)
                                { throw $error[0] }
                            }
                        }
                        else {
                            Write-Warning "Leaving existing SEPPmail outbound connector `"$($existingSMOutboundConn.Name)`" untouched."
                            $createOutbound = $false
                        }
                    }
                    else
                    {
                        if (!($force)) {
                            throw [System.Exception] "Outbound connector $($outbound.Name) already exists"
                        }
                        else {
                            Write-Verbose "Removing existing OutBound Connector $existingSMOutboundConn.Identity due to -force parameter"
                            try {
                                $existingSMOutboundConn | Remove-OutboundConnector -Confirm:$false # due to -force 
                            }
                            catch {
                                # Errormessage
                                $error[0]
                            }
                        }
                    }
                }
                else {
                    Write-Error "There are still transport-rules pointing to the outbound Connector $($existingSMOutboundConn.Identity) Use Remove-SC365rules -routing <routingmode>"
                }
            }
            else
            { Write-Verbose "No existing Outbound Connector found" }

            if($createOutbound -and (!($inboundonly)))
            {
                Write-Debug "Creating SEPPmail.cloud Outbound Connector $($param.Name)!"
                if ($PSCmdLet.ShouldProcess($($param.Name), 'Creating Outbound Connector'))
                {

                    $param.Comment += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                    New-OutboundConnector @param | Select-Object Identity,Enabled,WhenCreated,@{Name = 'Region'; Expression = {($_.TlsDomain.Split('.')[1])}}

                    if(!$?)
                    {throw $error[0]}
                }
            }
            #endregion - OutboundConnector

            #region - InboundConnector
            Write-Debug "Read Inbound Connector Settings"
            $param = $null
            $param = Get-SC365InboundConnectorSettings -routing $routing -option $option
        
            Write-verbose "if -disabled switch is used, the connector stays deactivated"
            if ($disabled) {
                $param.Enabled = $false
            }

            Write-Verbose "Read existing SEPPmail Inbound Connector from Exchange Online"
            $existingSMInboundConn = $allInboundConnectors | Where-Object Name -EQ $param.Name

            # only $false if the user says so interactively
            [bool]$createInbound = $true
            if ($existingSMInboundConn)
            {
                Write-Warning "Found existing SEPPmail.cloud inbound connector with name: `"$($existingSMInboundConn.Name)`", created `"$($existingSMInboundConn.WhenCreated)`" incoming SEPPmail is `"$($existingSMInboundConn.TlsSenderCertificateName)`""

                if (($InteractiveSession) -and (!($force)))
                {
                    [string] $tmp = $null
                    Do {
                        try {
                            [ValidateSet('y', 'Y', 'n', 'N')]$tmp = Read-Host -Prompt "Shall we delete and recreate the inbound connector (will only work if no rules use it)? (Y/N)"
                            break
                        }
                        catch {}
                    }
                    until ($?)

                    if ($tmp -eq 'y') {
                        $createInbound = $true

                        Write-Verbose "Removing existing SEPPmail.cloud Inbound Connector $($existingSMInboundConn.Name) !"
                        if ($PSCmdLet.ShouldProcess($($existingSMInboundConn.Name), 'Removing existing SEPPmail.cloud inbound Connector')) {
                            $existingSMInboundConn | Remove-InboundConnector -Confirm:$false # user already confirmed action

                            if (!$?)
                            { throw $error[0] }
                        }
                    }
                    else {
                        Write-Warning "Leaving existing SEPPmail.cloud Inbound Connector `"$($existingSMInboundConn.Name)`" untouched."
                        $createInbound = $false
                    }
                }
                else
                {
                    if (!($force)) {
                        throw [System.Exception] "Inbound connector $($param.Name) already exists"
                    }
                    else {
                        $existingSMInboundConn | Remove-InboundConnector -Confirm:$false # due to -force
                    }
                
                }
            }
            else
            {Write-Verbose "No existing Inbound Connector found"}

            if($createInbound)
            {
                Write-Verbose "Setting $TlscertificateName as TLSSendercertificate and IP addresses to region $region"
                if ($routing -eq 'parallel') {$param.RestrictDomainsToIPAddresses = $false}
                #60 if ($routing -eq 'parallel') {$param.SenderIPAddresses = $SEPPmailIPv4Range}
                if ($routing -eq 'inline') {$param.SenderDomains = 'smtp:' + '*' + ';1'} # smtp:*;1
                
                # Inline and parallel use same certificate
                $param.TlsSenderCertificateName = $TenantIdCertificateName
                Write-verbose "Inbound TlsSenderCertificateName is: $param.TLSSenderCertificatename"
                
                #region EFSkipIP in inbound connector
                if ($InboundEFSkipIPs){
                    [String[]]$EfSkipIPArray = $cloudConfig.GeoRegion.($region.Tolower()).IPv4AllowList + $cloudConfig.GeoRegion.($region.Tolower()).IPv6AllowList
                    $param.EFSkipIPs = $EfSkipIPArray
                } else {
                    Write-verbose "Inbound Connector $param.Name will be build WITHOUT IP-addresses in EFSkipIPs."
                }
                #endregion EFSkip In ibc

                #region Create Inbound Connector
                Write-Verbose "Creating SEPPmail.cloud Inbound Connector $($param.Name)!"
                if ($PSCmdLet.ShouldProcess($($param.Name), 'Creating Inbound Connector'))
                {

                    $param.Comment += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                    #[void](New-InboundConnector @param)
                    New-InboundConnector @param |Select-Object Identity,Enabled,WhenCreated,@{Name = 'Region'; Expression = $region}

                    if(!$?) {
                        throw $error[0]
                    } else {
                        #region - Add Region-based IP-range to hosted Connection Filter Policy AllowList
                        if ($option -eq 'AntiSpamAllowListing')
                        {
                            Write-Verbose "Adding SEPPmail.cloud to AllowList in 'Hosted Connection Filter Policy'"
                            Write-Verbose "Collecting existing AllowList"
                            $hcfp = Get-HostedConnectionFilterPolicy
                            [string[]]$existingAllowList = $hcfp.IPAllowList
                            Write-verbose "Adding SEPPmail.cloud IP ranges to HostedConnectionFilterPolicy $($hcfp.Id)"
                            if ($existingAllowList) {
                                $FinalIPList = ($existingAllowList + $SEPPmailIPv4Range)|sort-object -Unique
                            }
                            else {
                                $FinalIPList = $SEPPmailIPv4Range
                            }
                            Write-verbose "Adding IPaddress list with content $finalIPList to hosted connection filter policy $($hcfp.Id)"
                            if ($FinalIPList) {
                                [void](Set-HostedConnectionFilterPolicy -Identity $hcfp.Id -IPAllowList $finalIPList)
                            }
                        }
                        #endRegion - Hosted Connection Filter Policy AllowList
                    }
                }
                #endregion Create Inbound Connector
            }
            #endRegion - InboundConnector
    }
    end
    {

    }
}

<#
.SYNOPSIS
    Removes the SEPPmail inbound and outbound connectors
.DESCRIPTION
    Convenience function to remove the SEPPmail connectors
.EXAMPLE
    PS C:\> Remove-SC365Connectors -routing parallel
    Removes all SEPPmail Connectors from the exchange online environment in parallel mode.
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function Remove-SC365Connectors
{
    [CmdletBinding(SupportsShouldProcess=$true,
                   ConfirmImpact='Medium',
                   HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
                   )]
    Param
    (
        [Parameter(
            Mandatory = $true,
            Helpmessage = 'The routing tyoe of the connector to you want to remove'
        )]
        [ValidateSet('parallel','inline','p','i')]
        [String]$routing,
        
        [ValidateSet('AntiSpamAllowListing')]
        [String]$option
    )

    begin {
        if (!(Test-SC365ConnectionStatus))
        { throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" }
    
        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        
        if ($routing -eq 'p') {$routing = 'parallel'}
		if ($routing -eq 'i') {$routing = 'inline'}

    }
    process {
        $inbound = Get-SC365InboundConnectorSettings -routing $routing 
        $outbound = Get-SC365OutboundConnectorSettings -routing $routing
        $hcfp = Get-HostedConnectionFilterPolicy
    
        if($PSCmdlet.ShouldProcess($outbound.Name, "Remove SEPPmail outbound connector $($Outbound.Name)"))
        {
            if (Get-OutboundConnector -WarningAction SilentlyContinue | Where-Object {$_.Name -eq $($outbound.Name)})
            {
                Remove-OutboundConnector $outbound.Name -confirm:$false
            }
            else {
                Write-Warning 'No SEPPmail Outbound Connector found'
            }
        }
    
        if($PSCmdlet.ShouldProcess($inbound.Name, "Remove SEPPmail inbound connector $($inbound.Name)"))
        {
            $InboundConnector = Get-InboundConnector | Where-Object {$_.Name -eq $($inbound.Name)}
            if ($inboundConnector)
                {
                Write-Verbose 'Collect Inbound Connector IP for later AllowListremoval'
                
                [string]$InboundSEPPmailIP = $null
                if ($inboundConnector.TlsSenderCertificateName) {
                    [array]$InboundSEPPmailIP = $inboundConnector.SenderIPAddresses -split ' '
                }
                Remove-InboundConnector $inbound.Name -confirm:$false
    
                Write-Verbose "If Inbound Connector has been removed, remove also AllowListed IPs"
                if ((!($Option -like 'AntiSpamAllowListing')) -and (!(Get-InboundConnector | Where-Object Identity -eq $($inbound.Name))))
                {
                        Write-Verbose "Remove SEPPmail Appliance IP from AllowList in 'Hosted Connection Filter Policy'"
                        
                        Write-Verbose "Collecting existing AllowList"
                        [System.Collections.ArrayList]$existingAllowList = $hcfp.IPAllowList
                        Write-verbose "Removing SEPPmail Appliance IP $InboundSEPPmailIP from Policy $($hcfp.Id)"
                        if ($existingAllowList) {
                            foreach ($IP in $InboundSEPPmailIP) {
                                $existingAllowList.Remove($IP)
                            }
                            Set-HostedConnectionFilterPolicy -Identity $hcfp.Id -IPAllowList $existingAllowList
                            Write-Verbose "IP: $InboundSEPPmailIP removed from Hosted Connection Filter Policy $hcfp.Id"
                    }
                }
            }
            else 
            {
                Write-Warning 'No SEPPmail.Cloud Inbound Connector found'
            }
    
    }
    }
}



if (!(Get-Alias 'Set-SC365Connectors' -ErrorAction SilentlyContinue)) {
    New-Alias -Name Set-SC365Connectors -Value New-SC365Connectors
}

Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName SEPPmailCloudDomain -ScriptBlock $paramDomSB


# SIG # Begin signature block
# MIIL/AYJKoZIhvcNAQcCoIIL7TCCC+kCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4ePLrT4Rs2lFQ
# e7MIF+6qGjzxLwD57tWGXN6JvnIIhKCCCUAwggSZMIIDgaADAgECAhBxoLc2ld2x
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA335CNvig5ILnbAgkrF72uSqDQ
# mYtpXe820AaM3c4ZrzANBgkqhkiG9w0BAQEFAASCAQAkFhitY7UwQdbKDMgzvsx5
# AMDNDRwO0e3W3PDENjSl6k8DYqlF9minEeo+XjZcSKvP19STDtU9Pz8htcNa2ERu
# P6WkM3IY0UzmLlP5iVvwUY6FYxsOet28xrwQ3daaMc13cyvEzOF+KRbqSlJHc5iJ
# mZGCudjbjLjoriekmKCyaGUm9QJVZ7f+PmatdEPhHezhBPNdxPuz2vTvI1w3z60m
# eX8McwsqGfLKEIQuY+45EBlI7HVjeAXALU2Hf6J9WBllZgzQaBbfWoO+4ollm3P/
# 7v+VA1yZFvecZNvvY9uUnNrfHSJs4vcaURB9kHE8+OhV/+R1e9s2lquNY8pntudq
# SIG # End signature block
