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
.EXAMPLE
    Get-SC365Connectors -routing -inline -InboundOnly
    Shows Connectors in inline/inboundOnly Mode
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function Get-SC365Connectors
{
    [CmdletBinding(
        SupportsShouldProcess = $false,
        ConfirmImpact = 'Medium',
        DefaultParameterSetName = 'parallel',
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
     )]

    param
    (
        [Parameter(
            Mandatory = $true,
            HelpMessage = '`"inline`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft'
            )]
        [ValidateSet('parallel','inline','p','i')]
        [String] $routing,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'inline',
            HelpMessage = 'For routing type `"inline`", if only inbound service is used.'
        )]
        [switch]$inBoundOnly
    )

    begin {       
        if (!(Test-SC365ConnectionStatus)) { 
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
        }
        else {
            Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`" " 
        }
        if ($routing -eq 'p') { $routing = 'parallel' }
        if ($routing -eq 'i') { $routing = 'inline' }
    }
    process {
        $inbound = Get-SC365InboundConnectorSettings -Routing $routing
        $ibc = Get-InboundConnector $inbound.Name -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        if ($ibc | Where-Object Identity -eq $($inbound.Name))
        {
            $ibc | Add-Member -MemberType NoteProperty -Name "SC365ModuleVersion" -Value (Get-SC365moduleVersion $($ibc.comment))
            $ibc | Add-Member -MemberType NoteProperty -Name "Region" -Value (($ibc.TlsSenderCertificateName.Split('.')[1]))
            $ibc.PSObject.TypeNames.Insert(0, "SEPPmail.cloud.Connectors")
            $ibc
        }
        else 
        {
            Write-Warning "No SEPPmail.cloud Inbound Connector with Name `"$($inbound.Name)`" found - Wrong routing mode ?"
        }
        if ($inBoundOnly -eq $false) {
            $outbound = Get-SC365OutboundConnectorSettings -Routing $routing
            $obc = Get-OutboundConnector $outbound.Name -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            if ($obc | Where-Object Identity -eq $($outbound.Name)){
                $obc | Add-Member -MemberType NoteProperty -Name 'SC365ModuleVersion' -Value (Get-SC365moduleVersion $($obc.comment))
                $obc | Add-Member -MemberType NoteProperty -Name 'Region' -Value ($obc.TlsDomain.Split('.')[1])
                $obc.PSObject.TypeNames.Insert(0, "SEPPmail.cloud.Connectors")
                $obc
            }
            else {
                Write-Warning "No SEPPmail.cloud Outbound Connector with name `"$($outbound.Name)`" found - Wrong routing mode or inbound only ?"
            }
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
         DefaultParameterSetName = 'BothDirections',
         HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
     )]
    param
    (
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Bothdirections',
            HelpMessage = 'Default E-Mail domain of your Exchange Online tenant.'
            )]
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Default E-Mail domain of your Exchange Online tenant.'
            )]
        [Alias('domain','maildomain','primaryMailDomain')]
        [String] $SEPPmailCloudDomain,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'BothDirections',
            HelpMessage = 'Geographical region of the seppmail.cloud service, either "ch" or "de"'
        )]
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Geographical region of the seppmail.cloud service, either "ch" or "de"'
        )]
        [ValidateSet('ch','prv','de','dev')]
        [String]$region,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'BothDirections',
            HelpMessage = '`"seppmailcloud`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft'
            )]
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = '`"inline`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft'
            )]
        [ValidateSet('parallel','inline','p','i')]
        [String] $routing,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'For routingtype `"inline`", if only inbound service is used.'
        )]
        [switch]$inBoundOnly,

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
            ParameterSetName = 'BothDirections',
            HelpMessage = 'Force overwrite of existing connectors and ignore hybrid setup'
        )]
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Force overwrite of existing connectors and ignore hybrid setup'
        )]
        [switch]$force,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'BothDirections',
            HelpMessage = 'Use custom name instead of name from cloud config'
        )]
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'InBoundOnly',
            HelpMessage = 'Use custom name instead of name from cloud config'
        )]
        [String]$NamePrefix
    )

    begin
    {        
        if(!(Test-SC365ConnectionStatus)) {
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
        } else {
            Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        }
 
        if ($routing -eq 'p') {$routing = 'parallel'}
		if ($routing -eq 'i') {$routing = 'inline'}
        
        # Add Warning for INLINE Connectors
        if ($routing -eq 'inline') {
            Write-Warning "You are about to create a SEPPmail INLINE Setup which will affect ALL Domains of your Microsoft Tenant."
        }

        if ($SEPPmailCloudDomain) {
            if ((Confirm-SC365TenantDefaultDomain -ValidationDomain $SEPPmailCloudDomain) -eq $true) {
                Write-verbose "Domain is part of the tenant and the Default Domain"
            } else {
                if ((Confirm-SC365TenantDefaultDomain -ValidationDomain $SEPPmailCloudDomain) -eq $false) {
                    Write-verbose "Domain is part of the tenant"
                } else {
                    Write-Error "Domain is NOT Part of the tenant"
                    break
                }
            }
            Write-Verbose "Adding SEPPmail Support-Addresses to allowedSenders in HostedContentFilterPolicy"
            [string[]]$existingAllowedSenders = (Get-HostedContentFilterPolicy -Identity 'Default'|select-Object AllowedSenders).Allowedsenders.Sender|Select-Object -ExpandProperty Address
            [string[]]$SEPPmailAllowedSenders = @('support@seppmail.de','support@seppmail.com','support@seppmail.ch','servicedesk@seppmail.com')
            $allowedSenders = ($existingAllowedSenders + $SEPPmailAllowedSenders |Select-Object -Unique)

            Set-HostedContentFilterPolicy -Identity "Default" -AllowedSenders $allowedSenders
        } else {
            $SEPPmailCloudDomain = $tenantAcceptedDomains|Where-Object {$_.Default -eq $true}
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
        $regionConfig = (Get-SC365CloudConfig -Region $region)
        $SEPPmailIPv4Range = $regionConfig.IPv4AllowList
        
        if ($SEPPmailCloudDomain) {
            $TenantID = Get-SC365Tenantid -Maildomain $SEPPmailCloudDomain
        } else {
            $TenantID = Get-SC365Tenantid -maildomain ((Get-OrganizationConfig).Identity)
        }
        
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
        $existingSC365Rules = Get-TransportRule |Where-Object {($_.RouteMessageOutboundConnector -like '*SEPPmail.cloud*') -and ($_.Name -like '*SEPPmail.cloud*')} -ea 0
        #endregion collecting existing connectors and test for hybrid Setup

        Write-Debug "Look for ARC-Signature for SEPPmail.cloud and add if missing"
        try {
            [string]$ath = (Get-ArcConfig).ArcTrustedSealers
            if ($ath) {
                if ((!($ath.Contains('SEPPmail.cloud')))) {
                    Write-Verbose "ARC header `'$ath`' exists, but does not contain SEPPmail.cloud"
                    $athNew = $ath.Split(',') + 'SEPPmail.cloud'
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
            if (($existingSMOutboundConn) -and ((!$NamePrefix)))
            {
                if ($existingSC365Rules.Count -ge 1) {

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
                    Write-Error "There are still transport-rules pointing to the outbound Connector $($existingSMOutboundConn.Identity). `nRules are: $existingSc365Rules. `nUse `"Remove-SC365rules -routing parallel/inline`" to remove and restart New-SC365Connectors."
                }
            }
            else
            { Write-Verbose "No existing Outbound Connector found" }

            if($createOutbound -and (!($inBoundOnly)))
            {
                Write-Debug "Creating SEPPmail.cloud Outbound Connector $($param.Name)!"
                if ($PSCmdLet.ShouldProcess($($param.Name), 'Creating Outbound Connector'))
                {

                    $param.Comment += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                    if ($NamePrefix) {
                        $Param.Name = "$NamePrefix" + "$($Param.Name)"
                    }                    
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
            if (($existingSMInboundConn) -and ((!$NamePrefix)))
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
                
                #region Create Inbound Connector
                Write-Verbose "Creating SEPPmail.cloud Inbound Connector $($param.Name)!"
                if ($PSCmdLet.ShouldProcess($($param.Name), 'Creating Inbound Connector'))
                {

                    $param.Comment += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                    Write-Verbose "Check if parameter `$name was set and overwrite"
                    if ($namePrefix) {
                        $Param.Name = "$NamePrefix" + "$($Param.Name)"
                    }                    
                    New-InboundConnector @param |Select-Object Identity,Enabled,WhenCreated,@{Name = 'Region'; Expression = {($_.TlsSenderCertificateName.Split('.')[1])}}

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
    [CmdletBinding(
        SupportsShouldProcess = $true,
        ConfirmImpact = 'Medium',
        DefaultparameterSetname = 'parallel',
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
     )]

    param
    (
        [Parameter(
            Mandatory = $true,
            ParameterSetname = 'parallel',
            Helpmessage = '`"seppmailcloud`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft'
            )]
        [Parameter(
            Mandatory = $true,
            ParameterSetname = 'inline',
            Helpmessage = '`"inline`": mx points to SEPPmail.cloud, `"parallel`": mx points to Microsoft'
            )]
        [ValidateSet('parallel','inline','p','i')]
        [String] $routing,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'inline',
            HelpMessage = 'For routingtype `"inline`", if only inbound service is used.'
        )]
        [switch]$inBoundOnly,
           
        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'parallel',
            Helpmessage = '`Sets AntiSpam IP AllowList'
            )]
        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'inline',
            Helpmessage = '`Sets AntiSpam IP AllowList'
            )]
        [ValidateSet('AntiSpamAllowListing')]
        [String]$option = $null
    )

    begin {
        if (!(Test-SC365ConnectionStatus)){ 
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" 
        } else {
            Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`" " 
        }
        if ($routing -eq 'p') {$routing = 'parallel'}
		if ($routing -eq 'i') {$routing = 'inline'}
    }
    process {
        $inbound = Get-SC365InboundConnectorSettings -routing $routing 
        if (!($inboundOnly)) {$outbound = Get-SC365OutboundConnectorSettings -routing $routing}
        $hcfp = Get-HostedConnectionFilterPolicy
    
        if (!($InBoundOnly)) {
            if ($PSCmdlet.ShouldProcess($outbound.Name, "Remove SEPPmail outbound connector $($Outbound.Name)"))
            {
                if (Get-OutboundConnector -WarningAction SilentlyContinue | Where-Object {$_.Name -eq $($outbound.Name)})
                {
                    Remove-OutboundConnector $outbound.Name -confirm:$false
                }
                else {
                    Write-Warning 'No SEPPmail Outbound Connector found'
                }
            }    
        }
    
        if($PSCmdlet.ShouldProcess($inbound.Name, "Remove SEPPmail inbound connector $($inbound.Name)")) {
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
                #}
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
# MIIVzAYJKoZIhvcNAQcCoIIVvTCCFbkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPbIALDMLWJH7A
# FXUnepiPg/n6eYUspPX/dppI9SAVbqCCEggwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBJY8e6u4nbhWbNtSbjSeBqvG15
# iYnG4xYuE54FMKurQjANBgkqhkiG9w0BAQEFAASCAgAAZNRYWWzTsl3yvLcsvpuv
# Z0GKiBPawJ7dd5GALsZttY+QnoNeAGLi4mRmp9Jibm0lHfUOvAmJWs7bONpM/2qi
# h6QEhtLNsjtW/GbopN1Iqm09H9G5xau9M+9FpfvgakmjLKdJCTmLq1xZSWGrQ6Zh
# sfB+/Vf/r/wtPMckm1hT6236pWSg6uhMGEIRFBvKgUeQcteip72Dz6wTkTb0j24h
# ZZH6ZrdJlhgY67Y6ZsWFNvqknfvJptOiAba1g8l8kNvC2TnHHuZpR06KOIGS7aQg
# reNRZX8Xt85SGtvG4JMszR+p/gchWMlsEKq1pfo+/sLdLRF2BV6Dwi9Ba4QWBANG
# MDU5+rZRRL5kxlugH4UcIl9InB0QdKCQUwNboTCD7e3LurjgAeRtbthVCCTzPyvL
# QcEF1fDVOzonKLNruooQpK0T0P3IdpU9wvxQZIfP9+6gNbzn1GfxEpiEVlUDLtbJ
# BGlMHjEeiJBwTYGVeCRc7u5yqo0lXE+clJFP/aQ0bjJtZ3VMlXUUX+Xo6D65LYfD
# X9/r8DcrhWG1QUpAe4v8rGwSkHejL3v6VbcpaCR760iCqtLDkdHGooUothMxHjfT
# tYtw6IVl8HhuKAW8ZMjMiNMfNRVtl7nAfYNKrha1QNLedR9rz6Cm4+DSwT/EoC+E
# 8VWJyxeNTQ8z7/+Y387SUQ==
# SIG # End signature block
