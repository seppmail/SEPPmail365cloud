<#
.SYNOPSIS
    Read existing SEPPmail.cloud Exchange Online connectors
.DESCRIPTION
    SEPPmail.cloud uses 2 Connectors to transfer messages between SEPPmail.cloud and Exchange Online
    This commandlet will show existing connectors.
.EXAMPLE
    Get-SC365Connectors
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
        [ValidateSet('inline','parallel','microsoft','seppmail')]
        $routing
    )

    if (!(Test-SC365ConnectionStatus))
    { 
        throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
    }
    else {
        Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

        #rename pre-1.0.0 routing-modes
        if ($routing -eq 'seppmail') {
            $routing = 'inline'
        } else {
            $routing = 'parallel'
        }

        $inbound = Get-SC365InboundConnectorSettings -Routing $routing
        $outbound = Get-SC365OutboundConnectorSettings -Routing $routing
        $obc = Get-OutboundConnector $outbound.Name -WarningAction SilentlyContinue
        $ibc = Get-InboundConnector $inbound.Name

        if ($obc | Where-Object Identity -eq $($outbound.Name))
        {
            $obc
        }
        else {
            Write-Warning "No SEPPmail.cloud Outbound Connector with name `"$($outbound.Name)`" found"
        }
        if ($ibc | Where-Object Identity -eq $($inbound.Name))
        {
            $ibc

        }
        else 
        {
            Write-Warning "No SEPPmail.cloud Inbound Connector with Name `"$($inbound.Name)`" found"
        }
    }
}

<#
.SYNOPSIS
    Adds SEPPmail.cloud Exchange Online connectors
.DESCRIPTION
    SEPPmail.cloud uses 2 Connectors to transfer messages between SEPPmail.cloud and Exchange Online
    This commandlet will create the connectors for you, depending on the routing mode.
.EXAMPLE
    PS C:\> New-SC365Connectors -maildomain 'contoso.eu' -region 'ch' -routing 'inline'
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Switzerland and customers uses seppmail.cloud mailfilter. MX points to seppmail.cloud
.EXAMPLE
    PS C:\> New-SC365Connectors -maildomain 'contoso.eu' -region 'ch' -routing 'inline' -disabled
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Switzerland and customers uses seppmail.cloud mailfilter. MX points to seppmail.cloud.
    Connectors will be created in "disabled"-mode. You need to enable them manually.
.EXAMPLE
    PS C:\> New-SC365Connectors -maildomain 'contoso.eu' -region 'ch' -routing 'inline' -Confirm:$false -Force
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Switzerland and customers uses seppmail.cloud mailfilter. MX points to seppmail.cloud.
    Connectors will be created and existing connectors will be deleted without any further interaction.
.EXAMPLE
    PS C:\> New-SC365Connectors -maildomain 'contoso.eu' -routing 'parallel' -region 'de'
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Germany and customers uses Microsoft mailfilter. MX points to Microsoft.
.EXAMPLE
    PS C:\> New-SC365Connectors -maildomain 'contoso.eu' -routing 'parallel' -region 'de' -noInboundEFSkipIPs
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Germany and customers uses Microsoft mailfilter. MX points to Microsoft.
    In addition the IP-Addresses of SEPPmail.cloud are not listed in the "Enhanced Filter Skip list". This will impact SPAM of detection of MS Defender, USE WITH CARE!
.EXAMPLE
    PS C:\> New-SC365Connectors -maildomain 'contoso.eu' -routing 'parallel' -region 'de' -option NoAntiSpamAllowListing
    Creates Connectors for the maildomain contoso.eu, seppmail.cloud environment ist Germany and customers uses Microsoft mailfilter. MX points to Microsoft.
    In addition the IP-addresses of SEPPmail.cloud are not listed in the Default Hosted Connection Filter Policy. This will impact SPAM of detection of MS Defender, USE WITH CARE!
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
        [Alias('domain','maidomain')]
        [String] $primaryMailDomain,

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
        [ValidateSet('inline','parallel','seppmail','microsoft')]
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
            Helpmessage = 'Does not set IP-addresses of sending SEPPmail.cloud servers in EFSkipIPs-value in inbound connector.'
            )]
        [Parameter(
            Mandatory = $false,
            ParameterSetname = 'InBoundOnly',
            Helpmessage = 'Does not set IP-addresses of sending SEPPmail.cloud servers in EFSkipIPs in inbound connector'
            )]
        [switch]$noInboundEFSkipIPs = $false,

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
        [ValidateSet('NoAntiSpamAllowListing')]
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
 
        #region Preparing common setup
        Write-Verbose "Preparing values for Cloud configuration"

        #rename pre-1.0.0 routing-modes
        if ($routing -eq 'seppmail') {
            $routing = 'inline'
        } else {
            $routing = 'parallel'
        }


        Write-Verbose "Prepare smarthosts for e-Mail domain $primaryMailDomain"
        if ($routing -eq 'inline') {
            $OutboundSmartHost = ($primaryMailDomain.Replace('.','-')) + '.relay.seppmail.cloud'
        }
        if ($routing -eq 'parallel') {
            $OutboundSmartHost = ($primaryMailDomain.Replace('.','-')) + '.mail.seppmail.cloud'
        }
        
        Write-Verbose "Prepare GeoRegion configuration for region: $region"
        $CloudConfig = Get-Content "$PSScriptRoot\..\ExOConfig\CloudConfig\GeoRegion.json" -raw|Convertfrom-Json -AsHashtable
        $regionConfig = $cloudConfig.GeoRegion.($region.Tolower())
        $SEPPmailIPv4Range = $regionConfig.IPv4AllowList
        $TlsCertificateName = $regionConfig.TlsCertificate

        Write-Verbose "Set timestamp and Moduleversion for Comments"
        $Now = Get-Date
        $moduleVersion = $myInvocation.MyCommand.Version
        #endregion commonsetup

        #region collecting existing connectors and test for hybrid Setup
        Write-Verbose "Collecting existing connectors"
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
                    catch {}
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
        #endregion
    }

    process
    {
        #region - OutboundConnector
        Write-Verbose "Building Outbound parameters based on smarthost $outboundtlsdomain"
        #$outbound = Get-SC365OutboundConnectorSettings -routing $routing -Option $Option
        $param = Get-SC365OutboundConnectorSettings -Routing $routing -Option $option
        $param.SmartHosts = $OutboundSmartHost            
        $param.TlsDomain = $TlsCertificateName
 
        Write-verbose "if -disabled switch is used, the connector stays deactivated"
        if ($Disabled) {
            $param.Enabled = $false
        }

        Write-Verbose "Read existing SEPPmail.cloud outbound connector"
        $existingSMOutboundConn = $allOutboundConnectors | Where-Object Name -eq $param.Name
        # only $false if the user says so interactively
        
        [bool]$createOutBound = $true #Set Default Value
        #wait-debugger
        if ($existingSMOutboundConn)
        {
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
                        $existingSMOutboundConn | Remove-OutboundConnector -Confirm:$false # user confirmation action

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
                    Write-Verbose "Removing existing OutBound Connector $existingSMOutboundConn.Identity dur to -force parameter"
                    $existingSMOutboundConn | Remove-OutboundConnector -Confirm:$false # due to -force 
                }
            }
        }
        else
        { Write-Verbose "No existing Outbound Connector found" }

        if($createOutbound -and (!($inboundonly)))
        {
            Write-Verbose "Creating SEPPmail.cloud Outbound Connector $($param.Name)!"
            if ($PSCmdLet.ShouldProcess($($param.Name), 'Creating Outbound Connector'))
            {

                $param.Comment += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                $nobc = New-OutboundConnector @param
                $SC365ConnectorsHash = [ordered]@{
                    OBName                           = $nobc.Identity
                    OBEnabled                        = $nobc.Enabled
                    OBTlsDomain                      = $nobc.TlsDomain
                    OBTlsSettings                    = $nobc.TlsSettings
                    OBSmartHosts                     = $nobc.SmartHosts
                    OBOriginatingServer              = $nobc.OriginatingServer
                    OBOrganizationalUnitRootInternal = $nobc.OrganizationalUnitRootInternal
                }
                if(!$?)
                {throw $error[0]}
            }
        }
        #endregion - OutboundConnector

        #region - InboundConnector
        Write-Verbose "Read Inbound Connector Settings"
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
            $param.RestrictDomainsToIPAddresses = $false
            $param.RestrictDomainsToCertificate = $true
            $param.SenderIPAddresses = $SEPPmailIPv4Range
            $param.TlsSenderCertificateName = $TlsCertificateName

            #region EFSkipIP in inbound connector
            if ($NoInboundEFSkipIPs) {
                Write-Warning "Inbound Connector $param.Name will be build WITHOUT IP-addresses in EFSkipIPs. This will increase SPAM false-positives."
            } else {
                [String[]]$EfSkipIPArray = $cloudConfig.GeoRegion.($region.Tolower()).IPv4AllowList + $cloudConfig.GeoRegion.($region.Tolower()).IPv6AllowList
                $param.EFSkipIPs = $EfSkipIPArray
            }

            Write-Verbose "Creating SEPPmail.cloud Inbound Connector $($param.Name)!"
            if ($PSCmdLet.ShouldProcess($($param.Name), 'Creating Inbound Connector'))
            {

                $param.Comment += "`nCreated with SEPPmail365cloud PowerShell Module version $moduleVersion on $now"
                #[void](New-InboundConnector @param)
                $nibc = New-InboundConnector @param
                $SC365ConnectorsHash += [ordered]@{
                    IBName                           = $nibc.Identity
                    IBEnabled                        = $nibc.Enabled
                    IBTLSCertificate                 = $nibc.TlsSenderCertificateName
                    IBSenderIPAddresses              = $nibc.SenderIPAddresses
                    IBEFSkipIPs                      = $nibc.EFSkipIPs
                    IBOriginatingServer              = $nibc.OriginatingServer
                    IBOrganizationalUnitRootInternal = $nibc.OrganizationalUnitRootInternal
                }
                if(!$?) {
                    throw $error[0]
                } else {
                    #region - Add Region-based IP-range to hosted Connection Filter Policy AllowList
                    if (!($option -eq 'NoAntiSpamAllowListing'))
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
        }
        #endRegion - InboundConnector
    }

    end
    {
        $SC365Connectors = New-Object -TypeName PSobject -property $SC365ConnectorsHash
        $SC365Connectors
    }
}

<#
.SYNOPSIS
    Removes the SEPPmail inbound and outbound connectors
.DESCRIPTION
    Convenience function to remove the SEPPmail connectors
.EXAMPLE
    PS C:\> Remove-SC365Connectors
    Removes all SEPPmail Connectors from the exchange online environment.
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
        [ValidateSet('inline','parallel')]
        [String]$routing,
        
        [ValidateSet('NoAntiSpamAllowListing')]
        [String]$option
    )

    if (!(Test-SC365ConnectionStatus))
    { throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" }

    Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

    Write-Verbose "re-write pre-1.0.0 routing-modes of parametervalue $routing"
    if ($routing -eq 'seppmail') {
        $routing = 'inline'
    } else {
        $routing = 'parallel'
    }

    $inbound = Get-SC365InboundConnectorSettings -routing $routing 
    $outbound = Get-SC365OutboundConnectorSettings -routing $routing
    $hcfp = Get-HostedConnectionFilterPolicy

    if($PSCmdlet.ShouldProcess($outbound.Name, "Remove SEPPmail outbound connector $($Outbound.Name)"))
    {
        if (Get-OutboundConnector -WarningAction SilentlyContinue | Where-Object Identity -eq $($outbound.Name))
        {
            Remove-OutboundConnector $outbound.Name -confirm:$false
        }
        else {
            Write-Warning 'No SEPPmail Outbound Connector found'
        }
    }

    if($PSCmdlet.ShouldProcess($inbound.Name, "Remove SEPPmail inbound connector $($inbound.Name)"))
    {
        $InboundConnector = Get-InboundConnector | Where-Object Identity -eq $($inbound.Name)
        if ($inboundConnector)
            {
            Write-Verbose 'Collect Inbound Connector IP for later AllowListremoval'
            
            [string]$InboundSEPPmailIP = $null
            if ($inboundConnector.TlsSenderCertificateName) {
                [array]$InboundSEPPmailIP = $inboundConnector.SenderIPAddresses -split ' '
            }
            Remove-InboundConnector $inbound.Name -confirm:$false

            Write-Verbose "If Inbound Connector has been removed, remove also AllowListed IPs"
            if ((!($Option -like 'NoAntiSpamAllowListing')) -and (!(Get-InboundConnector | Where-Object Identity -eq $($inbound.Name))))
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
            Write-Warning 'No SEPPmail.cloud Inbound Connector found'
        }
    }
}

<#
.SYNOPSIS
    Backs up all existing connectors to individual json files
.DESCRIPTION
    Convenience function to perform a backup of all existing connectors
.EXAMPLE
    Backup-SC365Connectors -OutFolder "C:\temp"
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function Backup-SC365Connectors
{
    [CmdletBinding()]
    param
    (
        [Parameter(
             Mandatory = $true,
             HelpMessage = 'Folder in which the backed up configuration will be stored'
         )]
        [Alias('Folder','Path')]
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

        Get-InboundConnector | foreach-object{
            $n = $_.Name
            $n = $n -replace "[\[\]*\\/?:><`"]"

            $p = "$OutFolder\inbound_connector_$n.json"

            Write-Verbose "Backing up $($_.Name) to $p"
            ConvertTo-Json -InputObject $_ | Out-File $p
        }

        Get-OutboundConnector -WarningAction SilentlyContinue | foreach-object {
            $n = $_.Name
            $n = $n -replace "[\[\]*\\/?:><`"]"

            $p = "$OutFolder\outbound_connector_$n.json"
            Write-Verbose "Backing up $($_.Name) to $p"
            ConvertTo-Json -InputObject $_ | Out-File $p
        }
    }
}

if (!(Get-Alias 'Set-SC365Connectors' -ErrorAction SilentlyContinue)) {
    New-Alias -Name Set-SC365Connectors -Value New-SC365Connectors
}

Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName MailDomain -ScriptBlock $paramDomSB
#Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName region -ScriptBlock $paramRegionSB
#Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName routing -ScriptBlock $paramRoutingModeSB


# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbz63IuK2oez87ohmHZvBpRPi
# XQmggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# MRYEFHCxhDfQ9BGmG+ORBmBFLuKCuK7yMA0GCSqGSIb3DQEBAQUABIIBADpXa11c
# hm+eeXhB7LQDRgQukK1r5PWu+9OKIPXohLEvoav9qAWlRU+SXLp6wVMzPXqNygJr
# nsZmfkmgXOLM0aZ80DtgLO8F26TOsGaJW9TEQS8ChZ9pEc5t8nVhH2o7At3SgUvS
# D+RcB6bGwo/amTEEbzc6tP365CMbEr/u4gGFRwi5gqSDgJCWF3VgXMl+m77SxWzW
# dhwgrK2QHbha+OE09ZqN6zpuxE3ct5qiPT7o/d/7gqIgaM0AXWVpEHeNY6yM+WrX
# nYD07DXrIDxb9TAjY1u2vXZncEhSBfUwaLZMYXLAR2pqK90SEZOMC4awaW7iKsAS
# lLFybvhGvbaqZUA=
# SIG # End signature block
