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
        [ValidateSet('inline','parallel')]
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
        [ValidateSet('inline','parallel')]
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
        $TlsCertificateName = $regionConfig.TlsCertificate
        Write-Verbose "TLS Certificate is $TlsCertificateName"

        Write-Verbose 'Crafting Inbound Certificate Name'
        $ibTlsCertificateName = $SEPPmailCloudDomain.Split('.')[0] + '-' + $SEPPmailCloudDomain.Split('.')[1] + '.transport.seppmail.cloud'
        Write-verbose "IBC certificate Name is $ibTlsCertificateName"

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

        #endregion

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

            #region ibc
            if($createInbound)
            {
                Write-Verbose "Setting $TlscertificateName as TLSSendercertificate and IP addresses to region $region"
                if ($routing -eq 'parallel') {$param.RestrictDomainsToIPAddresses = $false}
                #60 if ($routing -eq 'parallel') {$param.SenderIPAddresses = $SEPPmailIPv4Range}
                if ($routing -eq 'inline') {$param.SenderDomains = 'smtp:' + '*' + ';1'} # smtp:*;1
                
                if ($routing -eq 'inline') 
                    {
                        $param.TlsSenderCertificateName = $IbTlsCertificateName
                    }
                else {
                        $param.TlsSenderCertificateName = $IbTlsCertificateName
                }

               #region EFSkipIP in inbound connector
                if ($InboundEFSkipIPs){
                    [String[]]$EfSkipIPArray = $cloudConfig.GeoRegion.($region.Tolower()).IPv4AllowList + $cloudConfig.GeoRegion.($region.Tolower()).IPv6AllowList
                    $param.EFSkipIPs = $EfSkipIPArray
                } else {
                    Write-verbose "Inbound Connector $param.Name will be build WITHOUT IP-addresses in EFSkipIPs."
                }
                #endregion EFSkip In ibc
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
            }
            #endRegion - InboundConnector
        #}

    }

    end
    {
        #$SC365Connectors = New-Object -TypeName PSobject -property $SC365ConnectorsHash
        #$SC365Connectors
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
        [ValidateSet('inline','parallel')]
        [String]$routing,
        
        [ValidateSet('AntiSpamAllowListing')]
        [String]$option
    )

    if (!(Test-SC365ConnectionStatus))
    { throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" }

    Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue

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

Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName SEPPmailCloudDomain -ScriptBlock $paramDomSB
Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName region -ScriptBlock $paramRegionSB
Register-ArgumentCompleter -CommandName New-SC365Connectors -ParameterName routing -ScriptBlock $paramRoutingModeSB


# SIG # Begin signature block
# MIIL1wYJKoZIhvcNAQcCoIILyDCCC8QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUpVZdK4urfWbWNl/tGZgcnGuo
# LCeggglAMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# MRYEFN2Yeh7ylCfDHlcwHsdTX/4fswbsMA0GCSqGSIb3DQEBAQUABIIBAHb8PzIj
# +Jm5m26AmkV1Jk3ahfGKr7zmQ0Eq3ZGCNhCr1GgpiHFWS4ZB9E3eIXTDvwxnfnED
# jVjvgQUwLAJ/gAYKSI/h68VRkfugRDwhGk/f6AXLNL+FR0BSK7JDAU/+dbmZ7t6K
# 0cA0dOnT8drPo+lha1vXSx1xgzb3+tAhuUOd7lkUWTwutFWEkpptTut9F1a1mF2a
# vtKm4hKrweRpcN8CXrpwbyRKUW7vWugd5uJM44Y0PD5ous7sNdRU80Sus9b5+f6C
# N5ce4PfRT08EFrGj78bvi9mgiRPONbFm1NAMrJ72Dz5uFgZ3TIjkqon62FKT8Cx0
# hhKQ9UN4fdAc6zw=
# SIG # End signature block
