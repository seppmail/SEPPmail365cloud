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
    Generates a report of the current Status of the Exchange Online environment
.DESCRIPTION
    The report will write all needed information of Exchange Online into an HTML file. This is useful for documentation and decisions for the integration. It also makes sense as some sort of snapshot documentation before and after an integration into seppmail.cloud
.EXAMPLE
    PS C:\> New-SC365ExoReport
    This reads relevant information of Exchange Online and writes a summary report in an HTML in the current directory
.EXAMPLE
    PS C:\> New-SC365ExoReport -FilePath '~/Desktop'
    -Filepath requires a relative path and may be used with or without filename (auto-generated filename)
.EXAMPLE
    PS C:\> New-SC365ExoReport -LiteralPath c:\temp\expreport.html
    Literalpath requires a full and valid path
.EXAMPLE
    PS C:\> New-SC365ExoReport -jsonBackup -FilePath '~/Desktop'
    JSONBackup writes a JSON file with all relevant configuration of the Exchange Online Tenanant in addition to the HTML report.
.INPUTS
    FilePath
.OUTPUTS
    HTML Report and JSON backup file
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function New-SC365ExOReport {
    [CmdletBinding(
        SupportsShouldProcess = $true,
                ConfirmImpact = 'Medium',
     DefaultParameterSetName  = 'FilePath',
                      HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
        )]
    Param (
        # Define output relative Filepath
        [Parameter(   
           Mandatory   = $false,
           HelpMessage = 'Relative path of the HTML report on disk',
           ParameterSetName = 'FilePath',
           Position = 0
        )]
        [Alias('Path')]
        [string]$filePath = '.',

        [Parameter(   
           Mandatory   = $false,
           HelpMessage = 'Literal path of the HTML report on disk',
           ParameterSetName = 'LiteralPath',
           Position = 0
        )]
        [string]$Literalpath = '.',

        [Parameter(   
           Mandatory   = $false,
           HelpMessage = 'Literal path of the HTML report on disk',
           ParameterSetName = 'LiteralPath',
           Position = 1
        )]
        [Parameter(   
            Mandatory   = $false,
            HelpMessage = 'Literal path of the HTML report on disk',
            ParameterSetName = 'FilePath',
            Position = 1
         )]
         [switch]$jsonBackup
 
    )

    begin
    {
        if (!(Test-SC365ConnectionStatus)){
            throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet" }
        else {
            Write-Information "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        }

        #region Filetest only if not $Literalpath is selected
        if ($PsCmdlet.ParameterSetName -eq "FilePath") {

            If (Test-Path $FilePath -PathType Container) {
                Write-Verbose "Filepath $Filepath is a directory"
                
                if (Test-Path (Split-Path (Resolve-Path $Filepath) -Parent)) {
                    Write-Verbose "Filepath $Filepath Container exists on disk, creating default ReportFilename"
                    $ReportFilename = New-SelfGeneratedReportName
                    $FinalPath = Join-Path -Path $filePath -ChildPath $ReportFilename
                } else {
                    throw [System.Exception] "$FilePath is not valid. Enter a valid filepath like ~\Desktop or c:\temp\expreport.html"
                }

                } else {
                    Write-Verbose "FilePath $Filepath is a Full Path including Filename"
                    if ((Split-Path $FilePath -Extension) -eq '.html') {
                        $FinalPath = $Filepath
                    } else {
                        throw [System.Exception] "$FilePath is not an HTML file. Enter a valid filepath like ~\Desktop or c:\temp\expreport.html"
                    }
                }
        }
        else {
        # Literalpath
            $SplitLiteralPath = Split-Path -Path $LiteralPath -Parent
            If (Test-Path -Path $SplitLiteralPath) {
                $finalPath = $LiteralPath
            } else {
                throw [System.Exception] "$LiteralPath does not exist. Enter a valid literal path like ~\exoreport.html or c:\temp\expreport.html"
            }
        }
        if ($jsonBackup) {
            $jsonPath = (Join-Path -Path (split-path $FinalPath -Parent) -ChildPath (split-path $FinalPath -leafbase)) + '.json'
        }
        #endregion

        #Region Design parameters
        $colorSEPPmailGreen = '#C7D400'
        $colorSEPPmailGrey = '#575757'
        $colorSEPPmailLightGrey = '#BEBEBE'

        $sectionStyle = @{
            Direction = 'column' 
            Margin = 2
            HeaderText = 'Exchange Online Status Report'
            HeaderBackGroundColor = $ColorSEPPmailgreen
            HeaderTextColor = $ColorSEPPmailGrey
            HeaderTextSize = 20
            BackgroundColor = $colorSEPPmailLightGrey
            BorderRadius = '5px'
        }
        $contentHeaderStyle = @{
            HeaderTextAlignment = 'center'
            HeaderTextColor = $ColorSEPPmailGreen 
            HeaderBackGroundColor = $colorSEPPmailGrey
            HeaderTextSize = 18
        }

        $contentBodyStyle = @{
            Margin = 7
            BorderRadius = '5px'
            JustifyContent = 'center'
            CanCollaps = $true
            BackgroundColor = 'White'
        }
        $tableStyle = @{
            Style = 'display' # 'cell-border', compact, display, hover, nowrap, order-column, row-border, stripe
            Buttons = 'copyHtml5','csvHtml5','excelHtml5','pdfHtml5','print'
            DisablePaging = $false
            DisableSearch = $false
            DisableOrdering = $false
            DisableResponsiveTable = $false
            SearchBuilderLocation=  'bottom'
            EnableColumnReorder = $true
            EnableRowReorder = $false
            HideFooter = $true
            AutoSize = $false
            TextWhenNoData = 'No data in Exchange Online tenant available.'
        }
        $helpTextStyle = @{
            FontSize = 11
            Color = $colorSEPPmailLightGrey
        }
        #endregion

    }

    process
    {
        try {
            # Initialize JsonDATA to be filled by Get-ExoHTMLData for Backup Purposes
            $script:JsonData = $null
            if ($pscmdlet.ShouldProcess("Target", "Operation")) {
                #"Whatis is $Whatif and `$pscmdlet.ShouldProcess is $($pscmdlet.ShouldProcess) "
                #For later Use
            }

            #region Collecting Report Header CreateUser
            if ($PSVersionTable.OS -like 'Microsoft Windows*') {
                $repUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            } else {
                $repUser = (hostname) + '/' + (whoami)
            }
            #endregion

            #region NEW Way of Collecting Data
            $ExoData = [ordered]@{}
            $ExoData['AccDom']=[ordered]@{
                VarNam = 'AccDom'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/get-accepteddomain'
                RawCmd = 'Get-AcceptedDomain'
                TabDat = 'Domainname,DomainType,Default,EmailOnly,ExternallyManaged,OutboundOnly,WhenCreated,WhenChanged'
                HdgTxt = 'Accepted Domains'
                HlpInf = 'The list of configured E-Mail-domains in this Tenant. The Tenant-Default-Domain is listed first. If the onmicrosoft.com domain is default, its highlighted in red.'
            }
            $ExoData['RemDom']=[ordered]@{
                VarNam = 'RemDom'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/get-remotedomain'
                RawCmd = 'Get-RemoteDomain'
                TabDat = 'DomainName,ContentType,IsInternal,SmtpDaneMandatoryModeEnabled,WhenCreated,WhenChanged'
                HdgTxt = 'Remote Domains'
                HlpInf = 'Remote Domains are used to control mail flow with more precision, apply message formatting and messaging policies and specify acceptable character sets for messages sent to and received from the remote domain'
            }
            $ExoData['OrgCfg']=[ordered]@{
                VarNam = 'OrgCfg'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/get-OrganizationConfig'
                RawCmd = 'Get-OrganizationConfig'
                TabDat = 'DisplayName,ExchangeVersion,SendFromAliasEnabled,AllowedMailboxRegions,DefaultMailboxRegion,DisablePlusAddressInRecipients,WhenCreated,WhenChanged'
                HdgTxt = 'Organizational Config'
                HlpInf = 'Some data around the physical location of your M365-Tenant'
            }
            $ExoData['TspCfg']=[ordered]@{
                VarNam = 'TspCfg'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-TransportConfig'
                RawCmd = 'Get-TransportConfig'
                TabDat = 'MaxSendSize,MaxReceiveSize,WhenCreated,WhenChanged'
                HdgTxt = 'Transport Configuration'
                HlpInf = 'View organization-wide transport configuration settings'
            }
            $ExoData['MxrRep']=[ordered]@{
                VarNam = 'MxrRep'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-MXRecordReport'
                RawCmd = '$accDom.DomainName |ForEach-Object {Get-MXRecordReport -Domain $_}'
                TabDat = 'Domain,HighestPriorityMailhostIpAddress,HighestPriorityMailhost,IsAcceptedDomain,Organization,PointsToService,RecordExists'
                HdgTxt = 'MX-Record Report'
                HlpInf = 'MX-Record DNS entries and IP-addresses of every accepted domain'
            }
            $ExoData['ArcSlr']=[ordered]@{
                VarNam = 'ArcSlr'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-ArcConfig'
                RawCmd = 'Get-ArcConfig'
                TabDat = ''
                HdgTxt = 'Trusted ARC Sealer Configuration'
                HlpInf = 'ARC is used to run SEPPmail.cloud or the SEPPmail Appliance in parallel mode with Exchange Online'
            }
            $ExoData['DkmSig']=[ordered]@{
                VarNam = 'dkmsig'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-DkimSigningConfig'
                RawCmd = 'Get-DkimSigningConfig'
                TabDat = 'Domain,Enabled,Status,Selector1CNAME,Selector2CNAME,WhenCreated,WhenChanged'
                HdgTxt = 'DKIM Signing Configuration'
                HlpInf = 'DKIM Keys per Domain, DNS entries contains a public key used to verify the digital signature of an email. Makes only sense if MX record points to Microsoft.'
            }
            $ExoData['DanSts']=[ordered]@{
                VarNam = 'dansts'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-Arc'
                RawCmd = 'Get-SmtpDaneInboundStatus -DomainName (($accDom|Where-Object {$_.Default -eq $true}).DomainName)'
                TabDat = ''
                HdgTxt = 'DANE Inbound Status for Default Domain'
                HlpInf = 'A DANE record is a DNSSEC-protected TLSA record that specifies the expected TLS certificate or certificate authority information for securely connecting to a server.'
            }
            $ExoData['ibdCon']=[ordered]@{
                VarNam = 'ibdcon'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-InboundConnector'
                RawCmd = 'Get-InboundConnector'
                TabDat = 'Identity,Enabled,ConnectorType,SenderDomains,SenderIPAddresses,TlsSenderCertificateName,EFSkipLastIP,EFSkipIPs,Comment,WhenCreated,WhenChanged'
                HdgTxt = 'Inbound Connectors'
                HlpInf = 'Connectivity for E-Mails flowing Inbound to Exchange Online'
            }
            $ExoData['obdCon']=[ordered]@{
                VarNam = 'obdCon'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-OutboundConnector'
                RawCmd = 'Get-OutboundConnector -IncludeTestModeConnectors:$true'
                TabDat = 'Identity,Enabled,ConnectorType,SmartHosts,TlsDomain,TlsSettings,RecipientDomains,Comment,WhenCreated,WhenChanged'
                HdgTxt = 'Outbound Connectors'
                HlpInf = 'Connectivity for E-Mails flowing Outbound from Exchange Online'
            }
            $ExoData['malFlw']=[ordered]@{
                VarNam = 'malFlw'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-MailFlowStatusReport'
                RawCmd = 'Get-MailFlowStatusReport -StartDate (Get-Date) -EndDate ((Get-Date).AddDays(1))'
                TabDat = ''
                HdgTxt = 'Mail Flow Status Report'
                HlpInf = 'E-Mails categorized by by severity, of the last 24 hours'
            }
            $ExoData['tapRls']=[ordered]@{
                VarNam = 'tapRls'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-TransportRule'
                RawCmd = 'Get-TransportRule'
                TabDat = 'Name,State,Mode,Priority,FromScope,SentToScope,StopRuleProcessing,ManuallyModified,Comments,Description,WhenCreated,WhenChanged'
                HdgTxt = 'E-Mail Transport Rules'
                HlpInf = 'Transport rules control mail flow by conditions and are important for the SEPPmail integration.'
            }
            $ExoData['apsPol']=[ordered]@{
                VarNam = 'apsPol'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-AntiPhishPolicy'
                RawCmd = 'Get-AntiPhishPolicy'
                TabDat = 'Identity,isDefault,IsValid,AuthenticationFailAction,WhenCreated,WhenChanged'
                HdgTxt = 'Anti-Phishig Policies'
                HlpInf = 'Anti-Phish Policies are a security measure designed to protect against phishing attacks by identifying and blocking emails'
            }
            $ExoData['MwfPol']=[ordered]@{
                VarNam = 'MwfPol'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-MalwareFilterPolicy'
                RawCmd = 'Get-MalwareFilterPolicy'
                TabDat = 'Identity,Action,IsDefault,Filetypes,WhenCreated,WhenChanged'
                HdgTxt = 'Anti-Malware Policies'
                HlpInf = 'Anti-Malware Policies are a security configuration that scans and blocks email messages containing malicious software, such as viruses or ransomware.'
            }
            $ExoData['hctFpl']=[ordered]@{
                VarNam = 'hctFpl'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-HostedContentFilterPolicy'
                RawCmd = 'Get-HostedContentFilterPolicy'
                TabDat = 'Name,IsDefault,ObjectState,MarkAsSpamSpfRecordHardFail,QuarantineRetentionPeriod,EndUserSpamNotificationFrequency,TestModeAction,IsValid,BulkSpamAction,PhishSpamAction,OriginatingServer,WhenCreated,WhenChanged'
                HdgTxt = 'Hosted Content Filter Policies'
                HlpInf = 'The HostedContentFilterPolicy in Exchange Online is a configuration that determines the filtering actions and thresholds for email content, including spam detection, safe sender lists, and quarantining, to protect against unwanted or malicious emails'
            }
            $ExoData['hcnFpl']=[ordered]@{
                VarNam = 'hcnFpl'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-HostedConnectionFilterPolicy'
                RawCmd = 'Get-HostedConnectionFilterPolicy'
                TabDat = ''
                HdgTxt = 'Hosted Connection Filter Policies'
                HlpInf = 'The HostedConnectionFilterPolicy in Exchange Online is a configuration that controls the connection filtering settings for incoming email, such as blocking or allowing specific IP addresses and domains, to manage spam and phishing protection.'
            }
            $ExoData['blkSnd']=[ordered]@{
                VarNam = 'blkSnd'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-BlockedSenderAddress'
                RawCmd = 'Get-BlockedSenderAddress'
                TabDat = ''
                HdgTxt = 'Blocked Sender Address List'
                HlpInf = 'The BlockedSenderAddress list in Exchange Online specifies individual email addresses that are explicitly blocked from sending messages to recipients in your organization, helping to prevent spam or unwanted emails from those addresses.'
            }
            $ExoData['hobFpl']=[ordered]@{
                VarNam = 'hobFpl'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-HostedOutboundSpamFilterPolicy'
                RawCmd = 'Get-HostedOutboundSpamFilterPolicy'
                TabDat = 'Name,IsDefault,Enabled,ActionWhenThresholdReached,WhenCreated,WhenChanged'
                HdgTxt = 'Hosted Outbound SPAMfilter Policies'
                HlpInf = 'The HostedOutboundSpamFilterPolicy in Exchange Online controls the filtering and management of outbound emails to detect and block potential spam or malicious messages sent from your organization, protecting your domains reputation.'
            }
            $ExoData['qarPol']=[ordered]@{
                VarNam = 'qarPol'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-QuarantinePolicy'
                RawCmd = 'Get-QuarantinePolicy'
                TabDat = 'Name,IsValid,QuarantinePolicyType,QuarantineRetentionDays,EndUserQuarantinePermissions,ESNEnabled,WhenCreated,WhenChanged'
                HdgTxt = 'Quarantine Policies'
                HlpInf = 'A quarantine policy in Exchange Online defines how quarantined emails are handled, including permissions for users to view, release, or report messages, and specifies notification settings for administrators and end users.'
            }
            $ExoData['iorCon']=[ordered]@{
                VarNam = 'iorCon'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-IntraOrganizationConnector'
                RawCmd = 'Get-IntraOrganizationConnector'
                TabDat = ''
                HdgTxt = 'Intra Organization Connectors'
                HlpInf = 'Intra-Organization Connectors in Exchange Online enable seamless mail flow, free/busy calendar sharing, and other organizational data sharing between different Exchange Online organizations or between Exchange Online and on-premises Exchange environments in a hybrid setup.'
            }
            $ExoData['hybMdc']=[ordered]@{
                VarNam = 'hybMdc'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/Get-HybridMailflowDatacenterIPs'
                RawCmd = 'Get-HybridMailflowDatacenterIPs'
                TabDat = ''
                HdgTxt = 'Hybrid Mailflow Datacenter IPs'
                HlpInf = 'List of IP addresses used by Microsoft datacenters for managing hybrid mail flow in an Exchange hybrid deployment'
            }
            <#$ExoData['nnnNnn']=[ordered]@{
                VarNam = 'nnnNnn'
                WebLnk = 'https://learn.microsoft.com/en-us/powershell/module/exchange/nnn'
                RawCmd = ''
                TabDat = ''
                HdgTxt = ''
                HlpInf = ''
            }#>
            $totalItems = $ExoData.Count
            $j = 0
            foreach ($ExoDataKey in $ExoData.Keys) {
                $InfoData = $ExoData[$ExoDataKey]
                $j++
                # Execute the RawCmd and store the raw result in a variable with 'Raw' postfix
                $rawVariableName = "$($InfoData.VarNam)Raw"
                try {
                    Set-Variable -Name $rawVariableName -Value (Invoke-Expression $InfoData.RawCmd) -Scope Script
                } catch {
                    Set-Variable -Name $rawVariableName -Value "$($_.Exception.Message)" -Scope Script
                }

                # Execute the RawCmd and pipe it to Select-Object with TabDat members
                $processedVariableName = $InfoData.VarNam
                Write-Progress -Activity 'Receiving Exchange Online Information by:'`
                    -Status "Processing $($infoData.RawCmd)" `
                    -PercentComplete (($j / $totalItems) * 100)

                if ([string]::IsNullOrWhiteSpace($InfoData.TabDat)) {
                    # If TabDat (Select-Object of Data) is empty, use the raw variable value
                    Set-Variable -Name $processedVariableName -Value (Get-Variable -Name $rawVariableName -ValueOnly) -Scope Script
                } else {
                    # Otherwise, process RawCmd with Select-Object and TabDat properties
                    try {
                        Set-Variable -Name $processedVariableName -Value (Invoke-Expression "$($InfoData.RawCmd) | Select-Object -Property $($InfoData.TabDat)") -Scope Script
                    } catch {
                        Set-Variable -Name $rawVariableName -Value "$($_.Exception.Message)" -Scope Script
                    }
                }
            }
            Write-Progress -Activity 'Receiving Exchange Online Information' -Status "Completed" -Completed

            #endregion

            #region Generate the HTML report
            $finalReport = New-HTML -HtmlData {
                New-HTMLImage -Source 'https://downloads.seppmail.com/wp-content/uploads/logo_seppmail_V1_Screen_M.png'  -Width '20%'
                #New-HTMLLogo -LogoPath '/Users/romanstadlmair/Desktop/NewReport/' -LeftLogoName 'SEPPmailLogo.png' -LeftLogoString '/Users/romanstadlmair/Desktop/NewReport/SEPPmailLogo.png'
                New-HTMLSection @sectionStyle -Headertext "Exchange Online Status Report for: $($OrgCfg.DisplayName)" -Content {    
                    New-HTMLContent @contentHeaderStyle @ContentbodyStyle -HeaderText 'Report Information' -Direction 'column' -Collapsed -Content {
                        $RawData =[ordered]@{
                            'Report created' = (Get-Date)
                            'Report created by' = $repUser
                            'FileName' = Split-Path $FinalPath -Leaf
                            'FilePath' = Split-Path $FinalPath -Parent
                            'Fullpath' = $FinalPath
                            'SEPPmail365cloud Module Version' = $Global:ModuleVersion
                            'Microsoft Tenant ID' = Get-SC365TenantID -maildomain (Get-AcceptedDomain|where-object InitialDomain -eq $true|select-object -expandproperty Domainname)
                        }
                        if ($jsonBackup) {$RawData.'Link to JSON File on Disk' = $JsonPath}
                        $RawDataNoHeader = [PSCustomObject]$RawData
                        New-HTMLTable -DataTable $rawDataNoHeader @TableStyle -TextWhenNoData 'Could not fetch this value' -EnableRowReorder   
                    }
                    New-HTMLContent @ContentHeaderStyle @contentBodyStyle -HeaderText 'General Setup' -Direction 'column' -Content {
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.AccDom.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "$($ExoData.AccDom.HlpInf)"}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.AccDom.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        Write-Verbose "Add Logic to detect if the default accepted domain is the onmicrosoft.com domain"
                        if (!($accDom|Get-Member -Name onmicrosoftIsDefault)) {
                            $accDom|Add-Member -MemberType NoteProperty -Name onmicrosoftIsDefault -Value 'False'
                        }
                        Foreach ($domain in $accDom) {
                            if (($domain.Default -eq $True) -and ($domain.DomainName -like '*.onmicrosoft.com') ){
                                $domain.onmicrosoftIsDefault = 'True'
                            }
                        }
                        New-HTMLTable -DataTable $accDom @tablestyle -ExcludeProperty 'onmicrosoftIsDefault' -DefaultSortColumn 'Default' -DefaultSortOrder 'Descending' -SearchBuilder {
                            New-HTMLTableCondition -Name 'Default' -ComparisonType string -Operator eq -Value 'True' -Row -FontWeight bold
                            New-HTMLTableCondition -Name 'onmicrosoftIsDefault' -ComparisonType bool -Operator eq -Value $true -row -Color red
                        }
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.RemDom.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "$($ExoData.RemDom.HlpInf)"}                
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.RemDom.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $RemDom @tablestyle -ExcludeProperty 'onmicrosoftIsDefault' -DefaultSortColumn 'Default' -DefaultSortOrder 'Descending'
                    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.OrgCfg.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "$($ExoData.OrgCfg.HlpInf)"}                
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.OrgCfg.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $OrgCfg @tableStyle
                    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.TspCfg.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "$($ExoData.TspCfg.HlpInf)"}                
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.TspCfg.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $TspCfg @tableStyle 
                    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.MxrRep.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "$($ExoData.MxrRep.HlpInf)"}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.MxrRep.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $MxrRep @tableStyle -DefaultSortColumn 'HighestPriorityMailhostIpAddress'
                    }
                    New-HTMLContent @ContentHeaderStyle @contentBodyStyle -HeaderText 'SMTP Security' -Direction 'row'-Content {
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.ArcSlr.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.ArcSlr.HlpInf)} -FontStyle italic -LineBreak 
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.ArcSlr.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "--------------------------------------------------------------------------------------------------------------------------------------------------------"}
                        New-HTMLText -Text $arcSlr.ArcTrustedSealers -FontSize 14 -LineBreak

                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.dkmSig.HdgTxt  -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.dkmSig.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.dkmSig.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $dkmSig @tableStyle -SearchBuilder {
                            New-HTMLTableCondition -Name 'Enabled' -ComparisonType string -Operator eq -Value 'False' -Color 'red' -row -FontWeight bold 
                        }
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.danSts.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.DanSts.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.danSts.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "--------------------------------------------------------------------------------------------------------------------------------------------------------"}
                        New-HTMLText -Text $DanSts -FontSize 14 -LineBreak
                    }
                    New-HTMLContent @contentHeaderStyle @ContentBodyStyle -Headertext 'External Connectivity' -Direction 'column' -Content {
                        if ($ibdcon) {
                            New-HTMLHeading -Heading h2 -HeadingText $ExoData.ibdcon.HdgTxt -Color $ColorSEPPmailGreen -Underline
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.ibdcon.HlpInf)}
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.ibdcon.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                            Write-verbose "Add Logic to detect if the correct EFIP logic is set"
                            if (!($ibdcon|Get-Member -Name EfSkipConfig)) {$ibdcon|Add-Member -MemberType NoteProperty -Name EfSkipConfig -Value 'undefined'}
                            Foreach ($ib in $ibdcon) {
                            if (($ib.Identity -like '*SEPPmail.cloud*') -and ($ib.ConnectorType -like 'OnPremises')) {
                                if ((!($ib.EFSkipIPs)) -and ($ib.EFSkipLastIP -eq $true)){
                                    $ib.EfSkipConfig = 'parallel'
                                }
                                if ((!(!($ib.EFSkipIPs))) -and ($ib.EFSkipLastIP -eq $true)){
                                    $ib.EfSkipConfig = 'EFSkipIPs not empty'
                                }
                                if ((!($ib.EFSkipIPs)) -and ($ib.EFSkipLastIP -eq $false)){
                                    $ib.EfSkipConfig = 'EFSkipLastIP is $false'
                                }
                                if ((!(!($ib.EFSkipIPs))) -and ($ib.EFSkipLastIP -eq $false)){
                                    $ib.EfSkipConfig = 'EFSkipLastIP is $false AND EFSkipIPs not empty'
                                }    
                            }
                            }
                            Write-Verbose "Add SEPPmail.cloud PowerShell Module version number to SEPPmail Connectors if available"
                            $IbcVersion = Get-SC365ModuleVersion -InputString $ibdcon.Comment
                            $ibdcon|Add-Member -membertype NoteProperty -Name SC365Version -value $IbcVersion.Tostring()
                            Write-Verbose "Create the IBC Data Table"
                            New-HTMLTable -DataTable $ibdcon @tableStyle -SearchBuilder {
                            New-HTMLTableCondition -Name 'Identity' -ComparisonType string -Operator like -Value 'SEPPmail' -FontWeight bold -Color $colorSEPPmailGreen -Row 
                            New-HTMLTableCondition -Name 'Identity' -ComparisonType string -Operator like -Value 'CodeTwo' -BackgroundColor GoldenYellow -Row
                            New-HTMLTableCondition -Name 'Identity' -ComparisonType string -Operator like -Value 'Exclaimer' -BackgroundColor GoldenYellow -Row
                            New-HTMLTableCondition -Name 'EfSkipConfig' -ComparisonType string -Operator like -Value 'EFSkip' -row -Color red
                            }
                        }
                        if ($obdCon) {    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.obdCon.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.obdCon.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.obdCon.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        Write-Verbose "Add SEPPmail.cloud PowerShell Module version number to SEPPmail Connectors if available"
                        $obdVersion = Get-SC365ModuleVersion -InputString $obdCon.Comment
                        $obdCon|Add-Member -membertype NoteProperty -Name SC365Version -value $obdVersion

                        New-HTMLTable -DataTable $obdCon @tableStyle -SearchBuilder {
                            New-HTMLTableCondition -Name 'Identity' -ComparisonType string -Operator like -Value 'CodeTwo' -BackgroundColor GoldenYellow -Row
                            New-HTMLTableCondition -Name 'Identity' -ComparisonType string -Operator like -Value 'Exclaimer' -BackgroundColor GoldenYellow -Row
                            New-HTMLTableCondition -Name 'Identity' -ComparisonType string -Operator like -Value 'SEPPmail' -FontWeight bold -Color $colorSEPPmailGreen -Row 
                        }
                        }
                        if ($malFlw) {
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.malFlw.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.malFlw.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.malFlw.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        #Mailflowstatusreport und Paging auf 20 Einträge (eingeklappt oder anpassen)
                        New-HTMLTable -DataTable $malFlw @tableStyle -PagingLength 20
                        } #FIXME: Else Messages incl Style
                    }
                    New-HTMLContent @contentHeaderStyle @ContentBodyStyle -HeaderText 'Transport Rules' -Direction 'column' -Content {
                        if ($tapRls) {
                            New-HTMLHeading -Heading h2 -HeadingText $ExoData.tapRls.HdgTxt -Color $ColorSEPPmailGreen -Underline
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.tapRls.HlpInf)}
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.tapRls.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                            Write-Verbose "Add SEPPmail.cloud PowerShell Module version number to SEPPmail Transportrules if available"
                            foreach ($rule in $tapRls) {
                                $tnrVersion = Get-SC365ModuleVersion -InputString $rule.Comments
                                $rule|Add-Member -membertype NoteProperty -Name SC365Version -value $tnrVersion
                            }
                            New-HTMLTable -DataTable $tapRls @tablestyle -DefaultSortColumn 'Name' -SearchBuilder {
                                New-HTMLTableCondition -Name 'Name' -ComparisonType string -Operator like -Value '[SEPPmail' -FontWeight bold -Color $colorSEPPmailGreen -Row #FIXME: doesnt match anymore :-)
                            }
                        } else {
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "No data found"}
                            
                        }
                    }
                    New-HTMLContent @contentHeaderStyle @ContentBodyStyle -HeaderText 'Defender Configuration' -Direction 'column' -Content {
                        if ($apsPol) {
                            New-HTMLHeading -Heading h2 -HeadingText $ExoData.apsPol.HdgTxt -Color $ColorSEPPmailGreen -Underline
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.apsPol.HlpInf)}
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.apsPol.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                            New-HTMLTable -DataTable $apsPol @tablestyle -DefaultSortColumn 'IsDefault' -DefaultSortOrder 'Descending' -SearchBuilder {
                                New-HTMLTableCondition -Name 'isDefault' -ComparisonType string -Operator eq -Value 'True' -FontWeight bold -row
                            }
                        }
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.MwfPol.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.mwfPol.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.MwfPol.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $MwfPol @tablestyle -DefaultSortColumn 'IsDefault' -DefaultSortOrder 'Descending' -SearchBuilder {
                            New-HTMLTableCondition -Name 'isDefault' -ComparisonType string -Operator eq -Value 'True' -FontWeight bold -row
                        }                    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.hctFpl.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.hctFpl.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.HctFpl.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $hctFpl @tablestyle -DefaultSortColumn 'IsDefault' -DefaultSortOrder 'Descending' -SearchBuilder {
                            New-HTMLTableCondition -Name 'isDefault' -ComparisonType string -Operator eq -Value 'True' -FontWeight bold -row    
                            New-HTMLTableCondition -Name 'MarkAsSpamSpfRecordHardFail' -ComparisonType string -Operator eq -Value 'On' -Row -Color Red -FontWeight bold
                        }

                        ##FIXME Im Parallel Modus darf die Default Policy NICHT aktiv sein ==> ROT
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.hcnFpl.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.hcnFpl.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.HcnFpl.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $hcnFpl @tablestyle -DefaultSortColumn 'IsDefault' -DefaultSortOrder 'Descending'  -SearchBuilder {
                            New-HTMLTableCondition -Name 'isDefault' -ComparisonType string -Operator eq -Value 'True' -FontWeight bold -row
                        }
                        if ($blkSnd) {
                            New-HTMLHeading -Heading h2 -HeadingText $ExoData.blkSnd.HdgTxt -Color $ColorSEPPmailGreen -Underline
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.hcnFpl.HlpInf)}
                            New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.blkSnd.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                            New-HTMLTable -DataTable $blkSnd @tablestyle
                        }
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.qarPol.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.qarPol.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.qarPol.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $qarPol @tablestyle -DefaultSortColumn 'IsDefault' -DefaultSortOrder 'Descending'
                    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.hobFpl.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.hobFpl.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.hobFpl.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $hobFpl @tablestyle -DefaultSortColumn 'IsDefault' -DefaultSortOrder 'Descending'
                    }
                    New-HTMLContent @contentHeaderStyle @ContentBodyStyle -Headertext 'Hybrid Information' -Direction 'column' -Collapsed -Content {
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.iorCon.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.iorCon.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.iorCon.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $iorCon @tablestyle
                    
                        New-HTMLHeading -Heading h2 -HeadingText $ExoData.hybMdc.HdgTxt -Color $ColorSEPPmailGreen -Underline
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output $($ExoData.hybMdc.HlpInf)}
                        New-HTMLTextBox @helpTextStyle -TextBlock {Write-Output "Link to the original CmdLet for further exploration <a href =`"$($ExoData.hybMdc.WebLnk)`" target=`"_blank`">CmdLet Help</a>"}                
                        New-HTMLTable -DataTable $hybMdc.DatacenterIPs @tableStyle
                    }
                }
            }
            #endregion
            # Write Report to Disk

            Write-Verbose "Write the Files to Disk"
            try {
                $finalReport|Out-File -FilePath $FinalPath -Force
                Write-Verbose "If JSONBackup is selected, write a JSON Backup"
                if ($jsonBackup) {
                    # Store json in the same location as HTML
                    #FIXME erzeuge RAWDATA
                    foreach ($ExoDataKey in $ExoData.Keys) {
                        $InfoData = $ExoData[$ExoDataKey]
                        $VarNamRawJSON = "$($InfoData.VarNam)" + "Raw"
                        $script:JsonData += (Get-Variable -Name $VarNamRawJSON | Select-Object -ExpandProperty Value)|Convertto-Json -Depth 5
                    }                 
                    $jsonData = Set-Content -Value $JsonData -Path $jsonPath -force
                }

            }
            catch{
                Write-Warning "Could not write report to $FinalPath"
                if ($IsWindows) {
                    $FinalPath = Join-Path -Path $env:localappdata -ChildPath $ReportFilename
                    if ($jsonBackup) {
                        $jsonpath = (Join-Path -Path (split-path $FinalPath -Parent) -ChildPath (split-path $FinalPath -leafbase)) + '.json'
                    }
                }
                if (($IsMacOs) -or ($isLinux)) {
                    $Finalpath = Join-Path -Path $env:HOME -ChildPath $ReportFilename
                    if ($jsonBackup) {
                        $jsonpath = (Join-Path -Path (split-path $FinalPath -Parent) -ChildPath (split-path $FinalPath -leafbase)) + '.json'
                    }
                }
                Write-Verbose "Writing report to $finalPath"
                try {
                    $finalReport|Out-File -FilePath $finalPath -Force
                    if ($jsonBackup) {
                        # Store json in the same location as HTML
                        Set-Content -Value $JsonData -Path $jsonPath -force
                    }
                }
                catch {
                    $error[0]
                }
            }
            if ($IsWindows) {
                Write-Information -MessageData "Opening $finalPath with default browser"
                Invoke-Expression "& '$finalpath'"
            }
            if (($IsMacOs) -or ($isLinux)) {
                "Report is stored on your disk at $finalpath. Open with your favorite browser."
                if ($jsonBackup) {
                    "Json Backup is stored on your disk at $jsonPath. Open with your favorite editor."
                }
            }
        }
        catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
        }
    }
    end {
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

<#
.SYNOPSIS
    Read Office/Microsoft365 Azure TenantID
.DESCRIPTION
    Every Exchange Online is part of some sort of Microsoft Subscription and each subscription has an Azure Active Directory included. We need the TenantId to identify managed domains in seppmail.cloud
.EXAMPLE
    PS C:\> Get-SC365TenantID -maildomain 'contoso.de'
    Explanation of what the example does
.INPUTS
    Maildomain as string (mandatory)
.OUTPUTS
    TenantID (GUID) as string
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
Function Get-SC365TenantID {
    [CmdLetBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    param (
        [Parameter(Mandatory=$true)]
        [Alias('SEPPmailCloudDomain')]
        [string]$maildomain
    )

    $uri = 'https://login.windows.net/' + $maildomain + '/.well-known/openid-configuration'
    $TenantId = (Invoke-WebRequest $uri| ConvertFrom-Json).token_endpoint.Split('/')[3]
    Return $tenantid
}

<#
.SYNOPSIS
    Test Exchange Online connectivity
.DESCRIPTION
    When staying in a Powershell Session with Exchange Online many things can occur to disturb the session. The Test-SC365connectivity CmdLet figures out if the session is still valid
.EXAMPLE
    PS C:\> Test-SC365ConnectionStatus
    Whithout any parameter the CmdLet emits just true or false
.EXAMPLE
    PS C:\> Test-SC365ConnectionStatus -verbose
    For deeper analisys of connectivity issues the verbose switch provides a lot of relevant information.
.EXAMPLE
    PS C:\> Test-SC365ConnectionStatus -showDefaultDomain
    ShowDefaultdomain will also emit the current default e-mail domain 
.EXAMPLE
    PS C:\> Test-SC365ConnectionStatus -Connect
    Connnect will try to connect via the standard method (web-browser) 
.INPUTS
    Inputs (if any)
.OUTPUTS
    true/false
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
function Test-SC365ConnectionStatus
{
    [CmdLetBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    Param
    (
        [Parameter(
            Mandatory=$false,
            HelpMessage = 'If turned on, the CmdLet will emit the current default domain'
        )]
        [switch]$showDefaultDomain,

        [Parameter(
            Mandatory=$false,
            HelpMessage = 'If turned on, the CmdLet will try to connect to Exchange Online is disconnected'
        )]
        [switch]$Connect

    )

    [bool]$isConnected = $false

    Write-Verbose "Check if module ExchangeOnlineManagement is imported"
    if(!(Get-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue))
    {
        Write-Warning "ExchangeOnlineManagement module not yet imported, importing ..."

        if(!(Import-Module ExchangeOnlineManagement -PassThru -ErrorAction SilentlyContinue))
        {throw [System.Exception] "ExchangeOnlineManagement module does not seem to be installed! Use 'Install-Module ExchangeOnlineManagement' to install.'"}
    }
    else
    {
        $ExoConnInfo = if (Get-Connectioninformation) {(Get-ConnectionInformation)[-1]}

        if ($ExoConnInfo) {
            Write-Verbose "Connected to Exchange Online Tenant $($ExoConnInfo.TenantID)"
            if ((!($global:tenantAcceptedDomains)) -or ((Get-AcceptedDomain) -ne $tenantAcceptedDomains) ) {
                try {
                    $global:tenantAcceptedDomains = Get-AcceptedDomain -Erroraction silentlycontinue
                }
                catch {
                    Write-Error "Cannot detect accepted domains, maybe disconnected. Connect to Exchange Online and load the module again!"
                    break
                }       
            }
            [datetime]$TokenExpiryTimeLocal = $ExoConnInfo.TokenExpiryTimeUTC.Datetime.ToLocalTime()
            $delta = New-TimeSpan -Start (Get-Date) -End $TokenExpiryTimeLocal
            $ticks = $delta.Ticks
            if ($ticks -like '-*') # Token expired
            {
                $isconnected = $false
                Write-Warning "You're not actively connected to your Exchange Online organization. TOKEN is EXPIRED"
                if(($InteractiveSession) -and ($Connect))# defined in public/Functions.ps1
                {
                    try
                    {
                        # throws an exception if authentication fails
                        Write-Verbose "Reconnecting to Exchange Online"
                        Connect-ExchangeOnline -SkipLoadingFormatData
                        $isConnected = $true
                    }
                    catch
                    {
                        throw [System.Exception] "Could not connect to Exchange Online, please retry."}
                }
                else {
                    $isConnected = $false
                }
                
            }
            else # Valid connection
            {
                $tokenLifeTime = [math]::Round($delta.TotalHours)
                Write-verbose "Active session token exipry time is $TokenExpiryTimeLocal (roughly $tokenLifeTime hours)"
                $tmpModuleName = Split-Path -Path $ExoConnInfo.ModuleName -Leaf
                Write-verbose "Active session Module name is $tmpModuleName"
                
                $isConnected = $true
                    
                [string] $Script:ExODefaultDomain = Get-AcceptedDomain | Where-Object{$_.Default} | Select-Object -ExpandProperty DomainName -First 1
                if ($showDefaultDomain) {"$Script:ExoDefaultdomain"}
            }
            } 
            else # No Connection 
            {
                if(($InteractiveSession) -and ($connect)) # defined in public/Functions.ps1
                {
                    try
                    {
                        # throws an exception if authentication fails
                        Write-Verbose "Connecting to Exchange Online"
                        Connect-ExchangeOnline -SkipLoadingFormatData
                    }
                    catch
                    {
                        throw [System.Exception] "Could not connect to Exchange Online, please retry."}
                }
                else {
                    $isConnected = $false
                }
            }
    }
    return $isConnected
}
<#
.SYNOPSIS
    Validates if a given strig (domainname) is the tenant default maildomain
.DESCRIPTION
    It takes the string and prooves if the domain is a member of the tenant. If this fails, the function returns an error and stops.
    If the domain is a tenant-member it returns true/false if its the domain default.
.NOTES
    - none -
.LINK
    - none -
.EXAMPLE
    Confirm-SC365TenantDefaultDomain -Domain 'contoso.eu'
    Returns either an error if the domain is NOT in the tenant, of true or false.
#>
function Confirm-SC365TenantDefaultDomain {
    param (
        [CmdLetBinding()]

        [Parameter(Mandatory = $true)]
        [String]$ValidationDomain
    )

    begin {
        $TenantDefaultDomain = $TenantAcceptedDomains.Where({$_.Default -eq $true}).DomainName     
    }
    process {
        If (!($TenantAcceptedDomains.DomainName -contains $ValidationDomain)) {
            Write-Error "$ValidationDomain is not member of the connected tenant. Retry using only tenant-domains (Use CmdLet: Get-AcceptedDomain)"
            break
        } else {
            if ($Validationdomain -eq $TenantDefaultDomain) {
                return $true
            } else {
                return $false
            }    
        }
    }
    end {
    }
}

<#
.SYNOPSIS
    Tracks Messages in Exchange Online connected with SEPPmail.cloud
.DESCRIPTION
    Combines information from messagetrace and messagetrace details, to emit information how a specific message went throught Exchangeonline
.NOTES
    - none - 
.LINK
    - none -
.EXAMPLE
    Get-SC365Messagetrace -MessageId '123@somedomain.com' -RecipientAddress 'bob@contoso.com'
#>
function Get-SC365MessageTrace {
    [CmdLetBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName=$true)]
        [String]$MessageId,
        
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName=$true)]
        [Alias('RecipientAddress')]
        [String]$Recipient
    )
    begin {
        $OldEaPreference = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
        Write-Information "This CmdLet is still under development" -InformationAction Continue
        
        Write-Verbose "Collecting SEPPmail Connectors for Messagetrace-Detail analysis"
        try {
                if (!($ibc = Get-Inboundconnector -Identity '[SEPPmail*')) {
                    Write-Error "Could not find SEPPmail Inbound-Connecor"
                }
                if (!($obc = Get-Outboundconnector -Identity '[SEPPmail*')) {
                    Write-Error "Could not find SEPPmail Outbound-Connecor"
                }
            }
        catch {
            Write-Error "Could not detect SEPPmail Connectors, aborting"
            break
            }
        $TenantAcceptedDomains = Get-AcceptedDomain
        }
   
    process {

        Write-Verbose "Retrieving initial Message-Trace id MessageID $MessageId for recipient $Recipient"
        Write-Progress -Activity "Loading message data" -Status "MessageTrace" -PercentComplete 0 -CurrentOperation "Start"

        $PlainMessageID = $MessageId.Trim('<','>')
        Write-Verbose "Formatting Parameterinput Messageid:$MessageId - adding < and > at beginning and end to filter property"
        if (!($MessageId.StartsWith('<'))) {$MessageId = "<" + $MessageId}
        if (!($MessageId.EndsWith('>'))) {$MessageId = $MessageId + ">"}
        Write-Verbose "MessageID after formatting is now $MessageId"

        Write-Progress -Activity "Loading message data" -Status "MessageTrace" -PercentComplete 40 -CurrentOperation "Messages loaded"
        $MessageTrace = Get-MessageTrace -MessageId $PlainMessageId -RecipientAddress $Recipient
        
        if ($MessageTrace.count -eq 1) {
            Write-Verbose "Try to find modified encrypted message"
            $EncMessageTrace = Get-MessageTrace -RecipientAddress $Recipient  | Where-Object {($_.MessageId -like '*' + $PlainMessageId + '>')}
            Write-verbose "Found message with modified MessageID $Messagetrace.MessageID"
            if ($EncMessageTrace.count -eq 2) {
                $MessageTrace = Get-MessageTrace -RecipientAddress $Recipient  | Where-Object {($_.MessageId -like '*' + $PlainMessageId + '>')}
            }
        }
        if (!($MessageTrace)){
            Write-Error "Could not find Message with ID $MessageID and recipient $recipient. Look for typos." 
            Write-Error "Message too young ? Try Get-Messagetrace"
            Write-Error "Message too old ? Try Search-MessageTrackingReport "
            break
        }
        try {
            Write-verbose "Test Maildirection, based on the fact that the $Recipient is part of TenantDomains"
            If (($TenantAcceptedDomains.DomainName).Contains(($Recipient -Split '@')[-1])) {
                $MailDirection = 'InBound'
            }
            else {
                $MailDirection = 'OutBound'
            }
        } 
        catch {
            Write-Error "Could not detech mail-direction of recipient-address $recipient. Check for typos and see error below."
            $error[0]
            break
        }

        Write-Verbose "Crafting basic MessageTraceInfo"

        $OutPutObject = [PSCustomObject][ordered]@{
            Subject                = if ($MessageTrace.count -eq 1) {$MessageTrace.Subject} else {$MessageTrace[0].Subject}
            Size                   = if ($MessageTrace.count -eq 1) {Convertto-SC365Numberformat -rawnumber $($MessageTrace.Size)} else {Convertto-SC365Numberformat -rawnumber $($MessageTrace[0].Size)}
            SenderAddresses        = if ($MessageTrace.count -eq 1) {$MessageTrace.SenderAddress} else {$MessageTrace[0].SenderAddress}
            MailDirection          = $MailDirection
            RoutingMode            = if (($ibc.identity -eq "[SEPPmail.cloud] Inbound-Parallel") -or ($obc.identity -eq "[SEPPmail.cloud] Outbound-Parallel")) {'Parallel'} else {'Inline'}
        } 
        Write-Verbose "Add MessageTraceId´s"
        foreach ($i in $Messagetrace) {Add-Member -InputObject $OutPutObject -MemberType NoteProperty -Name ('MessageTraceId' + ' Index'+ $I.Index) -Value $i.MessagetraceId}

        if ($MessageTrace.count -eq 1) {
            Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromIP -Value $MessageTrace.FromIP
            Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromDNS -Value (Resolve-DNS -query $MessageTrace.FromIP).Answers|Select-Object -ExpandProperty Address|Select-Object -ExpandProperty IPAddressToString
            if ($messagetrace.ToIp) {Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToIP -Value $MessageTrace.ToIP}
            if ($messagetrace.ToIp) {Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToDNS -Value (((Resolve-DNS -Query $MessageTrace.ToIP -QueryType PTR).Answers).PtrDomainName).Value}

        } else {
            if ($Messagetrace[0].FromIP) {
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromIP -Value $MessageTrace[0].FromIP
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromDNS -Value ((((Resolve-DNS -Query $MessageTrace[0].FromIP -QueryType PTR).Answers).PtrDomainName).Value)
            }
            else {
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromIP -Value '---empty---'
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromDNS -Value '---empty---'
            }
            if ($MessageTrace[0].ToIP) {
                Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToIP -Value $MessageTrace[0].ToIP
                Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToDNS -Value ((((Resolve-DNS -Query $MessageTrace[0].ToIP -QueryType PTR).Answers).PtrDomainName).Value)
            } else {
                Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToIP -Value '---empty---'
            }
        }
        Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name 'SplitLine' -Value "-------------------- MessageTrace DETAIL Info Starts Here --------------------"

        switch ($maildirection)
        {
            {(($_ -eq 'InBound') -and ($ibc.identity -match "\[SEPPmail\]") -and ($ibc.ConnectorType -eq 'OnPremises'))} 
            {
                # Im Parallel Mode kommt die Mail 2x, einmal von externem Host und einmal von SEPpmail, Index 0 und 1
                
                #$MessageTraceDetailExternal = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient
                $MTDExtReceived = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient -Event 'RECEIVE'
                $MTDExtExtSend = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient | Where-Object event -like '*send*'
                #$MessageTraceDetailSEPPmail = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient
                $MTDSEPPReceived = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient -Event 'RECEIVE'
                $MTDSEPPDelivered = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient -Event 'DELIVER'
                Write-Verbose "Crafting Inbound Connector Name"
                try {
                    $ibcName = (($MTDSEPPReceived.Data).Split(';') | Select-String 'S:InboundConnectorData=Name').ToString().Split('=')[-1]
                } 
                catch 
                {
                    $ibcName = '--- E-Mail did not go over SEPPmail Connector ---'
                }
                Write-Verbose "Preparing Output (Receive)Inbound-Parallel"
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExternalReceivedTime -Value $messageTrace[1].Received
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExternalReceivedSize -Value $messageTrace[1].Size
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToIP -Value $messageTrace[1].ToIP
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToDNS -Value ((((Resolve-DNS -Query $MessageTrace[1].ToIP -QueryType PTR).Answers).PtrDomainName).Value)
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExtMessageTraceId -Value $MessageTrace[1].MessageTraceId.Guid
                $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPMessageTraceId -Value $MessageTrace[0].MessageTraceId.Guid
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'FullTransportTime(s)' -Value (New-TimeSpan -Start $MTDExtReceived.Date -End $MTDSEPPDelivered.Date).Seconds
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'ExoTransportTime(s)' -Value (New-TimeSpan -Start $MTDExtReceived.Date -End $MTDExtExtSend.Date).Seconds
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'SEPPTransportTime(s)' -Value (New-TimeSpan -Start $MTDSEPPReceived.Date -End $MTDSEPPDelivered.Date).Seconds
                #$Outputobject | Add-Member -MemberType NoteProperty -Name SubmitDetail -Value $MTDSEPPDelivered.Detail # Boring
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExtSendDetail -Value $MTDExtExtSend.Detail
                $Outputobject | Add-Member -MemberType NoteProperty -Name InboundConnectorName -Value $ibcName
            }
            {(($_ -eq 'InBound') -and ($ibc.identity -match "\[SEPPmail\]")  -and ($ibc.ConnectorType -eq 'Partner'))} 
            {
                $MessageTraceDetail = Get-MessagetraceDetail -MessageTraceId $MessageTrace.MessageTraceId -Recipient $Recipient
                #$MTDReceived = $MessageTraceDetail|where-object {($_.Event -eq 'Received') -or ($_.Event -eq 'Empfangen')} 
                $MTDReceived = Get-MessagetraceDetail -MessageTraceId $MessageTrace.MessageTraceId -Recipient $Recipient -event 'received'
                #$MTDDelivered = $MessageTraceDetail|where-object {($_.Event -eq 'Delivered') -or ($_.Event -eq 'Zustellen')}
                Write-Verbose "Crafting Inbound Connector Name"
                try {
                    $ibcName = (($MTDReceived.Data).Split(';')|select-string 'S:InboundConnectorData=Name').ToString().Split('=')[-1]
                } catch {
                    $ibcName = '--- E-Mail did not go over SEPPmail Connector ---'
                }
                Write-Verbose "Preparing Output (Receive)Inbound-Inline"
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExternalReceivedTime -Value $messageTrace.Received
                #$Outputobject | Add-Member -MemberType NoteProperty -Name DeliveredDetail -Value $MTDDelivered.Detail # Boring Info
                $Outputobject | Add-Member -MemberType NoteProperty -Name ReceivedDetail -Value  $MTDReceived.Detail
                if ($MTReceived) {
                    $Outputobject | Add-Member -MemberType NoteProperty -Name 'ExoTransportTime (s)' -Value (New-TimeSpan -Start $MTReceived.Date -End $MTDelivered.Date).Seconds
                } else {
                    $outPutObject | Add-Member -MemberType NoteProperty -Name 'ExoTransportTime (s)' -Value '---cannot determine transporttime in inline mode---'
                }
                $Outputobject | Add-Member -MemberType NoteProperty -Name InboundConnectorName -Value $ibcName
            }
            {(($_ -eq 'OutBound') -and ($obc.identity -match "\[SEPPmail\]") -and ($ibc.ConnectorType -eq 'OnPremises'))}
            {
                # We take one of 2 Send/Receive Messagetraces from SEPPmail and get the details

                # Now this one has 3 Parts. 0= Recieve from Mailboxhost, 1 = SumbitMessage (Exo internal), 2 = Send to SEPPmail

                $MTDSEPPReceive = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient -Event 'receive'
                $MTDSEPPExtSend = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient |where-object Event -like  '*SEND*'                

                # $MTDSEPPSubmit = $MessageTraceDetailSEPPmail[1] Not interesting for us
                $MTDExtReceive = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient -Event 'receive'
                $MTDExtExtSend = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient |where-object Event -like  '*SEND*'
                try {
                    $obcName = (((($MTDSEPPExtSend.Data -Split '<') -replace ('>','')) -split (';') | select-String 'S:Microsoft.Exchange.Hygiene.TenantOutboundConnectorCustomData').ToString()).Split('=')[-1]
                }catch {
                    $obcName = "--- E-Mail did not go via a SEPPmail Connector ---"
                }
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToIP -Value $messageTrace[1].ToIP
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToDNS -Value ((((Resolve-DNS -Query $MessageTrace[1].ToIP -QueryType PTR).Answers).PtrDomainName).Value)
                $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPmailReceivedFromIP -Value $messageTrace[1].FromIP
                try { 
                    $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPmailReceivedFromDNS -Value ((((Resolve-DNS -Query $MessageTrace[1].FromIP -QueryType PTR).Answers).PtrDomainName).Value)
                } 
                catch {
                    Write-Information "Cannot Resolve $($messageTrace[1].FromIP)" -InformationAction Continue
                }
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'ExoTransPortTime(s)' -Value (New-TimeSpan -Start $MTDExtReceive.Date -End $MTDExtExtSend.Date).Seconds
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'SEPPmailTransPortTime(s)' -Value (New-TimeSpan -Start $MTDSEPPReceive.Date -End $MTDSEPPExtSend.Date).Seconds
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'FullTransPortTime(s)' -Value (New-TimeSpan -Start $MTDSEPPReceive.Date -End $MTDExtExtSend.Date).Seconds
                $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPReceiveDetail -Value $MTDSEPPReceive.Detail
                #$Outputobject | Add-Member -MemberType NoteProperty -Name SEPPSubmitDetail -Value $MTDSEPPSubmit.Detail # Boring
                $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPSendExtDetail -Value $MTDSEPPExtSend.Detail
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExtReceiveDetail -Value $MTDExtReceive.Detail
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExtSendDetail -Value $MTDExtExtSend.Detail
                $Outputobject | Add-Member -MemberType NoteProperty -Name OutboundConnectorName -Value $obcName
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExternalSendLatency -Value (((($MTDExtExtSend.Data -Split '<') -replace ('>','')) -split (';') | select-String 'S:ExternalSendLatency').ToString()).Split('=')[-1]
            }
            {(($_ -eq 'OutBound') -and ($obc.identity -match "\[SEPPmail\]") -and ($ibc.ConnectorType -eq 'Partner'))}
            {
                Write-Progress -Activity "Loading message data" -Status "MessageTrace" -PercentComplete 40 -CurrentOperation "Get-Messagetrace"
                $MessageTraceDetail = Get-MessagetraceDetail -MessageTraceId $MessageTrace.MessageTraceId -Recipient $Recipient
                # 1 = Empfangen/Receive
                $MTDReceive = $MessageTraceDetail|Where-Object {(($_.Event -eq 'Empfangen') -or ($_.Event -eq 'Receive'))}
                # 2 = Übermitteln/Submit
                # $MTDSubmit = $MessageTraceDetail|Where-Object {(($_.Event -eq 'Übermitteln') -or ($_.Event -eq 'Submit'))}
                # 3 = Send external/Extern senden
                $MTDExtSend = $MessageTraceDetail|Where-Object {(($_.Event -eq 'Send external') -or ($_.Event -eq 'Extern senden'))}
                Write-Progress -Activity "Loading message data" -Status "MessageTrace" -PercentComplete 80 -CurrentOperation "Preparing Output"
                Write-Verbose "Crafting Oubound Connector Name" 
                try {
                    $obcName = (((($MTDExtSend.Data -Split '<') -replace ('>','')) -split (';') | select-String 'S:Microsoft.Exchange.Hygiene.TenantOutboundConnectorCustomData').ToString()).Split('=')[-1]
                } catch {
                    $obcName = '--- E-Mail did not go over SEPPmail Connector ---'
                }
                Write-verbose "Adding Specific Outbound-Inline Data to output"
                if (($MTDReceive.Date) -and ($MTDExtSend.Date)) {$Outputobject | Add-Member -MemberType NoteProperty -Name 'ExoInternalTransportTime(s)' -Value (New-TimeSpan -Start $MTDReceive.Date -End $MTDExtSend.Date).Seconds}
                $Outputobject | Add-Member -MemberType NoteProperty -Name ReceiveDetail -Value $MTDReceive.Detail
                #$Outputobject | Add-Member -MemberType NoteProperty -Name SubmitDetail -Value $MTDSubmit.Detail # Keine Relevante Info
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExtSendDetail -Value $MTDExtSend.Detail
                $Outputobject | Add-Member -MemberType NoteProperty -Name OutboundConnectorName -Value $obcName
                if ($MTDExtSend) {$Outputobject | Add-Member -MemberType NoteProperty -Name ExternalSendLatency -Value (((($MTDExtSend.Data -Split '<') -replace ('>','')) -split (';') | select-String 'S:ExternalSendLatency').ToString()).Split('=')[-1]}
                Write-Progress -Activity "Loading message data" -Status "StatusMessage" -PercentComplete 100 -CurrentOperation "Done"
            }
            Default {
                $OutputObject |  Add-Member -MemberType NoteProperty -Name "SEPPmail Integration" -Value 'none found'
                Write-Verbose "E-Mail direction was $_, could not detect SEPPmail"
            }
        }
    }
    end {
        return $OutPutObject
        $ErrorActionPreference = $OldEaPreference
    }
}

<#
.SYNOPSIS
    Read Office/Microsoft365 Azure TenantID
.DESCRIPTION
    Every Exchange Online is part of some sort of Microsoft Subscription and each subscription has an Azure Active Directory included. We need the TenantId to identify managed domains in seppmail.cloud
.EXAMPLE
    PS C:\> Get-SC365TenantID -maildomain 'contoso.de'
    Explanation of what the example does
.INPUTS
    Maildomain as string (mandatory)
.OUTPUTS
    TenantID (GUID) as string
.NOTES
    See https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md for more
#>
Function Show-SC365Tenant {
    [CmdLetBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    param (
    )

    if(!(Test-SC365ConnectionStatus)) {
        throw [System.Exception] "You're not connected to Exchange Online - please connect prior to using this CmdLet"
    } else {
        Write-Verbose "Connected to Exchange Organization `"$Script:ExODefaultDomain`"" -InformationAction Continue
        try {
            $where = Get-OrganizationConfig
            $youare = [PSCustomobject]@{
                                 'Name' = $where.Name 
                          'DisplayName' = $where.DisplayName 
                 'DefaultMailboxRegion' = $where.DefaultMailboxRegion 
                'AllowedMailboxRegions' = $where.AllowedMailboxRegions 
                      'ExchangeVersion' = $where.ExchangeVersion
                 'SendFromAliasEnabled' = $where.SendFromAliasEnabled
                'IsTenantInGracePeriod' = $Where.IsTenantInGracePeriod        
                }
            return $youare
         } catch {
            Write-Error $Error[0].Message   
        }    
    }


}

<#
.SYNOPSIS
    Read DateTime of SEPPmail.cloud Setup
.DESCRIPTION
    Reads the Creation Time of the SEPPmail.cloud Inbound Connector and emits dateTime of the moment the setup occured.
.NOTES
    none
.LINK
    none
.EXAMPLE
    Get-SC365SetupTime 
    Montag, 20. März 2023 10:56:04
.EXAMPLE
    Get-Sc365SetupTime -verbose
    Montag, 20. März 2023 10:56:04
    VERBOSE: SEPPmail Cloud was created 13 days ago
    VERBOSE: Inbound Connector Comments writes install date is: 03/20/2023 10:56:04
#>
Function Get-SC365SetupTime {
    [CmdLetBinding()]

    param ()

    begin {}
    process {
        $ibc = Get-InboundConnector -Identity '[SEPPmail.Cloud]*'
        $ibc|Select-Object -ExpandProperty WhenCreated

        $days = (New-Timespan -Start (Get-Date $ibc.WhenCreated) -End (Get-Date)).Days
        Write-verbose "SEPPmail Cloud was created $days days ago"
        $commentstime = ((($ibc.comment|select-String 'Created with') -split 'version').trim()[-1] -split 'on').trim()[-1]
        Write-Verbose "Inbound Connector Comments writes install date is: $commentsTime"
    }
    end{}
}

<#
.SYNOPSIS
    Convert a DNS domain to IDNA Format
.DESCRIPTION
    The CmdLet uses the System.Globalization.IdnMapping .NET Method to convert a DNS Domain with special characters like ä, ü, or ß to the IDNA format.
.NOTES
    none
.LINK
    none
.EXAMPLE
    Get-Sc365SetupTime öbb.at
    xn--bb-eka.at
#>
function ConvertTo-SC365IDNA {
    [CmdLetBinding()]
    param (
        [Parameter(
                    Mandatory = $true,
            ValueFromPipeline = $true
            )]
        [string]$String
    )
    $idn = [System.Globalization.IdnMapping]::new()
    return $idn.GetAscii($String)
}
<#
.SYNOPSIS
    Convert a IDNA Formatted Domain to DNS
.DESCRIPTION
    The CmdLet uses the System.Globalization.IdnMapping .NET Method to convert a IDNA Domain with special characters like ä, ü, or ß to the original DNS format.
.NOTES
    none
.LINK
    none
.EXAMPLE
    Get-Sc365SetupTime xn--bb-eka.at
    öbb.at
#>
function ConvertFrom-SC365IDNA {
    [CmdLetBinding()]
    param (
        [Parameter(
                    Mandatory = $true,
            ValueFromPipeline = $true
            )]
        [string]$String
    )
    $idn = [System.Globalization.IdnMapping]::new()
    return $idn.GetUnicode($String)
}

function Update-SC365Setup {
    [CmdletBinding(
        SupportsShouldProcess = $true
    )]
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
        [string]$TempPrefix = 'temp',

        [Parameter(
            Mandatory = $false,
            HelpMessage = "Do not generate a report and JSON backup"
        )]
        [switch]$noReport

    )

    begin {
        $existEAValue = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
    }
    process {
        if (!($noReport)) {
        Write-verbose "Export Exo-Config as JSON"
        New-SC365ExOReport -jsonBackup
        }
        #region Infoblock
    Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "| This script is a helper and provides basic steps to upgrade your    |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "| SEPPmail.cloud/Exchange integration. It covers only STANDARD Setups!|" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "| The Script will:                                                    |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    1.) Check if there are any orphaned rule or connector objects    |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    2.) Rename SEPPmail.cloud Transport rules to `$backupName         |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    3.) Create Connectors with Temp Name                             |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    4.) Set (200) outbound transport rule to New-Connector           |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    5.) Rename SEPPmail.cloud Connectors to `$backupName              |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    6.) Attach old Transport rules to old Connector with BackupNam   |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    ----------------- OLD SETUP STILL RUNNING ------------------     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    7.) Rename NEW Connectors to original names                      |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    8.) Create new transport rules -PlacementPriority TOP            |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    -------------------- NEW SETUP RUNNING ----------------------    |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|    9.) Disable old rules                                            |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|   10.) Disable old connectors                                       |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|   11.) on -remove delete old transport rules and connectors         |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "| If you have any:                                                    |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|   - customizations to SEPPmail.cloud rules                          |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|   - other corporate transport rules                                 |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|   - disclaimer Services integrated via rules                        |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|   - or other special scenarios in your Exo-Tenant                   |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "| you need to adapt/change/post-configure the outcome of this script! |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "| DO NOT JUST FIRE IT UP AND HOPE THINGS ARE GOING TO WORK !!!!!!     |" -ForegroundColor Magenta -BackgroundColor Gray
    Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Magenta -BackgroundColor Gray
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
                Set-TransportRule -Identity $rule.Name -Name ($rule.Name -replace 'SEPPmail.Cloud',$BackupName)
            }
            #endregion rename existing rules to backup name

            #region 3 - create new connectors with temp Name
            Write-Verbose "3 - Creating new connectors with temp name" 
            $newConnectors = New-SC365connectors -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -region $DeplInfo.region -routing $DeplInfo.routing -NamePrefix $tempPrefix # -routing $($DeplInfo.Routing) -region $($DeplInfo.region) #FIXME:
            #endregion create new connectors with temp Name

            #region 4 - set outbound rule to new connector (Inline 200, parallel 1xx)
            Write-Verbose "4 - Set outbound rules to new connector" 
            $oldObc = Get-OutboundConnector -Identity '[SEPPmail.cloud] Outbound-*'
            [string]$tempObcName = $TempPrefix + $($OldObc.Identity)
            $rulesToChange = Get-TransportRule -Identity '[*' |Where-Object {$_.RouteMessageOutboundConnector -ne $null}
            foreach ($rule in $rulesToChange) {
                Set-TransportRule -Identity $($rule.Identity) -RouteMessageOutboundConnector $tempObcName
            }

            #endregion set outbound rule to new connector

            #region 5 - rename old connectors to backup Names
            Write-Verbose "5 - Rename existing SEPPmail.cloud Inbound Connector to $backupName"
            Write-Verbose "5a - Rename existing SEPPmail.cloud Inbound Connector"
            $oldIbc = Get-InboundConnector -Identity '[SEPPmail.cloud] Inbound-*' 
            Set-InboundConnector -Identity $($OldIbc.Identity) -Name ($($OldIbc.Identity) -replace 'SEPPmail.Cloud',$backupName)

            Write-Verbose "5b - Rename existing SEPPmail.cloud Outbound Connector"
            $oldObc = Get-OutBoundConnector -Identity "[SEPPmail.cloud] OutBound-*"
            Set-OutBoundConnector -Identity $($oldObc.Identity) -Name ($oldObc.Identity -replace 'SEPPmail.Cloud',$backupName)
            #endregion

            #region 6 - attach old outbound rules to old connector
            Write-Verbose "6 - Set outbound rule to old backup connector again"
            $bkpConnWildcard = "[" + $backupName + "]*"
            $bkpObc = Get-OutboundConnector -Identity $bkpConnWildcard
            foreach ($rule in $rulesToChange) {
                Set-TransportRule -Identity $($rule.Identity) -RouteMessageOutboundConnector $bkpObc
            }

            #endregion

            #region 7 - rename new connectors to final Names
            Write-Verbose "7 - Rename existing $tempPrefix connectors to final name"
            $finalObCName = ($newConnectors|Where-Object Identity -like '*OutBound*').Identity -replace "^$([regex]::Escape($tempPrefix))", ""
            Set-OutboundConnector -Identity ($newConnectors|Where-Object Identity -like '*OutBound*').Identity -Name $finalObcName

            $finalIbcName = ($newConnectors|Where-Object Identity -like '*Inbound*').Identity -replace "^$([regex]::Escape($tempPrefix))", ""
            Set-InBoundConnector -Identity ($newConnectors|Where-Object Identity -like '*InBound*').Identity -Name $finalIbcName
            #endregion

            #region 8 - create New Transport rules
            Write-Verbose "8 - Creating new Transport Rules" 
            New-SC365Rules -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -routing $DeplInfo.routing  -PlacementPriority Top
            #endregion

            #region 9 - disable old Transport rules
            $trWildcard = '[' + $BackupName + ']*'
            Write-Verbose "9 - Disable old Transport Rules"
            if ($PSCmdlet.ShouldProcess("Disabling Transport Rules matching $trWildcard")) {
                Get-TransportRule -Identity $trWildcard | Disable-TransportRule -Confirm:$false
            } 
            #endregion 9

            #region 10 - Disable old connectors
            Write-Verbose "10 - Disable old connectors" 
            Set-InBoundConnector -Identity $bkpConnWildcard -Enabled:$false
            Set-OutBoundConnector -Identity $bkpConnWildcard -Enabled:$false
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

Register-ArgumentCompleter -CommandName Get-SC365TenantId -ParameterName MailDomain -ScriptBlock $paramDomSB
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
