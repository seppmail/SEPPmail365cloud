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

        $DeplyoymentInfo = [PSCustomObject]@{
            DeployMentStatus    = $null
            SEPPmailCloudDomain = $Null
            region              = $null
            routing             = $null
            InBoundOnly         = $null
            CBCDeployed         = $Null
            CBCConnectorHost    = $null
            CBCDnsEntry         = $null
            CBCDnsTimeStamp     = $null
            InlineMXMatch       = $null
            MailHost            = $null
            GateHost            = $null
            RelayHost           = $null
        }

    }
    
    process {
        #region Select DefaultDomain
            if (!($SEPPmailCloudDomain)) {
                [String]$DNSHostDomain = $tenantAcceptedDomains |Where-Object 'Default' -eq $true |select-Object -ExpandProperty DomainName
                Write-Verbose "Extracted Default-Domain with name $SEPPmailcloudDomain from TenantDefaultDomains"
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

        <#    if (($SEPPmailCloudDomain) -and (!($tenantAcceptedDomains.DomainName.Contains($SEPPmailCloudDomain)))) {
               throw [System.Exception] "Domain $SEPPmailCloudDomain is not member of this tenant! Check for typos or connect to different tenant"
            }#>
        #endregion Select DefaultDomain

        #region Query SEPPmail routing-Hosts DNS records and detect routing mode and in/oitbound
            [string]$relayHost = $DnsHostDomain.Replace('.','-') + '.relay.seppmail.cloud'
             [string]$mailHost = $DnsHostDomain.Replace('.','-') + '.mail.seppmail.cloud'
             [string]$gateHost = $DnsHostDomain.Replace('.','-') + '.gate.seppmail.cloud'
            
        if (((Resolve-Dns -Query $GateHost).Answers)) {
            $routing = 'inline'
            if ((Resolve-Dns -Query $RelayHost).Answers) {                    
                $inBoundOnly = $false
                Write-Verbose "$Gatehost and $relayHost alive ==> InLine-bidirectional"
                $deploymentStatus = $true                
            } else {
                if (!((Resolve-Dns -Query $RelayHost).Answers)) {
                    $inBoundOnly = $true
                    $relayHost = $null
                    Write-Verbose "$Gatehost alive,$relayHost missing ==> InLine-InBound only"
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
                $mx = $mxFull
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
            if ($routing -eq 'inline') {
                [String[]]$GateIP = ((Resolve-Dns -Query $GateHost).Answers)|Select-Object -expand Address| Select-Object -expand IPAddresstoString
                Foreach ($IP in $GateIP) {if ($ch.GateIPs.Contains($Ip)) {$region = 'ch';break}}
                Foreach ($IP in $GateIP) {if ($de.GateIPs.Contains($Ip)) {$region = 'de';break}}
                Foreach ($IP in $GateIP) {if ($prv.GateIPs.Contains($Ip)) {$region = 'prv';break}}
            }
            if ($routing -eq 'parallel') {
               $MailIP = ((Resolve-Dns -Query $MailHost).Answers)|Select-Object -expand Address| Select-Object -expand IPAddresstoString
               Foreach ($ip in $mailIp) {if ($ch.MailIPs.Contains($Ip)) { $region = 'ch';break}}
               Foreach ($ip in $mailIp) {if ($de.MailIPs.Contains($Ip)) { $region = 'de';break}}
               Foreach ($ip in $mailIp) {if ($prv.MailIPs.Contains($IP)) { $region = 'prv';break}}
            }
        #endregion Cloud-IP-Addresses

        #region Check CBC Availability
            [String]$TenantID = Get-SC365TenantID -maildomain $DnsHostDomain -OutVariable "TenantID"
            $TenantIDHash = Get-SC365StringHash -String $TenantID
            [string]$hashedDomain =  $TenantIDHash + '.cbc.seppmail.cloud'
            if (((resolve-dns -query $hashedDomain -QueryType TXT).Answers)) {
               $CBCDeployed = $true
               Write-Verbose "$HashedDomain of TenantID $tenantId has a CBC entry"
            } else {
               $CBCDeployed = $false
               Write-Warning "Could not find TXT Entry for TenantID $TenantID of domain $DNSHostCloudDomain. Setup will most likely fail! Go to the SEPPmail.cloud-portal and check the deployment status."
            }
        #endregion CBC availability
    }
    end {
        $DeplyoymentInfo.DeployMentStatus = $DeploymentStatus
        $DeplyoymentInfo.Region = $region
        $DeplyoymentInfo.Routing = $routing
        $DeplyoymentInfo.InBoundOnly  = $inBoundOnly
        $DeplyoymentInfo.SEPPmailCloudDomain = $DNSHostDomain
        $DeplyoymentInfo.CBCDeployed = $CBCDeployed
        if ($DeplyoymentInfo.DeployMentTime) {$DeplyoymentInfo.DeployMentTime = $DeployMentTime}
        if ($region) {$DeplyoymentInfo.CBCConnectorHost = ($tenantId + ((Get-Variable $region).Value.TlsCertificate).Replace('*',''))}
        if ($CBCDeployed -eq $true) {$DeplyoymentInfo.CBCDnsEntry = ($TenantIDHash + '.cbc.seppmail.cloud')}
        if ($routing -eq 'inline') {$DeplyoymentInfo.InlineMXMatch = $MxMatch}
        if (($routing -eq 'inline') -and (!($inBoundOnly))) {$DeplyoymentInfo.RelayHost = $relayHost}
        if ($routing -eq 'inline') {$DeplyoymentInfo.GateHost = $gateHost}
        if ($routing -eq 'parallel') {$DeplyoymentInfo.MailHost = $MailHost}

        return $DeplyoymentInfo
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
.INPUTS
    FilePath
.OUTPUTS
    HTML Report
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
            Write-verbose 'Defining Function fo read Exo Data and return an info Message in HTML even if nothing is retrieved'
        }

        function New-SelfGeneratedReportName {
            Write-Verbose "Creating self-generated report filename."
            return ("{0:HHm-ddMMyyy}" -f (Get-Date)) + (Get-AcceptedDomain|where-object default -eq $true|select-object -expandproperty Domainname) + '.html'
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
        #endregion

        function Get-ExoHTMLData {
            param (
                [Parameter(
                      Mandatory = $true,
                    HelpMessage = 'Enter Cmdlte to ')]
                [string]$ExoCmd
            )
            try {
                $allCmd = $exoCmd.Split('|')[0].Trim()
                $htmlSelectCmd = $exoCmd.Split('|')[-1].Trim()

                $rawData = Invoke-Expression -Command $allCmd
                if ($null -eq $rawData) {
                    $ExoHTMLData = New-object -type PSobject -property @{Result = '--- no information available ---'}|Convertto-HTML -Fragment
                } else {
                    $ExoHTMLCmd = "{0}|{1}" -f  $allcmd,$htmlSelectCmd
                    $ExoHTMLData = Invoke-expression -Command $ExoHTMLCmd |Convertto-HTML -Fragment
                    if ($jsonBackup) {
                        $script:JsonData += '---' + $AllCmd + '---'|Convertto-Json
                        $script:JsonData += $rawData|ConvertTo-Json
                    }
                } 
                return $ExoHTMLData
            }
            catch {
                Write-Warning "Could not fetch data from command '$exoCmd'"
            }    
        }

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
            $mv = $myInvocation.MyCommand.Version
            $Top = "<p><h1>Exchange Online Report</h1><p>"
            $now = Get-Date
            if ($PSVersionTable.OS -like 'Microsoft Windows*') {
                $repUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            } else {
                $repUser = (hostname) + '/' + (whoami)
            }
            $RepCreationDateTime = "<p><body>Report created on: $now</body><p>"
            $RepCreatedBy = "<p><body>Report created by: $repUser</body><p>"
            $ReportFilename = Split-Path $FinalPath -Leaf
            $moduleVersion = "<p><body>SEPPmail365cloud Module Version: $mv</body><p>"
            $reportTenantID = Get-SC365TenantID -maildomain (Get-AcceptedDomain|where-object InitialDomain -eq $true|select-object -expandproperty Domainname)
            $TenantInfo = "<p><body>Microsoft O/M365 AzureAD Tenant ID: $reportTenantID</body><p>"
            Write-Verbose "Collecting Accepted Domains"
            $hSplitLine = '<p><h2>---------------------------------------------------------------------------------------------------------------------------</h2><p>'
            #region General infos
            $hGeneral =  '<p><h2>General Exchange Online and Subscription Information</h2><p>'
            
            $hA = '<p><h3>Accepted Domains</h3><p>'
            $A = Get-ExoHTMLData -ExoCmd 'Get-AcceptedDomain |select-object Domainname,DomainType,Default,EmailOnly,ExternallyManaged,OutboundOnly|Sort-Object -Descending Default '
            # Find out Office Configuration
            Write-Verbose "Collecting M365 Configuration"
            $hB = '<p><h3>ExO Configuration Details</h3><p>'
            $B = Get-ExoHTMLData -ExoCmd 'Get-OrganizationConfig |Select-Object DisplayName,ExchangeVersion,AllowedMailboxRegions,DefaultMailboxRegion,DisablePlusAddressInRecipients'

            # Find out possible Sending Limits for LFT
            Write-Verbose "Collecting Send and Receive limits for SEPPmail LFT configuration"
            $hP = '<p><h3>Send and Receive limits (for SEPPmail LFT configuration)</h3><p>'
            $P = Get-ExoHTMLData -ExoCmd  'Get-TransportConfig |Select-Object MaxSendSize,MaxReceiveSize'

            # Find out possible Office Message Encryption Settings
            Write-Verbose "Collecting Office Message Encryption Settings"
            $hP = '<p><h3>Office Message Encryption Settings</h3><p>'
            $P = Get-ExoHTMLData -ExoCmd 'Get-OMEConfiguration|Select-Object PSComputerName,TemplateName,OTPEnabled,SocialIdSignIn,ExternalMailExpiryInterval,Identity,IsValid'
            
            # Get MX Record Report for each domain
            $hO = '<p><h3>MX Record for each Domain</h3><p>'
            $O = $Null
            $oTemp = Get-AcceptedDomain
            Foreach ($AcceptedDomain in $oTemp.DomainName) {
                    $O += (Get-MxRecordReport -Domain $AcceptedDomain|Select-Object -Unique|Select-Object HighestPriorityMailhost,HighestPriorityMailhostIpAddress,Domain|Convertto-HTML -Fragment)
            }
            #endregion

            #region Security 
            $hSecurity = '<p><h2>Security related Information</h2><p>'
            $hC = '<p><h3>DKIM Settings</h3><p>'
            $C = Get-ExoHTMLData -ExoCmd 'Get-DkimSigningConfig|Select-Object Domain,Enabled,Status,Selector1CNAME,Selector2CNAME|sort-object Enabled -Descending'
            
            Write-Verbose "Collecting Phishing and Malware Policies"
            $hD = '<p><h3>Anti Phishing Policies</h3><p>'
            $D = Get-ExoHTMLData -ExoCmd 'Get-AntiPhishPolicy|Select-Object Identity,isDefault,IsValid,AuthenticationFailAction'
            
            $hE = '<p><h3>Anti Malware Policies</h3><p>'
            $E = Get-ExoHTMLData -ExoCmd 'Get-MalwareFilterPolicy|Select-Object Identity,Action,IsDefault,Filetypes'

            $hk = '<p><h3>Content Filter Policy</h3><p>'
            $k= Get-ExoHTMLData -ExoCmd 'Get-HostedContentFilterPolicy|Select-Object QuarantineRetentionPeriod,EndUserSpamNotificationFrequency,TestModeAction,IsValid,BulkSpamAction,PhishSpamAction,OriginatingServer'

            Write-Verbose "Blocked Sender Addresses"
            $hH = '<p><h3>Show Senders which are locked due to outbound SPAM</h3><p>'
            $h = Get-ExoHTMLData -ExoCmd 'Get-BlockedSenderAddress'
            
            Write-Verbose "Get Outbound SPAM Filter Policy"
            $hJ = '<p><h3>Outbound SPAM Filter Policy</h3><p>'
            $J = Get-ExoHTMLData -ExoCmd 'Get-HostedOutboundSpamFilterPolicy|Select-Object Name,IsDefault,Enabled,ActionWhenThresholdReached'
            
            Write-Verbose "Get Filter Policy"
            $hJ1 = '<p><h3>SPAM Filter Policy</h3><p>'
            $J1 = Get-ExoHTMLData -ExoCmd 'Get-HostedConnectionFilterPolicy|select-Object Name,IsDefault,Enabled,IPAllowList,IPBlockList'
            #endregion Security

            #region other connectors
            $hOtherConn = '<p><h2>Hybrid and other Connectors</h2><p>'
            Write-Verbose "Get-HybridMailflow"
            $hG = '<p><h3>Hybrid Mailflow Information</h3><p>'
            $g = Get-ExoHTMLData -ExoCmd 'Get-HybridMailflow'

            Write-Verbose "Get-IntraorgConnector"
            $hI = '<p><h3>Intra Org Connector Settings</h3><p>'
            $I = Get-ExoHTMLData -ExoCmd 'Get-IntraOrganizationConnector|Select-Object Identity,TargetAddressDomains,DiscoveryEndpoint,IsValid'
            #endregion

            #region connectors
            $hConnectors = '<p><h2>Existing Exchange Connectors</h2><p>'
            
            Write-Verbose "InboundConnectors"
            $hL = '<p><h3>Inbound Connectors</h3><p>'
            $L = Get-ExoHTMLData -ExoCmd 'Get-InboundConnector |Select-Object Identity,Enabled,SenderDomains,SenderIPAddresses,OrganizationalUnitRootInternal,TlsSenderCertificateName,OriginatingServer,IsValid'
            
            Write-Verbose "OutboundConnectors"
            $hM = '<p><h3>Outbound Connectors</h3><p>'
            $M = Get-ExoHTMLData -ExoCmd 'Get-OutboundConnector -IncludeTestModeConnectors:$true|Select-Object Identity,Enabled,SmartHosts,TlsDomain,TlsSettings,RecipientDomains,OriginatingServer,IsValid'
            #endregion connectors
            
            #region mailflow rules
            $hTransPortRules = '<p><h2>Existing Mailflow Rules</h2><p>'
            Write-Verbose "TransportRules"
            $hN = '<p><h3>Existing Transport Rules</h3><p>'
            $N = Get-ExoHTMLData -ExoCmd 'Get-TransportRule | select-object Name,State,Mode,Priority,FromScope,SentToScope,StopRuleProcessing'
            #endregion transport rules

            $HeaderLogo = [Convert]::ToBase64String((Get-Content -path $PSScriptRoot\..\HTML\SEPPmailLogo_T.png -AsByteStream))

            $LogoHTML = @"
<img src="data:image/jpg;base64,$($HeaderLogo)" style="left:150px alt="Exchange Online System Report">
"@

            $hEndOfReport = '<p><h2>--- End of Report ---</h2><p>'
            $style = Get-Content -Path $PSScriptRoot\..\HTML\SEPPmailReport.css
            $finalreport = Convertto-HTML -Body "$LogoHTML $Top $RepCreationDatetime $RepCreatedBy $moduleVersion $TenantInfo`
                   $hSplitLine $hGeneral $hSplitLine $hA $a $hB $b $hO $o`
                  $hSplitLine $hSecurity $hSplitLine $hC $c $hd $d $hE $e $hP $P $hH $H $hK $k $hJ $j $hJ1 $J1 `
                 $hSplitLine $hOtherConn $hSplitLine $hG $g $hI $i `
                $hSplitLine $hConnectors $hSplitLine $hL $l $hM $m `
            $hSplitLine $hTransPortRules $hSplitLine $hN $n $hEndofReport " -Title "SEPPmail365 Exo Report" -Head $style

            # Write Report to Disk
            try {
                $finalReport|Out-File -FilePath $FinalPath -Force
                if ($jsonBackup) {
                    # Store json in the same location as HTML
                    $jsonpath = (Join-Path -Path (split-path $FinalPath -Parent) -ChildPath (split-path $FinalPath -leafbase)) + '.json'
                    Set-Content -Value $JsonData -Path $jsonPath -force
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
            if (($IsMacOs) -or (isLinux)) {
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
        if ((!($InboundOnly)) -or (!($routing)) ) {
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
                     if ($DeploymentInfo.inBoundOnly -eq $true) {$inboundOnly = $true}
                    if ($DeploymentInfo.inBoundOnly -eq $false) {$inboundOnly = $false}
                     if ($null -eq $DeploymentInfo.inBoundOnly) {$inboundOnly = $false}
                }
            }
        } else {
            if ($deploymentInfo.routing -eq 'p') {$routing = 'parallel'}
            if ($deploymentInfo.routing -eq 'i') {$routing = 'inline'}
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
                Remove-SC365Rules -Routing $routing
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
    [CmdletBinding(
        SupportsShouldProcess = $true,
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
            [String]$SEPPmailCloudDomain,
    
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
            [ValidateSet('prv','de','ch')]
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
 
        if ((!($SeppmailcloudDomain)) -or (!($Region)) -or (!($routing)) ) {
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
                             if ($deploymentInfo.Region) {$Region = $deploymentInfo.Region} else {Write-Error "Could not autodetect region. Use manual parameters"; break}
                if ($DeploymentInfo.SEPPmailCloudDomain) {$SEPPmailCloudDomain = $DeploymentInfo.SEPPmailCloudDomain} else {Write-Error "Could not autodetect SEPPmailCloudDomain. Use manual parameters"; break}          
              if ($DeploymentInfo.inBoundOnly -eq $true) {$inboundOnly = $true}
             if ($DeploymentInfo.inBoundOnly -eq $false) {$inboundOnly = $false}
              if ($null -eq $DeploymentInfo.inBoundOnly) {$inboundOnly = $false}
                }
            }
        } else {
            if ($deploymentInfo.routing -eq 'p') {$routing = 'parallel'}
            if ($deploymentInfo.routing -eq 'i') {$routing = 'inline'}
 
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
            Write-Error "Setup removal failed. Try removing SEPPmail.cloud Rules and Connectors from Portal of with native CmdLets."
            break
        }

        try {
            if ($InBoundOnly -eq $true) {
                Write-Information '--- Creating inbound connector ---' -InformationAction Continue
                New-SC365Connectors -SEPPmailCloudDomain $SEPPmailCloudDomain -routing $routing -region $region -inboundonly:$true
            } else {
                Write-Information '--- Creating in and outbound connectors ---' -InformationAction Continue
                New-SC365Connectors -SEPPmailCloudDomain $SEPPmailCloudDomain -routing $routing -region $region
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
        if ($PSCmdLet.ShouldProcess($SEPPmailCloudDomain)) {
            Write-Information "--- Successfully created SEPPmail.cloud Setup for $seppmailclouddomain in region $region in $routing mode ---" -InformationAction Continue
            Write-Information "--- Wait a few minutes until changes are applied in the Microsoft cloud ---" -InformationAction Continue
            Write-Information "--- Afterwards, start testing E-Mails in and out ---" -InformationAction Continue
        }
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
        if ((!($InboundOnly)) -or (!($routing)) ) {
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
        if ($InBoundOnly -eq $true) {
            $smcConn = Get-SC365Connectors -Routing $routing -inboundonly:$true
        } else {
            $smcConn = Get-SC365Connectors -Routing $routing -inboundonly:$false
        }
        if ($InBoundOnly -eq $false) {
            $smcTRules = Get-SC365Rules -Routing $routing
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
            try {
                if (!($ibc = Get-Inboundconnector -Identity '[SEPPmail*')) {
                    Write-Error "Could not find SEPPmail.Cloud Inbound-Connecor"
                }
                if (!($obc = Get-Outboundconnector -Identity '[SEPPmail*')) {
                    Write-Error "Could not find SEPPmail.Cloud Outbound-Connecor"
                }
            }
            catch {
            Write-Error "Could not detect SEPPmail-Cloud Connectors, aborting"
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
            {($_ -eq 'InBound') -and ($ibc.identity -eq "[SEPPmail.cloud] Inbound-Parallel")} 
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
            {($_ -eq 'InBound') -and ($ibc.identity -eq "[SEPPmail.cloud] Inbound-Inline")} 
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
            {($_ -eq 'OutBound') -and ($obc.identity -eq "[SEPPmail.cloud] Outbound-Parallel")}
            {
                # We take one of 2 Send/Receive Messagetraces from SEPPmail and get the details

                # Now this one has 3 Parts. 0= Recieve from Mailboxhost, 1 = SumbitMessage (Exo internal), 2 = Send to SEPPmail
                $MTDSEPPReceive = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient -Event 'receive'
                # $MTDSEPPSubmit = $MessageTraceDetailSEPPmail[1] Not interesting for us
                $MTDSEPPExtSend = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient |where-object Event -like  '*SEND*'                
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
            {($_ -eq 'OutBound') -and ($obc.identity -eq "[SEPPmail.cloud] Outbound-Inline")}
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
Function Show-sc365Tenant {
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

Register-ArgumentCompleter -CommandName Get-SC365TenantId -ParameterName MailDomain -ScriptBlock $paramDomSB
Register-ArgumentCompleter -CommandName New-SC365Setup -ParameterName SEPPmailCloudDomain -ScriptBlock $paramDomSB


# SIG # Begin signature block
# MIIVzAYJKoZIhvcNAQcCoIIVvTCCFbkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9BZurMLn5uUYQ
# JuAJAWSrR1j0KCz5T4ndfZxIAjAfdKCCEggwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB478xPysvovXcXSAtfc8U6i4hK
# N+ueaJYQGUeGrVwIVTANBgkqhkiG9w0BAQEFAASCAgDNOyRIiumHRz4yde/N35NQ
# 1PZiQkRB1Ui4uuSXijGtZdA8Hoojhym4OarV6SzYYRhWFKcVncficA69sXRQQDgE
# Yb1z7Vu4O+ok7KwRLcl5Hd2kHjJy7luOlQeMEGi0bOWnml/jBq35nJSOlbbwzmxO
# 03R/jCj0xzn7dcwlXbcB5fchBDV+uZC1sX6cNZx3xO1oS5el2OCvKN9FxqPUjRRt
# 7ejyyp38s4WJ0SWXgN1B93TgP3GMELk0hJr2hyVUaBlpnW9/u3aHUbgh7+RyH4ah
# bErmBmd/KTfwR9zOg6H0D4tAcTCFsTMl6mkLfYWj0+qRT32vMrLAwe7MQkHP1cOq
# FN0YQIY6MU9jXpcmkkFC720YbYLCu4V4pBeCtu4XcDjOQLI4CB2j4hMSD4sfXFKR
# JUUB8IkEOI3Yq3Xrl0Hll7USOuxaOWYBZraZsdKCu1/12JOrta9PQRmmDuIQM/f4
# uXdejRtkaRigSZBN57X5xQAMg5oIC2RBVdXJWjoNs/Ub515jPwdsNCbK/pl+EuJf
# Z1+qcL5NVtf7Ppi2WYkZxSSZVElgNrntseofJsM2sgLVBq/o0uuZKfNskC2+Guth
# GEnp+Jgqu8bQfdJhMS406RlRVNgJ65B1ZuWBWmiUs6AVcIerUSfCHhrVUiEt8nr+
# jCUXMgyAqnYI+MFyqVUzKA==
# SIG # End signature block
