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
        [string]$LiteralPath = '.',

        [Parameter(   
           Mandatory   = $false,
           HelpMessage = 'URL Path of the header logo',
           ParameterSetName = 'LiteralPath',
           Position = 1
        )]
        [Parameter(   
            Mandatory   = $false,
            HelpMessage = 'URL Path of the header logo',
            ParameterSetName = 'FilePath',
            Position = 1
         )]
         [string]$LogoSource = 'https://downloads.seppmail.com/wp-content/uploads/logo_seppmail_V1_Screen_M.png',

         [Parameter(   
            Mandatory   = $false,
            HelpMessage = 'Scaling factor in % of the header logo',
            ParameterSetName = 'LiteralPath'
         )]
         [Parameter(   
             Mandatory   = $false,
             HelpMessage = 'Scaling factor in % of the header logo',
             ParameterSetName = 'FilePath'
          )]
          [ValidatePattern('^(100|[1-9][0-9]?)%$')]
          [string]$LogoWidth = '20%',

          [Parameter(   
            Mandatory   = $false,
            HelpMessage = 'URL when clicking the header logo',
            ParameterSetName = 'LiteralPath'
         )]
         [Parameter(   
             Mandatory   = $false,
             HelpMessage = 'URL when clicking the logo',
             ParameterSetName = 'FilePath'
          )]
          [string]$LogoUrl = 'https://www.seppmail.cloud',
 
        [Parameter(   
           Mandatory   = $false,
           HelpMessage = 'Literal path of the JSON backup on disk',
           ParameterSetName = 'LiteralPath'
        )]
        [Parameter(   
            Mandatory   = $false,
            HelpMessage = 'Literal path of the JSON backup on disk',
            ParameterSetName = 'FilePath'
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

                if ($FileLogo) {
                    New-HTMLLogo -LogoPath $logoPath
                } else {
                    New-HTMLImage -Source $LogoSource  -Width $LogoWidth -UrlLink $LogoUrl
                }
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
