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
           #Position = 0
        )]
        [Alias('Path')]
        [string]$filePath = '.',

        [Parameter(   
           Mandatory   = $false,
           HelpMessage = 'Literal path of the HTML report on disk',
           ParameterSetName = 'LiteralPath',
           Position = 0
        )]
        [string]$Literalpath = '.'
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
                $rawData = Invoke-Expression -Command $exoCmd
                if ($null -eq $rawData) {
                    $ExoHTMLData = New-object -type PSobject -property @{Result = '--- no information available ---'}|Convertto-HTML -Fragment
                } else {
                    $ExoHTMLData = $rawData|Convertto-HTML -Fragment
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
            $N = Get-ExoHTMLData -ExoCmd 'Get-TransportRule | select-object Name,State,Mode,Priority,FromScope,SentToScope'
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
            }
            catch{
                Write-Warning "Could not write report to $FinalPath"
                if ($IsWindows) {
                    $FinalPath = Join-Path -Path $env:localappdata -ChildPath $ReportFilename
                }
                if ($IsMacOs) {
                    $Finalpath = Join-Path -Path $env:HOME -ChildPath $ReportFilename
                }
                Write-Verbose "Writing report to $finalPath"
                try {
                    $finalReport|Out-File -FilePath $finalPath -Force
                }
                catch {
                    $error[0]
                }
            }

            if ($IsWindows) {
                Write-Information -MessageData "Opening $finalPath with default browser"
                Invoke-Expression "& '$finalpath'"
            }
            if ($IsMacOs) {
                "Report is stored on your disk at $finalpath. Open with your favorite browser."
            }
        }
        catch {
            throw [System.Exception] "Error: $($_.Exception.Message)"
        }
    }
    end {
    }
}

function Remove-SC365Setup {
    [CmdletBinding()]
    param()

    Begin {}
    Process {
        Remove-SC365Rules
        Remove-SC365Connectors
    }
    End{}
}

function New-SC365Setup {
    [CmdletBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md'
    )]

    # Specifies a path to one or more locations.
    param(
        [Parameter(
            Mandatory=$true,
            Position=0,
            HelpMessage="All Domains included / booked in the SEPPmail.cloud")]
            [Alias("domain")]
            [ValidateNotNullOrEmpty()]
        [String]$SEPPmailCloudDomain,

        [Parameter(
            Mandatory=$true,
            Position=1,
            HelpMessage="Inline routing via SEPPmail (MX ==> SEPPmail), or routing via Microsoft (MX ==> Microsoft)")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('parallel','inline','p','i')]
        [String]$routing,
    
        [Parameter(
            Mandatory=$true,
            Position=0,
            HelpMessage="Physical location of your data")]
            [ValidateSet('prv','de','ch')]
        [String]$region
        )

    Begin {
        if ($routing -eq 'p') {$routing = 'parallel'}
        if ($routing -eq 'i') {$routing = 'inline'}
    }
    Process {
        New-SC365Connectors -SEPPmailCloudDomain $SEPPmailCloudDomain -routing $routing -region $region
        New-SC365Rules -SEPPmailCloudDomain $SEPPmailCloudDomain -routing $routing
    }
    End{
        Write-Information "Wait a few minutes until changes are applied in the Microsoft cloud"
        Write-Information "Afterwards, start testing E-Mails in and out"
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
        [Parameter(
            Mandatory=$true,
            Alias = 'MailDomain')]
        [string]$SEPPmailCloudDomain
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

function Get-SC365MessageTrace {
    [CmdLetBinding(
        HelpURI = 'https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md#setup-the-integration'
    )]
    param (
        [Parameter(Mandatory = $true)]
        [String]$MessageId,
        
        [Parameter(Mandatory = $true)]
        [Alias('RecipientAddress')]
        [String]$Recipient
    )
    begin {
        Write-Information "This CmdLet is still under development"
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
        Write-Verbose "Retrieving Tenant-Domains"
        $TenantDomains = (Get-AcceptedDomain).DomainName

        Write-Verbose "Retrieving initial Message-Trace id MessageID $MessageId for recipient $Recipient"
        Write-Progress -Activity "Loading message data" -Status "MessageTrace" -PercentComplete 0 -CurrentOperation "Start"

        $PlainMessageID = $MessageId.Trim('<','>')
        Write-Verbose "Formatting Parameterinput Messageid:$MessageId - adding < and > at beginning and end to filter property"
        if (!($MessageId.StartsWith('<'))) {$MessageId = "<" + $MessageId}
        if (!($MessageId.EndsWith('<'))) {$MessageId = $MessageId + ">"}
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
            Write-Error "Could not find Message with ID $MessageID and recipient $recipient. Look for typos. Message too old ? Try Search-MessageTrackingReport or Get-Messagetrace"
            break
        }
        try {
            Write-verbose "Test Maildirection, based on the fact that the $Recipient is part of TenantDomains"
            If ($TenantDomains.Contains(($Recipient -Split '@')[-1])) {
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

        Write-Verbose "Crafting basic MessateTraceInfo"

        $OutPutObject = [PSCustomObject][ordered]@{
            Subject                = if ($MessageTrace.count -eq 1) {$MessageTrace.Subject} else {$MessageTrace[0].Subject}
            Size                   = if ($MessageTrace.count -eq 1) {$MessageTrace.Size} else {$MessageTrace[0].Size} #|{$_/1KB} | ToString('.0') + ' kB'
            SenderAddresses        = if ($MessageTrace.count -eq 1) {$MessageTrace.SenderAddress} else {$MessageTrace[0].SenderAddress}
            MailDirection          = $MailDirection
            RoutingMode            = if (($ibc.identity -eq "[SEPPmail.cloud] Inbound-Parallel") -or ($obc.identity -eq "[SEPPmail.cloud] Outbound-Parallel")) {'Parallel'} else {'Inline'}
        } 
        Write-Verbose "Add MessageTraceId´s"
        foreach ($i in $Messagetrace) {Add-Member -InputObject $OutPutObject -MemberType NoteProperty -Name ('MessageTraceId' + ' Index'+ $I.Index) -Value $i.MessagetraceId}

        if ($MessageTrace.count -eq 1) {
            Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromIP -Value $MessageTrace.FromIP
            Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromDNS -Value (Resolve-SC365DNSname -IPAddress $MessageTrace.FromIP)
            if ($messagetrace.ToIp) {Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToIP -Value $MessageTrace.ToIP}
            if ($messagetrace.ToIp) {Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToDNS -Value (Resolve-SC365DNSname -IPAddress $MessageTrace.ToIP)}

        } else {
            if ($Messagetrace[0].FromIP) {
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromIP -Value $MessageTrace[0].FromIP
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromDNS -Value (Resolve-SC365DNSname -IPAddress $MessageTrace[0].FromIP)
            }
            else {
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromIP -Value '---empty---'
                Add-Member -InputObject $OutputObject -membertype NoteProperty -Name ExternalFromDNS -Value '---empty---'
            }
            if ($MessageTrace[0].ToIP) {
                Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToIP -Value $MessageTrace[0].ToIP
                Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToDNS -Value (Resolve-SC365DNSname -IPAddress $MessageTrace[0].ToIP)
            } else {
                Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name ExternalToIP -Value '---empty---'
            }
        }
        Add-Member -InputObject $OutPutObject -membertype NoteProperty -Name 'SplitLine' -Value "-------------------- MessageTrace DETAIL Info Starts Here --------------------"

    }
    
    process {
        switch ($maildirection)
        {
            {($_ -eq 'InBound') -and ($ibc.identity -eq "[SEPPmail.cloud] Inbound-Parallel")} 
            {
                # Im Parallel Mode kommt die Mail 2x, einmal von externem Host und einmal von SEPpmail, Index 0 und 1
                
                $MessageTraceDetailExternal = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient
                $MTDExtReceived = $MessageTraceDetailExternal[0]
                $MTDExtExtSend = $MessageTraceDetailExternal[1]
                $MessageTraceDetailSEPPmail = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient
                $MTDSEPPReceived = $MessageTraceDetailSEPPmail[0]
                $MTDSEPPDelivered = $MessageTraceDetailSEPPmail[1]
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
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToDNS -Value (Resolve-SC365DNSName -IPAddress $messageTrace[1].ToIP)
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
                $MTDReceived = $MessageTraceDetail|where-object {($_.Event -eq 'Received') -or ($_.Event -eq 'Empfangen')} 
                $MTDDelivered = $MessageTraceDetail|where-object {($_.Event -eq 'Delivered') -or ($_.Event -eq 'Zustellen')}
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
                #if ($MessageTrace.count -eq 1) {
                 #   $MessageTraceDetailSEPPmail = Get-MessagetraceDetail -MessageTraceId $MessageTrace.MessageTraceId -Recipient $Recipient
                #} else {
                    $MessageTraceDetailSEPPmail = Get-MessagetraceDetail -MessageTraceId $MessageTrace[1].MessageTraceId -Recipient $Recipient
                #}
                
                # Now this one has 3 Parts. 0= Recieve from Mailboxhost, 1 = SumbitMessage (Exo internal), 2 = Send to SEPPmail
                $MTDSEPPReceive = $MessageTraceDetailSEPPmail[0]
                # $MTDSEPPSubmit = $MessageTraceDetailSEPPmail[1] Not interesting for us
                $MTDSEPPExtSend = $MessageTraceDetailSEPPmail[2]
                
                $MessageTraceDetailExternal = Get-MessagetraceDetail -MessageTraceId $MessageTrace[0].MessageTraceId -Recipient $Recipient
                $MTDExtReceive = $MessageTraceDetailExternal[0]
                $MTDExtExtSend = $MessageTraceDetailExternal[1]
                try {
                    $obcName = (((($MTDSEPPExtSend.Data -Split '<') -replace ('>','')) -split (';') | select-String 'S:Microsoft.Exchange.Hygiene.TenantOutboundConnectorCustomData').ToString()).Split('=')[-1]
                }catch {
                    $obcName = "--- E-Mail did not go via a SEPPmail Connector ---"
                }
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToIP -Value $messageTrace[1].ToIP
                $Outputobject | Add-Member -MemberType NoteProperty -Name FromExternalSendToDNS -Value (Resolve-SC365DNSname -IPAddress $messageTrace[1].ToIP)
                $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPmailReceivedFromIP -Value $messageTrace[1].FromIP
                try { 
                    $Outputobject | Add-Member -MemberType NoteProperty -Name SEPPmailReceivedFromDNS -Value (Resolve-SC365DNSname -IPAddress $messageTrace[1].FromIP)
                } 
                catch {
                    Write-Information "Cannot Resolve $($messageTrace[1].FromIP)"
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
                $MTDSubmit = $MessageTraceDetail|Where-Object {(($_.Event -eq 'Übermitteln') -or ($_.Event -eq 'Submit'))}
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
                $Outputobject | Add-Member -MemberType NoteProperty -Name 'ExoInternalTransportTime(s)' -Value (New-TimeSpan -Start $MTDReceive.Date -End $MTDExtSend.Date).Seconds
                $Outputobject | Add-Member -MemberType NoteProperty -Name ReceiveDetail -Value $MTDReceive.Detail
                #$Outputobject | Add-Member -MemberType NoteProperty -Name SubmitDetail -Value $MTDSubmit.Detail # Keine Relevante Info
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExtSendDetail -Value $MTDExtSend.Detail
                $Outputobject | Add-Member -MemberType NoteProperty -Name OutboundConnectorName -Value $obcName
                $Outputobject | Add-Member -MemberType NoteProperty -Name ExternalSendLatency -Value (((($MTDExtSend.Data -Split '<') -replace ('>','')) -split (';') | select-String 'S:ExternalSendLatency').ToString()).Split('=')[-1]
                Write-Progress -Activity "Loading message data" -Status "StatusMessage" -PercentComplete 100 -CurrentOperation "Done"
            }

        }


    }
        #endregion Send/Outbound
    end {
        #$SC365MessageTrace = New-Object -TypeName pscustomobject -ArgumentList $SC365MessageTraceHT
        return $OutPutObject
        #$SC365MessageTraceHT
    }
}

Register-ArgumentCompleter -CommandName Get-SC365TenantId -ParameterName MailDomain -ScriptBlock $paramDomSB

# SIG # Begin signature block
# MIIL/AYJKoZIhvcNAQcCoIIL7TCCC+kCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAWiKCo+z+7yPkf
# WrhDSX9CZCCN344G1x3SMtNDdktG2KCCCUAwggSZMIIDgaADAgECAhBxoLc2ld2x
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
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDaxqgFcLygE5NNKe0ksdqVlf1u
# xA4aGxBgbbhG4Z3K0jANBgkqhkiG9w0BAQEFAASCAQAZJo0ZDtM3siO8P7PJqfGD
# c8BSbUlrxLFKyfVFyweglMJ8uBGYbSglEBRUuJefq3iW7CeAEBvC5DJjqR3gorsV
# x2XxWvuJaGG8zLxoeiCHMsXqUd1yKjVS9n/HD3jef230oZZb+47egI9S99qeCaMv
# Ic7ZhqMfu58rGTnBDFCbjdKvUWJLMPYTpp7Gs1cnMHixxHL9YU9M/UHGg3O0S1my
# F7kz4lmZtld5lI9ozIigYpOgkHddlPcXgEJp7V38l8vrsA86YURYJs+tH0JnM9fX
# No9BC7SyqYaLYKmFe0wtNyIisYnqLCw5lV9hd9v69jn8l1I1gPvJMVGpJJPfAzDi
# SIG # End signature block
