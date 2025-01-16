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
        $ibc = Get-InboundConnector -Identity '[SEPPmail.Cloud]*' | select-object -first 1
        $ibc|Select-Object -ExpandProperty WhenCreated

        $days = (New-Timespan -Start (Get-Date $ibc.WhenCreated) -End (Get-Date)).Days
        Write-Output "SEPPmail Cloud was created $days days ago"
        $commentstime = ((($ibc.comment|select-String 'Created with') -split 'version').trim()[-1] -split 'on').trim()[-1]
        Write-Verbose "Inbound Connector Comments writes install date is: $commentsTime"
    }
    end{}
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

Register-ArgumentCompleter -CommandName Get-SC365TenantId -ParameterName MailDomain -ScriptBlock $paramDomSB
